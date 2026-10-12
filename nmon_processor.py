#!/usr/bin/env python3
"""
NMON Processing Engine
======================
Processes one or more IBM NMON files and generates a consolidated
multi-sheet Excel workbook ready for SAP sizing and capacity analysis.

Usage
-----
    python nmon_processor.py [file1.nmon file2.nmon ...] [options]

Options
-------
    --output DIR   Directory to write the output workbook (default: current dir)
    --input  DIR   Directory to scan for *.nmon files (used when no files given)

If no files or directory are supplied the current working directory is scanned.
"""

import re
import sys
import os
import glob
import argparse
import warnings
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional

import numpy as np
import pandas as pd
from tqdm import tqdm

# openpyxl for rich formatting; xlsxwriter used via pandas for bulk writes
try:
    import openpyxl
    from openpyxl.styles import (
        PatternFill, Font, Alignment, Border, Side,
        numbers as xl_numbers,
    )
    from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("openpyxl is required.  Install with:  pip install openpyxl")

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
RED_FILL    = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")
GREEN_FILL  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
ALT_FILL    = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

HEADER_FONT  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
BODY_FONT    = Font(name="Calibri", size=10)
BOLD_FONT    = Font(name="Calibri", bold=True, size=10)
TITLE_FONT   = Font(name="Calibri", bold=True, size=14, color="1F3864")

THIN_SIDE   = Side(style="thin", color="D3D3D3")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE,
                     top=THIN_SIDE, bottom=THIN_SIDE)

NMON_TS_FORMATS = [
    "%d-%b-%Y %H:%M:%S",
    "%d-%b-%Y %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
]

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _pct(arr: pd.Series, p: float) -> float:
    a = arr.dropna()
    return float(np.percentile(a, p)) if len(a) else float("nan")


def _safe(val, default=None):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return default
    return val


def _parse_ts(raw: str) -> Optional[datetime]:
    raw = raw.strip()
    for fmt in NMON_TS_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return None


def _to_float(v) -> Optional[float]:
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return None


def _extract_sid(text: str) -> str:
    """
    Try to pull a 3-char SAP SID from hostname / metadata text.
    Returns '' when nothing convincing is found.
    """
    # Common patterns: SID= or /SID/ or a standalone 3-char uppercase word
    # surrounded by non-alpha chars
    patterns = [
        r'\bSID[=:\s]+([A-Z][A-Z0-9]{2})\b',
        r'[_\-/\s]([A-Z][A-Z0-9]{2})[_\-/\s]',
        r'^([A-Z][A-Z0-9]{2})[_\-]',
    ]
    for pat in patterns:
        m = re.search(pat, text.upper())
        if m:
            candidate = m.group(1)
            # Reject obvious non-SID tokens
            if candidate not in ("IBM", "CPU", "MEM", "NET", "SMP", "AIX", "SAP",
                                  "TMP", "VAR", "OPT", "USR", "SYS", "LOG", "RUN",
                                  "ETH", "ENT", "HDK", "SDA", "SDB", "SDC", "LVM"):
                return candidate
    return ""


def _auto_col_width(ws, min_width=8, max_width=50):
    """Auto-size every column of a worksheet."""
    for col in ws.columns:
        best = min_width
        for cell in col:
            try:
                best = max(best, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(best + 2, max_width)


def _apply_header(ws, row: int = 1):
    """Apply header styling to the given row."""
    for cell in ws[row]:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _write_df(ws, df: pd.DataFrame, start_row: int = 1):
    """Write a DataFrame to a worksheet starting at start_row (1-based)."""
    # Header
    for ci, col in enumerate(df.columns, 1):
        ws.cell(row=start_row, column=ci, value=col)
    _apply_header(ws, start_row)
    # Data rows
    for ri, row_data in enumerate(df.itertuples(index=False), start_row + 1):
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font   = BODY_FONT
            cell.border = THIN_BORDER
            # Alternating row fill
            if ri % 2 == 0:
                cell.fill = ALT_FILL


def _pct_col_cf(ws, col_letter: str, start_row: int, end_row: int,
                yellow_lo: float, red_lo: float):
    """Apply yellow/red conditional formatting to a percentage column."""
    rng = f"{col_letter}{start_row}:{col_letter}{end_row}"
    ws.conditional_formatting.add(
        rng,
        CellIsRule(operator="between",
                   formula=[str(yellow_lo), str(red_lo - 0.001)],
                   fill=YELLOW_FILL)
    )
    ws.conditional_formatting.add(
        rng,
        CellIsRule(operator="greaterThanOrEqual",
                   formula=[str(red_lo)],
                   fill=RED_FILL)
    )


# ─────────────────────────────────────────────────────────────────────────────
# NMON PARSER
# ─────────────────────────────────────────────────────────────────────────────

class NmonFile:
    """
    Parses a single NMON file into structured Python objects.
    All raw sections are stored as lists of raw token rows keyed by section tag.
    """

    def __init__(self, path: str):
        self.path     = path
        self.hostname = ""
        self.sid      = ""

        # Identity / hardware
        self.lpar_name     = ""
        self.serial        = ""
        self.machine_type  = ""
        self.processor_model = ""
        self.processor_type  = ""
        self.os_name       = ""
        self.os_version    = ""
        self.physical_cpus : Optional[int]   = None
        self.virtual_cpus  : Optional[int]   = None
        self.entitled_cpus : Optional[float] = None
        self.smc_name      = ""          # Shared Memory / Processor Pool name
        self.installed_mem_gb: Optional[float] = None
        self.virt_tech     = ""

        # Timestamps
        self.ts_start: Optional[datetime] = None
        self.ts_end  : Optional[datetime] = None

        # Timestamp index  {snap_id → datetime}
        self._timestamps: dict = {}

        # Parsed data frames (populated during parse)
        self.cpu_raw:    pd.DataFrame = pd.DataFrame()
        self.mem_raw:    pd.DataFrame = pd.DataFrame()
        self.paging_raw: pd.DataFrame = pd.DataFrame()
        self.disk_raw:   pd.DataFrame = pd.DataFrame()
        self.net_raw:    pd.DataFrame = pd.DataFrame()
        self.lpar_raw:   pd.DataFrame = pd.DataFrame()

        # Warnings collected during parse
        self.warnings: list = []

    # ── main parse entry point ────────────────────────────────────────────────

    def parse(self) -> None:
        with open(self.path, "r", encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()

        # Bucket all lines by their section tag
        sections: dict = {}
        for line in lines:
            line = line.rstrip("\n\r")
            if not line.strip():
                continue
            parts = line.split(",")
            tag   = parts[0].strip().upper()
            if tag not in sections:
                sections[tag] = []
            sections[tag].append(parts)

        self._parse_identity(sections)
        self._parse_timestamps(sections)
        self._parse_cpu(sections)
        self._parse_mem(sections)
        self._parse_disk(sections)
        self._parse_net(sections)
        self._infer_collection_period()

    # ── identity / metadata ───────────────────────────────────────────────────

    def _parse_identity(self, sections: dict) -> None:
        aaa = sections.get("AAA", [])
        bbb = sections.get("BBB", [])

        meta = {}
        for row in aaa:
            if len(row) >= 3:
                meta[row[1].strip().upper()] = row[2].strip()

        for row in bbb:
            if len(row) >= 3:
                meta[row[1].strip().upper()] = ",".join(row[2:]).strip()

        self.hostname = meta.get("HOST", meta.get("HOSTNAME", ""))
        if not self.hostname and "NODENAME" in meta:
            self.hostname = meta["NODENAME"]
        self.hostname = self.hostname.upper()

        self.os_name    = meta.get("OS",  meta.get("OSNAME",    "AIX"))
        self.os_version = meta.get("VER", meta.get("OSVERSION", ""))
        self.serial     = meta.get("SERIAL", meta.get("SERIALNUMBER", ""))
        self.machine_type   = meta.get("MACHINE", meta.get("MACHINETYPE", ""))
        self.processor_model = meta.get("PROCESSOR", meta.get("PROCESSORTYPE", ""))
        self.processor_type  = meta.get("PROCESSORTYPE", "")
        self.lpar_name   = meta.get("LPAR", meta.get("LPARNAME", self.hostname))
        self.smc_name    = meta.get("SHAREDPOOL", meta.get("SMCNAME", ""))
        self.virt_tech   = meta.get("VIRTUALIZATION",
                                    "PowerVM LPAR" if self.os_name.upper() == "AIX" else "")

        # CPUs
        try:
            self.physical_cpus = int(meta.get("PHYSICALCPU", meta.get("PYSICALCPU",
                                     meta.get("CPUS", meta.get("PHYSICALCPUS", "")))))
        except (ValueError, TypeError):
            pass

        try:
            self.virtual_cpus = int(meta.get("LOGICALCPU",
                                    meta.get("LOGICALCPUS",
                                    meta.get("VIRTUALCPUS",
                                    meta.get("VCPUS", "")))))
        except (ValueError, TypeError):
            pass

        try:
            self.entitled_cpus = float(meta.get("ENTITLED",
                                       meta.get("ENTITLEDCAPACITY",
                                       meta.get("ENTITLEDCPU", ""))))
        except (ValueError, TypeError):
            pass

        # Memory — look in MEM header or AAA
        try:
            mem_mb = float(meta.get("MEMORY", meta.get("TOTALMEM",
                           meta.get("REALMEM", ""))))
            self.installed_mem_gb = round(mem_mb / 1024, 2)
        except (ValueError, TypeError):
            pass

        # SID — try hostname, LPAR name, metadata
        for candidate_text in (self.hostname, self.lpar_name,
                                meta.get("SID", ""), meta.get("SAPSID", "")):
            sid = _extract_sid(candidate_text)
            if sid:
                self.sid = sid
                break

    # ── timestamps ────────────────────────────────────────────────────────────

    def _parse_timestamps(self, sections: dict) -> None:
        for row in sections.get("ZZZZ", []):
            # ZZZZ,T0001,23:00:00,01-JAN-2024
            # ZZZZ,T0001,01-JAN-2024 23:00:00
            if len(row) < 3:
                continue
            snap_id = row[1].strip()
            if len(row) >= 4:
                ts_str = f"{row[3].strip()} {row[2].strip()}"
            else:
                ts_str = row[2].strip()
            ts = _parse_ts(ts_str)
            if ts:
                self._timestamps[snap_id] = ts

        if self._timestamps:
            vals = list(self._timestamps.values())
            self.ts_start = min(vals)
            self.ts_end   = max(vals)

    # ── timestamp lookup helper ───────────────────────────────────────────────

    def _ts(self, snap_id: str) -> Optional[datetime]:
        return self._timestamps.get(snap_id)

    # ── CPU ───────────────────────────────────────────────────────────────────

    def _parse_cpu(self, sections: dict) -> None:
        """
        CPU data can appear in multiple sections:
          CPU_ALL   — per-logical-CPU breakdown (we use the aggregate)
          CPU_TOTAL — overall system totals (most reliable)
          LPAR      — LPAR-level capacity metrics
          PROC      — processor utilisation (fallback)

        Priority: CPU_ALL  (aggregate row) → CPU_TOTAL → PROC
        """
        records   = []
        lpar_rows = []
        warnings  = []

        # ── Parse LPAR section for entitled / virtual / physical CPUs ─────────
        lpar_header = None
        for row in sections.get("LPAR", []):
            tag = row[0].strip().upper()
            if tag == "LPAR" and len(row) > 2 and row[1].strip().upper() == "LPAR":
                lpar_header = [x.strip().upper() for x in row[2:]]
            elif tag == "LPAR" and lpar_header:
                snap_id = row[1].strip() if len(row) > 1 else ""
                ts = self._ts(snap_id)
                vals = row[2:]
                d = dict(zip(lpar_header, vals))
                entitled = _to_float(d.get("ENTITLED", d.get("ENTITLEDCAPACITY", None)))
                virtual  = _to_float(d.get("VIRTUALCPUS", d.get("VCPUS", None)))
                physical = _to_float(d.get("PHYSICALCPU", d.get("PYSICALCPU", None)))
                # Fallback: derive physical from busy% * entitled
                lpar_rows.append({
                    "snap_id":        snap_id,
                    "ts":             ts,
                    "entitled":       entitled,
                    "virtual":        virtual,
                    "physical_raw":   physical,
                })
                # Patch top-level fields if not already set
                if self.entitled_cpus is None and entitled is not None:
                    self.entitled_cpus = entitled
                if self.virtual_cpus is None and virtual is not None:
                    try:
                        self.virtual_cpus = int(virtual)
                    except (ValueError, TypeError):
                        pass

        # Index LPAR rows by snap_id for fast lookup
        lpar_by_snap = {r["snap_id"]: r for r in lpar_rows}

        # ── Parse CPU_ALL ─────────────────────────────────────────────────────
        cpu_header = None
        cpu_source = None

        for section_name in ("CPU_ALL", "CPU_TOTAL", "PROC"):
            rows = sections.get(section_name, [])
            if not rows:
                continue
            for row in rows:
                tag = row[0].strip().upper()
                if tag == section_name and len(row) > 2 and row[1].strip().upper() == section_name:
                    cpu_header = [x.strip().upper() for x in row[2:]]
                    cpu_source = section_name
                elif tag == section_name and cpu_header and cpu_source == section_name:
                    snap_id = row[1].strip() if len(row) > 1 else ""
                    ts = self._ts(snap_id)
                    vals = row[2:]
                    d    = dict(zip(cpu_header, vals))

                    def _get(*keys):
                        for k in keys:
                            v = _to_float(d.get(k))
                            if v is not None:
                                return v
                        return None

                    user  = _get("USER%", "USER", "USR%", "USER_PCT")
                    sys_  = _get("SYS%",  "SYS",  "SYSTEM%", "SYSTEM_PCT")
                    wait  = _get("WAIT%", "WAIT", "WAI%")
                    idle  = _get("IDLE%", "IDLE", "IDL%")

                    # Derive missing values
                    if idle is None and user is not None:
                        idle = max(0.0, 100.0 - (user or 0) - (sys_ or 0) - (wait or 0))
                    busy = (100.0 - idle) if idle is not None else None
                    if busy is None and user is not None:
                        busy = (user or 0) + (sys_ or 0) + (wait or 0)

                    # LPAR capacity metrics for this snap
                    lp = lpar_by_snap.get(snap_id, {})
                    entitled = _safe(lp.get("entitled"), self.entitled_cpus)
                    virtual  = _safe(lp.get("virtual"),  self.virtual_cpus)
                    phys_raw = lp.get("physical_raw")
                    if phys_raw is None and busy is not None and entitled is not None:
                        phys_raw = round(busy * entitled / 100.0, 4)

                    records.append({
                        "Timestamp":        ts,
                        "SID":              self.sid,
                        "Hostname":         self.hostname,
                        "User_Pct":         round(user  or 0, 2),
                        "System_Pct":       round(sys_  or 0, 2),
                        "Wait_Pct":         round(wait  or 0, 2),
                        "Idle_Pct":         round(idle  or 0, 2),
                        "Busy_Pct":         round(busy  or 0, 2),
                        "LPAR_Physical_CPU":round(phys_raw, 4) if phys_raw is not None else None,
                        "LPAR_Entitled_CPU":entitled,
                        "LPAR_Virtual_CPUs":virtual,
                    })
            if cpu_source:
                break  # found a CPU section — stop trying fallbacks

        if not records:
            warnings.append(("CPU", "No CPU data found (CPU_ALL/CPU_TOTAL/PROC sections missing)"))
        elif cpu_source != "CPU_ALL":
            warnings.append(("CPU", f"CPU data sourced from {cpu_source} (CPU_ALL not found)"))

        self.cpu_raw = pd.DataFrame(records) if records else pd.DataFrame(
            columns=["Timestamp","SID","Hostname","User_Pct","System_Pct",
                     "Wait_Pct","Idle_Pct","Busy_Pct",
                     "LPAR_Physical_CPU","LPAR_Entitled_CPU","LPAR_Virtual_CPUs"])

        # Build LPAR_Capacity from merged LPAR section rows + CPU busy
        lpar_cap_records = []
        for lp in lpar_rows:
            snap = lp["snap_id"]
            ts   = lp["ts"]
            ent  = lp.get("entitled")
            virt = lp.get("virtual")
            phys = lp.get("physical_raw")
            # If physical_raw was missing from LPAR section, look it up from CPU
            if phys is None and ent is not None:
                cpu_snap = next(
                    (r for r in records if r["Timestamp"] == ts), None
                )
                if cpu_snap and cpu_snap.get("Busy_Pct") is not None:
                    phys = round(cpu_snap["Busy_Pct"] * ent / 100.0, 4)
            lpar_cap_records.append({
                "Timestamp":        ts,
                "Hostname":         self.hostname,
                "LPAR_Physical_CPU":round(phys, 4) if phys is not None else None,
                "LPAR_Entitled_CPU":ent,
                "LPAR_Virtual_CPUs":virt,
            })

        # If no LPAR section rows but we have CPU rows, derive LPAR capacity from there
        if not lpar_cap_records and records:
            for r in records:
                if r.get("LPAR_Physical_CPU") is not None:
                    lpar_cap_records.append({
                        "Timestamp":        r["Timestamp"],
                        "Hostname":         self.hostname,
                        "LPAR_Physical_CPU":r["LPAR_Physical_CPU"],
                        "LPAR_Entitled_CPU":r["LPAR_Entitled_CPU"],
                        "LPAR_Virtual_CPUs":r["LPAR_Virtual_CPUs"],
                    })

        self.lpar_raw = pd.DataFrame(lpar_cap_records) if lpar_cap_records else pd.DataFrame(
            columns=["Timestamp","Hostname","LPAR_Physical_CPU",
                     "LPAR_Entitled_CPU","LPAR_Virtual_CPUs"])

        for tag, msg in warnings:
            self.warnings.append({"Hostname": self.hostname, "Section": tag,
                                  "Issue": msg, "Severity": "Warning"})

    # ── Memory ────────────────────────────────────────────────────────────────

    def _parse_mem(self, sections: dict) -> None:
        records       = []
        paging_records = []
        mem_header    = None
        pg_header     = None

        for row in sections.get("MEM", []):
            tag = row[0].strip().upper()
            if tag == "MEM" and len(row) > 2 and row[1].strip().upper() == "MEM":
                mem_header = [x.strip().upper() for x in row[2:]]
            elif tag == "MEM" and mem_header:
                snap_id = row[1].strip() if len(row) > 1 else ""
                ts      = self._ts(snap_id)
                vals    = row[2:]
                d       = dict(zip(mem_header, vals))

                def _m(*keys):
                    for k in keys:
                        v = _to_float(d.get(k))
                        if v is not None:
                            return v
                    return None

                real_mb  = _m("REALMEM(MB)", "REALMEM", "TOTALMEM", "REAL_MB")
                free_mb  = _m("FREEMEM(MB)", "FREEMEM",  "FREE_MB")
                used_mb  = None
                if real_mb is not None and free_mb is not None:
                    used_mb = max(0.0, real_mb - free_mb)

                total_gb = round(real_mb / 1024, 3) if real_mb is not None else None
                used_gb  = round(used_mb / 1024, 3) if used_mb is not None else None
                free_gb  = round(free_mb / 1024, 3) if free_mb is not None else None
                used_pct = round(used_mb / real_mb * 100, 2) \
                           if (real_mb and real_mb > 0 and used_mb is not None) else None

                # Try to patch installed memory from first row
                if self.installed_mem_gb is None and total_gb is not None:
                    self.installed_mem_gb = total_gb

                records.append({
                    "Timestamp":      ts,
                    "Hostname":       self.hostname,
                    "Total_Memory_GB":total_gb,
                    "Used_Memory_GB": used_gb,
                    "Free_Memory_GB": free_gb,
                    "Used_Memory_Pct":used_pct,
                })

        if not records:
            self.warnings.append({"Hostname": self.hostname, "Section": "MEM",
                                  "Issue": "No MEM section data found", "Severity": "Warning"})

        # Paging
        for row in sections.get("PAGING", []):
            tag = row[0].strip().upper()
            if tag == "PAGING" and len(row) > 2 and row[1].strip().upper() == "PAGING":
                pg_header = [x.strip().upper() for x in row[2:]]
            elif tag == "PAGING" and pg_header:
                snap_id = row[1].strip() if len(row) > 1 else ""
                ts      = self._ts(snap_id)
                vals    = row[2:]
                d       = dict(zip(pg_header, vals))

                def _p(*keys):
                    for k in keys:
                        v = _to_float(d.get(k))
                        if v is not None:
                            return v
                    return None

                paging_records.append({
                    "Timestamp":  ts,
                    "Hostname":   self.hostname,
                    "PageIn":     _p("PAGEIN", "PAGE_IN", "PGIN", "PAGEINS"),
                    "PageOut":    _p("PAGEOUT", "PAGE_OUT", "PGOUT", "PAGEOUTS"),
                    "PageRate":   _p("PAGINGRATE", "PAGING", "PGRATE"),
                })

        self.mem_raw    = pd.DataFrame(records) if records else pd.DataFrame(
            columns=["Timestamp","Hostname","Total_Memory_GB",
                     "Used_Memory_GB","Free_Memory_GB","Used_Memory_Pct"])
        self.paging_raw = pd.DataFrame(paging_records) if paging_records else pd.DataFrame(
            columns=["Timestamp","Hostname","PageIn","PageOut","PageRate"])

    # ── Disk ──────────────────────────────────────────────────────────────────

    def _parse_disk(self, sections: dict) -> None:
        # We want DISKBUSY primarily; use DISKREAD/DISKWRITE/DISKXFER for throughput
        disk_headers: dict = {}
        disk_data: dict    = {}   # {section → [rows]}

        for sec in ("DISKBUSY", "DISKREAD", "DISKWRITE", "DISKXFER", "DISKBSIZE"):
            rows = sections.get(sec, [])
            for row in rows:
                tag = row[0].strip().upper()
                if tag == sec and len(row) > 2 and row[1].strip().upper() == sec:
                    disk_headers[sec] = [x.strip() for x in row[2:]]
                elif tag == sec and sec in disk_headers:
                    snap_id = row[1].strip() if len(row) > 1 else ""
                    disk_data.setdefault(sec, []).append(
                        (snap_id, row[2:])
                    )

        if not disk_headers:
            self.warnings.append({"Hostname": self.hostname, "Section": "DISK",
                                  "Issue": "No DISKBUSY/DISKREAD/DISKWRITE sections found",
                                  "Severity": "Info"})
            self.disk_raw = pd.DataFrame(
                columns=["Timestamp","Hostname","Disk_Name","Busy_Pct",
                         "Read_KB_Sec","Write_KB_Sec","Transfers_Sec"])
            return

        # Build a combined row per (snap, disk_name)
        combo: dict = {}   # (snap_id, disk_name) → {fields}
        headers_busy = disk_headers.get("DISKBUSY", [])
        for snap_id, vals in disk_data.get("DISKBUSY", []):
            ts = self._ts(snap_id)
            for disk, v in zip(headers_busy, vals):
                key = (snap_id, disk.strip())
                combo[key] = {
                    "Timestamp": ts,
                    "Hostname":  self.hostname,
                    "Disk_Name": disk.strip(),
                    "Busy_Pct":  _to_float(v),
                    "Read_KB_Sec":   None,
                    "Write_KB_Sec":  None,
                    "Transfers_Sec": None,
                }

        for sec, col in (("DISKREAD","Read_KB_Sec"),
                         ("DISKWRITE","Write_KB_Sec"),
                         ("DISKXFER","Transfers_Sec")):
            hdrs = disk_headers.get(sec, [])
            for snap_id, vals in disk_data.get(sec, []):
                ts = self._ts(snap_id)
                for disk, v in zip(hdrs, vals):
                    key = (snap_id, disk.strip())
                    if key not in combo:
                        combo[key] = {
                            "Timestamp": ts,
                            "Hostname":  self.hostname,
                            "Disk_Name": disk.strip(),
                            "Busy_Pct":  None,
                            "Read_KB_Sec":  None,
                            "Write_KB_Sec": None,
                            "Transfers_Sec":None,
                        }
                    combo[key][col] = _to_float(v)

        self.disk_raw = pd.DataFrame(list(combo.values())) if combo else pd.DataFrame(
            columns=["Timestamp","Hostname","Disk_Name","Busy_Pct",
                     "Read_KB_Sec","Write_KB_Sec","Transfers_Sec"])

    # ── Network ───────────────────────────────────────────────────────────────

    def _parse_net(self, sections: dict) -> None:
        net_header = None
        records    = []

        for row in sections.get("NET", []):
            tag = row[0].strip().upper()
            if tag == "NET" and len(row) > 2 and row[1].strip().upper() == "NET":
                net_header = [x.strip() for x in row[2:]]
            elif tag == "NET" and net_header:
                snap_id = row[1].strip() if len(row) > 1 else ""
                ts      = self._ts(snap_id)
                vals    = row[2:]
                for iface_metric, v in zip(net_header, vals):
                    # Header format: "eth0-read-KB/s" or "eth0-write-KB/s"
                    m = re.match(r"(.+?)-(read|write).*", iface_metric, re.IGNORECASE)
                    if not m:
                        continue
                    iface     = m.group(1)
                    direction = m.group(2).lower()
                    val       = _to_float(v)
                    # Find or create existing record for this snap+interface
                    existing = next(
                        (r for r in records
                         if r["Timestamp"] == ts and r["Interface"] == iface),
                        None
                    )
                    if existing is None:
                        existing = {
                            "Timestamp":   ts,
                            "Hostname":    self.hostname,
                            "Interface":   iface,
                            "Read_KB_Sec": None,
                            "Write_KB_Sec":None,
                        }
                        records.append(existing)
                    if direction == "read":
                        existing["Read_KB_Sec"]  = val
                    else:
                        existing["Write_KB_Sec"] = val

        if not records:
            self.warnings.append({"Hostname": self.hostname, "Section": "NET",
                                  "Issue": "No NET section data found", "Severity": "Info"})

        self.net_raw = pd.DataFrame(records) if records else pd.DataFrame(
            columns=["Timestamp","Hostname","Interface","Read_KB_Sec","Write_KB_Sec"])

    # ── infer collection period ───────────────────────────────────────────────

    def _infer_collection_period(self) -> None:
        # If timestamps already set from ZZZZ, do nothing.
        # Otherwise fall back to file mtime.
        if self.ts_start and self.ts_end:
            return
        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(self.path))
            if self.ts_start is None:
                self.ts_start = mtime
            if self.ts_end is None:
                self.ts_end = mtime
        except OSError:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# NMON PROCESSOR — orchestrates multiple files and builds the workbook
# ─────────────────────────────────────────────────────────────────────────────

class NmonProcessor:

    def __init__(self, output_dir: str = "."):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.nmon_files: list[NmonFile] = []
        self.warnings:   list[dict]     = []

        # Consolidated frames (populated in analyse())
        self.cpu_raw:    pd.DataFrame = pd.DataFrame()
        self.lpar_raw:   pd.DataFrame = pd.DataFrame()
        self.mem_raw:    pd.DataFrame = pd.DataFrame()
        self.paging_raw: pd.DataFrame = pd.DataFrame()
        self.disk_raw:   pd.DataFrame = pd.DataFrame()
        self.net_raw:    pd.DataFrame = pd.DataFrame()

        # Summary frames
        self.cpu_summary:    pd.DataFrame = pd.DataFrame()
        self.lpar_summary:   pd.DataFrame = pd.DataFrame()
        self.mem_summary:    pd.DataFrame = pd.DataFrame()
        self.disk_summary:   pd.DataFrame = pd.DataFrame()
        self.net_summary:    pd.DataFrame = pd.DataFrame()
        self.kpi_summary:    pd.DataFrame = pd.DataFrame()
        self.exec_summary:   pd.DataFrame = pd.DataFrame()
        self.sys_inventory:  pd.DataFrame = pd.DataFrame()

    # ── file discovery ────────────────────────────────────────────────────────

    def discover(self, paths: list) -> int:
        """
        Expand a list of file paths / directory paths into individual .nmon files.
        Returns total count.
        """
        found = []
        print("\n[1/8] Discovering NMON files…")
        for p in tqdm(paths, desc="  Scanning", unit="path"):
            p = Path(p)
            if p.is_dir():
                found.extend(p.rglob("*.nmon"))
                found.extend(p.rglob("*.NMON"))
            elif p.is_file() and p.suffix.lower() == ".nmon":
                found.append(p)
        found = sorted(set(found))
        print(f"  → Found {len(found)} NMON file(s).")
        self._paths = found
        return len(found)

    # ── parse ─────────────────────────────────────────────────────────────────

    def parse_all(self) -> None:
        print("\n[2/8] Parsing NMON files…")
        for path in tqdm(self._paths, desc="  Parsing", unit="file"):
            tqdm.write(f"      {path.name}")
            try:
                nf = NmonFile(str(path))
                nf.parse()
                self.nmon_files.append(nf)
                self.warnings.extend(nf.warnings)
            except Exception as exc:
                self.warnings.append({
                    "Hostname": path.stem, "Section": "PARSE",
                    "Issue": f"Failed to parse file: {exc}", "Severity": "Error"
                })

        if not self.nmon_files:
            raise RuntimeError("No NMON files could be parsed. Aborting.")

    # ── consolidate raw frames ────────────────────────────────────────────────

    def _consolidate(self) -> None:
        cpu_parts  = [nf.cpu_raw  for nf in self.nmon_files if not nf.cpu_raw.empty]
        lpar_parts = [nf.lpar_raw for nf in self.nmon_files if not nf.lpar_raw.empty]
        mem_parts  = [nf.mem_raw  for nf in self.nmon_files if not nf.mem_raw.empty]
        pg_parts   = [nf.paging_raw for nf in self.nmon_files if not nf.paging_raw.empty]
        disk_parts = [nf.disk_raw for nf in self.nmon_files if not nf.disk_raw.empty]
        net_parts  = [nf.net_raw  for nf in self.nmon_files if not nf.net_raw.empty]

        if cpu_parts:
            self.cpu_raw = pd.concat(cpu_parts, ignore_index=True)
        if lpar_parts:
            self.lpar_raw = pd.concat(lpar_parts, ignore_index=True)
        if mem_parts:
            self.mem_raw = pd.concat(mem_parts, ignore_index=True)
        if pg_parts:
            self.paging_raw = pd.concat(pg_parts, ignore_index=True)
        if disk_parts:
            self.disk_raw = pd.concat(disk_parts, ignore_index=True)
        if net_parts:
            self.net_raw = pd.concat(net_parts, ignore_index=True)

    # ── CPU analysis ──────────────────────────────────────────────────────────

    def _analyse_cpu(self) -> None:
        if self.cpu_raw.empty:
            return
        df = self.cpu_raw.dropna(subset=["Busy_Pct"])
        rows = []
        for host, grp in df.groupby("Hostname"):
            bp = grp["Busy_Pct"]
            peak_ts = None
            if not bp.empty:
                idx = bp.idxmax()
                peak_ts = grp.loc[idx, "Timestamp"] if idx in grp.index else None
            rows.append({
                "SID":          grp["SID"].iloc[0] if "SID" in grp.columns else "",
                "Hostname":     host,
                "CPU_Avg_Pct":  round(bp.mean(), 2),
                "CPU_Min_Pct":  round(bp.min(),  2),
                "CPU_Max_Pct":  round(bp.max(),  2),
                "CPU_P95_Pct":  round(_pct(bp, 95), 2),
                "CPU_P99_Pct":  round(_pct(bp, 99), 2),
                "Peak_Timestamp": peak_ts,
            })
        self.cpu_summary = pd.DataFrame(rows)

    # ── LPAR capacity analysis ────────────────────────────────────────────────

    def _analyse_lpar(self) -> None:
        if self.lpar_raw.empty:
            return
        df = self.lpar_raw.dropna(subset=["LPAR_Physical_CPU"])
        rows = []
        for host, grp in df.groupby("Hostname"):
            pc = grp["LPAR_Physical_CPU"]
            peak_ts = None
            if not pc.empty:
                idx = pc.idxmax()
                peak_ts = grp.loc[idx, "Timestamp"] if idx in grp.index else None
            rows.append({
                "Hostname":            host,
                "Physical_CPU_Avg":    round(pc.mean(), 4),
                "Physical_CPU_Max":    round(pc.max(),  4),
                "Physical_CPU_P95":    round(_pct(pc, 95), 4),
                "Peak_Physical_CPU_TS":peak_ts,
                "Entitled_CPUs":       grp["LPAR_Entitled_CPU"].dropna().mean(),
                "Virtual_CPUs":        grp["LPAR_Virtual_CPUs"].dropna().mean(),
            })
        self.lpar_summary = pd.DataFrame(rows)

    # ── Memory analysis ───────────────────────────────────────────────────────

    def _analyse_mem(self) -> None:
        pg_totals: dict = {}
        if not self.paging_raw.empty:
            pg = self.paging_raw.dropna(subset=["PageOut"])
            for host, grp in pg.groupby("Hostname"):
                pg_totals[host] = {
                    "Total_Page_In":     grp["PageIn"].sum() if "PageIn" in grp else 0,
                    "Total_Page_Out":    grp["PageOut"].sum(),
                    "Avg_Page_Rate":     grp["PageRate"].mean() if "PageRate" in grp else None,
                    "Max_Page_Rate":     grp["PageRate"].max()  if "PageRate" in grp else None,
                }

        rows = []
        if self.mem_raw.empty:
            return
        df = self.mem_raw.dropna(subset=["Used_Memory_Pct"])
        for host, grp in df.groupby("Hostname"):
            mp = grp["Used_Memory_Pct"]
            pg = pg_totals.get(host, {})
            rows.append({
                "Hostname":          host,
                "Avg_Memory_Pct":    round(mp.mean(), 2),
                "Peak_Memory_Pct":   round(mp.max(),  2),
                "P95_Memory_Pct":    round(_pct(mp, 95), 2),
                "Total_Memory_GB":   round(grp["Total_Memory_GB"].dropna().max() or 0, 2),
                "Total_Page_In":     pg.get("Total_Page_In",  0),
                "Total_Page_Out":    pg.get("Total_Page_Out", 0),
                "Avg_Page_Rate":     pg.get("Avg_Page_Rate"),
                "Max_Page_Rate":     pg.get("Max_Page_Rate"),
            })
        self.mem_summary = pd.DataFrame(rows)

    # ── Disk analysis ─────────────────────────────────────────────────────────

    def _analyse_disk(self) -> None:
        if self.disk_raw.empty:
            return
        df = self.disk_raw
        rows = []
        for (host, disk), grp in df.groupby(["Hostname", "Disk_Name"]):
            bp = grp["Busy_Pct"].dropna()
            rd = grp["Read_KB_Sec"].dropna()
            wr = grp["Write_KB_Sec"].dropna()
            rows.append({
                "Hostname":        host,
                "Disk_Name":       disk,
                "Avg_Busy_Pct":    round(bp.mean(), 2) if len(bp) else None,
                "Peak_Busy_Pct":   round(bp.max(),  2) if len(bp) else None,
                "P95_Busy_Pct":    round(_pct(bp, 95), 2) if len(bp) else None,
                "Avg_Read_KB_Sec": round(rd.mean(), 2) if len(rd) else None,
                "Avg_Write_KB_Sec":round(wr.mean(), 2) if len(wr) else None,
                "Peak_Read_KB_Sec":round(rd.max(),  2) if len(rd) else None,
                "Peak_Write_KB_Sec":round(wr.max(), 2) if len(wr) else None,
            })
        self.disk_summary = pd.DataFrame(rows)

    # ── Network analysis ──────────────────────────────────────────────────────

    def _analyse_net(self) -> None:
        if self.net_raw.empty:
            return
        df = self.net_raw.copy()
        df["Total_KB_Sec"] = (df["Read_KB_Sec"].fillna(0) + df["Write_KB_Sec"].fillna(0))
        rows = []
        for (host, iface), grp in df.groupby(["Hostname", "Interface"]):
            rd = grp["Read_KB_Sec"].dropna()
            wr = grp["Write_KB_Sec"].dropna()
            total = grp["Total_KB_Sec"].dropna()
            rows.append({
                "Hostname":        host,
                "Interface":       iface,
                "Avg_Read_KB_Sec": round(rd.mean(),    2) if len(rd) else None,
                "Avg_Write_KB_Sec":round(wr.mean(),    2) if len(wr) else None,
                "Peak_Read_KB_Sec":round(rd.max(),     2) if len(rd) else None,
                "Peak_Write_KB_Sec":round(wr.max(),    2) if len(wr) else None,
                "Avg_Total_KB_Sec":round(total.mean(), 2) if len(total) else None,
                "Peak_Total_KB_Sec":round(total.max(), 2) if len(total) else None,
                "P95_Total_KB_Sec": round(_pct(total, 95), 2) if len(total) else None,
            })
        self.net_summary = pd.DataFrame(rows)

    # ── System inventory ──────────────────────────────────────────────────────

    def _build_inventory(self) -> None:
        rows = []
        for nf in self.nmon_files:
            rows.append({
                "SID":               nf.sid,
                "Hostname":          nf.hostname,
                "LPAR":              nf.lpar_name,
                "Machine_Type":      nf.machine_type,
                "Processor_Model":   nf.processor_model,
                "Processor_Type":    nf.processor_type,
                "Operating_System":  nf.os_name,
                "OS_Version":        nf.os_version,
                "Serial_Number":     nf.serial,
                "Physical_CPUs":     nf.physical_cpus,
                "Virtual_CPUs":      nf.virtual_cpus,
                "Entitled_CPUs":     nf.entitled_cpus,
                "Installed_Memory_GB": nf.installed_mem_gb,
                "Virtualization":    nf.virt_tech,
                "Collection_Start":  nf.ts_start,
                "Collection_End":    nf.ts_end,
                "Source_File":       Path(nf.path).name,
            })
        self.sys_inventory = pd.DataFrame(rows)

    # ── KPI summary ───────────────────────────────────────────────────────────

    def _build_kpi(self) -> None:
        # Pivot all per-host summaries into one row per hostname
        hosts = sorted({nf.hostname for nf in self.nmon_files})
        cpu_idx  = self.cpu_summary.set_index("Hostname")  if not self.cpu_summary.empty  else pd.DataFrame()
        lpar_idx = self.lpar_summary.set_index("Hostname") if not self.lpar_summary.empty else pd.DataFrame()
        mem_idx  = self.mem_summary.set_index("Hostname")  if not self.mem_summary.empty  else pd.DataFrame()

        # Disk: highest busy disk per host
        disk_max: dict = {}
        if not self.disk_summary.empty:
            for host, grp in self.disk_summary.groupby("Hostname"):
                bp = grp["Peak_Busy_Pct"].dropna()
                disk_max[host] = round(bp.max(), 2) if len(bp) else None

        # Network: highest peak throughput per host
        net_max: dict = {}
        if not self.net_summary.empty:
            for host, grp in self.net_summary.groupby("Hostname"):
                tp = grp["Peak_Total_KB_Sec"].dropna()
                net_max[host] = round(tp.max(), 2) if len(tp) else None

        # Paging totals per host
        pg_out: dict = {}
        if not self.mem_summary.empty:
            for _, row in self.mem_summary.iterrows():
                pg_out[row["Hostname"]] = row.get("Total_Page_Out", 0)

        # Collection period per host
        coll: dict = {}
        for nf in self.nmon_files:
            coll[nf.hostname] = (nf.ts_start, nf.ts_end)

        rows = []
        for host in hosts:
            cpu  = cpu_idx.loc[host]  if host in cpu_idx.index  else {}
            lpar = lpar_idx.loc[host] if host in lpar_idx.index else {}
            mem  = mem_idx.loc[host]  if host in mem_idx.index  else {}
            rows.append({
                "Hostname":           host,
                "CPU_Avg_Pct":        cpu.get("CPU_Avg_Pct"),
                "CPU_P95_Pct":        cpu.get("CPU_P95_Pct"),
                "CPU_Max_Pct":        cpu.get("CPU_Max_Pct"),
                "Physical_CPU_Avg":   lpar.get("Physical_CPU_Avg"),
                "Physical_CPU_P95":   lpar.get("Physical_CPU_P95"),
                "Physical_CPU_Max":   lpar.get("Physical_CPU_Max"),
                "Memory_Avg_Pct":     mem.get("Avg_Memory_Pct"),
                "Memory_P95_Pct":     mem.get("P95_Memory_Pct"),
                "Memory_Max_Pct":     mem.get("Peak_Memory_Pct"),
                "Disk_Max_Busy_Pct":  disk_max.get(host),
                "Network_Max_KB_Sec": net_max.get(host),
                "Total_Page_Out":     pg_out.get(host, 0),
                "Collection_Start":   coll.get(host, (None, None))[0],
                "Collection_End":     coll.get(host, (None, None))[1],
            })
        self.kpi_summary = pd.DataFrame(rows)

    # ── Executive summary ─────────────────────────────────────────────────────

    def _build_exec_summary(self) -> None:
        inv_idx = self.sys_inventory.set_index("Hostname") \
                  if not self.sys_inventory.empty else pd.DataFrame()
        rows = []
        for _, kpi in self.kpi_summary.iterrows():
            host = kpi["Hostname"]
            inv  = inv_idx.loc[host] if host in inv_idx.index else {}

            comments = []
            cpu_p95 = kpi.get("CPU_P95_Pct")
            mem_p95 = kpi.get("Memory_P95_Pct")
            dsk_p95 = kpi.get("Disk_Max_Busy_Pct")   # using peak as proxy
            pg_out  = kpi.get("Total_Page_Out", 0) or 0

            if cpu_p95 and cpu_p95 > 80:
                comments.append("Review CPU sizing")
            if mem_p95 and mem_p95 > 85:
                comments.append("Memory pressure detected")
            if dsk_p95 and dsk_p95 > 70:
                comments.append("Storage bottleneck risk")
            if pg_out > 0:
                comments.append("Investigate memory pressure (paging detected)")
            if not comments:
                comments.append("No significant bottlenecks detected")

            rows.append({
                "SID":                    inv.get("SID", ""),
                "Hostname":               host,
                "Virtual_CPUs":           inv.get("Virtual_CPUs"),
                "Entitled_CPUs":          inv.get("Entitled_CPUs"),
                "CPU_Avg_Pct":            kpi.get("CPU_Avg_Pct"),
                "CPU_P95_Pct":            kpi.get("CPU_P95_Pct"),
                "CPU_Max_Pct":            kpi.get("CPU_Max_Pct"),
                "Physical_CPU_Avg":       kpi.get("Physical_CPU_Avg"),
                "Physical_CPU_P95":       kpi.get("Physical_CPU_P95"),
                "Physical_CPU_Max":       kpi.get("Physical_CPU_Max"),
                "Memory_Avg_Pct":         kpi.get("Memory_Avg_Pct"),
                "Memory_P95_Pct":         kpi.get("Memory_P95_Pct"),
                "Memory_Max_Pct":         kpi.get("Memory_Max_Pct"),
                "Highest_Disk_Busy_Pct":  kpi.get("Disk_Max_Busy_Pct"),
                "Highest_Network_KB_Sec": kpi.get("Network_Max_KB_Sec"),
                "Total_Page_Out":         pg_out,
                "Comments":               " | ".join(comments),
            })
        self.exec_summary = pd.DataFrame(rows)

    # ── run full analysis pipeline ────────────────────────────────────────────

    def analyse(self) -> None:
        steps = [
            ("[3/8] CPU analysis…",     self._analyse_cpu),
            ("[4/8] Memory analysis…",  self._analyse_mem),
            ("[5/8] Disk analysis…",    self._analyse_disk),
            ("[6/8] Network analysis…", self._analyse_net),
        ]
        print("\n[2b/8] Consolidating parsed data…")
        self._consolidate()
        for label, fn in tqdm(steps, desc="  Analysing", unit="step"):
            tqdm.write(f"  {label}")
            fn()

        print("\n[7/8] Building KPI and summary tables…")
        self._analyse_lpar()
        self._build_inventory()
        self._build_kpi()
        self._build_exec_summary()

    # ── output filename ───────────────────────────────────────────────────────

    def _output_filename(self) -> str:
        sids = sorted({nf.sid for nf in self.nmon_files if nf.sid})
        if len(sids) == 1:
            prefix = sids[0]
        elif len(sids) > 1:
            prefix = "MULTI"
        else:
            prefix = "NMON"
        return f"{prefix}_NMON_Analysis.xlsx"

    # ── Excel generation ──────────────────────────────────────────────────────

    def write_excel(self) -> str:
        fname = self.output_dir / self._output_filename()
        print(f"\n[8/8] Writing Excel workbook → {fname}")

        wb = openpyxl.Workbook()
        wb.remove(wb.active)   # remove default sheet

        sheet_order = [
            ("Executive_Summary",  self.exec_summary),
            ("System_Inventory",   self.sys_inventory),
            ("CPU_Raw",            self.cpu_raw),
            ("CPU_Summary",        self.cpu_summary),
            ("LPAR_Capacity",      self.lpar_raw),
            ("Memory_Raw",         self.mem_raw),
            ("Memory_Summary",     self.mem_summary),
            ("Disk_Raw",           self.disk_raw),
            ("Disk_Summary",       self.disk_summary),
            ("Network_Raw",        self.net_raw),
            ("Network_Summary",    self.net_summary),
            ("KPI_Summary",        self.kpi_summary),
            ("Warnings",           pd.DataFrame(self.warnings) if self.warnings
                                   else pd.DataFrame(columns=["Hostname","Section",
                                                               "Issue","Severity"])),
        ]

        for sheet_name, df in tqdm(sheet_order, desc="  Sheets", unit="sheet"):
            tqdm.write(f"      {sheet_name}")
            ws = wb.create_sheet(title=sheet_name)
            self._write_sheet(ws, sheet_name, df)

        wb.save(str(fname))
        return str(fname)

    # ── sheet writer dispatcher ───────────────────────────────────────────────

    def _write_sheet(self, ws, name: str, df: pd.DataFrame) -> None:
        if df is None or df.empty:
            ws.append(["No data available"])
            ws["A1"].font = BOLD_FONT
            return

        df = df.copy()

        # Ensure Timestamp columns are native datetime (not Timestamp objects)
        for col in df.select_dtypes(include=["datetime64", "datetimetz"]).columns:
            df[col] = pd.to_datetime(df[col]).dt.to_pydatetime()
            df[col] = df[col].apply(lambda x: x if pd.notnull(x) else None)

        _write_df(ws, df)

        max_row = ws.max_row
        max_col = ws.max_column

        # Freeze top row
        ws.freeze_panes = "A2"

        # Auto-filter
        ws.auto_filter.ref = ws.dimensions

        # Auto-size columns
        _auto_col_width(ws)

        # Sheet-specific conditional formatting
        col_map = {cell.value: get_column_letter(cell.column)
                   for cell in ws[1]}

        cf_specs = {
            "Executive_Summary": [
                ("CPU_P95_Pct",            80, 90),
                ("CPU_Max_Pct",            80, 90),
                ("Memory_P95_Pct",         80, 90),
                ("Memory_Max_Pct",         80, 90),
                ("Highest_Disk_Busy_Pct",  70, 90),
            ],
            "CPU_Raw": [
                ("Busy_Pct",  80, 90),
            ],
            "CPU_Summary": [
                ("CPU_Avg_Pct", 80, 90),
                ("CPU_P95_Pct", 80, 90),
                ("CPU_Max_Pct", 80, 90),
            ],
            "Memory_Raw": [
                ("Used_Memory_Pct", 80, 90),
            ],
            "Memory_Summary": [
                ("Avg_Memory_Pct",  80, 90),
                ("Peak_Memory_Pct", 80, 90),
                ("P95_Memory_Pct",  80, 90),
            ],
            "Disk_Raw": [
                ("Busy_Pct", 70, 90),
            ],
            "Disk_Summary": [
                ("Avg_Busy_Pct",  70, 90),
                ("Peak_Busy_Pct", 70, 90),
                ("P95_Busy_Pct",  70, 90),
            ],
            "KPI_Summary": [
                ("CPU_Avg_Pct",   80, 90),
                ("CPU_P95_Pct",   80, 90),
                ("CPU_Max_Pct",   80, 90),
                ("Memory_Avg_Pct",80, 90),
                ("Memory_P95_Pct",80, 90),
                ("Memory_Max_Pct",80, 90),
                ("Disk_Max_Busy_Pct", 70, 90),
            ],
        }

        for col_name, yellow_lo, red_lo in cf_specs.get(name, []):
            letter = col_map.get(col_name)
            if letter:
                _pct_col_cf(ws, letter, 2, max(max_row, 2), yellow_lo, red_lo)

        # Number formatting for percentage columns
        pct_cols = {c for c in col_map if "pct" in c.lower() or "%" in c.lower()}
        for col_name in pct_cols:
            letter = col_map.get(col_name)
            if not letter:
                continue
            for row in ws.iter_rows(min_row=2, max_row=max_row,
                                    min_col=ws[f"{letter}1"].column,
                                    max_col=ws[f"{letter}1"].column):
                for cell in row:
                    if cell.value is not None:
                        cell.number_format = "0.00"

        # Number formatting for GB / numeric columns
        gb_cols = {c for c in col_map if "gb" in c.lower() or "memory" in c.lower()
                   or "cpu" in c.lower() or "kb" in c.lower()}
        for col_name in gb_cols:
            letter = col_map.get(col_name)
            if not letter:
                continue
            for row in ws.iter_rows(min_row=2, max_row=max_row,
                                    min_col=ws[f"{letter}1"].column,
                                    max_col=ws[f"{letter}1"].column):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0.00"

        # Timestamp columns — ISO datetime display
        ts_cols = {c for c in col_map if "timestamp" in c.lower()
                   or "start" in c.lower() or "end" in c.lower()}
        for col_name in ts_cols:
            letter = col_map.get(col_name)
            if not letter:
                continue
            for row in ws.iter_rows(min_row=2, max_row=max_row,
                                    min_col=ws[f"{letter}1"].column,
                                    max_col=ws[f"{letter}1"].column):
                for cell in row:
                    if cell.value:
                        cell.number_format = "YYYY-MM-DD HH:MM:SS"

        # Warnings sheet: colour Severity column
        if name == "Warnings" and "Severity" in col_map:
            sev_letter = col_map["Severity"]
            for row in ws.iter_rows(min_row=2, max_row=max_row,
                                    min_col=ws[f"{sev_letter}1"].column,
                                    max_col=ws[f"{sev_letter}1"].column):
                for cell in row:
                    v = str(cell.value or "").lower()
                    if v == "error":
                        cell.fill = RED_FILL
                    elif v == "warning":
                        cell.fill = YELLOW_FILL
                    elif v == "info":
                        cell.fill = GREEN_FILL

    # ── summary report ────────────────────────────────────────────────────────

    def print_summary(self, output_path: str) -> None:
        n_files   = len(self.nmon_files)
        n_hosts   = len({nf.hostname for nf in self.nmon_files})
        n_warn    = len(self.warnings)
        n_records = sum([
            len(self.cpu_raw), len(self.mem_raw),
            len(self.disk_raw), len(self.net_raw),
        ])
        starts = [nf.ts_start for nf in self.nmon_files if nf.ts_start]
        ends   = [nf.ts_end   for nf in self.nmon_files if nf.ts_end]
        period = ""
        if starts and ends:
            period = f"{min(starts).strftime('%Y-%m-%d %H:%M')}  →  {max(ends).strftime('%Y-%m-%d %H:%M')}"

        print("\n" + "═" * 60)
        print("  NMON PROCESSING COMPLETE")
        print("═" * 60)
        print(f"  Files processed   : {n_files}")
        print(f"  Servers detected  : {n_hosts}")
        print(f"  Collection period : {period or 'N/A'}")
        print(f"  Total records     : {n_records:,}")
        print(f"  Warnings          : {n_warn}")
        print(f"  Output file       : {output_path}")
        print("═" * 60 + "\n")

    # ── main entry point ──────────────────────────────────────────────────────

    def run(self, paths: list) -> str:
        n = self.discover(paths)
        if n == 0:
            raise RuntimeError("No .nmon files found in the provided paths.")
        self.parse_all()
        self.analyse()
        output_path = self.write_excel()
        self.print_summary(output_path)
        return output_path


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Process IBM NMON files and generate a consolidated Excel workbook.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "files", nargs="*",
        help="One or more .nmon files or directories to scan"
    )
    parser.add_argument(
        "--input", "-i", dest="input_dir", default=None,
        help="Directory to scan for *.nmon files (alternative to listing files)"
    )
    parser.add_argument(
        "--output", "-o", dest="output_dir", default=".",
        help="Directory to write the output workbook (default: current directory)"
    )
    args = parser.parse_args()

    paths = list(args.files)
    if args.input_dir:
        paths.append(args.input_dir)
    if not paths:
        paths = ["."]

    proc = NmonProcessor(output_dir=args.output_dir)
    try:
        proc.run(paths)
    except RuntimeError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
