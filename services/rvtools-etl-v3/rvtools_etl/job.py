"""
vmiq-rvtools-etl v3 — multi-estate, multi-vCenter, all xlsx files
==================================================================
Scans /uploads/ and loads EVERY .xlsx file found.

Directory structure:
  /uploads/
  ├── ITEAST/
  │   ├── vcenter-a.xlsx       ← any filename, just needs .xlsx
  │   ├── vcenter-b.xlsx
  │   ├── vcenter-prod-01.xlsx
  │   └── ... (all 21 files)
  ├── ITWEST/
  │   ├── vcenter-x.xlsx
  │   └── vcenter-y.xlsx
  └── ITCENTRAL/
      └── vcenter-z.xlsx

Rules:
  - Every .xlsx file in an estate subfolder is processed
  - Root level .xlsx files go to estate 'DEFAULT'
  - 'archive' subfolders are skipped
  - Files starting with '~' (Excel temp) are skipped
  - After a successful load, file is moved to archive/
  - If one file fails, others continue processing
"""

import os
import sys
import uuid
import asyncio
import logging
import shutil
from datetime import date, datetime
from pathlib import Path

import asyncpg

from .parser import RVToolsParser
from .writer import SnapshotWriter

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("rvtools-etl")


def get_config() -> dict:
    required = ["DATABASE_URL"]
    missing = [k for k in required if not os.getenv(k)]
    if missing:
        log.error(f"Missing required environment variables: {missing}")
        sys.exit(1)

    return {
        "database_url":     os.environ["DATABASE_URL"],
        "upload_dir":       Path(os.getenv("RVTOOLS_UPLOAD_DIR", "/uploads")),
        "keep_archives":    int(os.getenv("RVTOOLS_KEEP_ARCHIVES", "30")),
        "sheet_vms":        os.getenv("RVTOOLS_SHEET_VMS",        "vInfo"),
        "sheet_hosts":      os.getenv("RVTOOLS_SHEET_HOSTS",      "vHost"),
        "sheet_clusters":   os.getenv("RVTOOLS_SHEET_CLUSTERS",   "vCluster"),
        "sheet_datastores": os.getenv("RVTOOLS_SHEET_DATASTORES", "vDatastore"),
        "sheet_disks":      os.getenv("RVTOOLS_SHEET_DISKS",      "vDisk"),
        "sheet_snapshots":  os.getenv("RVTOOLS_SHEET_SNAPSHOTS",  "vSnapshot"),
    }


def find_all_xlsx(upload_dir: Path) -> list[tuple[Path, str]]:
    """
    Find every .xlsx file under upload_dir.
    Returns list of (filepath, estate_name) tuples.

    - Estate name = name of the immediate subdirectory
    - Root level files → estate 'DEFAULT'
    - Skips: archive/ folders, hidden files, ~$ temp files
    """
    files = []

    if not upload_dir.exists():
        log.error(f"Upload directory does not exist: {upload_dir}")
        return files

    # ── Estate subdirectories ──────────────────────────────────
    for subdir in sorted(upload_dir.iterdir()):
        if not subdir.is_dir():
            continue
        # Skip archive dirs and hidden dirs
        if subdir.name.lower() == "archive" or subdir.name.startswith("."):
            continue

        estate = subdir.name.upper()
        estate_files = sorted(subdir.glob("*.xlsx"))
        # Skip Excel temp files (~$filename.xlsx)
        estate_files = [f for f in estate_files
                        if not f.name.startswith("~")]

        if not estate_files:
            log.warning(f"Estate folder '{estate}' is empty — no .xlsx files")
            continue

        for f in estate_files:
            files.append((f, estate))
            log.info(f"  Found: [{estate}] {f.name}")

    # ── Root level files ───────────────────────────────────────
    root_files = sorted([
        f for f in upload_dir.glob("*.xlsx")
        if not f.name.startswith("~")
    ])
    for f in root_files:
        files.append((f, "DEFAULT"))
        log.info(f"  Found: [DEFAULT] {f.name}")

    return files


def archive_file(file_path: Path, keep: int) -> None:
    """
    Move processed file to archive/ subfolder with date stamp.
    Keeps last N archives per vCenter.
    Example: vcenter-a.xlsx → archive/vcenter-a_20260509.xlsx
    """
    archive_dir = file_path.parent / "archive"
    archive_dir.mkdir(exist_ok=True)

    date_stamp   = datetime.now().strftime("%Y%m%d_%H%M%S")
    archive_name = archive_dir / f"{file_path.stem}_{date_stamp}.xlsx"

    shutil.move(str(file_path), str(archive_name))
    log.info(f"  Archived → archive/{archive_name.name}")

    # Keep only last N archives for this filename stem
    stem_archives = sorted(
        archive_dir.glob(f"{file_path.stem}_*.xlsx"),
        reverse=True
    )
    for old in stem_archives[keep:]:
        old.unlink()
        log.info(f"  Deleted old archive: {old.name}")


async def log_pipeline_start(pool, run_id: str, file_count: int) -> None:
    await pool.execute("""
        INSERT INTO pipeline_run_log
            (id, run_date, pipeline, status, started_at, records_in)
        VALUES ($1, $2, 'rvtools', 'running', NOW(), $3)
    """, run_id, date.today(), file_count)


async def log_pipeline_end(pool, run_id: str, status: str,
                            records_out: int = 0,
                            error: str = None) -> None:
    await pool.execute("""
        UPDATE pipeline_run_log
        SET status=$2, finished_at=NOW(),
            records_out=$3, error_message=$4
        WHERE id=$1
    """, run_id, status, records_out, error)


async def main() -> None:
    cfg        = get_config()
    upload_dir = cfg["upload_dir"]

    log.info("=" * 60)
    log.info("RVTools ETL — multi-estate, all xlsx files")
    log.info(f"Upload dir : {upload_dir}")
    log.info(f"Run date   : {date.today()}")
    log.info("=" * 60)

    # ── 1. Discover all xlsx files ─────────────────────────────
    log.info("Scanning for xlsx files...")
    files = find_all_xlsx(upload_dir)

    if not files:
        log.error(
            f"No .xlsx files found under {upload_dir}\n"
            f"Expected:\n"
            f"  /uploads/ITEAST/vcenter-a.xlsx\n"
            f"  /uploads/ITEAST/vcenter-b.xlsx\n"
            f"  /uploads/ITWEST/vcenter-x.xlsx"
        )
        sys.exit(1)

    log.info(f"Total files to process: {len(files)}")

    # Group by estate for cleaner logging
    estates: dict[str, list] = {}
    for f, estate in files:
        estates.setdefault(estate, []).append(f)
    for estate, elist in estates.items():
        log.info(f"  {estate}: {len(elist)} file(s)")

    # ── 2. Connect to database ─────────────────────────────────
    log.info("Connecting to PostgreSQL...")
    pool = await asyncpg.create_pool(
        cfg["database_url"],
        min_size=2, max_size=8,
        command_timeout=60,
    )
    log.info("Connected.")

    run_id = str(uuid.uuid4())
    await log_pipeline_start(pool, run_id, len(files))

    # ── 3. Process every file ──────────────────────────────────
    sheet_config = {
        "vms":        cfg["sheet_vms"],
        "hosts":      cfg["sheet_hosts"],
        "clusters":   cfg["sheet_clusters"],
        "datastores": cfg["sheet_datastores"],
        "disks":      cfg["sheet_disks"],
        "snapshots":  cfg["sheet_snapshots"],
    }

    total_vms = 0
    succeeded = []
    failed    = []

    for idx, (file_path, estate_name) in enumerate(files, 1):
        label    = f"[{estate_name}] {file_path.name}"
        size_mb  = file_path.stat().st_size / 1024 / 1024

        log.info("-" * 60)
        log.info(f"File {idx}/{len(files)}: {label} ({size_mb:.1f} MB)")

        try:
            # Parse
            parser = RVToolsParser(
                filepath=file_path,
                sheet_config=sheet_config,
                estate=estate_name,
            )
            data = parser.parse()

            vm_count = len(data["vms"])
            log.info(
                f"  Parsed  — vCenters: {len(data['vcenters'])}, "
                f"Clusters: {len(data['clusters'])}, "
                f"Hosts: {len(data['hosts'])}, "
                f"VMs: {vm_count}, "
                f"Datastores: {len(data['datastores'])}"
            )

            if vm_count == 0:
                raise ValueError(
                    f"No VMs parsed from '{cfg['sheet_vms']}' sheet. "
                    f"Check sheet name and column headers."
                )

            # Write to DB
            writer = SnapshotWriter(pool)
            await writer.write(data)
            total_vms += vm_count

            # Archive the processed file
            archive_file(file_path, cfg["keep_archives"])

            succeeded.append(label)
            log.info(f"  ✓ Done  — {vm_count} VMs loaded")

        except Exception as exc:
            log.error(f"  ✗ FAILED: {exc}", exc_info=True)
            failed.append(f"{label}: {exc}")
            # Continue — don't stop processing remaining files

    # ── 4. Summary ─────────────────────────────────────────────
    log.info("=" * 60)
    log.info(f"RVTools ETL complete")
    log.info(f"  Succeeded : {len(succeeded)}/{len(files)} files")
    log.info(f"  Failed    : {len(failed)}/{len(files)} files")
    log.info(f"  Total VMs : {total_vms}")

    if failed:
        log.warning("Failed files:")
        for f in failed:
            log.warning(f"  ✗ {f}")

    status = "success" if not failed else \
             ("partial" if succeeded else "failed")

    await log_pipeline_end(
        pool, run_id, status,
        records_out=total_vms,
        error="\n".join(failed) if failed else None,
    )

    await pool.close()

    if status == "failed":
        sys.exit(1)

    log.info("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
