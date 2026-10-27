"""
Conversion engines for XLSB -> XLSX.

This module deliberately has NO Streamlit imports so it can be tested and
reused on its own.

Engines (tried in this order when "auto" is selected):

1. excel        Microsoft Excel via COM (Windows + desktop Excel only).
                Highest fidelity. Used when running locally on Windows.
2. libreoffice  Headless LibreOffice (`soffice`). Good fidelity (formatting,
                formulas, number formats). This is the engine used on
                Streamlit Community Cloud (installed through packages.txt).
3. pyxlsb       Pure Python fallback (pyxlsb + xlsxwriter). Copies cell
                VALUES only: no formulas, styles, charts or date formats.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

ENGINE_EXCEL = "excel"
ENGINE_LIBREOFFICE = "libreoffice"
ENGINE_PYXLSB = "pyxlsb"
ENGINE_ORDER = [ENGINE_EXCEL, ENGINE_LIBREOFFICE, ENGINE_PYXLSB]

ENGINE_LABELS = {
    ENGINE_EXCEL: "Microsoft Excel",
    ENGINE_LIBREOFFICE: "LibreOffice",
    ENGINE_PYXLSB: "Python (pyxlsb)",
}

XLSX_FILE_FORMAT = 51  # Excel's xlOpenXMLWorkbook

_WINDOWS_SOFFICE_PATHS = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)

# LibreOffice cannot run two conversions on the same profile at once, so all
# LibreOffice calls in this process are serialised.
_LO_LOCK = threading.Lock()
_LO_PROFILE_DIR = Path(tempfile.gettempdir()) / "xlsb2xlsx_lo_profile"


@dataclass
class ConversionResult:
    ok: bool
    input_name: str
    output_path: str = ""
    size_bytes: int = 0
    seconds: float = 0.0
    engine: str = ""
    error: str = ""


# --------------------------------------------------------------------------
# Environment helpers
# --------------------------------------------------------------------------

def is_hosted_cloud() -> bool:
    """True when running on Streamlit Community Cloud (or a similar host)."""
    return (
        Path("/mount/src").exists()
        or os.environ.get("HOME", "") == "/home/adminuser"
        or bool(os.environ.get("STREAMLIT_SHARING_MODE"))
    )


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    if os.name == "nt":
        for candidate in _WINDOWS_SOFFICE_PATHS:
            if Path(candidate).is_file():
                return candidate
    return None


def _excel_available() -> tuple[bool, str]:
    if os.name != "nt":
        return False, "Windows only"
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "pywin32 not installed"
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application"))
    except OSError:
        return False, "Excel is not installed"
    return True, "Ready"


def _libreoffice_available() -> tuple[bool, str]:
    path = find_soffice()
    if path:
        return True, "Ready"
    return False, "LibreOffice not found"


def _pyxlsb_available() -> tuple[bool, str]:
    try:
        import pyxlsb  # noqa: F401
        import xlsxwriter  # noqa: F401
    except ImportError as exc:
        return False, f"Missing package: {exc.name}"
    return True, "Ready"


def engine_status() -> dict[str, tuple[bool, str]]:
    """Return {engine_key: (available, detail)} for every engine."""
    return {
        ENGINE_EXCEL: _excel_available(),
        ENGINE_LIBREOFFICE: _libreoffice_available(),
        ENGINE_PYXLSB: _pyxlsb_available(),
    }


# --------------------------------------------------------------------------
# Engine implementations (each raises on failure)
# --------------------------------------------------------------------------

def _convert_excel(src: Path, dst: Path, timeout: int) -> None:
    import pythoncom
    import win32com.client

    excel = None
    workbook = None
    pythoncom.CoInitialize()
    try:
        # Separate instance: never attach to the user's interactive Excel.
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        workbook = excel.Workbooks.Open(
            Filename=str(src.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
        )
        workbook.SaveAs(
            Filename=str(dst.resolve()),
            FileFormat=XLSX_FILE_FORMAT,
            AddToMru=False,
            Local=True,
        )
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.DisplayAlerts = False
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _convert_libreoffice(src: Path, dst: Path, timeout: int) -> None:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) was not found on this machine.")

    with _LO_LOCK, tempfile.TemporaryDirectory(prefix="x2x_lo_") as work:
        outdir = Path(work) / "out"
        outdir.mkdir()
        _LO_PROFILE_DIR.mkdir(parents=True, exist_ok=True)

        # Convert a copy with a plain ASCII name so unusual characters in the
        # original filename can never confuse the soffice command line.
        staged = Path(work) / "input.xlsb"
        shutil.copyfile(src, staged)

        cmd = [
            soffice,
            f"-env:UserInstallation={_LO_PROFILE_DIR.as_uri()}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nofirststartwizard",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            str(outdir),
            str(staged),
        ]
        env = {**os.environ, "HOME": os.environ.get("HOME") or work}

        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout,
                env=env,
            )
        except subprocess.TimeoutExpired:
            raise TimeoutError(
                f"LibreOffice did not finish within {timeout} seconds."
            ) from None

        produced = outdir / "input.xlsx"
        if not produced.is_file() or produced.stat().st_size == 0:
            lines = [
                ln.strip()
                for ln in (proc.stderr or proc.stdout or "").splitlines()
                if ln.strip()
            ]
            raise RuntimeError(
                "LibreOffice could not convert this file (it may be corrupt, "
                "password-protected or not a real .xlsb)."
                + (f" Details: {lines[0][:200]}" if lines else "")
            )
        shutil.move(str(produced), str(dst))


def _convert_pyxlsb(src: Path, dst: Path, timeout: int) -> None:
    import xlsxwriter
    from pyxlsb import open_workbook

    with open_workbook(str(src)) as wb_in:
        wb_out = xlsxwriter.Workbook(
            str(dst),
            {"constant_memory": True, "nan_inf_to_errors": True},
        )
        try:
            for sheet_name in wb_in.sheets:
                ws_out = wb_out.add_worksheet(sheet_name)
                with wb_in.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows(sparse=True):
                        for cell in row:
                            value = cell.v
                            if value is None:
                                continue
                            # Explicit writers: a text value that begins with
                            # "=" must stay text, never become a formula.
                            if isinstance(value, bool):
                                ws_out.write_boolean(cell.r, cell.c, value)
                            elif isinstance(value, (int, float)):
                                ws_out.write_number(cell.r, cell.c, value)
                            else:
                                ws_out.write_string(cell.r, cell.c, str(value))
        finally:
            wb_out.close()


_ENGINE_FUNCS = {
    ENGINE_EXCEL: _convert_excel,
    ENGINE_LIBREOFFICE: _convert_libreoffice,
    ENGINE_PYXLSB: _convert_pyxlsb,
}


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------

def resolve_engine_chain(choice: str) -> list[str]:
    """Turn a UI choice ('auto' or an engine key) into engines to try."""
    status = engine_status()
    if choice == "auto":
        return [e for e in ENGINE_ORDER if status[e][0]]
    return [choice] if status.get(choice, (False, ""))[0] else []


def convert_file(
    src: Path,
    dst: Path,
    engine_choice: str = "auto",
    timeout: int = 300,
) -> ConversionResult:
    """
    Convert one .xlsb file to .xlsx.

    With engine_choice='auto' the engines are tried in priority order and the
    next one is used if the previous fails. The source file is never modified.
    """
    src, dst = Path(src), Path(dst)
    started = time.monotonic()

    def fail(message: str, engine: str = "") -> ConversionResult:
        return ConversionResult(
            ok=False,
            input_name=src.name,
            engine=engine,
            seconds=round(time.monotonic() - started, 2),
            error=message,
        )

    if not src.is_file():
        return fail(f"Input file not found: {src}")
    if src.suffix.lower() != ".xlsb":
        return fail("Input file must have an .xlsb extension.")
    if src.resolve() == dst.resolve():
        return fail("Input and output paths must be different.")

    chain = resolve_engine_chain(engine_choice)
    if not chain:
        return fail(
            "No conversion engine is available "
            f"(requested: {engine_choice}). See the sidebar for details."
        )

    dst.parent.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []

    for engine in chain:
        try:
            if dst.exists():
                dst.unlink()
            _ENGINE_FUNCS[engine](src, dst, timeout)

            if not dst.is_file() or dst.stat().st_size == 0:
                raise RuntimeError("The engine produced an empty output file.")

            return ConversionResult(
                ok=True,
                input_name=src.name,
                output_path=str(dst),
                size_bytes=dst.stat().st_size,
                seconds=round(time.monotonic() - started, 2),
                engine=engine,
            )
        except Exception as exc:  # noqa: BLE001 - report every failure kind
            errors.append(f"{ENGINE_LABELS[engine]}: {type(exc).__name__}: {exc}")
            try:
                if dst.exists():
                    dst.unlink()
            except OSError:
                pass

    return fail(" | ".join(errors), engine=ENGINE_LABELS[chain[-1]])


# --------------------------------------------------------------------------
# Path utilities
# --------------------------------------------------------------------------

def safe_relative_path(upload_name: str) -> Path:
    """Sanitise a browser-supplied name; reject absolute paths and '..'."""
    normalized = upload_name.replace("\\", "/")
    relative = PurePosixPath(normalized)
    if relative.is_absolute() or ".." in relative.parts:
        raise ValueError(f"Unsafe upload path: {upload_name}")
    if not relative.parts:
        raise ValueError("The uploaded file has no filename.")
    return Path(*relative.parts)


def list_xlsb_files(directory: Path) -> tuple[list[Path], int]:
    """
    Return (xlsb files directly inside `directory`, count of other files).

    Only .xlsb is converted. Everything else is ignored. Excel's temporary
    lock files (names starting with '~$') are ignored as well.
    """
    xlsb: list[Path] = []
    ignored = 0
    for item in sorted(directory.iterdir(), key=lambda p: p.name.casefold()):
        if not item.is_file():
            continue
        if item.suffix.lower() == ".xlsb" and not item.name.startswith("~$"):
            xlsb.append(item)
        else:
            ignored += 1
    return xlsb, ignored


def unique_output_path(folder: Path, relative: Path, used: set[str]) -> Path:
    """Return folder/relative (suffix .xlsx), adding ' (2)', ' (3)' on clashes."""
    candidate = relative.with_suffix(".xlsx")
    counter = 2
    while candidate.as_posix().casefold() in used:
        candidate = relative.with_name(f"{relative.stem} ({counter}).xlsx")
        counter += 1
    used.add(candidate.as_posix().casefold())
    return folder / candidate
