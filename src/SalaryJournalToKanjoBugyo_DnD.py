# -*- coding: utf-8 -*-
"""
SalaryJournalToKanjoBugyo_DnD.py

ドラッグ＆ドロップで対象ファイルを受け取るためのGUI。
ボタンは1つのみ配置し、3入力ファイルから step0023 と勘定奉行CSV出力までを自動実行する。
"""

from __future__ import annotations

import ctypes
from datetime import datetime
import os
import re
import shutil
import subprocess
import sys
import traceback
from ctypes import wintypes
from pathlib import Path
from typing import List, Optional, Tuple

import win32api
import win32con
import win32gui

if not hasattr(win32con, "DEFAULT_GUI_FONT"):
    win32con.DEFAULT_GUI_FONT = 17


class DRAWITEMSTRUCT(ctypes.Structure):
    _fields_ = [
        ("CtlType", wintypes.UINT),
        ("CtlID", wintypes.UINT),
        ("itemID", wintypes.UINT),
        ("itemAction", wintypes.UINT),
        ("itemState", wintypes.UINT),
        ("hwndItem", wintypes.HWND),
        ("hDC", wintypes.HDC),
        ("rcItem", wintypes.RECT),
        ("itemData", ctypes.c_void_p),
    ]


BUTTON_ID_BASE: int = 1001
BUTTON_HEIGHT: int = 28
BUTTON_MARGIN: int = 10
ACTION_BUTTON_COLOR = (0xC0, 0xE8, 0xFF)
BUTTON_LABELS: Tuple[str, ...] = ("再実行",)
INSTRUCTION_FONT_HEIGHT: int = -17
INSTRUCTION_FONT_FACE: str = "Meiryo UI"


g_main_window_handle: Optional[int] = None
g_action_button_handles: List[int] = []
g_action_button_brush_handle: Optional[int] = None
g_default_gui_font_handle: Optional[int] = None
g_instruction_font_handle: Optional[int] = None
g_last_dropped_files: List[str] = []


def show_message_box(pszMessage: str, pszTitle: str) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONINFORMATION
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(iOwnerWindowHandle, pszMessage, pszTitle, iMessageBoxType)


def show_error_message_box(pszMessage: str, pszTitle: str) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONERROR
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(iOwnerWindowHandle, pszMessage, pszTitle, iMessageBoxType)


def append_error_log(pszMessage: str) -> None:
    pszOutputPath: str = os.path.join(
        os.path.dirname(__file__),
        "SalaryJournalToKanjoBugyo_DnD_error.txt",
    )
    with open(pszOutputPath, "a", encoding="utf-8", newline="") as objFile:
        objFile.write(pszMessage + "\n")


def resolve_company_or_division_directory(objDroppedPaths: Optional[List[Path]] = None) -> Path:
    objScriptDirectory: Path = Path(__file__).resolve().parent
    objTargetModePath: Path = objScriptDirectory / "company_or_division.txt"
    objScriptDirectory.mkdir(parents=True, exist_ok=True)
    objScriptModePath: Path = objScriptDirectory / "company_or_division.txt"
    if not objTargetModePath.exists():
        if objScriptModePath.is_file():
            shutil.copy2(str(objScriptModePath), str(objTargetModePath))
        else:
            objTargetModePath.write_text("", encoding="utf-8")
    elif not objTargetModePath.is_file():
        objTargetModePath.write_text("", encoding="utf-8")
    return objScriptDirectory


def append_fallback_status_log(
    objDroppedPaths: List[Path],
    objStatuses: dict[str, str],
    pszSelectedModeOrException: str,
) -> None:
    objOutputDirectory: Path = resolve_company_or_division_directory(objDroppedPaths)
    objOutputPath: Path = objOutputDirectory / "log_DnD_input_count_fallback_success_failure.txt"
    objLogLine: str = "\t".join([
        datetime.now().isoformat(timespec="seconds"),
        str(len(objDroppedPaths)),
        objStatuses.get("combined4", "未実行"),
        objStatuses.get("legacy", "未実行"),
        objStatuses.get("make_rawdata", "未実行"),
        pszSelectedModeOrException,
    ])
    with objOutputPath.open("a", encoding="utf-8", newline="") as objFile:
        objFile.write(objLogLine + "\n")


def report_exception(pszContext: str, exc: Exception) -> None:
    pszTraceback = traceback.format_exc()
    append_error_log(
        "\n".join(
            [
                f"[Error] {pszContext}",
                f"Detail: {exc}",
                pszTraceback,
            ]
        )
    )
    show_error_message_box(
        f"Error: {pszContext}. Detail = {exc}\n(See error log for traceback.)",
        "SalaryJournalToKanjoBugyo_DnD",
    )


def ensure_default_gui_font_handle() -> Optional[int]:
    global g_default_gui_font_handle
    if g_default_gui_font_handle is None:
        iFontId: int = getattr(win32con, "DEFAULT_GUI_FONT", 17)
        g_default_gui_font_handle = win32gui.GetStockObject(iFontId)
    return g_default_gui_font_handle


def ensure_instruction_font_handle() -> Optional[int]:
    global g_instruction_font_handle
    if g_instruction_font_handle is None:
        try:
            g_instruction_font_handle = ctypes.windll.gdi32.CreateFontW(
                INSTRUCTION_FONT_HEIGHT,
                0,
                0,
                0,
                win32con.FW_NORMAL,
                0,
                0,
                0,
                win32con.SHIFTJIS_CHARSET,
                win32con.OUT_DEFAULT_PRECIS,
                win32con.CLIP_DEFAULT_PRECIS,
                win32con.CLEARTYPE_QUALITY,
                win32con.DEFAULT_PITCH | win32con.FF_DONTCARE,
                INSTRUCTION_FONT_FACE,
            )
        except Exception:
            g_instruction_font_handle = None
    if not g_instruction_font_handle:
        return ensure_default_gui_font_handle()
    return g_instruction_font_handle


def ensure_action_button_brush() -> Optional[int]:
    global g_action_button_brush_handle
    if g_action_button_brush_handle is None:
        iRed, iGreen, iBlue = ACTION_BUTTON_COLOR
        g_action_button_brush_handle = win32gui.CreateSolidBrush(win32api.RGB(iRed, iGreen, iBlue))
    return g_action_button_brush_handle


def create_action_buttons(iWindowHandle: int) -> None:
    global g_action_button_handles
    g_action_button_handles = []
    ensure_default_gui_font_handle()
    for iIndex, pszLabel in enumerate(BUTTON_LABELS):
        iButtonHandle: int = win32gui.CreateWindowEx(
            win32con.WS_EX_NOPARENTNOTIFY,
            "BUTTON",
            pszLabel,
            win32con.WS_TABSTOP | win32con.WS_VISIBLE | win32con.WS_CHILD | win32con.BS_OWNERDRAW,
            0,
            0,
            120,
            BUTTON_HEIGHT,
            iWindowHandle,
            BUTTON_ID_BASE + iIndex,
            win32api.GetModuleHandle(None),
            None,
        )
        iFontHandle = win32gui.GetStockObject(win32con.DEFAULT_GUI_FONT)
        win32gui.SendMessage(iButtonHandle, win32con.WM_SETFONT, iFontHandle, True)
        g_action_button_handles.append(iButtonHandle)
        win32gui.ShowWindow(iButtonHandle, win32con.SW_SHOWNORMAL)
    update_action_button_layout(iWindowHandle)


def update_action_button_layout(iWindowHandle: int) -> None:
    if not g_action_button_handles:
        return
    objClientRect = win32gui.GetClientRect(iWindowHandle)
    iButtonWidth: int = objClientRect[2] - BUTTON_MARGIN * 2
    if iButtonWidth < 120:
        iButtonWidth = 120
    iButtonX: int = BUTTON_MARGIN
    iButtonSpacing: int = BUTTON_MARGIN
    iTotalButtonsHeight: int = BUTTON_HEIGHT * len(g_action_button_handles) + iButtonSpacing * (len(g_action_button_handles) - 1)
    iButtonY: int = objClientRect[3] - BUTTON_MARGIN - iTotalButtonsHeight
    if iButtonY < BUTTON_MARGIN:
        iButtonY = BUTTON_MARGIN

    for iButtonHandle in g_action_button_handles:
        win32gui.MoveWindow(
            iButtonHandle,
            iButtonX,
            iButtonY,
            iButtonWidth,
            BUTTON_HEIGHT,
            True,
        )
        iButtonY += BUTTON_HEIGHT + iButtonSpacing


def draw_instruction_text(iWindowHandle: int) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)

    iFontHandle = ensure_instruction_font_handle()
    iPreviousFontHandle = None
    if iFontHandle:
        iPreviousFontHandle = win32gui.SelectObject(iDeviceContextHandle, iFontHandle)

    iMargin: int = 5
    objClientRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )

    pszInstructionText: str = (
        "次の4ファイルをドラッグ＆ドロップしてください。\n"
        "1) 管理会計：工数*.xlsx\n"
        "2) 作成用データ：工数*.xlsx\n"
        "3) 作成用データ：支給・控除等一覧表_給与_*.csv\n"
        "4) 前払通勤交通費按分表(4~9月・10~3月)_*.xlsx\n"
        "4入力一致時は旧→新の統合ルートでstep0023・勘定奉行CSVまで自動実行します。\n"
        "（従来の3入力ルートも引き続き利用できます）"
    )
    iDrawTextFormat: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
    win32gui.DrawText(iDeviceContextHandle, pszInstructionText, -1, objClientRect, iDrawTextFormat)

    if iPreviousFontHandle:
        win32gui.SelectObject(iDeviceContextHandle, iPreviousFontHandle)
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


def handle_action_button_click() -> int:
    if not g_last_dropped_files:
        show_error_message_box("Error: 先に必要ファイルをドラッグ＆ドロップしてください。", "SalaryJournalToKanjoBugyo_DnD")
        return 1
    objDroppedPaths: List[Path] = [Path(pszPath).resolve() for pszPath in g_last_dropped_files]
    pszMode, objSelectedPaths = resolve_drop_inputs_with_fallback(objDroppedPaths)
    if pszMode == "combined4":
        run_combined_four_input_flow(objSelectedPaths)
    elif pszMode == "legacy":
        run_legacy_three_stage_flow(objSelectedPaths)
    else:
        run_make_rawdata_three_stage_flow([str(objPath) for objPath in objSelectedPaths])
    return 0


def run_command_by_script_name(
    pszScriptName: str,
    objWorkingDirectory: Path,
    objArguments: List[str],
) -> None:
    objScriptPath: Path = Path(__file__).resolve().parent / pszScriptName
    if not objScriptPath.exists():
        raise FileNotFoundError(f"Cmd script not found: {objScriptPath}")
    objCompleted = subprocess.run(
        [sys.executable, str(objScriptPath), *objArguments],
        cwd=str(objWorkingDirectory),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if objCompleted.returncode != 0:
        raise RuntimeError(
            "Cmd failed: {0}\nstdout:\n{1}\nstderr:\n{2}".format(
                " ".join(objArguments),
                objCompleted.stdout,
                objCompleted.stderr,
            )
        )


def run_command_by_make_rawdata_script(objWorkingDirectory: Path, objArguments: List[str]) -> None:
    run_command_by_script_name("make_rawdata_manhour_salary_Cmd.py", objWorkingDirectory, objArguments)


def run_command_by_parttime_script(objWorkingDirectory: Path, objArguments: List[str]) -> None:
    run_command_by_script_name("parttime_salary_allocation_table_Cmd.py", objWorkingDirectory, objArguments)


def resolve_legacy_drop_inputs(objDroppedPaths: List[Path]) -> tuple[Path, Path, Path]:
    def is_legacy_manhour_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "作成用データ" in pszName and "工数" in pszName

    def is_legacy_management_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "管理会計" in pszName and "工数" in pszName

    def is_legacy_salary_csv(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".csv") and "支給・控除等一覧表" in pszName and "給与" in pszName

    def select_last_candidate(objCandidates: List[Path], pszLabel: str) -> Optional[Path]:
        if not objCandidates:
            return None
        if len(objCandidates) > 1:
            append_error_log(
                "Warning: legacy {0} candidates are multiple; selected last dropped file: {1}".format(
                    pszLabel,
                    objCandidates[-1],
                )
            )
        return objCandidates[-1]

    objManhourCandidates: List[Path] = []
    objManagementCandidates: List[Path] = []
    objSalaryCandidates: List[Path] = []
    for objPath in objDroppedPaths:
        if is_legacy_manhour_xlsx(objPath):
            objManhourCandidates.append(objPath)
            continue
        if is_legacy_management_xlsx(objPath):
            objManagementCandidates.append(objPath)
            continue
        if is_legacy_salary_csv(objPath):
            objSalaryCandidates.append(objPath)
            continue

    objManhourPath: Optional[Path] = select_last_candidate(objManhourCandidates, "manhour xlsx")
    objManagementPath: Optional[Path] = select_last_candidate(objManagementCandidates, "management xlsx")
    objSalaryPath: Optional[Path] = select_last_candidate(objSalaryCandidates, "salary csv")
    if objManhourPath is None or objManagementPath is None or objSalaryPath is None:
        raise ValueError("legacy inputs are not complete.")
    return objManhourPath, objManagementPath, objSalaryPath


def resolve_combined_four_input_drop_inputs(objDroppedPaths: List[Path]) -> tuple[Path, Path, Path, Path]:
    def is_legacy_manhour_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "作成用データ" in pszName and "工数" in pszName

    def is_legacy_management_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "管理会計" in pszName and "工数" in pszName

    def is_legacy_salary_csv(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".csv") and "支給・控除等一覧表" in pszName and "給与" in pszName

    def is_prepaid_commute_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "前払通勤交通費按分表" in pszName

    def select_last_candidate(objCandidates: List[Path], pszLabel: str) -> Optional[Path]:
        if not objCandidates:
            return None
        if len(objCandidates) > 1:
            append_error_log(
                "Warning: combined4 {0} candidates are multiple; selected last dropped file: {1}".format(
                    pszLabel,
                    objCandidates[-1],
                )
            )
        return objCandidates[-1]

    objManhourCandidates: List[Path] = []
    objManagementCandidates: List[Path] = []
    objSalaryCandidates: List[Path] = []
    objPrepaidCandidates: List[Path] = []
    objIgnoredFiles: List[Path] = []
    for objPath in objDroppedPaths:
        if is_legacy_manhour_xlsx(objPath):
            objManhourCandidates.append(objPath)
            continue
        if is_legacy_management_xlsx(objPath):
            objManagementCandidates.append(objPath)
            continue
        if is_legacy_salary_csv(objPath):
            objSalaryCandidates.append(objPath)
            continue
        if is_prepaid_commute_xlsx(objPath):
            objPrepaidCandidates.append(objPath)
            continue
        objIgnoredFiles.append(objPath)

    objManagementPath: Optional[Path] = select_last_candidate(objManagementCandidates, "management xlsx")
    objManhourPath: Optional[Path] = select_last_candidate(objManhourCandidates, "manhour xlsx")
    objSalaryPath: Optional[Path] = select_last_candidate(objSalaryCandidates, "salary csv")
    objPrepaidPath: Optional[Path] = select_last_candidate(objPrepaidCandidates, "prepaid commute xlsx")

    if objIgnoredFiles:
        append_error_log(
            "Ignored dropped files (combined4):\n{0}".format(
                "\n".join([str(objPath) for objPath in objIgnoredFiles])
            )
        )
    append_error_log(
        "Accepted dropped files (combined4):\n{0}".format(
            "\n".join([str(objPath) for objPath in objDroppedPaths])
        )
    )

    objMissingLabels: List[str] = []
    if objManagementPath is None:
        objMissingLabels.append("管理会計：工数*.xlsx")
    if objManhourPath is None:
        objMissingLabels.append("作成用データ：工数*.xlsx")
    if objSalaryPath is None:
        objMissingLabels.append("作成用データ：支給・控除等一覧表_給与_*.csv")
    if objPrepaidPath is None:
        objMissingLabels.append("前払通勤交通費按分表*.xlsx")
    if objMissingLabels:
        raise ValueError("必要ファイルが不足しています:\n\n{0}".format("\n".join(objMissingLabels)))

    return objManagementPath, objManhourPath, objSalaryPath, objPrepaidPath


def resolve_make_rawdata_drop_inputs(objDroppedPaths: List[Path]) -> tuple[Path, Path, Path]:
    def is_step0005_tsv(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return (
            pszLower.endswith(".tsv")
            and "新_ローデータ_シート" in pszName
            and "step0005" in pszLower
        )

    def is_salary_step0001_tsv(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return (
            pszLower.endswith(".tsv")
            and "支給・控除等一覧表" in pszName
            and "給与" in pszName
            and "step0001" in pszLower
        )

    def is_prepaid_commute_xlsx(objPath: Path) -> bool:
        pszName: str = objPath.name
        pszLower: str = pszName.lower()
        return pszLower.endswith(".xlsx") and "前払通勤交通費按分表" in pszName

    def select_last_candidate(objCandidates: List[Path], pszLabel: str) -> Optional[Path]:
        if not objCandidates:
            return None
        if len(objCandidates) > 1:
            append_error_log(
                "Warning: {0} candidates are multiple; selected last dropped file: {1}".format(
                    pszLabel,
                    objCandidates[-1],
                )
            )
        return objCandidates[-1]

    objStep0005Candidates: List[Path] = []
    objSalaryStep0001Candidates: List[Path] = []
    objPrepaidXlsxCandidates: List[Path] = []
    objIgnoredFiles: List[Path] = []
    for objPath in objDroppedPaths:
        if is_step0005_tsv(objPath):
            objStep0005Candidates.append(objPath)
            continue
        if is_salary_step0001_tsv(objPath):
            objSalaryStep0001Candidates.append(objPath)
            continue
        if is_prepaid_commute_xlsx(objPath):
            objPrepaidXlsxCandidates.append(objPath)
            continue
        objIgnoredFiles.append(objPath)

    objStep0005Path: Optional[Path] = select_last_candidate(objStep0005Candidates, "step0005 TSV")
    objSalaryStep0001Path: Optional[Path] = select_last_candidate(objSalaryStep0001Candidates, "給与 step0001 TSV")
    objPrepaidXlsxPath: Optional[Path] = select_last_candidate(objPrepaidXlsxCandidates, "前払通勤交通費按分表 XLSX")

    objMissingLabels: List[str] = []
    if objStep0005Path is None:
        objMissingLabels.append("step0005 TSV")
    if objSalaryStep0001Path is None:
        objMissingLabels.append("給与 step0001 TSV")
    if objPrepaidXlsxPath is None:
        objMissingLabels.append("前払通勤交通費按分表 XLSX")

    if objIgnoredFiles:
        append_error_log(
            "Ignored dropped files:\n{0}".format(
                "\n".join([str(objPath) for objPath in objIgnoredFiles])
            )
        )
    append_error_log(
        "Accepted dropped files:\n{0}".format(
            "\n".join([str(objPath) for objPath in objDroppedPaths])
        )
    )

    if objMissingLabels:
        raise ValueError(
            "必要ファイルが不足しています:\n\n{0}".format(
                "\n".join(objMissingLabels)
            )
        )
    return objStep0005Path, objSalaryStep0001Path, objPrepaidXlsxPath


def list_tsv_files(objDirectory: Path) -> set[Path]:
    return {objPath.resolve() for objPath in objDirectory.glob("*.tsv")}


def list_csv_files(objDirectory: Path) -> set[Path]:
    return {objPath.resolve() for objPath in objDirectory.glob("*.csv")}


def find_new_file_by_pattern(objBefore: set[Path], objDirectory: Path, objPattern: str) -> Path:
    objRegex = re.compile(objPattern)
    objAfter: set[Path] = list_tsv_files(objDirectory)
    objCreated: List[Path] = [objPath for objPath in objAfter if objPath not in objBefore]
    objCandidates: List[Path] = [objPath for objPath in objCreated if objRegex.match(objPath.name) is not None]
    if not objCandidates:
        objCandidates = [objPath for objPath in objAfter if objRegex.match(objPath.name) is not None]
    if not objCandidates:
        raise FileNotFoundError(f"Could not find output by pattern: {objPattern}")
    objCandidates.sort(key=lambda objPath: objPath.stat().st_mtime, reverse=True)
    return objCandidates[0]


def find_make_rawdata_generated_sheet_tsv(
    objWorkingDirectory: Path,
    objPrepaidXlsxPath: Path,
    objBeforeStep2: set[Path],
) -> Path:
    return find_new_file_by_pattern(
        objBeforeStep2,
        objWorkingDirectory,
        rf"^{re.escape(objPrepaidXlsxPath.stem)}_.*\.tsv$",
    )


def verify_make_rawdata_required_outputs(
    objWorkingDirectory: Path,
    objBeforeStep3: set[Path],
    objBeforeCsv: set[Path],
) -> tuple[Path, Path, Path]:
    objStep0023Path: Path = find_new_file_by_pattern(
        objBeforeStep3,
        objWorkingDirectory,
        r"^新_ローデータ_シート_step0023_\d{4}年(?:04-09月|10-03月)_\d{2}月_前払通勤交通費按分表\.tsv$",
    )
    objUtf8CsvCandidates: List[Path] = [
        objPath for objPath in objWorkingDirectory.glob("通勤費プロジェクト振替_*_勘定奉行用.csv")
        if objPath.resolve() not in objBeforeCsv
    ]
    objSjisCsvCandidates: List[Path] = [
        objPath for objPath in objWorkingDirectory.glob("通勤費プロジェクト振替_*_勘定奉行用_sjis.csv")
        if objPath.resolve() not in objBeforeCsv
    ]
    if not objUtf8CsvCandidates:
        objUtf8CsvCandidates = list(objWorkingDirectory.glob("通勤費プロジェクト振替_*_勘定奉行用.csv"))
    if not objSjisCsvCandidates:
        objSjisCsvCandidates = list(objWorkingDirectory.glob("通勤費プロジェクト振替_*_勘定奉行用_sjis.csv"))
    if not objUtf8CsvCandidates or not objSjisCsvCandidates:
        raise FileNotFoundError("Could not find output accounting CSV pair (utf-8 / sjis).")
    objUtf8CsvCandidates.sort(key=lambda objPath: objPath.stat().st_mtime, reverse=True)
    objSjisCsvCandidates.sort(key=lambda objPath: objPath.stat().st_mtime, reverse=True)
    return objStep0023Path, objUtf8CsvCandidates[0], objSjisCsvCandidates[0]


def resolve_generated_step0005_and_salary_step0001(
    objWorkingDirectory: Path,
    objBeforeTsv: set[Path],
) -> tuple[Path, Path]:
    objStep0005Path: Path = find_new_file_by_pattern(
        objBeforeTsv,
        objWorkingDirectory,
        r"^新_ローデータ_シート_step0005_\d{4}年\d{2}月\.tsv$",
    )
    objSalaryStep0001Path: Path = find_new_file_by_pattern(
        objBeforeTsv,
        objWorkingDirectory,
        r"^支給・控除等一覧表_給与_step0001_.+\.tsv$",
    )
    return objStep0005Path, objSalaryStep0001Path


def run_make_rawdata_from_generated_inputs(
    objStep0005Path: Path,
    objSalaryStep0001Path: Path,
    objPrepaidXlsxPath: Path,
) -> tuple[Path, Path, Path]:
    objWorkingDirectory: Path = objStep0005Path.parent
    for objPath in (objSalaryStep0001Path, objPrepaidXlsxPath):
        if objPath.parent != objWorkingDirectory:
            append_error_log(
                "Selected files are in different directories:\n{0}\n{1}\n{2}".format(
                    objStep0005Path,
                    objSalaryStep0001Path,
                    objPrepaidXlsxPath,
                )
            )
            raise ValueError("採用されたファイルは同じフォルダ内にある必要があります。")

    run_command_by_make_rawdata_script(
        objWorkingDirectory,
        [objStep0005Path.name, objSalaryStep0001Path.name],
    )

    objBeforeStep2: set[Path] = list_tsv_files(objWorkingDirectory)
    run_command_by_make_rawdata_script(objWorkingDirectory, [objPrepaidXlsxPath.name])
    objPrepaidSheetTsvPath: Path = find_make_rawdata_generated_sheet_tsv(
        objWorkingDirectory,
        objPrepaidXlsxPath,
        objBeforeStep2,
    )

    objBeforeStep3: set[Path] = list_tsv_files(objWorkingDirectory)
    objBeforeCsv: set[Path] = list_csv_files(objWorkingDirectory)
    run_command_by_make_rawdata_script(
        objWorkingDirectory,
        [objPrepaidSheetTsvPath.name],
    )
    return verify_make_rawdata_required_outputs(
        objWorkingDirectory,
        objBeforeStep3,
        objBeforeCsv,
    )


def run_combined_four_input_flow(
    objSelectedPaths: tuple[Path, Path, Path, Path],
) -> None:
    objManagementXlsxPath, objManhourXlsxPath, objSalaryCsvPath, objPrepaidXlsxPath = objSelectedPaths
    objWorkingDirectory: Path = objManagementXlsxPath.parent
    for objPath in (objManhourXlsxPath, objSalaryCsvPath, objPrepaidXlsxPath):
        if objPath.parent != objWorkingDirectory:
            raise ValueError("採用された4ファイルは同じフォルダ内にある必要があります。")

    objBeforeLegacyTsv: set[Path] = list_tsv_files(objWorkingDirectory)
    run_legacy_three_stage_flow((objManhourXlsxPath, objManagementXlsxPath, objSalaryCsvPath))
    objStep0005Path: Path
    objSalaryStep0001Path: Path
    objStep0005Path, objSalaryStep0001Path = resolve_generated_step0005_and_salary_step0001(
        objWorkingDirectory,
        objBeforeLegacyTsv,
    )

    objStep0023Path: Path
    objUtf8CsvPath: Path
    objSjisCsvPath: Path
    objStep0023Path, objUtf8CsvPath, objSjisCsvPath = run_make_rawdata_from_generated_inputs(
        objStep0005Path,
        objSalaryStep0001Path,
        objPrepaidXlsxPath,
    )

    show_message_box(
        "4入力統合ルートの処理が完了しました。\n\n"
        + f"step0005: {objStep0005Path.name}\n"
        + f"給与step0001: {objSalaryStep0001Path.name}\n"
        + f"step0023: {objStep0023Path.name}\n"
        + f"CSV(UTF-8): {objUtf8CsvPath.name}\n"
        + f"CSV(SJIS): {objSjisCsvPath.name}",
        "SalaryJournalToKanjoBugyo_DnD",
    )


def build_make_rawdata_error_report(
    pszContext: str,
    exc: Exception,
    objFilePaths: Optional[List[str]] = None,
) -> str:
    objLines: List[str] = [f"Error: {pszContext}", f"Detail: {exc}"]
    if objFilePaths:
        objLines.append("Input files:")
        objLines.extend(objFilePaths)
    return "\n".join(objLines)


def resolve_drop_inputs_with_fallback(objDroppedPaths: List[Path]) -> tuple[str, tuple[Path, ...]]:
    objStatuses: dict[str, str] = {
        "combined4": "未実行",
        "legacy": "未実行",
        "make_rawdata": "未実行",
    }
    try:
        objCombined = resolve_combined_four_input_drop_inputs(objDroppedPaths)
        objStatuses["combined4"] = "成功"
        append_error_log("Detected combined4 dropped-file set. Running combined4 flow.")
        append_fallback_status_log(objDroppedPaths, objStatuses, "combined4")
        return "combined4", objCombined
    except Exception:
        objStatuses["combined4"] = "失敗"
    try:
        objLegacy = resolve_legacy_drop_inputs(objDroppedPaths)
        objStatuses["legacy"] = "成功"
        append_error_log("Detected legacy dropped-file set. Running legacy flow.")
        append_fallback_status_log(objDroppedPaths, objStatuses, "legacy")
        return "legacy", objLegacy
    except Exception:
        objStatuses["legacy"] = "失敗"
    try:
        objMakeRawdata = resolve_make_rawdata_drop_inputs(objDroppedPaths)
        objStatuses["make_rawdata"] = "成功"
        append_error_log("Detected make_rawdata dropped-file set. Running make_rawdata flow.")
        append_fallback_status_log(objDroppedPaths, objStatuses, "make_rawdata")
        return "make_rawdata", objMakeRawdata
    except Exception as objException:
        objStatuses["make_rawdata"] = "失敗"
        append_fallback_status_log(
            objDroppedPaths,
            objStatuses,
            f"{objException.__class__.__name__}: {objException}",
        )
        raise


def run_legacy_three_stage_flow(objSelectedPaths: tuple[Path, Path, Path]) -> None:
    objManhourXlsxPath, objManagementXlsxPath, objSalaryCsvPath = objSelectedPaths
    objWorkingDirectory: Path = objManhourXlsxPath.parent
    for objPath in (objManagementXlsxPath, objSalaryCsvPath):
        if objPath.parent != objWorkingDirectory:
            raise ValueError("採用された3ファイルは同じフォルダ内にある必要があります。")

    objBeforeStep1: set[Path] = list_tsv_files(objWorkingDirectory)
    run_command_by_parttime_script(objWorkingDirectory, [objManagementXlsxPath.name])
    objManagementTsvPath: Path = find_new_file_by_pattern(
        objBeforeStep1,
        objWorkingDirectory,
        r"^管理会計：工数.*_.*\.tsv$",
    )

    objBeforeStep2: set[Path] = list_tsv_files(objWorkingDirectory)
    run_command_by_parttime_script(objWorkingDirectory, [objManhourXlsxPath.name])
    objManhourSheetTsvPath: Path = find_new_file_by_pattern(
        objBeforeStep2,
        objWorkingDirectory,
        rf"^{re.escape(objManhourXlsxPath.stem)}_.*\.tsv$",
    )

    run_command_by_parttime_script(objWorkingDirectory, [objManhourSheetTsvPath.name])
    objManhourHmmssTsvPath: Path = objWorkingDirectory / f"{objManhourSheetTsvPath.stem}_h_mm_ss.tsv"
    if not objManhourHmmssTsvPath.exists():
        raise FileNotFoundError(f"h_mm_ss TSV not found: {objManhourHmmssTsvPath}")

    objBeforeStep4: set[Path] = list_tsv_files(objWorkingDirectory)
    run_command_by_parttime_script(objWorkingDirectory, [objManhourHmmssTsvPath.name])
    objNewRawdataStep0001Path: Path = find_new_file_by_pattern(
        objBeforeStep4,
        objWorkingDirectory,
        r"^新_ローデータ_シート_step0001_\d{4}年\d{2}月\.tsv$",
    )

    objBeforeStep5: set[Path] = list_tsv_files(objWorkingDirectory)
    run_command_by_parttime_script(objWorkingDirectory, [objSalaryCsvPath.name])
    objSalaryStep0001Path: Path = find_new_file_by_pattern(
        objBeforeStep5,
        objWorkingDirectory,
        r"^支給・控除等一覧表_給与_step0001_.+\.tsv$",
    )

    run_command_by_parttime_script(
        objWorkingDirectory,
        [objNewRawdataStep0001Path.name, objSalaryStep0001Path.name],
    )
    objNewRawdataStep0002Path: Path = objWorkingDirectory / objNewRawdataStep0001Path.name.replace(
        "_step0001_",
        "_step0002_",
        1,
    )
    if not objNewRawdataStep0002Path.exists():
        raise FileNotFoundError(f"step0002 TSV not found: {objNewRawdataStep0002Path}")

    run_command_by_parttime_script(
        objWorkingDirectory,
        [
            objNewRawdataStep0002Path.name,
            objSalaryStep0001Path.name,
            objManagementTsvPath.name,
        ],
    )
    objStep0005Path: Path = objWorkingDirectory / objNewRawdataStep0002Path.name.replace("_step0002_", "_step0005_", 1)
    if not objStep0005Path.exists():
        raise FileNotFoundError(f"step0005 TSV not found: {objStep0005Path}")

    show_message_box(
        "給与、法定福利、定期代按分用の新_ローデータ_シートを作成しました。\n\n"
        + f"step0005: {objStep0005Path.name}",
        "SalaryJournalToKanjoBugyo_DnD",
    )


def run_make_rawdata_three_stage_flow(objFilePaths: List[str]) -> None:
    objDroppedPaths: List[Path] = [Path(pszPath).resolve() for pszPath in objFilePaths]
    objStep0005Path: Path
    objSalaryStep0001Path: Path
    objPrepaidXlsxPath: Path
    objStep0005Path, objSalaryStep0001Path, objPrepaidXlsxPath = resolve_make_rawdata_drop_inputs(objDroppedPaths)

    objStep0023Path: Path
    objUtf8CsvPath: Path
    objSjisCsvPath: Path
    objStep0023Path, objUtf8CsvPath, objSjisCsvPath = run_make_rawdata_from_generated_inputs(
        objStep0005Path,
        objSalaryStep0001Path,
        objPrepaidXlsxPath,
    )

    show_message_box(
        "前払通勤交通費按分〜step0023〜勘定奉行CSV出力までの処理が完了しました。\n\n"
        + f"step0023: {objStep0023Path.name}\n"
        + f"CSV(UTF-8): {objUtf8CsvPath.name}\n"
        + f"CSV(SJIS): {objSjisCsvPath.name}",
        "SalaryJournalToKanjoBugyo_DnD",
    )


def handle_drop_files(iDropHandle: int) -> int:
    global g_last_dropped_files
    iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)
    objFiles: List[str] = []
    for iIndex in range(iFileCount):
        objFiles.append(win32api.DragQueryFile(iDropHandle, iIndex))
    win32api.DragFinish(iDropHandle)

    if not objFiles:
        show_error_message_box("Error: ファイルが取得できませんでした。", "SalaryJournalToKanjoBugyo_DnD")
        return 1

    g_last_dropped_files = list(objFiles)
    objDroppedPaths: List[Path] = [Path(pszPath).resolve() for pszPath in objFiles]
    pszMode, objSelectedPaths = resolve_drop_inputs_with_fallback(objDroppedPaths)
    if pszMode == "combined4":
        run_combined_four_input_flow(objSelectedPaths)
    elif pszMode == "legacy":
        run_legacy_three_stage_flow(objSelectedPaths)
    else:
        run_make_rawdata_three_stage_flow([str(objPath) for objPath in objSelectedPaths])
    return 0


def window_proc(iWindowHandle: int, iMessage: int, iWparam: int, iLparam: int) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        if not g_action_button_handles:
            create_action_buttons(iWindowHandle)
        return 0

    if iMessage == win32con.WM_DROPFILES:
        try:
            handle_drop_files(iWparam)
        except Exception as exc:
            show_error_message_box(
                build_make_rawdata_error_report("failed to handle dropped files", exc, g_last_dropped_files),
                "SalaryJournalToKanjoBugyo_DnD",
            )
            report_exception("failed to handle dropped files", exc)
        return 0

    if iMessage == win32con.WM_COMMAND:
        iButtonId: int = win32api.LOWORD(iWparam)
        if iButtonId == BUTTON_ID_BASE:
            try:
                handle_action_button_click()
            except Exception as exc:
                report_exception("failed to run action button handler", exc)
            return 0

    if iMessage == win32con.WM_DRAWITEM:
        objDrawItem = DRAWITEMSTRUCT.from_address(iLparam)
        if objDrawItem.CtlID == BUTTON_ID_BASE:
            iDeviceContextHandle: int = objDrawItem.hDC
            objRect = objDrawItem.rcItem
            objRectTuple: tuple[int, int, int, int] = (
                objRect.left,
                objRect.top,
                objRect.right,
                objRect.bottom,
            )
            iBrushHandle = ensure_action_button_brush()
            if iBrushHandle:
                win32gui.FillRect(iDeviceContextHandle, objRectTuple, iBrushHandle)

            iFontHandle = win32gui.GetStockObject(win32con.DEFAULT_GUI_FONT)
            iPreviousFont = win32gui.SelectObject(iDeviceContextHandle, iFontHandle)
            pszText: str = win32gui.GetWindowText(objDrawItem.hwndItem)
            win32gui.DrawText(
                iDeviceContextHandle,
                pszText,
                -1,
                objRectTuple,
                win32con.DT_CENTER | win32con.DT_VCENTER | win32con.DT_SINGLELINE,
            )
            if iPreviousFont:
                win32gui.SelectObject(iDeviceContextHandle, iPreviousFont)
            if objDrawItem.itemState & win32con.ODS_FOCUS:
                win32gui.DrawFocusRect(iDeviceContextHandle, objRectTuple)
            return 1

    if iMessage == win32con.WM_SIZE:
        update_action_button_layout(iWindowHandle)
        return 0

    if iMessage == win32con.WM_CTLCOLORBTN:
        iBrushHandle = ensure_action_button_brush()
        if iBrushHandle:
            return iBrushHandle

    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0

    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def register_window_class(pszWindowClassName: str) -> int:
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    objWndClass = win32gui.WNDCLASS()
    objWndClass.hInstance = iInstanceHandle
    objWndClass.lpszClassName = pszWindowClassName
    objWndClass.lpfnWndProc = window_proc
    objWndClass.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
    objWndClass.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    objWndClass.hbrBackground = win32con.COLOR_WINDOW + 1

    iClassAtom: int = win32gui.RegisterClass(objWndClass)
    return iClassAtom


def create_main_window(pszWindowClassName: str, pszWindowTitle: str) -> int:
    global g_main_window_handle
    iInstanceHandle: int = win32api.GetModuleHandle(None)

    iWindowStyle: int = (
        win32con.WS_OVERLAPPED
        | win32con.WS_CAPTION
        | win32con.WS_SYSMENU
        | win32con.WS_MINIMIZEBOX
    )
    iWindowExStyle: int = win32con.WS_EX_ACCEPTFILES

    iWindowPosX: int = win32con.CW_USEDEFAULT
    iWindowPosY: int = win32con.CW_USEDEFAULT
    iDesktopWidth: int = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    iDesktopHeight: int = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    iWindowWidth: int = int(iDesktopWidth * 0.5)
    iWindowHeight: int = int(iDesktopHeight * 0.5)

    iWindowHandle: int = win32gui.CreateWindowEx(
        iWindowExStyle,
        pszWindowClassName,
        pszWindowTitle,
        iWindowStyle,
        iWindowPosX,
        iWindowPosY,
        iWindowWidth,
        iWindowHeight,
        0,
        0,
        iInstanceHandle,
        None,
    )
    g_main_window_handle = iWindowHandle

    if not g_action_button_handles:
        create_action_buttons(iWindowHandle)

    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)
    update_action_button_layout(iWindowHandle)
    win32gui.InvalidateRect(iWindowHandle, None, True)

    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
    )

    win32gui.DragAcceptFiles(iWindowHandle, True)
    return iWindowHandle


def main() -> None:
    pszWindowClassName: str = "SalaryJournalToKanjoBugyoDndWindowClass"
    pszWindowTitle: str = "SalaryJournalToKanjoBugyo (Drag & Drop)"

    try:
        register_window_class(pszWindowClassName)
    except Exception as exc:
        report_exception("failed to register window class", exc)
        return

    try:
        create_main_window(pszWindowClassName, pszWindowTitle)
    except Exception as exc:
        report_exception("failed to create main window", exc)
        return

    try:
        win32gui.PumpMessages()
    except Exception as exc:
        report_exception("unexpected exception in message loop", exc)


if __name__ == "__main__":
    main()
