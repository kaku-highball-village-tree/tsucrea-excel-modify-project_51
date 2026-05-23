# -*- coding: utf-8 -*-
"""
DivisionSellGeneralAdminCost_Allocation_DnD.py

division別販管費配賦用の入力ファイルをドラッグ＆ドロップで受け取り、
DivisionSellGeneralAdminCost_Allocation_Cmd.py を実行するGUI。
"""

from __future__ import annotations

import ctypes
from datetime import datetime
import os
import subprocess
import sys
import traceback
from ctypes import wintypes
from typing import List, Optional

import win32api
import win32con
import win32gui

if not hasattr(win32con, "DEFAULT_GUI_FONT"):
    win32con.DEFAULT_GUI_FONT = 17

INSTRUCTION_FONT_HEIGHT: int = -17
INSTRUCTION_FONT_FACE: str = "Meiryo UI"

g_main_window_handle: Optional[int] = None
g_default_gui_font_handle: Optional[int] = None
g_instruction_font_handle: Optional[int] = None


def show_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONINFORMATION
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def show_error_message_box(
    pszMessage: str,
    pszTitle: str,
) -> None:
    iOwnerWindowHandle: int = g_main_window_handle or win32gui.GetForegroundWindow()
    iMessageBoxType: int = (
        win32con.MB_OK
        | win32con.MB_ICONERROR
        | win32con.MB_TASKMODAL
        | win32con.MB_SETFOREGROUND
    )
    win32gui.MessageBox(
        iOwnerWindowHandle,
        pszMessage,
        pszTitle,
        iMessageBoxType,
    )


def append_error_log(pszMessage: str) -> None:
    pszOutputPath: str = os.path.join(
        os.path.dirname(__file__),
        "DivisionSellGeneralAdminCost_Allocation_DnD_error.txt",
    )
    with open(pszOutputPath, "a", encoding="utf-8", newline="") as objFile:
        objFile.write(pszMessage + "\n")


def report_exception(pszContext: str, exc: Exception) -> None:
    pszTraceback: str = traceback.format_exc()
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
        "DivisionSellGeneralAdminCost_Allocation_DnD",
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
    return g_instruction_font_handle


def run_division_sell_general_admin_cost_allocation(
    objFilePaths: List[str],
) -> int:
    if not objFilePaths:
        show_error_message_box(
            "Error: no files were dropped.",
            "DivisionSellGeneralAdminCost_Allocation_DnD",
        )
        return 1

    pszScriptPath: str = os.path.join(
        os.path.dirname(__file__),
        "DivisionSellGeneralAdminCost_Allocation_Cmd.py",
    )
    if not os.path.exists(pszScriptPath):
        pszErrorMessage: str = (
            "Error: DivisionSellGeneralAdminCost_Allocation_Cmd.py not found. Path = "
            + pszScriptPath
        )
        append_error_log(pszErrorMessage)
        show_error_message_box(pszErrorMessage, "DivisionSellGeneralAdminCost_Allocation_DnD")
        return 1

    objCommand: List[str] = [sys.executable, pszScriptPath] + objFilePaths
    try:
        objResult = subprocess.run(
            objCommand,
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:  # noqa: BLE001
        report_exception(
            "unexpected exception while running DivisionSellGeneralAdminCost_Allocation_Cmd.py",
            exc,
        )
        return 1

    if objResult.returncode != 0:
        pszStdErr: str = objResult.stderr
        if pszStdErr.strip() == "":
            pszStdErr = "Process exited with non-zero return code and no stderr output."
        pszErrorMessage: str = (
            "Error: DivisionSellGeneralAdminCost_Allocation_Cmd.py exited with non-zero return code.\n\n"
            + "Return code = "
            + str(objResult.returncode)
            + "\n\n"
            + "stderr:\n"
            + pszStdErr
        )
        append_error_log(pszErrorMessage)
        show_error_message_box(pszErrorMessage, "DivisionSellGeneralAdminCost_Allocation_DnD")
        return objResult.returncode

    pszStdOut: str = objResult.stdout.strip()
    if pszStdOut == "":
        pszStdOut = "DivisionSellGeneralAdminCost_Allocation_Cmd.py finished successfully."

    objStdOutLines: List[str] = [pszLine.strip() for pszLine in pszStdOut.splitlines() if pszLine.strip() != ""]
    iStep0001Count: int = 0
    iStep0002Count: int = 0
    iStep0003Count: int = 0
    iWarningCount: int = 0
    objOutputDirectories: List[str] = []
    objSeenDirectories = set()
    for pszLine in objStdOutLines:
        if pszLine.startswith("Processed PL TSV count:"):
            try:
                iStep0001Count = int(pszLine.split(":", 1)[1].strip())
            except ValueError:
                iStep0001Count = 0
        elif pszLine.startswith("Processed step0002 TSV count:"):
            try:
                iStep0002Count = int(pszLine.split(":", 1)[1].strip())
            except ValueError:
                iStep0002Count = 0
        elif pszLine.startswith("Processed step0003 TSV count:"):
            try:
                iStep0003Count = int(pszLine.split(":", 1)[1].strip())
            except ValueError:
                iStep0003Count = 0
        elif pszLine.startswith("Warning error file count:"):
            try:
                iWarningCount = int(pszLine.split(":", 1)[1].strip())
            except ValueError:
                iWarningCount = 0
        elif pszLine.startswith("Output("):
            pszPath = pszLine.split(":", 1)[1].strip() if ":" in pszLine else ""
            if pszPath != "":
                pszDirectory: str = os.path.dirname(pszPath)
                if pszDirectory not in objSeenDirectories:
                    objOutputDirectories.append(pszDirectory)
                    objSeenDirectories.add(pszDirectory)

    pszResultPath: str = os.path.join(
        os.path.dirname(__file__),
        "DivisionSellGeneralAdminCost_Allocation_DnD_result.txt",
    )
    objResultLines: List[str] = [
        "JOB: DivisionSellGeneralAdminCost_Allocation_DnD.py",
        "DATE: " + datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        f"INPUT_FILE_COUNT: {len(objFilePaths)}",
        f"STEP0001_COUNT: {iStep0001Count}",
        f"STEP0002_COUNT: {iStep0002Count}",
        f"STEP0003_COUNT: {iStep0003Count}",
        f"WARNING_ERROR_FILE_COUNT: {iWarningCount}",
        "",
        "DETAIL_STDOUT:",
        pszStdOut,
    ]
    with open(pszResultPath, "w", encoding="utf-8", newline="") as objResultFile:
        objResultFile.write("\n".join(objResultLines) + "\n")

    pszOutputDirectoryText: str = ""
    if len(objOutputDirectories) == 1:
        pszOutputDirectoryText = objOutputDirectories[0]
    elif len(objOutputDirectories) >= 2:
        pszOutputDirectoryText = "複数フォルダ（詳細は結果一覧を参照）"
    else:
        pszOutputDirectoryText = os.path.dirname(__file__)

    pszMessage: str = (
        "成功しました。\n\n"
        + f"step0001: {iStep0001Count}件 / step0002: {iStep0002Count}件 / step0003: {iStep0003Count}件\n\n"
        + f"出力先: {pszOutputDirectoryText}\n\n"
        + "結果一覧を保存しました。"
    )
    show_message_box(pszMessage, "DivisionSellGeneralAdminCost_Allocation_DnD")
    return 0


def draw_instruction_text(
    iWindowHandle: int,
) -> None:
    iDeviceContextHandle, objPaintStruct = win32gui.BeginPaint(iWindowHandle)
    objClientRect = win32gui.GetClientRect(iWindowHandle)

    iFontHandle: Optional[int] = ensure_instruction_font_handle()
    iPreviousFontHandle: Optional[int] = None
    if iFontHandle:
        iPreviousFontHandle = win32gui.SelectObject(
            iDeviceContextHandle,
            iFontHandle,
        )

    iMargin: int = 8
    objDrawRect = (
        objClientRect[0] + iMargin,
        objClientRect[1] + iMargin,
        objClientRect[2] - iMargin,
        objClientRect[3] - iMargin,
    )
    pszInstructionText: str = (
        "division別販管費配賦用の入力ファイルを、このウィンドウにドラッグ＆ドロップしてください。\n"
        "複数ファイルをまとめてドラッグ＆ドロップできます。\n"
        "DivisionSellGeneralAdminCost_Allocation_Cmd.py にファイルパスを渡して処理します。"
    )
    iDrawTextFormat: int = win32con.DT_LEFT | win32con.DT_TOP | win32con.DT_WORDBREAK
    win32gui.DrawText(
        iDeviceContextHandle,
        pszInstructionText,
        -1,
        objDrawRect,
        iDrawTextFormat,
    )

    if iPreviousFontHandle:
        win32gui.SelectObject(
            iDeviceContextHandle,
            iPreviousFontHandle,
        )
    win32gui.EndPaint(iWindowHandle, objPaintStruct)


def window_proc(
    iWindowHandle: int,
    iMessage: int,
    iWparam: int,
    iLparam: int,
) -> int:
    if iMessage == win32con.WM_CREATE:
        win32gui.DragAcceptFiles(iWindowHandle, True)
        return 0

    if iMessage == win32con.WM_DROPFILES:
        iDropHandle: int = iWparam
        iFileCount: int = win32api.DragQueryFile(iDropHandle, -1)

        objFiles: List[str] = []
        for iIndex in range(iFileCount):
            pszFilePath: str = win32api.DragQueryFile(iDropHandle, iIndex)
            objFiles.append(pszFilePath)

        win32api.DragFinish(iDropHandle)

        if iFileCount <= 0:
            show_error_message_box(
                "Error: no files were dropped.",
                "DivisionSellGeneralAdminCost_Allocation_DnD",
            )
            return 0

        run_division_sell_general_admin_cost_allocation(objFiles)
        return 0

    if iMessage == win32con.WM_PAINT:
        draw_instruction_text(iWindowHandle)
        return 0

    if iMessage == win32con.WM_DESTROY:
        win32gui.PostQuitMessage(0)
        return 0

    return win32gui.DefWindowProc(iWindowHandle, iMessage, iWparam, iLparam)


def register_window_class(
    pszWindowClassName: str,
) -> int:
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


def create_main_window(
    pszWindowClassName: str,
    pszWindowTitle: str,
) -> int:
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
    iWindowWidth: int = int(iDesktopWidth * 0.618)
    iWindowHeight: int = int(iDesktopHeight * 0.618)

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

    win32gui.ShowWindow(iWindowHandle, win32con.SW_SHOWNORMAL)
    win32gui.UpdateWindow(iWindowHandle)

    win32gui.SetWindowPos(
        iWindowHandle,
        win32con.HWND_TOPMOST,
        0,
        0,
        0,
        0,
        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE,
    )

    win32gui.DragAcceptFiles(iWindowHandle, True)
    return iWindowHandle


def main() -> None:
    pszWindowClassName: str = "DivisionSellGeneralAdminCostAllocationDndWindowClass"
    pszWindowTitle: str = "DivisionSellGeneralAdminCost Allocation (Drag & Drop)"

    try:
        register_window_class(pszWindowClassName)
    except Exception as exc:
        report_exception("failed to register window class", exc)
        return

    try:
        create_main_window(
            pszWindowClassName,
            pszWindowTitle,
        )
    except Exception as exc:
        report_exception("failed to create main window", exc)
        return

    try:
        win32gui.PumpMessages()
    except Exception as exc:
        report_exception("unexpected exception in message loop", exc)


if __name__ == "__main__":
    main()
