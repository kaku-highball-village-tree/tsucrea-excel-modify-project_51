# -*- coding: utf-8 -*-
"""
DivisionSellGeneralAdminCost_Allocation_Cmd.py

受け取った入力ファイル群（SellGeneralAdminCost_Allocation_Cmd_0002.py と同等）から、
損益計算書TSVを対象に 1列目と「販売費及び一般管理費計」列のみを残した
Div販管費TSVを出力する。

さらに、step0001 の出力TSVを読み込み、
「スタートアップコミュニティDiv販管費」〜「C008_新規プロポーザル」
の2列目合計を算出して、各終端行の直下に「Div販管費」行を挿入した
step0002 TSV を出力する。
"""

from __future__ import annotations

import csv
import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Optional, Tuple


PL_TSV_PATTERN = re.compile(r"^損益計算書_(\d{4})年(\d{2})月_A∪B_プロジェクト名_C∪D_vertical\.tsv$")
STEP0001_TSV_PATTERN = re.compile(r"^損益計算書_step0001_(\d{4})年(\d{2})月_A∪B_C∪D_Div販管費_vertical\.tsv$")
STEP0002_TSV_PATTERN = re.compile(r"^損益計算書_step0002_(\d{4})年(\d{2})月_A∪B_C∪D_Div販管費_vertical\.tsv$")
MANHOUR_TSV_PATTERN = re.compile(
    r"^工数_(\d{4})年(\d{2})月_step0014_各プロジェクトの計上カンパニー名_工数_カンパニーの工数\.tsv$"
)
TARGET_COLUMN_NAME: str = "販売費及び一般管理費計"
RANGE_START_LABEL: str = "スタートアップコミュニティDiv販管費"
RANGE_END_LABEL: str = "C008_新規プロポーザル"
PROJECT_CODE_PATTERN = re.compile(r"((?:P\d{5}|[^P]\d{3})_)")


def append_error_log(pszMessage: str) -> None:
    pszOutputPath: str = os.path.join(
        os.path.dirname(__file__),
        "DivisionSellGeneralAdminCost_Allocation_Cmd_error.txt",
    )
    with open(pszOutputPath, "a", encoding="utf-8", newline="") as objFile:
        objFile.write(pszMessage + "\n")


def write_numeric_warning_error_file(
    pszInputPath: str,
    objWarnings: List[Dict[str, str]],
) -> str:
    pszNow: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    append_error_log(f"[{pszNow}] ERROR START")
    append_error_log("JOB: DivisionSellGeneralAdminCost_Allocation_Cmd.py")
    append_error_log("ERROR_TYPE: NON_NUMERIC_VALUE_REPLACED_WITH_ZERO")
    append_error_log("INPUT_FILE: " + pszInputPath)
    for objWarning in objWarnings:
        append_error_log("ROW_INDEX: " + objWarning["row_index"])
        append_error_log("VALUE: " + objWarning["value"])
        append_error_log("REPLACED_VALUE: 0")
    append_error_log("RESULT: WARNING")
    pszEnd: str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    append_error_log(f"[{pszEnd}] ERROR END")

    pszBaseName: str = os.path.basename(pszInputPath)
    pszErrorPath: str = os.path.join(os.path.dirname(pszInputPath), pszBaseName + "_error.txt")
    objLines: List[str] = [
        "JOB: DivisionSellGeneralAdminCost_Allocation_Cmd.py",
        "ERROR_TYPE: NON_NUMERIC_VALUE_REPLACED_WITH_ZERO",
        "INPUT_FILE: " + pszInputPath,
    ]
    for objWarning in objWarnings:
        objLines.append("ROW_INDEX: " + objWarning["row_index"])
        objLines.append("VALUE: " + objWarning["value"])
        objLines.append("REPLACED_VALUE: 0")
    objLines.append("RESULT: WARNING")
    with open(pszErrorPath, "w", encoding="utf-8", newline="") as objFile:
        objFile.write("\n".join(objLines) + "\n")
    return pszErrorPath


def is_pl_tsv_file_name(pszBaseName: str) -> bool:
    return PL_TSV_PATTERN.fullmatch(pszBaseName) is not None


def is_step0001_tsv_file_name(pszBaseName: str) -> bool:
    return STEP0001_TSV_PATTERN.fullmatch(pszBaseName) is not None


def is_step0002_tsv_file_name(pszBaseName: str) -> bool:
    return STEP0002_TSV_PATTERN.fullmatch(pszBaseName) is not None


def is_manhour_tsv_file_name(pszBaseName: str) -> bool:
    return MANHOUR_TSV_PATTERN.fullmatch(pszBaseName) is not None


def build_step0001_output_file_name(pszInputBaseName: str) -> Optional[str]:
    objMatch = PL_TSV_PATTERN.fullmatch(pszInputBaseName)
    if objMatch is None:
        return None
    iYear: str = objMatch.group(1)
    iMonth: str = objMatch.group(2)
    return f"損益計算書_step0001_{iYear}年{iMonth}月_A∪B_C∪D_Div販管費_vertical.tsv"


def build_step0002_output_file_name(pszStep0001BaseName: str) -> Optional[str]:
    objMatch = STEP0001_TSV_PATTERN.fullmatch(pszStep0001BaseName)
    if objMatch is None:
        return None
    iYear: str = objMatch.group(1)
    iMonth: str = objMatch.group(2)
    return f"損益計算書_step0002_{iYear}年{iMonth}月_A∪B_C∪D_Div販管費_vertical.tsv"


def build_step0003_output_file_name(pszStep0002BaseName: str) -> Optional[str]:
    objMatch = STEP0002_TSV_PATTERN.fullmatch(pszStep0002BaseName)
    if objMatch is None:
        return None
    iYear: str = objMatch.group(1)
    iMonth: str = objMatch.group(2)
    return f"損益計算書_step0003_{iYear}年{iMonth}月_A∪B_C∪D_Div販管費_vertical.tsv"


def extract_year_month_from_name(pszBaseName: str) -> Optional[str]:
    objPatterns = [PL_TSV_PATTERN, STEP0001_TSV_PATTERN, STEP0002_TSV_PATTERN, MANHOUR_TSV_PATTERN]
    for objPattern in objPatterns:
        objMatch = objPattern.fullmatch(pszBaseName)
        if objMatch is not None:
            return f"{objMatch.group(1)}年{objMatch.group(2)}月"
    return None


def extract_project_code(pszText: str) -> Optional[str]:
    objMatch = PROJECT_CODE_PATTERN.search(pszText)
    if objMatch is None:
        return None
    return objMatch.group(1)


def find_target_column_index(objHeader: List[str]) -> Optional[int]:
    for iIndex, pszColumn in enumerate(objHeader):
        if pszColumn == TARGET_COLUMN_NAME:
            return iIndex
    return None


def process_one_pl_tsv_to_step0001(pszInputPath: str) -> str:
    pszBaseName: str = os.path.basename(pszInputPath)
    pszOutputBaseName: Optional[str] = build_step0001_output_file_name(pszBaseName)
    if pszOutputBaseName is None:
        raise ValueError(f"invalid PL TSV file name: {pszBaseName}")

    pszOutputPath: str = os.path.join(os.path.dirname(pszInputPath), pszOutputBaseName)

    with open(pszInputPath, "r", encoding="utf-8", newline="") as objInputFile:
        objReader = csv.reader(objInputFile, delimiter="\t")
        objRows: List[List[str]] = list(objReader)

    if not objRows:
        raise ValueError(f"empty TSV: {pszInputPath}")

    objHeader: List[str] = objRows[0]
    if len(objHeader) == 0:
        raise ValueError(f"header is empty: {pszInputPath}")

    iTargetIndex: Optional[int] = find_target_column_index(objHeader)
    if iTargetIndex is None:
        raise ValueError(
            f"required column not found: {TARGET_COLUMN_NAME} in {pszInputPath}"
        )

    iFirstColumnIndex: int = 0
    objOutputRows: List[List[str]] = []
    for objRow in objRows:
        pszFirstValue: str = objRow[iFirstColumnIndex] if len(objRow) > iFirstColumnIndex else ""
        pszTargetValue: str = objRow[iTargetIndex] if len(objRow) > iTargetIndex else ""
        objOutputRows.append([pszFirstValue, pszTargetValue])

    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        objWriter = csv.writer(objOutputFile, delimiter="\t", lineterminator="\n")
        objWriter.writerows(objOutputRows)

    return pszOutputPath


def parse_numeric_value(
    pszValue: str,
    iRowIndex: int,
    objWarnings: List[Dict[str, str]],
) -> float:
    pszTrimmed: str = pszValue.strip()
    if pszTrimmed == "":
        return 0.0

    pszNormalized: str = pszTrimmed.replace(",", "")
    try:
        return float(pszNormalized)
    except ValueError:
        objWarnings.append({"row_index": str(iRowIndex), "value": pszValue})
        return 0.0


def process_one_step0001_tsv_to_step0002(pszStep0001Path: str) -> Tuple[Optional[str], Optional[str]]:
    pszBaseName: str = os.path.basename(pszStep0001Path)
    pszOutputBaseName: Optional[str] = build_step0002_output_file_name(pszBaseName)
    if pszOutputBaseName is None:
        raise ValueError(f"invalid step0001 TSV file name: {pszBaseName}")

    with open(pszStep0001Path, "r", encoding="utf-8", newline="") as objInputFile:
        objReader = csv.reader(objInputFile, delimiter="\t")
        objRows: List[List[str]] = list(objReader)

    if not objRows:
        return None, None

    objOutputRows: List[List[str]] = []
    bInRange: bool = False
    fCurrentRangeSum: float = 0.0
    iInsertedCount: int = 0
    objWarnings: List[Dict[str, str]] = []

    for iRowIndex, objRow in enumerate(objRows):
        objOutputRows.append(objRow)
        pszFirstColumn: str = objRow[0] if len(objRow) >= 1 else ""
        pszSecondColumn: str = objRow[1] if len(objRow) >= 2 else ""

        if pszFirstColumn == RANGE_START_LABEL:
            bInRange = True
            fCurrentRangeSum = 0.0

        if bInRange:
            fCurrentRangeSum += parse_numeric_value(pszSecondColumn, iRowIndex, objWarnings)

        if bInRange and pszFirstColumn == RANGE_END_LABEL:
            objOutputRows.append(["Div販管費", str(int(fCurrentRangeSum))])
            iInsertedCount += 1
            bInRange = False
            fCurrentRangeSum = 0.0

    if iInsertedCount == 0:
        return None, None

    pszOutputPath: str = os.path.join(os.path.dirname(pszStep0001Path), pszOutputBaseName)
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        objWriter = csv.writer(objOutputFile, delimiter="\t", lineterminator="\n")
        objWriter.writerows(objOutputRows)

    pszErrorPath: Optional[str] = None
    if objWarnings:
        pszErrorPath = write_numeric_warning_error_file(pszStep0001Path, objWarnings)

    return pszOutputPath, pszErrorPath


def load_manhour_map(pszManhourPath: str) -> Dict[str, str]:
    with open(pszManhourPath, "r", encoding="utf-8", newline="") as objInputFile:
        objReader = csv.reader(objInputFile, delimiter="\t")
        objRows: List[List[str]] = list(objReader)

    objMap: Dict[str, str] = {}
    for objRow in objRows:
        if len(objRow) < 3:
            continue
        pszProjectCode: Optional[str] = extract_project_code(objRow[0])
        if pszProjectCode is None:
            continue
        objMap[pszProjectCode] = objRow[2]
    return objMap


def process_one_step0002_with_manhour_to_step0003(
    pszStep0002Path: str,
    pszManhourPath: str,
) -> str:
    pszStep0002BaseName: str = os.path.basename(pszStep0002Path)
    pszOutputBaseName: Optional[str] = build_step0003_output_file_name(pszStep0002BaseName)
    if pszOutputBaseName is None:
        raise ValueError(f"invalid step0002 TSV file name: {pszStep0002BaseName}")

    objManhourMap: Dict[str, str] = load_manhour_map(pszManhourPath)

    with open(pszStep0002Path, "r", encoding="utf-8", newline="") as objInputFile:
        objReader = csv.reader(objInputFile, delimiter="\t")
        objRows: List[List[str]] = list(objReader)

    objOutputRows: List[List[str]] = []
    for iRowIndex, objRow in enumerate(objRows):
        objNewRow: List[str] = list(objRow)
        while len(objNewRow) < 3:
            objNewRow.append("")
        if iRowIndex == 0:
            objNewRow[2] = "工数"
            objOutputRows.append(objNewRow)
            continue
        pszProjectCode: Optional[str] = extract_project_code(objNewRow[0] if objNewRow else "")
        if pszProjectCode is not None and pszProjectCode in objManhourMap:
            objNewRow[2] = objManhourMap[pszProjectCode]
        elif objNewRow[2] == "":
            objNewRow[2] = "0:00:00"
        objOutputRows.append(objNewRow)

    pszOutputPath: str = os.path.join(os.path.dirname(pszStep0002Path), pszOutputBaseName)
    with open(pszOutputPath, "w", encoding="utf-8", newline="") as objOutputFile:
        objWriter = csv.writer(objOutputFile, delimiter="\t", lineterminator="\n")
        objWriter.writerows(objOutputRows)
    return pszOutputPath


def build_year_month_path_map(objPaths: List[str]) -> Dict[str, str]:
    objMap: Dict[str, str] = {}
    for pszPath in objPaths:
        pszYearMonth: Optional[str] = extract_year_month_from_name(os.path.basename(pszPath))
        if pszYearMonth is None:
            continue
        objMap[pszYearMonth] = pszPath
    return objMap


def run_step0003_with_maps(
    objStep0002Map: Dict[str, str],
    objManhourMap: Dict[str, str],
) -> Tuple[List[str], List[str]]:
    objStep0003OutputPaths: List[str] = []
    objSkippedYearMonths: List[str] = []
    for pszYearMonth, pszStep0002Path in sorted(objStep0002Map.items()):
        if pszYearMonth not in objManhourMap:
            objSkippedYearMonths.append(pszYearMonth)
            continue
        pszStep0003Path: str = process_one_step0002_with_manhour_to_step0003(
            pszStep0002Path,
            objManhourMap[pszYearMonth],
        )
        objStep0003OutputPaths.append(pszStep0003Path)
    return objStep0003OutputPaths, objSkippedYearMonths


def main() -> int:
    objInputFiles: List[str] = sys.argv[1:]
    if not objInputFiles:
        print("Error: no input files specified.", file=sys.stderr)
        return 1

    objPlInputFiles: List[str] = []
    objStep0002InputFiles: List[str] = []
    objManhourInputFiles: List[str] = []
    for pszPath in objInputFiles:
        pszBaseName: str = os.path.basename(pszPath)
        if is_pl_tsv_file_name(pszBaseName):
            objPlInputFiles.append(pszPath)
        elif is_step0002_tsv_file_name(pszBaseName):
            objStep0002InputFiles.append(pszPath)
        elif is_manhour_tsv_file_name(pszBaseName):
            objManhourInputFiles.append(pszPath)

    if objStep0002InputFiles and objManhourInputFiles:
        objStep0002Map: Dict[str, str] = build_year_month_path_map(objStep0002InputFiles)
        objManhourMap: Dict[str, str] = build_year_month_path_map(objManhourInputFiles)
        try:
            objStep0003OutputPaths, objSkippedYearMonths = run_step0003_with_maps(
                objStep0002Map,
                objManhourMap,
            )
        except Exception as exc:  # noqa: BLE001
            print(f"Error: failed to process step0003. Detail = {exc}", file=sys.stderr)
            return 1

        print(f"Processed step0003 TSV count: {len(objStep0003OutputPaths)}")
        for pszOutputPath in objStep0003OutputPaths:
            print(f"Output(step0003): {pszOutputPath}")
        if objSkippedYearMonths:
            print(f"Skipped step0003 count: {len(objSkippedYearMonths)}")
            for pszYearMonth in objSkippedYearMonths:
                print(f"Skipped(step0003): {pszYearMonth}")
        return 0

    if not objPlInputFiles:
        print("Error: no PL TSV files matched the expected format.", file=sys.stderr)
        return 1

    objStep0001OutputPaths: List[str] = []
    for pszPlPath in objPlInputFiles:
        if not os.path.isfile(pszPlPath):
            print(f"Error: input file not found: {pszPlPath}", file=sys.stderr)
            return 1
        try:
            pszStep0001Path: str = process_one_pl_tsv_to_step0001(pszPlPath)
        except Exception as exc:  # noqa: BLE001
            print(f"Error: failed to process {pszPlPath}. Detail = {exc}", file=sys.stderr)
            return 1
        objStep0001OutputPaths.append(pszStep0001Path)

    objStep0002OutputPaths: List[str] = []
    objSkippedStep0001Paths: List[str] = []
    objErrorPaths: List[str] = []
    for pszStep0001Path in objStep0001OutputPaths:
        try:
            objStep0002Result = process_one_step0001_tsv_to_step0002(pszStep0001Path)
        except Exception as exc:  # noqa: BLE001
            print(f"Error: failed to process step0002 {pszStep0001Path}. Detail = {exc}", file=sys.stderr)
            return 1
        pszStep0002Path, pszErrorPath = objStep0002Result
        if pszStep0002Path is None:
            objSkippedStep0001Paths.append(pszStep0001Path)
            continue
        objStep0002OutputPaths.append(pszStep0002Path)
        if pszErrorPath is not None:
            objErrorPaths.append(pszErrorPath)

    print(f"Processed PL TSV count: {len(objStep0001OutputPaths)}")
    for pszOutputPath in objStep0001OutputPaths:
        print(f"Output(step0001): {pszOutputPath}")

    print(f"Processed step0002 TSV count: {len(objStep0002OutputPaths)}")
    for pszOutputPath in objStep0002OutputPaths:
        print(f"Output(step0002): {pszOutputPath}")

    if objSkippedStep0001Paths:
        print(f"Skipped step0002 count: {len(objSkippedStep0001Paths)}")
        for pszPath in objSkippedStep0001Paths:
            print(f"Skipped(step0002): {pszPath}")

    if objErrorPaths:
        print(f"Warning error file count: {len(objErrorPaths)}")
        for pszPath in objErrorPaths:
            print(f"WarningErrorFile: {pszPath}")

    if objManhourInputFiles:
        objStep0002Map = build_year_month_path_map(objStep0002OutputPaths)
        objManhourMap = build_year_month_path_map(objManhourInputFiles)
        try:
            objStep0003OutputPaths, objSkippedYearMonths = run_step0003_with_maps(
                objStep0002Map,
                objManhourMap,
            )
        except Exception as exc:  # noqa: BLE001
            print(f"Error: failed to process step0003. Detail = {exc}", file=sys.stderr)
            return 1
        print(f"Processed step0003 TSV count: {len(objStep0003OutputPaths)}")
        for pszOutputPath in objStep0003OutputPaths:
            print(f"Output(step0003): {pszOutputPath}")
        if objSkippedYearMonths:
            print(f"Skipped step0003 count: {len(objSkippedYearMonths)}")
            for pszYearMonth in objSkippedYearMonths:
                print(f"Skipped(step0003): {pszYearMonth}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
