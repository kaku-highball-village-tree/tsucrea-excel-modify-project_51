# -*- coding: utf-8 -*-
"""
DivisionSellGeneralAdminCost_Allocation_Cmd.py

受け取った入力ファイル群（SellGeneralAdminCost_Allocation_Cmd_0002.py と同等）から、
損益計算書TSVを対象に 1列目と「販売費及び一般管理費計」列のみを残した
Div販管費TSVを出力する。
"""

from __future__ import annotations

import csv
import os
import re
import sys
from typing import List, Optional


PL_TSV_PATTERN = re.compile(r"^損益計算書_(\d{4})年(\d{2})月_A∪B_プロジェクト名_C∪D_vertical\.tsv$")
TARGET_COLUMN_NAME: str = "販売費及び一般管理費計"


def is_pl_tsv_file_name(pszBaseName: str) -> bool:
    return PL_TSV_PATTERN.fullmatch(pszBaseName) is not None


def build_output_file_name(pszInputBaseName: str) -> Optional[str]:
    objMatch = PL_TSV_PATTERN.fullmatch(pszInputBaseName)
    if objMatch is None:
        return None
    iYear: str = objMatch.group(1)
    iMonth: str = objMatch.group(2)
    return f"損益計算書_{iYear}年{iMonth}月_A∪B_C∪D_Div販管費_vertical.tsv"


def find_target_column_index(objHeader: List[str]) -> Optional[int]:
    for iIndex, pszColumn in enumerate(objHeader):
        if pszColumn == TARGET_COLUMN_NAME:
            return iIndex
    return None


def process_one_pl_tsv(pszInputPath: str) -> str:
    pszBaseName: str = os.path.basename(pszInputPath)
    pszOutputBaseName: Optional[str] = build_output_file_name(pszBaseName)
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


def main() -> int:
    objInputFiles: List[str] = sys.argv[1:]
    if not objInputFiles:
        print("Error: no input files specified.", file=sys.stderr)
        return 1

    objPlInputFiles: List[str] = []
    for pszPath in objInputFiles:
        pszBaseName: str = os.path.basename(pszPath)
        if is_pl_tsv_file_name(pszBaseName):
            objPlInputFiles.append(pszPath)

    if not objPlInputFiles:
        print("Error: no PL TSV files matched the expected format.", file=sys.stderr)
        return 1

    objOutputPaths: List[str] = []
    for pszPlPath in objPlInputFiles:
        if not os.path.isfile(pszPlPath):
            print(f"Error: input file not found: {pszPlPath}", file=sys.stderr)
            return 1
        try:
            pszOutputPath: str = process_one_pl_tsv(pszPlPath)
        except Exception as exc:  # noqa: BLE001
            print(f"Error: failed to process {pszPlPath}. Detail = {exc}", file=sys.stderr)
            return 1
        objOutputPaths.append(pszOutputPath)

    print(f"Processed PL TSV count: {len(objOutputPaths)}")
    for pszOutputPath in objOutputPaths:
        print(f"Output: {pszOutputPath}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
