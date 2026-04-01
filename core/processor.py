from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, List
import shutil

import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


LogFn = Callable[[str], None]


class ProcessingError(Exception):
    """业务处理异常。"""


@dataclass
class ProcessStats:
    folder1_files: int = 0
    folder2_files: int = 0
    folder1_sheets: int = 0
    folder2_sheets: int = 0
    rows_spot_bond: int = 0
    rows_pledged_repo: int = 0
    rows_agreement_repo: int = 0


COLUMN_MAPPING1 = {
    "成交序号": "成交序号",
    "成交编号": "成交编号",
    "业务日期": "业务日期",
    "股东代码": "股东代码",
    "委托方向": "委托方向",
    "证券代码": "证券代码",
    "交易市场": "交易市场",
    "成交数量": "成交数量",
    "成交均价": "成交均价",
    "成交金额": "成交金额",
    "业务分类": "业务分类",
    "证券类别": "证券类别",
    "证券名称": "证券名称",
}

COLUMN_MAPPING2 = {
    "成交序号": "成交序号",
    "成交编号": "成交编号",
    "业务日期": "业务日期",
    "股东代码": "股东代码",
    "委托方向": "委托方向",
    "交易市场": "交易市场",
    "质押数量(张/份)": "质押数量(张/份)",
    "成交最低价": "成交最低价",
    "成交金额": "成交金额",
    "质押券代码": "质押券代码",
    "回购天数": "回购天数",
}

SPOT_BOND_OUTPUT_COLUMNS = [
    "成交序号",
    "成交编号",
    "业务日期",
    "股东代码",
    "委托方向",
    "证券代码",
    "交易市场",
    "成交数量",
    "成交均价",
    "成交金额",
    "业务分类",
]

AGREEMENT_REPO_OUTPUT_COLUMNS = [
    "成交序号",
    "成交编号",
    "业务日期",
    "股东代码",
    "委托方向",
    "交易市场",
    "质押数量(张/份)",
    "成交最低价",
    "成交金额",
    "质押券代码",
    "回购天数",
]


class ExcelFlowProcessor:
    def __init__(self, logger: LogFn | None = None):
        self.logger = logger or (lambda msg: None)

    def process(
        self,
        folder1: str | Path,
        folder2: str | Path,
        template_file: str | Path,
        output_file: str | Path,
    ) -> ProcessStats:
        folder1 = Path(folder1)
        folder2 = Path(folder2)
        template_file = Path(template_file)
        output_file = Path(output_file)

        self._validate_paths(folder1, folder2, template_file, output_file)

        stats = ProcessStats()
        self.logger("开始读取 Excel 文件...")

        merged_data1, f1_count, s1_count = self._collect_from_folder(folder1, COLUMN_MAPPING1)
        merged_data2, f2_count, s2_count = self._collect_from_folder(folder2, COLUMN_MAPPING2)

        stats.folder1_files = f1_count
        stats.folder2_files = f2_count
        stats.folder1_sheets = s1_count
        stats.folder2_sheets = s2_count

        self.logger(f"成交信息pb：读取文件 {f1_count} 个，sheet {s1_count} 个。")
        self.logger(f"协议回购成交信息pb：读取文件 {f2_count} 个，sheet {s2_count} 个。")

        merged_data1 = self._transform_data1(merged_data1)
        merged_data2 = self._transform_data2(merged_data2)

        stats.rows_spot_bond = len(merged_data1)
        stats.rows_pledged_repo = len(merged_data1)
        stats.rows_agreement_repo = len(merged_data2)

        self.logger("开始写入模板文件...")
        self._write_to_template(template_file, output_file, merged_data1, merged_data2)

        self.logger("处理完成。")
        return stats

    def _validate_paths(
        self,
        folder1: Path,
        folder2: Path,
        template_file: Path,
        output_file: Path,
    ) -> None:
        if not folder1.exists() or not folder1.is_dir():
            raise ProcessingError(f"成交信息pb 文件夹不存在或不可用: {folder1}")
        if not folder2.exists() or not folder2.is_dir():
            raise ProcessingError(f"协议回购成交信息pb 文件夹不存在或不可用: {folder2}")
        if not template_file.exists() or not template_file.is_file():
            raise ProcessingError(f"模板文件不存在: {template_file}")
        if template_file.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ProcessingError("模板文件必须是 .xlsx 或 .xlsm 格式。")
        if output_file.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ProcessingError("输出文件必须是 .xlsx 或 .xlsm 格式。")

        output_file.parent.mkdir(parents=True, exist_ok=True)

    def _collect_from_folder(
        self,
        folder: Path,
        column_mapping: dict,
    ) -> tuple[pd.DataFrame, int, int]:
        excel_files = sorted(
            [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".xlsm"}]
        )
        merged_frames: List[pd.DataFrame] = []
        total_sheets = 0

        for file_path in excel_files:
            self.logger(f"读取文件: {file_path.name}")
            try:
                xls = pd.ExcelFile(file_path)
            except Exception as exc:
                raise ProcessingError(f"无法打开 Excel 文件 {file_path}: {exc}") from exc

            for sheet_name in xls.sheet_names:
                total_sheets += 1
                try:
                    sheet_df = pd.read_excel(file_path, sheet_name=sheet_name)
                except Exception as exc:
                    raise ProcessingError(f"读取文件 {file_path.name} 的 sheet {sheet_name} 失败: {exc}") from exc

                if sheet_df.empty:
                    self.logger(f"  - 跳过空sheet: {sheet_name}")
                    continue

                missing = [col for col in column_mapping.keys() if col not in sheet_df.columns]
                if missing:
                    self.logger(
                        f"  - 跳过sheet: {sheet_name}（缺少列: {', '.join(missing)}）"
                    )
                    continue

                filtered_df = sheet_df[list(column_mapping.keys())].copy()
                filtered_df.columns = [column_mapping.get(col, col) for col in filtered_df.columns]
                merged_frames.append(filtered_df)

        if not merged_frames:
            return pd.DataFrame(columns=list(column_mapping.values())), len(excel_files), total_sheets

        return pd.concat(merged_frames, ignore_index=True), len(excel_files), total_sheets

    def _transform_data1(self, data: pd.DataFrame) -> pd.DataFrame:
        if data.empty:
            return pd.DataFrame(columns=SPOT_BOND_OUTPUT_COLUMNS)

        transformed = data.copy()
        transformed.replace("固定收益平台", "上交所固定收益平台", inplace=True)

        mask_e = transformed["证券名称"].astype(str).str.contains("E", na=False) & (
            transformed["交易市场"] == "深交所A"
        )
        transformed.loc[mask_e, "业务分类"] = "固收收益平台"

        mask_jj = (transformed["交易市场"] == "深交所A") & (transformed["证券类别"] != "可转债")
        transformed.loc[mask_jj, "业务分类"] = "固收收益平台"

        transformed = transformed.drop(columns=["证券名称", "证券类别"], errors="ignore")
        transformed = transformed.reindex(columns=SPOT_BOND_OUTPUT_COLUMNS)
        return transformed

    def _transform_data2(self, data: pd.DataFrame) -> pd.DataFrame:
        if data.empty:
            return pd.DataFrame(columns=AGREEMENT_REPO_OUTPUT_COLUMNS)

        transformed = data.copy()
        transformed.replace({"正回购续做": "正回购", "逆回购续做": "逆回购"}, inplace=True)
        transformed = transformed.reindex(columns=AGREEMENT_REPO_OUTPUT_COLUMNS)
        return transformed

    def _write_to_template(
        self,
        template_file: Path,
        output_file: Path,
        data1: pd.DataFrame,
        data2: pd.DataFrame,
    ) -> None:
        if template_file.resolve() != output_file.resolve():
            shutil.copy2(template_file, output_file)
            self.logger(f"已复制模板到输出文件: {output_file}")

        try:
            book = load_workbook(output_file)
        except Exception as exc:
            raise ProcessingError(f"打开输出文件失败: {exc}") from exc

        self._write_regular_sheet(book, "现券买卖", data1, max_col=11)
        # 修复点：质押式回购应写入协议回购数据（merged_data2）而不是 merged_data1
        self._write_regular_sheet(book, "质押式回购", data2, max_col=11)
        self._write_agreement_sheet(book, "协议回购", data2)

        try:
            book.save(output_file)
        except Exception as exc:
            raise ProcessingError(f"保存输出文件失败: {exc}") from exc
        finally:
            book.close()

    def _require_sheet(self, book: Workbook, sheet_name: str):
        if sheet_name not in book.sheetnames:
            raise ProcessingError(f"模板缺少工作表：{sheet_name}")
        return book[sheet_name]

    def _clear_sheet(self, sheet, min_row: int, max_col: int) -> None:
        for row in sheet.iter_rows(min_row=min_row, max_col=max_col, max_row=sheet.max_row):
            for cell in row:
                cell.value = None

    def _write_regular_sheet(self, book: Workbook, sheet_name: str, data: pd.DataFrame, max_col: int) -> None:
        sheet = self._require_sheet(book, sheet_name)
        self._clear_sheet(sheet, min_row=4, max_col=max_col)

        for row_num, row_data in enumerate(data.itertuples(index=False), 4):
            for col_num, cell_value in enumerate(row_data, 1):
                sheet.cell(row=row_num, column=col_num, value=cell_value)
        self.logger(f"  - 已写入[{sheet_name}] {len(data)} 行")

    def _write_agreement_sheet(self, book: Workbook, sheet_name: str, data: pd.DataFrame) -> None:
        sheet = self._require_sheet(book, sheet_name)
        self._clear_sheet(sheet, min_row=4, max_col=12)

        target_columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "L"]
        for row_num, row_data in enumerate(data.itertuples(index=False), 4):
            for col_letter, cell_value in zip(target_columns, row_data):
                sheet[f"{col_letter}{row_num}"] = cell_value

        self.logger(f"  - 已写入[{sheet_name}] {len(data)} 行")
