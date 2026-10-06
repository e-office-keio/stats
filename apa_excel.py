"""
APA Style Excel Report Generator using openpyxl and matplotlib.
(Japanese Title and Notation Support)
"""

import io
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyxlImage

def create_apa_border():
    """APAスタイル用ボーダー定義"""
    thick_side = Side(border_style="medium", color="000000")
    thin_side = Side(border_style="thin", color="000000")
    return thick_side, thin_side

def format_apa_table(ws, start_row, title, df, note=None, table_num=1):
    """
    指定したワークシートにDFをAPA7thスタイルの表（日本語表記）として書き込む
    
    Parameters:
    - ws: openpyxl Worksheet
    - start_row: 開始行番号 (1-indexed)
    - title: 表のタイトル (str)
    - df: 表示するpandas DataFrame
    - note: 表下の注釈テキスト (str)
    - table_num: 表番号 (int)
    
    Returns:
    - next_row: 次の要素を書き込める開始行番号
    """
    thick_side, thin_side = create_apa_border()
    font_family = "游ゴシック"  # 日本語標準フォント
    
    current_row = start_row
    
    # 1. 表番号 (太字)
    ws.cell(row=current_row, column=1, value=f"表 {table_num}")
    ws.cell(row=current_row, column=1).font = Font(name=font_family, size=11, bold=True)
    current_row += 1
    
    # 2. 表タイトル (斜体)
    ws.cell(row=current_row, column=1, value=title)
    ws.cell(row=current_row, column=1).font = Font(name=font_family, size=11, italic=True)
    current_row += 1
    
    header_row = current_row
    num_cols = len(df.columns) + 1  # インデックス + 列数
    
    # ヘッダー行の書き込み
    idx_name = df.index.name if df.index.name else ""
    cell = ws.cell(row=header_row, column=1, value=str(idx_name))
    cell.font = Font(name=font_family, size=11, bold=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = Border(top=thick_side, bottom=thin_side)
    
    for c_idx, col_name in enumerate(df.columns, start=2):
        cell = ws.cell(row=header_row, column=c_idx, value=str(col_name))
        cell.font = Font(name=font_family, size=11, bold=True)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = Border(top=thick_side, bottom=thin_side)
        
    current_row += 1
    
    # データ行の書き込み
    for r_idx, (idx_val, row) in enumerate(df.iterrows()):
        r = current_row + r_idx
        # インデックス列
        idx_cell = ws.cell(row=r, column=1, value=str(idx_val))
        idx_cell.font = Font(name=font_family, size=11)
        idx_cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # 数値・文字列データ列
        for c_idx, val in enumerate(row, start=2):
            val_cell = ws.cell(row=r, column=c_idx)
            val_cell.font = Font(name=font_family, size=11)
            
            if pd.isna(val):
                val_cell.value = "-"
                val_cell.alignment = Alignment(horizontal="center", vertical="center")
            elif isinstance(val, (int, np.integer)):
                val_cell.value = int(val)
                val_cell.number_format = "#,##0"
                val_cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (float, np.floating)):
                val_cell.value = float(val)
                if abs(val) < 0.001 and val != 0:
                    val_cell.number_format = "0.000"
                else:
                    val_cell.number_format = "0.00"
                val_cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                val_cell.value = str(val)
                val_cell.alignment = Alignment(horizontal="left", vertical="center")
                
    data_end_row = current_row + len(df) - 1
    
    # テーブル最下部に太枠線を適用
    for c_idx in range(1, num_cols + 1):
        cell = ws.cell(row=data_end_row, column=c_idx)
        cell.border = Border(bottom=thick_side)
        
    current_row = data_end_row + 1
    
    # 3. 注釈 (Note -> 注)
    if note:
        note_text = note if note.startswith("注.") else f"注. {note}"
        note_cell = ws.cell(row=current_row, column=1, value=note_text)
        note_cell.font = Font(name=font_family, size=10, italic=True)
        current_row += 1
        
    current_row += 1
    
    # 列幅の調整
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= start_row and cell.row <= current_row:
                val_str = str(cell.value or "")
                # 日本語全角文字の幅を考慮（全角は約2文字分）
                len_count = sum(2 if ord(c) > 256 else 1 for c in val_str)
                max_len = max(max_len, len_count)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        
    return current_row


def add_image_to_sheet(ws, img_bytes, cell_location="H2"):
    """
    BytesIO画像データをExcelシートの指定位置に挿入する
    """
    img_bytes.seek(0)
    img = OpenPyxlImage(img_bytes)
    img.width = 550
    img.height = 380
    ws.add_image(img, cell_location)


def build_full_excel_report(analysis_results):
    """
    複数の分析結果（辞書形式）を受け取り、オープンピクセルWorkbookを作成してBytesIOで返す
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # デフォルトシート削除
    
    table_counter = 1
    for item in analysis_results:
        sheet_name = item.get("sheet_name", f"表_{table_counter}")
        sheet_name = sheet_name[:31]  # Excelのシート名31文字制限
        ws = wb.create_sheet(title=sheet_name)
        
        ws.views.sheetView[0].showGridLines = True
        
        df = item.get("df")
        title = item.get("title", "分析結果")
        note = item.get("note", None)
        fig_bytes = item.get("fig_bytes", None)
        
        if df is not None and not df.empty:
            next_row = format_apa_table(ws, start_row=2, title=title, df=df, note=note, table_num=table_counter)
            table_counter += 1
        
        if fig_bytes is not None:
            add_image_to_sheet(ws, fig_bytes, cell_location="H2")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
