"""
APA Style Excel Report Generator using openpyxl and matplotlib.
"""

import io
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyxlImage

def create_apa_border():
    """APAスタイル用ボーダー定義"""
    thick_side = Side(border_style="medium", color="000000")
    thin_side = Side(border_style="thin", color="000000")
    
    header_top = Border(top=thick_side, bottom=thin_side)
    header_bottom = Border(bottom=thin_side)
    table_top = Border(top=thick_side)
    table_bottom = Border(bottom=thick_side)
    
    return thick_side, thin_side

def format_apa_table(ws, start_row, title, df, note=None, table_num=1):
    """
    指定したワークシートにDFをAPA7thスタイルの表として書込み・フォーマットする関数
    
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
    thick_side = Side(border_style="medium", color="000000")
    thin_side = Side(border_style="thin", color="000000")
    
    current_row = start_row
    
    # 1. Table Number (Bold)
    ws.cell(row=current_row, column=1, value=f"Table {table_num}")
    ws.cell(row=current_row, column=1).font = Font(name="Calibri", size=11, bold=True)
    current_row += 1
    
    # 2. Table Title (Italic)
    ws.cell(row=current_row, column=1, value=title)
    ws.cell(row=current_row, column=1).font = Font(name="Calibri", size=11, italic=True)
    current_row += 1
    
    header_row = current_row
    num_cols = len(df.columns) + 1  # インデックス + 列数
    
    # Header Line Write
    # インデックス名または空文字
    idx_name = df.index.name if df.index.name else ""
    cell = ws.cell(row=header_row, column=1, value=str(idx_name))
    cell.font = Font(name="Calibri", size=11, bold=True)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = Border(top=thick_side, bottom=thin_side)
    
    for c_idx, col_name in enumerate(df.columns, start=2):
        cell = ws.cell(row=header_row, column=c_idx, value=str(col_name))
        cell.font = Font(name="Calibri", size=11, bold=True)
        # 数値ヘッダーは右寄せ、文字列ヘッダーは左寄せ/中央寄せ
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = Border(top=thick_side, bottom=thin_side)
        
    current_row += 1
    data_start_row = current_row
    
    # Data Rows Write
    for r_idx, (idx_val, row) in enumerate(df.iterrows()):
        r = current_row + r_idx
        # Index col
        idx_cell = ws.cell(row=r, column=1, value=str(idx_val))
        idx_cell.font = Font(name="Calibri", size=11)
        idx_cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # Values
        for c_idx, val in enumerate(row, start=2):
            val_cell = ws.cell(row=r, column=c_idx)
            val_cell.font = Font(name="Calibri", size=11)
            
            if pd.isna(val):
                val_cell.value = "-"
                val_cell.alignment = Alignment(horizontal="center", vertical="center")
            elif isinstance(val, (int, np.integer)):
                val_cell.value = int(val)
                val_cell.number_format = "#,##0"
                val_cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (float, np.floating)):
                val_cell.value = float(val)
                # 小数点フォーマット
                if abs(val) < 0.001 and val != 0:
                    val_cell.number_format = "0.000"
                else:
                    val_cell.number_format = "0.00"
                val_cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                val_cell.value = str(val)
                val_cell.alignment = Alignment(horizontal="left", vertical="center")
                
    data_end_row = current_row + len(df) - 1
    
    # Apply Table Bottom Border to last data row
    for c_idx in range(1, num_cols + 1):
        cell = ws.cell(row=data_end_row, column=c_idx)
        cell.border = Border(bottom=thick_side)
        
    current_row = data_end_row + 1
    
    # 3. Note (If exists)
    if note:
        note_cell = ws.cell(row=current_row, column=1, value=f"Note. {note}")
        note_cell.font = Font(name="Calibri", size=10, italic=True)
        current_row += 1
        
    current_row += 1  # 空間確保
    
    # 列幅調整
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= start_row and cell.row <= current_row:
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
    return current_row


def add_image_to_sheet(ws, img_bytes, cell_location="G2"):
    """
    BytesIO画像データをExcelシートの指定位置に挿入する
    """
    img_bytes.seek(0)
    img = OpenPyxlImage(img_bytes)
    # サイズ微調整 (例: 500x350)
    img.width = 550
    img.height = 380
    ws.add_image(img, cell_location)


def build_full_excel_report(analysis_results):
    """
    複数の分析結果（辞書形式）を受け取り、オープンピクセルWorkbookを作成してBytesIOで返す
    analysis_results: list of dict
      [
        {
          "sheet_name": "基本統計量",
          "title": "Descriptive Statistics",
          "df": dataframe,
          "note": "注釈文字列",
          "fig_bytes": BytesIO (optional)
        },
        ...
      ]
    """
    wb = openpyxl.Workbook()
    # デフォルトのシート削除
    wb.remove(wb.active)
    
    table_counter = 1
    for item in analysis_results:
        sheet_name = item.get("sheet_name", f"Table_{table_counter}")
        # シート名文字制限 (31文字)
        sheet_name = sheet_name[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        # 背景色を白（Gridlinesを有効）に設定
        ws.views.sheetView[0].showGridLines = True
        
        df = item.get("df")
        title = item.get("title", "Analysis Result")
        note = item.get("note", None)
        fig_bytes = item.get("fig_bytes", None)
        
        if df is not None and not df.empty:
            next_row = format_apa_table(ws, start_row=2, title=title, df=df, note=note, table_num=table_counter)
            table_counter += 1
        
        if fig_bytes is not None:
            # テーブルの横（G列付近）に画像を配置
            add_image_to_sheet(ws, fig_bytes, cell_location="H2")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
