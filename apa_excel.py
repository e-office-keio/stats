"""
APA Style Excel Report Generator with Native Editable Excel Charts & Image fallback.
(Full Japanese & Font Fix Support)
"""

import io
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.chart import BarChart, LineChart, Reference, Series

def create_apa_border():
    """APAスタイル用ボーダー定義"""
    thick_side = Side(border_style="medium", color="000000")
    thin_side = Side(border_style="thin", color="000000")
    return thick_side, thin_side

def format_apa_table(ws, start_row, title, df, note=None, table_num=1):
    """
    指定したワークシートにDFをAPA7thスタイルの表（日本語表記）として書き込む
    
    Returns:
    - data_info: {"start_row": int, "end_row": int, "num_cols": int}
    - current_row: 次の要素を書き込める開始行番号
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
    data_start_row = current_row
    
    # データ行の書き込み
    for r_idx, (idx_val, row) in enumerate(df.iterrows()):
        r = current_row + r_idx
        idx_cell = ws.cell(row=r, column=1, value=str(idx_val))
        idx_cell.font = Font(name=font_family, size=11)
        idx_cell.alignment = Alignment(horizontal="left", vertical="center")
        
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
                len_count = sum(2 if ord(c) > 256 else 1 for c in val_str)
                max_len = max(max_len, len_count)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        
    data_info = {
        "header_row": header_row,
        "data_start_row": data_start_row,
        "data_end_row": data_end_row,
        "num_cols": num_cols
    }
    return data_info, current_row


def add_native_excel_chart(ws, data_info, title="グラフ"):
    """
    Excel上でダブルクリック＆自在に編集可能な「ネイティブグラフ」を作成して挿入
    """
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "数値 / 度数"
    chart.x_axis.title = "カテゴリ"
    
    # データ範囲とカテゴリ名（X軸）の参照
    data_ref = Reference(ws, min_col=2, min_row=data_info["header_row"], max_col=data_info["num_cols"], max_row=data_info["data_end_row"])
    cats_ref = Reference(ws, min_col=1, min_row=data_info["data_start_row"], max_row=data_info["data_end_row"])
    
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    
    chart.width = 16
    chart.height = 10
    
    ws.add_chart(chart, "H2")


def add_image_to_sheet(ws, img_bytes, cell_location="H2"):
    """
    画像データ(bytes または BytesIO)をExcelシートの指定位置に挿入する
    """
    if isinstance(img_bytes, bytes):
        img_buf = io.BytesIO(img_bytes)
    else:
        img_bytes.seek(0)
        img_buf = img_bytes
    img = OpenPyxlImage(img_buf)
    img.width = 550
    img.height = 380
    ws.add_image(img, cell_location)


def build_full_excel_report(analysis_results):
    """
    複数の分析結果（辞書形式）を受け取り、オープンピクセルWorkbookを作成してBytesIOで返す。
    ネィティブ編集可能グラフと画像グラフの両方を埋め込み。
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
        
        data_info = None
        if df is not None and not df.empty:
            data_info, next_row = format_apa_table(ws, start_row=2, title=title, df=df, note=note, table_num=table_counter)
            table_counter += 1
            
            # 単純集計・クロス集計・記述統計量等の場合はExcelネイティブ編集可能チャートを追加
            try:
                if "単純集計" in sheet_name or "クロス" in sheet_name or "基本統計量" in sheet_name or "t検定" in sheet_name or "ANOVA" in sheet_name:
                    add_native_excel_chart(ws, data_info, title=f"{title} (編集可能グラフ)")
            except Exception:
                pass  # 万が一ネイティブチャート作成不可時は画像フォールバック
        
        if fig_bytes is not None:
            # ネイティブグラフがある場合はその下（H18）に画像配置、ない場合はH2に配置
            img_pos = "H18" if data_info and ("単純集計" in sheet_name or "クロス" in sheet_name or "基本統計量" in sheet_name or "t検定" in sheet_name or "ANOVA" in sheet_name) else "H2"
            add_image_to_sheet(ws, fig_bytes, cell_location=img_pos)
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
