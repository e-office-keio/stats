"""
APA Style Excel Report Generator with Native Editable Excel Charts for All Analyses.
(Full Japanese, Editable Charts & Font Fix Support)
"""

import io
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.chart import BarChart, LineChart, ScatterChart, Reference, Series


def create_apa_border():
    """APAスタイル用ボーダー定義"""
    thick_side = Side(border_style="medium", color="000000")
    thin_side = Side(border_style="thin", color="000000")
    return thick_side, thin_side


def format_apa_table(ws, start_row, title, df, note=None, table_num=1, is_factor_analysis=False):
    """
    指定したワークシートにDFをAPA7thスタイルの表（日本語表記）として書き込む
    - 因子分析表の場合：主因子負荷量を太字化、統計値行の上に区切り罫線を挿入
    """
    thick_side, thin_side = create_apa_border()
    font_family = "游ゴシック"
    
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
    
    # 因子分析の判定
    factor_cols_indices = [i + 2 for i, c in enumerate(df.columns) if "第" in str(c) and "因子" in str(c)]
    stat_row_keywords = ["因子寄与", "寄与率", "累積寄与率", "分散説明率", "負荷量二乗和"]
    first_stat_row_idx = None
    
    for r_idx, idx_val in enumerate(df.index):
        if any(kw in str(idx_val) for kw in stat_row_keywords):
            first_stat_row_idx = r_idx
            break
            
    # データ行の書き込み
    for r_idx, (idx_val, row) in enumerate(df.iterrows()):
        r = current_row + r_idx
        idx_cell = ws.cell(row=r, column=1, value=str(idx_val))
        
        is_stat_row = any(kw in str(idx_val) for kw in stat_row_keywords)
        idx_cell.font = Font(name=font_family, size=11, bold=is_stat_row)
        idx_cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # 因子分析の統計値行の境界線（観測変数と寄与率の間の区切り罫線）
        is_first_stat = (r_idx == first_stat_row_idx)
        
        # 観測変数行での主因子判定
        max_col_for_row = None
        if is_factor_analysis and not is_stat_row and factor_cols_indices:
            factor_vals = []
            for c_pos in factor_cols_indices:
                val = row.iloc[c_pos - 2]
                try:
                    factor_vals.append((abs(float(val)), c_pos))
                except Exception:
                    pass
            if factor_vals:
                max_col_for_row = max(factor_vals, key=lambda x: x[0])[1]
        
        for c_idx, val in enumerate(row, start=2):
            val_cell = ws.cell(row=r, column=c_idx)
            
            # 主因子の負荷量は太字
            is_bold = is_stat_row or (c_idx == max_col_for_row)
            val_cell.font = Font(name=font_family, size=11, bold=is_bold)
            
            if is_first_stat:
                val_cell.border = Border(top=thin_side)
                idx_cell.border = Border(top=thin_side)
                
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
        # 既存のtopボーダーを維持しつつbottomに太枠線を設定
        top_b = cell.border.top if cell.border else Side(border_style=None)
        cell.border = Border(top=top_b, bottom=thick_side)
        
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


def add_image_to_sheet(ws, img_bytes, cell_location="H2"):
    """画像データ(bytes)をExcelシートに挿入"""
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
    全分析結果からAPAスタイルのExcelワークブックを構築する。
    すべての分析手法について、Excel上で自在に編集可能なネイティブチャートを作成して挿入。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # デフォルトシート削除
    
    table_counter = 1
    for item in analysis_results:
        sheet_name = item.get("sheet_name", f"表_{table_counter}")
        sheet_name = sheet_name[:31]  # 31文字制限
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True
        
        df = item.get("df")
        title = item.get("title", "分析結果")
        note = item.get("note", None)
        fig_bytes = item.get("fig_bytes", None)
        extra_df = item.get("extra_df", None)
        extra_title = item.get("extra_title", "補助データ")
        
        data_info = None
        has_native_chart = False
        
        is_factor = "因子分析" in sheet_name
        if df is not None and not df.empty:
            data_info, next_row = format_apa_table(
                ws, start_row=2, title=title, df=df, note=note, 
                table_num=table_counter, is_factor_analysis=is_factor
            )
            table_counter += 1
            
            try:
                # 1. 単純集計: 棒グラフ
                if "単純集計" in sheet_name:
                    chart = BarChart()
                    chart.type = "col"
                    chart.style = 10
                    chart.title = f"{title}"
                    chart.y_axis.title = "度数 (N)"
                    chart.x_axis.title = "カテゴリ"
                    chart.legend = None
                    chart.width, chart.height = 16, 10
                    
                    data_ref = Reference(ws, min_col=2, min_row=data_info["header_row"], max_col=2, max_row=data_info["data_end_row"])
                    cats_ref = Reference(ws, min_col=1, min_row=data_info["data_start_row"], max_row=data_info["data_end_row"])
                    chart.add_data(data_ref, titles_from_data=True)
                    chart.set_categories(cats_ref)
                    ws.add_chart(chart, "H2")
                    has_native_chart = True

                # 2. 基本統計量 / ヒストグラム分布: 縦棒グラフ
                elif "分布_" in sheet_name or "ヒストグラム" in sheet_name:
                    chart = BarChart()
                    chart.type = "col"
                    chart.style = 10
                    chart.title = f"{title}"
                    chart.y_axis.title = "度数 (N)"
                    chart.x_axis.title = "階級 (区間)"
                    chart.legend = None
                    chart.gapWidth = 10
                    chart.width, chart.height = 16, 10
                    
                    data_ref = Reference(ws, min_col=2, min_row=data_info["header_row"], max_col=2, max_row=data_info["data_end_row"])
                    cats_ref = Reference(ws, min_col=1, min_row=data_info["data_start_row"], max_row=data_info["data_end_row"])
                    chart.add_data(data_ref, titles_from_data=True)
                    chart.set_categories(cats_ref)
                    ws.add_chart(chart, "H2")
                    has_native_chart = True

                # 3. クロス集計 / 構成比: 100%積み上げ棒グラフ
                elif "構成比" in sheet_name:
                    chart = BarChart()
                    chart.type = "col"
                    chart.grouping = "stacked"
                    chart.overlap = 100
                    chart.title = f"{title}"
                    chart.y_axis.title = "構成比 (%)"
                    chart.x_axis.title = "グループ / カテゴリ"
                    chart.width, chart.height = 16, 10
                    
                    data_ref = Reference(ws, min_col=2, min_row=data_info["header_row"], max_col=data_info["num_cols"], max_row=data_info["data_end_row"])
                    cats_ref = Reference(ws, min_col=1, min_row=data_info["data_start_row"], max_row=data_info["data_end_row"])
                    chart.add_data(data_ref, titles_from_data=True)
                    chart.set_categories(cats_ref)
                    ws.add_chart(chart, "H2")
                    has_native_chart = True

                # 4. 独立t検定 & 一元配置分散分析: 各従属変数の群別平均値棒グラフ
                elif ("t検定" in sheet_name or "分散分析" in sheet_name) and "二元配置" not in sheet_name and "対応あり" not in sheet_name:
                    m_sd_cols = [c for c in df.columns if "M (SD)" in str(c)]
                    if m_sd_cols:
                        chart_start_base = data_info["data_end_row"] + 4
                        ws.cell(row=chart_start_base - 1, column=1, value="[ グラフ用平均値データ ]").font = Font(name="游ゴシック", size=10, bold=True, italic=True)
                        
                        num_groups = len(m_sd_cols)
                        group_names = [str(c).replace(" M (SD)", "") for c in m_sd_cols]
                        
                        for r_idx, (dep_var, row) in enumerate(df.iterrows()):
                            curr_start_row = chart_start_base + (r_idx * 4)
                            ws.cell(row=curr_start_row, column=1, value="従属変数")
                            for c_idx, grp_name in enumerate(group_names, start=2):
                                ws.cell(row=curr_start_row, column=c_idx, value=grp_name)
                                
                            curr_data_row = curr_start_row + 1
                            ws.cell(row=curr_data_row, column=1, value=str(dep_var))
                            for c_idx, col_name in enumerate(m_sd_cols, start=2):
                                val_str = str(row[col_name])
                                try:
                                    mean_val = float(val_str.split("(")[0].strip())
                                    ws.cell(row=curr_data_row, column=c_idx, value=mean_val)
                                except Exception:
                                    ws.cell(row=curr_data_row, column=c_idx, value=0.0)
                                    
                            chart_cell = f"H{2 + r_idx * 16}"
                            chart = BarChart()
                            chart.type = "col"
                            chart.style = 10
                            chart.title = f"{dep_var} - 平均値比較"
                            chart.y_axis.title = "平均値 (M)"
                            chart.x_axis.title = "グループ"
                            chart.width, chart.height = 15, 9.5
                            
                            data_ref = Reference(ws, min_col=2, min_row=curr_start_row, max_col=num_groups + 1, max_row=curr_data_row)
                            cats_ref = Reference(ws, min_col=1, min_row=curr_data_row, max_row=curr_data_row)
                            chart.add_data(data_ref, titles_from_data=True)
                            chart.set_categories(cats_ref)
                            ws.add_chart(chart, chart_cell)
                            
                        has_native_chart = True

                # 5. 対応のあるt検定: 条件間変化の折れ線グラフ
                elif "対応ありt" in sheet_name:
                    # M (SD) 列の抽出
                    m_sd_cols = [c for c in df.columns if "M (SD)" in str(c)]
                    if len(m_sd_cols) >= 2:
                        helper_start = data_info["data_end_row"] + 3
                        ws.cell(row=helper_start, column=1, value="条件名")
                        ws.cell(row=helper_start, column=2, value="平均値 (M)")
                        
                        for c_idx, col_name in enumerate(m_sd_cols):
                            val_str = str(df.iloc[0][col_name])
                            try:
                                m_val = float(val_str.split("(")[0].strip())
                            except Exception:
                                m_val = 0.0
                            cond_label = str(col_name).replace(" M (SD)", "")
                            ws.cell(row=helper_start + c_idx + 1, column=1, value=cond_label)
                            ws.cell(row=helper_start + c_idx + 1, column=2, value=m_val)
                            
                        chart = LineChart()
                        chart.title = f"{title}"
                        chart.style = 13
                        chart.y_axis.title = "平均値 (M)"
                        chart.x_axis.title = "測定条件"
                        chart.legend = None
                        chart.width, chart.height = 15, 9.5
                        
                        data_ref = Reference(ws, min_col=2, min_row=helper_start, max_row=helper_start + len(m_sd_cols))
                        cats_ref = Reference(ws, min_col=1, min_row=helper_start + 1, max_row=helper_start + len(m_sd_cols))
                        chart.add_data(data_ref, titles_from_data=True)
                        chart.set_categories(cats_ref)
                        ws.add_chart(chart, "H2")
                        has_native_chart = True

                # 6. 二元配置分散分析: 交互作用折れ線グラフ
                elif "二元配置ANOVA" in sheet_name and extra_df is not None:
                    # extra_df (セル別平均値表) から平均値を抽出してヘルパー表を作成
                    helper_start = next_row + len(extra_df) + 6
                    ws.cell(row=helper_start - 1, column=1, value="[ 交互作用プロット用平均値データ ]").font = Font(name="游ゴシック", size=10, bold=True, italic=True)
                    
                    factor1_name = str(extra_df.index.name).split("\\")[0].strip() if "\\" in str(extra_df.index.name) else "要因1"
                    ws.cell(row=helper_start, column=1, value=factor1_name)
                    
                    for c_i, col_name in enumerate(extra_df.columns, start=2):
                        ws.cell(row=helper_start, column=c_i, value=str(col_name))
                        
                    for r_i, (idx_name, r_row) in enumerate(extra_df.iterrows()):
                        curr_r = helper_start + r_i + 1
                        ws.cell(row=curr_r, column=1, value=str(idx_name))
                        for c_i, col_name in enumerate(extra_df.columns, start=2):
                            v_str = str(r_row[col_name])
                            try:
                                m_val = float(v_str.split("(")[0].strip())
                            except Exception:
                                m_val = 0.0
                            ws.cell(row=curr_r, column=c_i, value=m_val)
                            
                    chart = LineChart()
                    chart.title = f"二元配置ANOVA 交互作用プロット"
                    chart.style = 13
                    chart.y_axis.title = "平均値 (M)"
                    chart.x_axis.title = factor1_name
                    chart.width, chart.height = 16, 10
                    
                    data_ref = Reference(ws, min_col=2, min_row=helper_start, max_col=len(extra_df.columns) + 1, max_row=helper_start + len(extra_df))
                    cats_ref = Reference(ws, min_col=1, min_row=helper_start + 1, max_row=helper_start + len(extra_df))
                    chart.add_data(data_ref, titles_from_data=True)
                    chart.set_categories(cats_ref)
                    ws.add_chart(chart, "H2")
                    has_native_chart = True

                # 7. 重回帰分析: 標準化偏回帰係数 (Beta) 比較棒グラフ
                elif "回帰分析" in sheet_name:
                    beta_col = None
                    for c_i, col in enumerate(df.columns, start=2):
                        if "標準化" in str(col) or "β" in str(col) or "Beta" in str(col):
                            beta_col = c_i
                            break
                    if beta_col:
                        # 切片を除く説明変数行のみで棒グラフを作成
                        start_r = data_info["data_start_row"]
                        if "切片" in str(df.index[0]) or "定数項" in str(df.index[0]):
                            start_r += 1
                        if start_r <= data_info["data_end_row"]:
                            chart = BarChart()
                            chart.type = "bar"  # 横棒
                            chart.style = 10
                            chart.title = "標準化偏回帰係数 (β)"
                            chart.x_axis.title = "標準化係数 (β)"
                            chart.y_axis.title = "説明変数"
                            chart.legend = None
                            chart.width, chart.height = 15, 9.5
                            
                            data_ref = Reference(ws, min_col=beta_col, min_row=data_info["header_row"], max_col=beta_col, max_row=data_info["data_end_row"])
                            cats_ref = Reference(ws, min_col=1, min_row=start_r, max_row=data_info["data_end_row"])
                            chart.add_data(data_ref, titles_from_data=True)
                            chart.set_categories(cats_ref)
                            ws.add_chart(chart, "H2")
                            has_native_chart = True

                # 8. ロジスティック回帰: オッズ比 (OR) 比較横棒グラフ
                elif "ロジスティック" in sheet_name:
                    or_col = None
                    for c_i, col in enumerate(df.columns, start=2):
                        if "オッズ比" in str(col) or "OR" in str(col):
                            or_col = c_i
                            break
                    if or_col:
                        start_r = data_info["data_start_row"]
                        if "切片" in str(df.index[0]) or "定数項" in str(df.index[0]):
                            start_r += 1
                        if start_r <= data_info["data_end_row"]:
                            chart = BarChart()
                            chart.type = "bar"
                            chart.style = 10
                            chart.title = "オッズ比 (Odds Ratio)"
                            chart.x_axis.title = "オッズ比 (OR)"
                            chart.y_axis.title = "説明変数"
                            chart.legend = None
                            chart.width, chart.height = 15, 9.5
                            
                            data_ref = Reference(ws, min_col=or_col, min_row=data_info["header_row"], max_col=or_col, max_row=data_info["data_end_row"])
                            cats_ref = Reference(ws, min_col=1, min_row=start_r, max_row=data_info["data_end_row"])
                            chart.add_data(data_ref, titles_from_data=True)
                            chart.set_categories(cats_ref)
                            ws.add_chart(chart, "H2")
                            has_native_chart = True

                # 9. 因子分析: スクリープロット (固有値推移折れ線グラフ)
                elif "因子分析" in sheet_name:
                    scree_df = item.get("scree_df", None)
                    if scree_df is not None and not scree_df.empty:
                        helper_start = data_info["data_end_row"] + 4
                        ws.cell(row=helper_start - 1, column=1, value="[ スクリープロット用データ (固有値) ]").font = Font(name="游ゴシック", size=10, bold=True, italic=True)
                        
                        ws.cell(row=helper_start, column=1, value="因子番号")
                        ws.cell(row=helper_start, column=2, value="固有値")
                        ws.cell(row=helper_start, column=3, value="カイザー基準 (1.0)")
                        
                        for s_i, (f_name, s_row) in enumerate(scree_df.iterrows()):
                            c_row = helper_start + s_i + 1
                            ws.cell(row=c_row, column=1, value=str(f_name))
                            ws.cell(row=c_row, column=2, value=float(s_row["固有値"]))
                            ws.cell(row=c_row, column=3, value=float(s_row["基準値 (1.0)"]))
                            
                        chart = LineChart()
                        chart.title = "スクリープロット (固有値の推移)"
                        chart.style = 13
                        chart.y_axis.title = "固有値 (Eigenvalue)"
                        chart.x_axis.title = "因子番号"
                        chart.width, chart.height = 16, 10
                        
                        data_ref = Reference(ws, min_col=2, min_row=helper_start, max_col=3, max_row=helper_start + len(scree_df))
                        cats_ref = Reference(ws, min_col=1, min_row=helper_start + 1, max_row=helper_start + len(scree_df))
                        chart.add_data(data_ref, titles_from_data=True)
                        chart.set_categories(cats_ref)
                        ws.add_chart(chart, "H2")
                        has_native_chart = True

                # 10. 信頼性分析: 項目削除時αの比較横棒グラフ
                elif "信頼性" in sheet_name:
                    alpha_col = None
                    for c_i, col in enumerate(df.columns, start=2):
                        if "項目削除時" in str(col) or "alpha" in str(col).lower():
                            alpha_col = c_i
                            break
                    if alpha_col:
                        chart = BarChart()
                        chart.type = "bar"
                        chart.style = 10
                        chart.title = "項目削除時 クロンバックのα係数"
                        chart.x_axis.title = "α係数"
                        chart.y_axis.title = "尺度項目"
                        chart.legend = None
                        chart.width, chart.height = 15, max(8.0, len(df) * 0.6)
                        
                        data_ref = Reference(ws, min_col=alpha_col, min_row=data_info["header_row"], max_col=alpha_col, max_row=data_info["data_end_row"])
                        cats_ref = Reference(ws, min_col=1, min_row=data_info["data_start_row"], max_row=data_info["data_end_row"])
                        chart.add_data(data_ref, titles_from_data=True)
                        chart.set_categories(cats_ref)
                        ws.add_chart(chart, "H2")
                        has_native_chart = True

            except Exception:
                has_native_chart = False
                
        # 補助テーブル (extra_df) の出力 (二元配置ANOVAなど)
        if extra_df is not None and not extra_df.empty:
            extra_start_row = next_row + 2
            format_apa_table(ws, start_row=extra_start_row, title=extra_title, df=extra_df, note=None, table_num=table_counter)
            table_counter += 1

        # ネイティブグラフが作成できなかった場合、または相関等の場合は画像を配置
        if fig_bytes is not None and not has_native_chart:
            add_image_to_sheet(ws, fig_bytes, cell_location="H2")
            
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
