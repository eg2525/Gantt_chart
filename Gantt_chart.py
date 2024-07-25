import streamlit as st
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from openpyxl.utils import get_column_letter
import io

# Streamlitアプリの開始
st.title('ガントチャート作成アプリ')

# ファイルアップロード
uploaded_file = st.file_uploader("ファイルをアップロードしてください", type=["csv"])

if uploaded_file is not None:
    # データフレームを読み込む
    df = pd.read_csv(uploaded_file, encoding='cp932')

    # 不要な列を削除
    df = df.drop(['レコードの開始行', '標題', '担当社員', '担当者', 'レコード番号', '更新者', '作成者', '更新日時', '作成日時', 'ステータス', 'プロジェクトコード', '関連者',
                  '削除依頼', '削除依頼者', '削除依頼理由', '削除依頼日', 'ルックアップ(被相続人)', '被相続人:顧客コード', '顧客名&ﾌﾘｶﾞﾅ', 'サーバーアドレス', 'ルックアップ(相続人)',
                  '相続人:顧客名', '相続人:顧客コード', '作業予定者', '総予定工数', '工程リスト', '解約事由', '解約日'], axis=1)

    # データを日付形式に変換
    df['開始予定日'] = pd.to_datetime(df['開始予定日'], errors='coerce')
    df['終了予定日'] = pd.to_datetime(df['終了予定日'], errors='coerce')
    df['相続開始日'] = pd.to_datetime(df['相続開始日'], errors='coerce')

    # '工程'列のユニークな値を取得し、チェックボックスを作成
    unique_tasks = df['工程'].unique()
    selected_tasks = st.multiselect('表示する工程を選択してください', unique_tasks, default=list(unique_tasks))

    if st.button('ガントチャート作成'):
        # 相続開始日をカレンダーの開始日とし、その週の月曜日に設定
        inheritance_start = df['相続開始日'].min()
        calendar_start = inheritance_start - pd.to_timedelta(inheritance_start.weekday(), unit='D')
        calendar_end = df['終了予定日'].max()
        calendar_days = pd.date_range(start=calendar_start, end=calendar_end, freq='W-MON')

        # Excelファイルを作成
        wb = Workbook()
        ws = wb.active
        ws.title = 'ガントチャート'

        # 色の設定
        colors = ['0DACDC', 'ECDA2F', 'F11D1A']
        # 罫線の設定
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 日付を横一列に連続的に並べる
        current_row = 1
        ws.cell(row=current_row, column=1, value='作業名')

        month_groups = calendar_days.to_period('M').unique()
        for month in month_groups:
            month_days = calendar_days[calendar_days.to_period('M') == month]
            for i, day in enumerate(month_days, start=2):
                cell = ws.cell(row=current_row, column=i, value=day.strftime('%Y-%m-%d'))
                cell.font = Font(bold=True)  # 太文字
                cell.border = thin_border  # 罫線を追加
                cell.alignment = Alignment(horizontal='center', vertical='center')
                ws.column_dimensions[get_column_letter(i)].width = 15  # 列幅を設定

            # 選択された工程に関連する作業名をフィルタリング
            filtered_df = df[df['工程'].isin(selected_tasks)]

            # 各作業名を行ヘッダーに設定
            task_rows = {task: idx + current_row + 1 for idx, task in enumerate(filtered_df['作業名'].unique())}
            for task, row in task_rows.items():
                cell = ws.cell(row=row, column=1, value=task)
                cell.font = Font(bold=True)  # 太文字
                cell.border = thin_border  # 罫線を追加
                cell.alignment = Alignment(horizontal='center', vertical='center')
                ws.row_dimensions[row].height = 20  # 行高さを設定

            # 各作業のセルに色をつける
            filtered_df = filtered_df.dropna(subset=['開始予定日', '終了予定日'])
            for index, row in filtered_df.iterrows():
                task_row = task_rows[row['作業名']]
                start_col = (row['開始予定日'] - calendar_start).days // 7 + 2
                end_col = (row['終了予定日'] - calendar_start).days // 7 + 2
                task_weeks = int(end_col - start_col + 1)  # 整数に変換
                part_length = task_weeks // 3
                part_remainder = task_weeks % 3

                for i in range(start_col, end_col + 1):
                    cell = ws.cell(row=task_row, column=i)
                    cell.border = thin_border  # 罫線を追加
                    # 色の段階を設定
                    if i < start_col + part_length + (1 if part_remainder > 0 else 0):
                        fill_color = colors[0]
                    elif i < start_col + 2 * part_length + (2 if part_remainder > 1 else 1):
                        fill_color = colors[1]
                    else:
                        fill_color = colors[2]
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                    if i == end_col:
                        cell.value = '提出期限'
                        cell.alignment = Alignment(horizontal='center', vertical='center')

            # 次の月に移動する前に2行の空行を追加
            current_row = max(task_rows.values()) + 2

        # 最初の2行が空行であれば削除
        if all(ws.cell(row=1, column=col).value is None for col in range(1, ws.max_column + 1)) and all(ws.cell(row=2, column=col).value is None for col in range(1, ws.max_column + 1)):
            ws.delete_rows(1, 2)

        # 列幅を自動調整（作業名の列は最低150px）
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # 列名を取得
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max((max_length + 2), 20) if column == 'A' else (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 20)

        # バイナリデータとしてExcelファイルを保存
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # ダウンロードボタンを表示
        st.download_button(label='ガントチャートをダウンロード', data=output, file_name='GanttChart.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
