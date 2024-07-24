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
    df = df.drop(['レコードの開始行','標題','担当社員','担当者','レコード番号','更新者', '作成者', '更新日時', '作成日時', 'ステータス','プロジェクトコード', '関連者',
             '削除依頼', '削除依頼者', '削除依頼理由',
           '削除依頼日', 'ルックアップ(被相続人)', '被相続人:顧客コード', '顧客名&ﾌﾘｶﾞﾅ','サーバーアドレス', 'ルックアップ(相続人)', '相続人:顧客名', '相続人:顧客コード', '作業予定者','総予定工数', '工程リスト',
           '解約事由', '解約日',
           ], axis=1)

    # データを日付形式に変換
    df['開始予定日'] = pd.to_datetime(df['開始予定日'], errors='coerce')
    df['終了予定日'] = pd.to_datetime(df['終了予定日'], errors='coerce')

    # '工程'列のユニークな値を取得し、チェックボックスを作成
    unique_tasks = df['工程'].unique()
    selected_tasks = st.multiselect('表示する工程を選択してください', unique_tasks, default=list(unique_tasks))

    if st.button('ガントチャート作成'):
        # 選択された工程に関連する作業名をフィルタリング
        filtered_df = df[df['工程'].isin(selected_tasks)]

        # 欠損値を含む行を削除
        filtered_df = filtered_df.dropna(subset=['開始予定日', '終了予定日'])

        # カレンダーの開始日と終了日を設定
        calendar_start = filtered_df['開始予定日'].min()
        calendar_end = filtered_df['終了予定日'].max()
        calendar_days = pd.date_range(start=calendar_start, end=calendar_end)

        # Excelファイルを作成
        wb = Workbook()
        ws = wb.active
        ws.title = 'ガントチャート'

        # 色の設定
        colors = ['0DACDC', 'ECDA2F', 'F11D1A']
        # 罫線の設定
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 日付を左一列に連続的に並べる
        ws.cell(row=1, column=1, value='日付')
        for i, day in enumerate(calendar_days, start=2):
            cell = ws.cell(row=i, column=1, value=day.strftime('%Y-%m-%d'))
            cell.border = thin_border  # 罫線を追加
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # 各作業名を列ヘッダーに設定
        task_columns = {task: idx+2 for idx, task in enumerate(filtered_df['作業名'].unique())}
        for task, col in task_columns.items():
            cell = ws.cell(row=1, column=col, value=task)
            cell.font = Font(bold=True)  # 太文字
            cell.border = thin_border  # 罫線を追加
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col)].width = 20  # 列幅を設定

        # 各作業のセルに色をつける
        for index, row in filtered_df.iterrows():
            task_col = task_columns[row['作業名']]
            start_idx = (row['開始予定日'] - calendar_start).days + 2
            end_idx = (row['終了予定日'] - calendar_start).days + 2
            task_days = int(end_idx - start_idx + 1)  # 整数に変換
            part_length = task_days // 3
            part_remainder = task_days % 3

            for i in range(start_idx, end_idx + 1):
                cell = ws.cell(row=i, column=task_col)
                cell.border = thin_border  # 罫線を追加
                # 色の段階を設定
                if i < start_idx + part_length + (1 if part_remainder > 0 else 0):
                    fill_color = colors[0]
                elif i < start_idx + 2 * part_length + (2 if part_remainder > 1 else 1):
                    fill_color = colors[1]
                else:
                    fill_color = colors[2]
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                if i == end_idx:
                    cell.value = '提出期限'
                    cell.alignment = Alignment(horizontal='center', vertical='center')

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
