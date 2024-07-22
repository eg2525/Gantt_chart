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
    df['開始予定日'] = pd.to_datetime(df['開始予定日'])
    df['終了予定日'] = pd.to_datetime(df['終了予定日'])

    # '工程'列のユニークな値を取得し、チェックボックスを作成
    unique_tasks = df['工程'].unique()
    selected_tasks = st.multiselect('表示する工程を選択してください', unique_tasks, default=list(unique_tasks))

    if st.button('ガントチャート作成'):
        # カレンダーの開始日と終了日を設定
        calendar_start = df['開始予定日'].min()
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

        # 週ごとのヘッダー行を作成
        for week_start in calendar_days:
            week_end = week_start + pd.Timedelta(days=6)
            week_row = ws.max_row + 2

            # ヘッダー行を作成
            header_cell = ws.cell(row=week_row, column=1, value='作業名')
            header_cell.border = thin_border
            header_cell.alignment = Alignment(horizontal='center', vertical='center')
            for i, day in enumerate(pd.date_range(start=week_start, end=week_end), start=2):
                col_letter = get_column_letter(i)
                cell = ws.cell(row=week_row, column=i, value=day.strftime('%Y-%m-%d'))
                cell.font = Font(bold=True)  # 太文字
                cell.border = thin_border  # 罫線を追加
                ws.column_dimensions[col_letter].width = 15  # 列幅を設定

            # 各タスクの行を作成
            task_row = week_row + 1
            for index, row in df.iterrows():
                # '工程'列のフィルタリングを追加
                if row['工程'] in selected_tasks and row['終了予定日'] >= week_start and row['開始予定日'] <= week_end:
                    task_cell = ws.cell(row=task_row, column=1, value=row['作業名'])
                    task_cell.border = thin_border
                    task_cell.alignment = Alignment(horizontal='center', vertical='center')
                    task_days = (row['終了予定日'] - row['開始予定日']).days + 1
                    part_length = task_days // 3
                    part_remainder = task_days % 3
                    for i, day in enumerate(pd.date_range(start=week_start, end=week_end), start=2):
                        if row['開始予定日'] <= day <= row['終了予定日']:
                            cell = ws.cell(row=task_row, column=i)
                            cell.border = thin_border  # 罫線を追加
                            # 最終日に「提出期限」と表示
                            if day == row['終了予定日']:
                                cell.value = '提出期限'
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            # 色の段階を設定
                            if day < row['開始予定日'] + pd.Timedelta(days=part_length + (1 if part_remainder > 0 else 0)):
                                fill_color = colors[0]
                            elif day < row['開始予定日'] + pd.Timedelta(days=2 * part_length + (2 if part_remainder > 1 else 1)):
                                fill_color = colors[1]
                            else:
                                fill_color = colors[2]
                            cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
                    task_row += 1

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
