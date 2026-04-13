import openpyxl
from datetime import datetime
import os

def create_excel_report():
    # 保存先ディレクトリの作成
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)
    
    filename = f"{output_dir}/report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # ワークブックの作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "AutoReport"

    # ヘッダーとデータの書き込み
    data = [
        ["実行日時", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["ステータス", "Success"],
        ["項目", "値"],
        ["Python Version", "3.11"],
        ["Library", "openpyxl"]
    ]

    for row in data:
        ws.append(row)

    # 書式設定（少し本気感を出します）
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill

    # 保存
    wb.save(filename)
    print(f"File created: {filename}")

if __name__ == "__main__":
    create_excel_report()
