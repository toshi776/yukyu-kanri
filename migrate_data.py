#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
有給データ移行スクリプト
各事業所のExcelファイルから利用者データを抽出してCSVに出力
"""

import pandas as pd
import os
import re
import csv
from pathlib import Path

def extract_number_from_filename(filename):
    """ファイル名から番号を抽出"""
    match = re.match(r'^(\d+)', filename)
    return int(match.group(1)) if match else 0

def get_remaining_vacation_days(excel_file):
    """Excelファイルの管理簿タブのJ列最下行から残有給日数を取得"""
    try:
        # まず管理簿シートを探す
        xl_file = pd.ExcelFile(excel_file)
        sheet_name = None
        
        # 管理簿タブを探す
        for sheet in xl_file.sheet_names:
            if '管理簿' in sheet:
                sheet_name = sheet
                break
        
        # 管理簿タブが見つからない場合は最初のシートを使用
        if sheet_name is None:
            sheet_name = xl_file.sheet_names[0]
        
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # J列（10番目の列、0-indexed で9）の最後の値を取得
        if len(df.columns) > 9:  # J列が存在するか確認
            j_column = df.iloc[:, 9]  # J列を取得
            # NaNでない最後の値を取得
            non_nan_values = j_column.dropna()
            if len(non_nan_values) > 0:
                last_value = non_nan_values.iloc[-1]
                # 数値の場合のみ返す
                if pd.notnull(last_value) and isinstance(last_value, (int, float)):
                    return int(last_value)
        return 0
    except Exception as e:
        print(f"エラー: {excel_file} - {e}")
        return 0

def extract_name_from_filename(filename):
    """ファイル名から氏名を抽出"""
    # 拡張子を除去
    name_part = filename.replace('.xlsx', '')
    # 番号とピリオドを除去
    match = re.match(r'^\d+\.(.+)', name_part)
    return match.group(1) if match else name_part

def process_office(office_path, office_code):
    """事業所のデータを処理"""
    results = []
    
    # 数字で始まるExcelファイルのみを対象
    excel_files = [f for f in os.listdir(office_path) 
                   if f.endswith('.xlsx') and re.match(r'^\d+', f)]
    
    # ファイル名の数字順にソート
    excel_files.sort(key=extract_number_from_filename)
    
    for filename in excel_files:
        file_path = os.path.join(office_path, filename)
        
        # ファイル名から情報を抽出
        file_number = extract_number_from_filename(filename)
        user_name = extract_name_from_filename(filename)
        
        # 利用者番号を生成（例：R01, P07）
        user_id = f"{office_code}{file_number:02d}"
        
        # 残有給日数を取得
        remaining_days = get_remaining_vacation_days(file_path)
        
        results.append({
            '利用者番号': user_id,
            '利用者名': user_name,
            '残有給日数': remaining_days,
            '備考': ''
        })
        
        print(f"処理完了: {user_id} - {user_name} - {remaining_days}日")
    
    return results

def main():
    """メイン処理"""
    base_path = "エクセルデータ"
    
    # 事業所の設定
    offices = {
        'ライズ': 'R',
        'パロン': 'P', 
        'シエル': 'S',
        'EBISU': 'E'
    }
    
    all_data = []
    
    print("データ移行を開始します...")
    
    for office_name, office_code in offices.items():
        office_path = os.path.join(base_path, office_name)
        
        if os.path.exists(office_path):
            print(f"\n{office_name}の処理を開始...")
            office_data = process_office(office_path, office_code)
            all_data.extend(office_data)
            print(f"{office_name}完了: {len(office_data)}名")
        else:
            print(f"警告: {office_path} が見つかりません")
    
    # CSVファイルに出力
    output_file = "有給管理_移行データ.csv"
    
    with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
        fieldnames = ['利用者番号', '利用者名', '残有給日数', '備考']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for row in all_data:
            writer.writerow(row)
    
    print(f"\n移行完了!")
    print(f"出力ファイル: {output_file}")
    print(f"総データ数: {len(all_data)}名")
    
    # 事業所別の集計
    for office_name, office_code in offices.items():
        count = len([d for d in all_data if d['利用者番号'].startswith(office_code)])
        print(f"{office_name}: {count}名")

if __name__ == "__main__":
    main()