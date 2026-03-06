import pandas as pd
import numpy as np
import io

def load_daily_data(file, val_name="count"):
    """
    Load daily order/terminal data from the user-provided complex Excel format.
    Assumes:
    - Row 1: "年月日" prefixes
    - Row 2: "x月"
    - Row 3: Day numbers (1, 2, 3...)
    - Row 4: Column names like "販売店名", "受付担当者", "オーダ数" / "端末販売数"
    - Row 5+: Data
    
    Returns a long-format DataFrame: [shop, name, date, count]
    """
    try:
        # Read Excel skipping the first 3 lines (0-indexed 0,1,2).
        # We want the 4th row (index 3) to be the header.
        df_raw = pd.read_excel(file, header=3)
        
        # We need "販売店名" and "受付担当者". The rest are metric columns per day.
        # But wait, looking at the image:
        # Col A: 販売店名
        # Col B: 受付担当者
        # Col C, D, E... are "オーダ数" or "端末販売数" corresponding to the days in row 3.
        
        # Because the metric names are generic ("オーダ数"), they might be duplicated.
        # Let's read without header to get the exact day numbers from row 3.
        file.seek(0)
        df_full = pd.read_excel(file, header=None)
        
        # Row index 2 (3rd row) contains the days (1, 2, 3, 4...) starting from Col C (index 2)
        days_row = df_full.iloc[2].values
        
        # Row index 3 (4th row) contains the headers "販売店名", "受付担当者", "オーダ数"...
        headers_row = df_full.iloc[3].values
        
        # Extract data rows (index 4 onwards)
        data_rows = df_full.iloc[4:].copy()
        
        records = []
        for _, row in data_rows.iterrows():
            shop = row.iloc[0]
            name = row.iloc[1]
            
            # Skip if name is empty/nan
            if pd.isna(name) or str(name).strip() == "":
                continue
                
            shop_str = str(shop).strip() if pd.notna(shop) else ""
            name_str = str(name).strip()
            
            # Iterate through day columns
            for col_idx in range(2, len(row)):
                day_val = days_row[col_idx]
                if pd.isna(day_val):
                    continue
                    
                count_val = row.iloc[col_idx]
                # If it's a number, record it
                if pd.notna(count_val) and isinstance(count_val, (int, float)):
                    records.append({
                        "shop": shop_str,
                        "name": name_str,
                        "date": int(day_val),
                        val_name: int(count_val)
                    })
                    
        return pd.DataFrame(records), None
        
    except Exception as e:
        return None, f"Excelファイルの読み込みに失敗しました: {str(e)}"
