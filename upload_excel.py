import pandas as pd
import requests
import json

# ==================== 設定區 ====================
EXCEL_FILE = "2026人票總表.xlsx"  # 你的 Excel 檔名
CAMP_YEAR = "2026"           # 📅 這一屆營隊的年份 (可自行修改，如 "2027")
# ⚠️ 請替換成你的 Firebase Realtime Database 網址（注意：路徑加入了年份變數，記得結尾要有 /.json）
FIREBASE_URL = f"https://cijiguozhongkuailechengz-26e8d-default-rtdb.asia-southeast1.firebasedatabase.app/{CAMP_YEAR}.json"
# ===============================================

def upload_data():
    try:
        # 🟢 【全新修改】動態詢問使用者要抓取幾列工作人員資料
        print("--------------------------------------------------")
        user_input = input("🔢 請輸入工作人員預計讀取的列數上限\n(提示: 直接按 [Enter] 將預設截取前 100 列): ").strip()
        
        # 防呆驗證：確認使用者輸入的是否為純數字，如果不是則自動套用 100 列
        if user_input.isdigit():
            STAFF_ROW_LIMIT = int(user_input)
        else:
            STAFF_ROW_LIMIT = 100
            print("💡 偵測到未輸入或格式非純數字，自動套用預設值：前 100 列。")
        print("--------------------------------------------------")

        print(f"🔄 正在讀取 Excel 檔案，準備匯入 {CAMP_YEAR} 年份資料，並過濾 {CAMP_YEAR} 精簡版欄位...")
        excel_file_obj = pd.ExcelFile(EXCEL_FILE)
        sheet_names = excel_file_obj.sheet_names
        print(f"📋 偵測到 Excel 內的分頁有: {sheet_names}")
        
        # 自動模糊匹配分頁名稱
        camper_sheet = [s for s in sheet_names if "學員" in s]
        staff_sheet = [s for s in sheet_names if "工作人員" in s or "工人" in s or "staff" in s.lower()]
        
        # ✨【欄位精簡】工作人員標準保留欄位已全面剔除：安單
        CAMPER_TARGET_COLS = ["組別", "身分", "姓名", "性別", "報到", "發票", "同意書", "繳費狀況", "報到備註", "待確認事項"]
        STAFF_TARGET_COLS = ["組別", "報到", "職務", "姓名", "生理性別", "身分別"]

        # 1. 讀取與處理【學員】資料
        if camper_sheet:
            print(f"📖 匯入學員分頁：[{camper_sheet[0]}]...")
            camper_df = pd.read_excel(EXCEL_FILE, sheet_name=camper_sheet[0])
            camper_df.columns = camper_df.columns.astype(str).str.strip()
            
            # 自動過濾重複標題列
            if "組別" in camper_df.columns:
                camper_df = camper_df[camper_df["組別"].astype(str).str.strip() != "組別"]
        else:
            print("⚠️ 提示：找不到『學員』分頁，建立空外殼。")
            camper_df = pd.DataFrame(columns=CAMPER_TARGET_COLS)

        # 2. 讀取與處理【工作人員】資料
        if staff_sheet:
            print(f"📖 匯入工作人員分頁：[{staff_sheet[0]}]...")
            staff_df = pd.read_excel(EXCEL_FILE, sheet_name=staff_sheet[0], header=1)
            staff_df.columns = staff_df.columns.astype(str).str.strip()
            
            # 自動修復合併儲存格
            if "組別" in staff_df.columns:
                staff_df["組別"] = staff_df["組別"].apply(lambda x: None if pd.isna(x) or str(x).strip() in ["", "nan"] else str(x).strip())
                staff_df["組別"] = staff_df["組別"].ffill()
                print("⚡ 已成功自動修復【工作人員】的合併儲存格組別欄位！")
            
            print(f"✂️ 已啟動防擾機制：工作人員僅精準截取前 {STAFF_ROW_LIMIT} 列資料。")
            staff_df = staff_df.head(STAFF_ROW_LIMIT)
        else:
            print("⚠️ 提示：找不到『工作人員』分頁，建立空外殼。")
            staff_df = pd.DataFrame(columns=STAFF_TARGET_COLS)
        
        # 3. 兼容性處理：合併生理性別
        if "生理性別" not in staff_df.columns and not staff_df.empty:
            if "生理" in staff_df.columns: staff_df["生理性別"] = staff_df["生理"]
            elif "性別" in staff_df.columns: staff_df["生理性別"] = staff_df["性別"]
            else: staff_df["生理性別"] = ""

        # 4. 補齊確認欄位
        for col in CAMPER_TARGET_COLS:
            if col not in camper_df.columns: camper_df[col] = ""
        for col in STAFF_TARGET_COLS:
            if col not in staff_df.columns: staff_df[col] = ""
            
        camper_df = camper_df[CAMPER_TARGET_COLS].copy()
        staff_df = staff_df[STAFF_TARGET_COLS].copy()
        
        # 5. 處理日期與空值型態
        for df in [camper_df, staff_df]:
            for col in df.columns:
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    df[col] = df[col].dt.strftime('%Y-%m-%d')
                else:
                    df[col] = df[col].apply(lambda x: x.strftime('%Y-%m-%d') if hasattr(x, 'strftime') else x)

        # 精準紀錄學員繳費原本就為空的位置
        is_payment_null = camper_df["繳費狀況"].isna() | (camper_df["繳費狀況"].astype(str).str.strip() == "")

        camper_df = camper_df.fillna("")
        staff_df = staff_df.fillna("")

        # 剔除姓名為空的整列資料
        camper_df = camper_df[camper_df["姓名"].astype(str).str.strip() != ""]
        staff_df = staff_df[staff_df["姓名"].astype(str).str.strip() != ""]

        # 將紀錄為空的學員繳費欄位還原為 None (null)
        camper_df["繳費狀況"] = camper_df["繳費狀況"].astype(object)
        camper_df.loc[is_payment_null, "繳費狀況"] = None

        # 6. 現場全新空白備註初始化
        camper_df["備註"] = ""
        staff_df["備註"] = ""

        # 7. 轉換布林值（打勾方塊）
        # ✨ 升級防呆：不論 Excel 欄位填 True、1、v、V、o、O、是 還是 已交，通通都能精準轉為 True
        for col in ["報到", "發票", "同意書"]:
            if col in camper_df.columns:
                camper_df[col] = camper_df[col].apply(
                    lambda x: True if str(x).upper().strip() in ["TRUE", "1.0", "1", "V", "O", "是", "已交", "打勾"] else False
                )
                
        if "報到" in staff_df.columns:
            staff_df["報到"] = staff_df["報到"].apply(
                lambda x: True if str(x).upper().strip() in ["TRUE", "1.0", "1", "V", "O", "是", "已交", "打勾"] else False
            )

        # 8. 打包轉為字典
        camper_list = camper_df.to_dict(orient="records")
        staff_list = staff_df.to_dict(orient="records")

        print(f"📊 【資料分析結果】安全清洗後，成功匯入 {len(camper_list)} 筆學員、 {len(staff_list)} 筆工作人員。")

        payload = {
            "campers": camper_list,
            "staff": staff_list
        }
        
        print("🚀 正在將無個資安全數據推送到 Firebase 雲端...")
        response = requests.put(FIREBASE_URL, json=payload)
        
        if response.status_code == 200:
            print(f"🎉 【完美成功】{CAMP_YEAR} 年度無個資純淨版名單已同步至 Firebase！")
        else:
            print(f"❌ 【失敗】HTTP {response.status_code}: {response.text}")
            
    except Exception as e:
        print(f"❌ 發生未知錯誤: {e}")

if __name__ == "__main__":
    upload_data()