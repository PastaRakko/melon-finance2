import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import numpy as np
import io
from github import Github

st.set_page_config(page_title="哈蜜瓜收支表", page_icon="🍈", layout="wide")
st.title("🍈哈蜜瓜收支與季繳管理表")

# ==========================================
# 1. 讀取 Excel 檔案
# ==========================================
@st.cache_data
def load_excel():
    df_trans = pd.read_excel("哈蜜瓜收支表.xlsx", sheet_name='收支明細')
    df_members = pd.read_excel("哈蜜瓜收支表.xlsx", sheet_name='季繳名單')
    
    # 自動清除所有欄位名稱前後的隱藏空白字元
    df_trans.columns = df_trans.columns.str.strip()
    df_members.columns = df_members.columns.str.strip()
   
    return df_trans, df_members

try:
    df_trans, df_members = load_excel()
except Exception as e:
    st.error("讀取 Excel 失敗，請確認檔案與分頁是否存在。")
    st.stop()

# ==========================================
# 2. 系統初始化 (將兩份資料都放入 session_state)
# ==========================================
if 'members_df' not in st.session_state:
    df_mem = df_members.dropna(subset=['姓名']).copy()
    
    # 確保日期格式正確
    df_mem['季繳開始日期'] = pd.to_datetime(df_mem['季繳開始日期'], errors='coerce')
    today = pd.to_datetime(datetime.today().date())
    days_diff = (today - df_mem['季繳開始日期']).dt.days
    
    df_mem['剩餘次數'] = np.where(
        df_mem['季繳開始日期'].notna(),
        10 - np.ceil(days_diff / 7),
        0 # 如果沒填開始日期，預設顯示為 0
    ).astype(int)
    
    st.session_state.members_df = df_mem[['姓名', '繳費日期', '季繳開始日期', '剩餘次數']]

if 'trans_df' not in st.session_state:
    df_t = df_trans.copy()
    # 清理收支明細 (將文字轉為數字，並去除資料內的逗號與可能殘留的符號)
    if '金額' in df_t.columns:
        df_t['金額'] = pd.to_numeric(df_t['金額'].astype(str).replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    if '日期' in df_t.columns:
        df_t['日期'] = pd.to_datetime(df_t['日期'], errors='coerce')
    st.session_state.trans_df = df_t

current_trans = st.session_state.trans_df

# ------------------------------------------
# 動態產生下拉選單
# ------------------------------------------
item_options = ["場地費", "臨打", "季繳", "羽毛球", "活動費", "其他"]
if '項目' in current_trans.columns:
    for x in current_trans['項目'].dropna().unique():
        if x not in item_options and str(x).strip() != "":
            item_options.append(x)

# 經手人選單
handler_options = ["櫃台", "妙", "齊"]
if '經手人' in current_trans.columns:
    for x in current_trans['經手人'].dropna().unique():
        if x not in handler_options and str(x).strip() != "":
            handler_options.append(x)

# ==========================================
# 左側邊欄：一鍵雲端存檔
# ==========================================
with st.sidebar:
    st.header("💾雲端同步存檔")
    st.info("💡完成資料修改後，請務必點擊下方按鈕以更新")
    
    if st.button("👆確認更新並存檔", use_container_width=True):
        with st.spinner("⌛正在上傳資料中，請稍候..."):
            try:
                # 1. 產生最新的 Excel 檔案內容到記憶體
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    st.session_state.trans_df.to_excel(writer, sheet_name='收支明細', index=False)
                    st.session_state.members_df.to_excel(writer, sheet_name='季繳名單', index=False)
                excel_data = output.getvalue()
                
                # 2. 呼叫 Secrets 裡面的鑰匙連線至 GitHub
                g = Github(st.secrets["GITHUB_TOKEN"])
                repo = g.get_repo(st.secrets["REPO_NAME"])
                
                # 3. 取得原本的檔案並進行覆蓋
                file_path = "哈蜜瓜收支表.xlsx"
                contents = repo.get_contents(file_path)
                
                commit_message = f"自動存檔: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                repo.update_file(contents.path, commit_message, excel_data, contents.sha)
                
                st.success("✅存檔成功！資料已完成同步")
            except Exception as e:
                st.error(f"❌存檔失敗：請檢查 Secrets 設定或 Token 權限。錯誤細節：{e}")

# ==========================================
# 3. 建立網頁分頁
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["💰收支概況", "📊記帳與管理", "🏸季繳追蹤", "ℹ️後台資訊"])


# ------------------------------------------
# 分頁 1：💰收支概況
# ------------------------------------------
with tab1:
    st.subheader("💰總財務概況")
    
    if '類別' in current_trans.columns and '金額' in current_trans.columns:
        total_income = current_trans[current_trans['類別'] == '收入']['金額'].sum()
        total_expense = current_trans[current_trans['類別'] == '支出']['金額'].sum()
        net_balance = total_income - total_expense

        # 安全挪用款
        if '項目' in current_trans.columns:
            venue_expenses = current_trans[(current_trans['類別'] == '支出') & (current_trans['項目'] == '場地費')]['金額']
            weekly_venue_fee = venue_expenses.mode()[0] if not venue_expenses.empty else 760
        else:
            weekly_venue_fee = 380 * 2
        safe_buffer = weekly_venue_fee * 2 

        # 季繳預留金
        total_remaining = st.session_state.members_df['剩餘次數'].sum()
        cost_per_time = 220
        quarterly_reserve = total_remaining * cost_per_time

        col1, col2, col3 = st.columns(3)
        with col2:
            st.metric("💵總結餘金額", f"${net_balance:,.0f}")
        with col1:
            st.success(f"### 🛡️可挪用金額: **${safe_buffer:,.0f}**")
        with col3:
            st.metric("💳季繳預留金", f"${quarterly_reserve:,.0f}")
            
    st.divider()
    
    st.subheader("📊收支圖表分析")
    
    if '類別' in current_trans.columns and '金額' in current_trans.columns and '日期' in current_trans.columns:
        col1, col2 = st.columns(2)
        with col1:
            summary = current_trans.groupby('類別')['金額'].sum().reset_index()
            fig_pie = px.pie(summary, values='金額', names='類別', title="總收入 vs 總支出",
                             color='類別', color_discrete_map={'收入':'#28a745', '支出':'#dc3545'})
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col2:
            plot_df = current_trans.dropna(subset=['日期']).copy()
            plot_df['月份'] = plot_df['日期'].dt.strftime('%Y-%m')

            monthly_summary = plot_df.groupby(['月份', '類別'])['金額'].sum().reset_index()
            
            fig_bar = px.bar(monthly_summary, x='月份', y='金額', color='類別', barmode='group', title="每月收支變化",
                             color_discrete_map={'收入':'#28a745', '支出':'#dc3545'})
            # 確保 X 軸顯示為純文字類別 (解決時間格式跑版)
            fig_bar.update_layout(xaxis_type='category')
            st.plotly_chart(fig_bar, use_container_width=True)


# ------------------------------------------
# 分頁 2：📊記帳與管理
# ------------------------------------------
with tab2:
    col_left, col_right = st.columns([1, 2])
    
    # 左側：新增收支紀錄
    with col_left:
        st.subheader("✍️新增收支紀錄")
        with st.form("add_transaction_form", clear_on_submit=True):
            t_date = st.date_input("📅日期", datetime.today())
            t_type = st.selectbox("類別", ["收入", "支出"])
            t_item = st.selectbox("項目", item_options)
            t_amount = st.number_input("金額 ($)", min_value=0, step=1)
            t_handler = st.selectbox("經手人", handler_options)
            t_note = st.text_input("備註")
            
            submitted = st.form_submit_button("➕確認新增帳目", use_container_width=True)
            if submitted:
                if not t_item:
                    st.warning("⚠️請填寫「項目」欄位！")
                else:
                    new_record = pd.DataFrame([{
                        "日期": pd.to_datetime(t_date),
                        "類別": t_type,
                        "項目": t_item,
                        "金額": t_amount,
                        "經手人": t_handler,
                        "備註": t_note
                    }])
                    st.session_state.trans_df = pd.concat([st.session_state.trans_df, new_record], ignore_index=True)
                    st.success(f"✅已成功記帳：{t_type} - {t_item} ${t_amount}")
                    st.rerun()

    # 右側：收支總表與編輯
    with col_right:
        st.subheader("📝收支總表")
        # 將標題列切分為左右，讓按鈕靠右對齊
        col_title, col_dl = st.columns([7, 3])
        with col_title:
            st.subheader("📝收支總表")
        with col_dl:
            # 準備 Excel 檔案
            dl_output = io.BytesIO()
            with pd.ExcelWriter(dl_output, engine='openpyxl') as writer:
                st.session_state.trans_df.to_excel(writer, sheet_name='收支明細', index=False)
                st.session_state.members_df.to_excel(writer, sheet_name='季繳名單', index=False)
            
            st.download_button(
                label="📥下載 Excel",
                data=dl_output.getvalue(),
                file_name=f"哈蜜瓜收支表_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_tab2"
            )

        # 初始化動態 Key (用於取消刪除時強制刷新前端畫面)
        if "trans_key_version" not in st.session_state:
            st.session_state.trans_key_version = 0
        
        if "日期" in st.session_state.trans_df.columns:
            display_df = st.session_state.trans_df.sort_values(by="日期", ascending=False).copy()
        else:
            display_df = st.session_state.trans_df.copy()
            st.warning(f"⚠️找不到『日期』欄位，目前的欄位有：{list(display_df.columns)}。請確認 Excel 的標題列位置。")

        # 處理空值避免顯示 None
        for col in display_df.select_dtypes(include=['object']).columns:
            display_df[col] = display_df[col].fillna("")
            
        # 插入刪除打勾欄位
        display_df.insert(0, "🗑️刪除", False)
        
        # ✨ 動態抓取第一欄名稱，完全避開 Emoji 字串辨識問題
        del_col = display_df.columns[0]
            
        edited_trans = st.data_editor(
            display_df, 
            use_container_width=True, 
            num_rows="fixed",
            hide_index=True,
            key=f"trans_editor_{st.session_state.trans_key_version}",
            column_config={
                del_col: st.column_config.CheckboxColumn("刪除", default=False, width="small"),
                "日期": st.column_config.DateColumn("日期", format="YYYY-MM-DD"),
                "類別": st.column_config.SelectboxColumn("類別", options=["收入", "支出"], required=True),
                "項目": st.column_config.SelectboxColumn("項目", options=item_options),
                "金額": st.column_config.NumberColumn("金額", format="%d"),
                "經手人": st.column_config.SelectboxColumn("經手人", options=handler_options),
                "備註": st.column_config.TextColumn("備註")
            }
        )   
        
        # 處理打勾刪除與雙擊修改邏輯
        if edited_trans[del_col].any():
            st.warning("⚠️發現已勾選的項目，確定要刪除嗎？")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("❎確認刪除", key="confirm_del_trans", use_container_width=True):
                    # 完全不使用字串，直接過濾掉打勾的資料，並用 .iloc[:, 1:] 排除打勾欄
                    st.session_state.trans_df = edited_trans[~edited_trans[del_col]].iloc[:, 1:].reset_index(drop=True)
                    st.session_state.trans_key_version += 1 
                    st.rerun()

            with col_btn2:
                if st.button("🔙取消刪除", key="cancel_del_trans", use_container_width=True):
                    st.session_state.trans_key_version += 1
                    st.rerun()
        else:
            # 用 .iloc[:, 1:] 比對，不依賴字串名稱
            orig_check = display_df.iloc[:, 1:]
            edit_check = edited_trans.iloc[:, 1:]
            if not edit_check.equals(orig_check):
                st.session_state.trans_df = edit_check
                st.rerun()

# ------------------------------------------
# 分頁 3：🏸季繳追蹤
# ------------------------------------------
with tab3:
    col_left, col_right = st.columns([1, 2])
    
    # 左側：新增季繳人員
    with col_left:
        st.subheader("➕新增季繳人員")
        with st.form("add_member_form"):
            new_name = st.text_input("👤姓名", placeholder="必填")
            new_date = st.date_input("📅繳費日期", datetime.today())
            new_start_date = st.date_input("📅季繳開始日期", datetime.today(), help="⚠️僅能選擇禮拜二")
                
            submit_new_member = st.form_submit_button("➕確認新增人員", use_container_width=True)
            
            if submit_new_member:
                if not new_name.strip():
                    st.warning("⚠️請填寫「姓名」欄位！")
                elif new_start_date.weekday() != 1:
                    weekdays_zh = ["一", "二", "三", "四", "五", "六", "日"]
                    wrong_day = weekdays_zh[new_start_date.weekday()]
                    st.error(f"⚠️【日期錯誤】「季繳開始日期」必須為星期二！您選擇的 {new_start_date.strftime('%Y-%m-%d')} 是星期{wrong_day}。")
                else:
                    today_dt = pd.to_datetime(datetime.today().date())
                    start_dt = pd.to_datetime(new_start_date)
                    days_diff_new = (today_dt - start_dt).days
                    calculated_times = int(10 - np.ceil(days_diff_new / 7))
                    
                    new_member_data = pd.DataFrame([{
                        "姓名": new_name,
                        "繳費日期": pd.to_datetime(new_date),
                        "季繳開始日期": start_dt,
                        "剩餘次數": calculated_times
                    }])
                    st.session_state.members_df = pd.concat(
                        [st.session_state.members_df, new_member_data], 
                        ignore_index=True
                    )
                    st.success(f"✅已成功新增球友：{new_name} (開始日: {new_start_date}, 系統自動計算剩餘 {calculated_times} 次)")
                    st.rerun() 

    # 右側：季繳追蹤清單與編輯
    with col_right:
        st.subheader("📋季繳追蹤清單")
        
        # 將標題列切分為左右，讓按鈕靠右對齊
        col_title, col_dl = st.columns([7, 3])
        with col_title:
            st.subheader("📋季繳追蹤清單")
        with col_dl:
            # 準備 Excel 檔案
            dl_output = io.BytesIO()
            with pd.ExcelWriter(dl_output, engine='openpyxl') as writer:
                st.session_state.trans_df.to_excel(writer, sheet_name='收支明細', index=False)
                st.session_state.members_df.to_excel(writer, sheet_name='季繳名單', index=False)
                
            st.download_button(
                label="📥下載 Excel",
                data=dl_output.getvalue(),
                file_name=f"哈蜜瓜收支表_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_tab3"
            )

        # 初始化動態 Key (用於取消刪除時強制刷新前端畫面)
        if "members_key_version" not in st.session_state:
            st.session_state.members_key_version = 0
        
        def update_status(times):
            if times <= 0: return "⛔已結束"
            elif times <= 2: return "⚠️剩餘2次，可提醒繳費"
            else: return "🟢進行中"

        display_df = st.session_state.members_df.copy()
        display_df['狀態提醒'] = display_df['剩餘次數'].apply(update_status)
          
        # 插入刪除打勾欄位
        display_df.insert(0, "🗑️刪除", False)
        
        # ✨ 動態抓取第一欄名稱，完全避開 Emoji 字串辨識問題
        del_col_mem = display_df.columns[0]

        def highlight_zero(row):
            if row['剩餘次數'] <= 0:
                return ['text-decoration: line-through; color: #888888;'] * len(row)
            return [''] * len(row)

        edited_members = st.data_editor(
            display_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            key=f"members_editor_{st.session_state.members_key_version}",
            column_config={
                del_col_mem: st.column_config.CheckboxColumn("刪除", default=False, width="small"),
                "繳費日期": st.column_config.DateColumn("繳費日期", format="YYYY-MM-DD"),
                "季繳開始日期": st.column_config.DateColumn("季繳開始日期", format="YYYY-MM-DD"),
                "剩餘次數": st.column_config.NumberColumn("剩餘次數", format="%d")
            }
        )

        # 處理打勾刪除與雙擊修改邏輯
        if edited_members[del_col_mem].any():
            st.warning("⚠️發現已勾選的項目，確定要刪除嗎？")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("❎確認刪除", key="confirm_del_mem", use_container_width=True):
                    # 篩選掉打勾的列，並只保留需要的四個欄位
                    st.session_state.members_df = edited_members[~edited_members[del_col_mem]][['姓名', '繳費日期', '季繳開始日期', '剩餘次數']].reset_index(drop=True)
                    st.session_state.members_key_version += 1
                    st.rerun()
            with col_btn2:
                if st.button("🔙取消刪除", key="cancel_del_mem", use_container_width=True):
                    st.session_state.members_key_version += 1
                    st.rerun()
        else:
            # 直接選取特定欄位比對，不依賴字串名稱來 drop
            orig_check = display_df[['姓名', '繳費日期', '季繳開始日期', '剩餘次數']]
            edit_check = edited_members[['姓名', '繳費日期', '季繳開始日期', '剩餘次數']]
            if not edit_check.equals(orig_check):
                st.session_state.members_df = edit_check
                st.rerun()


# ------------------------------------------
# 分頁 4：ℹ️後台資訊
# ------------------------------------------
with tab4:
    st.subheader("ℹ️後台資訊")
    st.info("💡提示：如需修改以下內容，請通知團主")
    
    st.markdown("""
    ### 💰收費標準
    * **季繳球友：** 預繳10次共2200元
    * **臨打球友：** 每次250元、只打1小時每次150元
    * **羽毛球費：** 每桶610元
    
    ### 🎉球團福利
    * **當月壽星優惠：** 只要為打過一次的球友，便可享一次臨打免費+飲料任挑
       
    ### 🧑‍💼預估設定
    * **總結餘金額：** 收入-支出
    * **季繳預留金：** 預留2周場地費並預設租借2個場地
    * **可挪用金額：** 總結餘金額-季繳預留金

    ### 📈統計結果【2026.4~2026.9】
    * 每次開團平均人次：12.43人
    * 每次開團場地平均使用時數：3.78hrs
    * 每次開團分攤時數：0.314hr/人場
    """)
