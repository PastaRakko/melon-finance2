import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timezone, timedelta
TW_TZ = timezone(timedelta(hours=8))
import numpy as np
import io
from github import Github

st.set_page_config(page_title="哈蜜瓜收支表", page_icon="🍈", layout="wide")

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

# 輔助函式：取得依日期排序好的收支明細表 (由新到舊，且格式化為純日期 YYYY-MM-DD)
def get_sorted_trans_df(df):
    temp_df = df.copy()
    if '日期' in temp_df.columns:
        temp_df['日期'] = pd.to_datetime(temp_df['日期'], errors='coerce')
        temp_df = temp_df.sort_values(by='日期', ascending=False, na_position='last').reset_index(drop=True)
        temp_df['日期'] = temp_df['日期'].dt.strftime('%Y-%m-%d')
    return temp_df

# 輔助函式：將季繳名單的日期欄位也轉為純日期字串
def get_clean_members_df(df):
    temp_df = df.copy()
    for date_col in ['繳費日期', '季繳開始日期']:
        if date_col in temp_df.columns:
            temp_df[date_col] = pd.to_datetime(temp_df[date_col], errors='coerce').dt.strftime('%Y-%m-%d')
    return temp_df

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
    if '金額' in df_t.columns:
        df_t['金額'] = pd.to_numeric(df_t['金額'].astype(str).replace(r'[\$,]', '', regex=True), errors='coerce').fillna(0)
    # 初始化時立即進行日期排序與格式化
    st.session_state.trans_df = get_sorted_trans_df(df_t)

current_trans = st.session_state.trans_df

# ------------------------------------------
# 動態產生下拉選單
# ------------------------------------------
item_options = ["場地費", "臨打", "季繳", "羽毛球", "活動費", "其他"]
if '項目' in current_trans.columns:
    for x in current_trans['項目'].dropna().unique():
        if x not in item_options and str(x).strip() != "":
            item_options.append(x)

handler_options = ["櫃台", "妙", "齊"]
if '經手人' in current_trans.columns:
    for x in current_trans['經手人'].dropna().unique():
        if x not in handler_options and str(x).strip() != "":
            handler_options.append(x)

# ==========================================
# 快捷按鈕存檔與重新讀取 (標題旁邊)
# ==========================================

st.markdown("""
<style>
div.stButton > button[kind="primary"] {
    background-color: #0d6efd !important;
    color: white !important;
    border-color: #0d6efd !important;
}
div.stButton > button[kind="primary"]:hover {
    background-color: #0b5ed7 !important;
    border-color: #0a58ca !important;
}
</style>
""", unsafe_allow_html=True)

col_title, col_save, col_reload = st.columns([6, 2, 2], vertical_alignment="bottom")
with col_title:
    st.title("🍈哈蜜瓜收支與季繳管理表")
with col_save:
    if st.button("💾確認更新並存檔", type="primary", use_container_width=True):
        with st.spinner("⌛正在上傳資料中，請稍候..."):
            try:
                now_str = datetime.now(TW_TZ).strftime('%Y-%m-%d %H:%M:%S')
                
                # 存檔前整理資料格式（純日期 YYYY-MM-DD，收支明細按日期排序）
                save_trans_df = get_sorted_trans_df(st.session_state.trans_df)
                save_members_df = get_clean_members_df(st.session_state.members_df)
                
                st.session_state.trans_df = save_trans_df
                
                # 1. 產生最新的 Excel 檔案內容到記憶體
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    save_trans_df.to_excel(writer, sheet_name='收支明細', index=False)
                    save_members_df.to_excel(writer, sheet_name='季繳名單', index=False)
                excel_data = output.getvalue()
                
                # 2. 呼叫 Secrets 連線至 GitHub 覆蓋舊檔
                g = Github(st.secrets["GITHUB_TOKEN"])
                repo = g.get_repo(st.secrets["REPO_NAME"])
                
                file_path = "哈蜜瓜收支表.xlsx"
                contents = repo.get_contents(file_path)
                
                commit_message = f"自動存檔: {now_str}"
                repo.update_file(contents.path, commit_message, excel_data, contents.sha)
                
                st.session_state.last_updated = now_str
                st.success("✅存檔成功！資料已完成同步")
            except Exception as e:
                st.error(f"❌存檔失敗：請檢查 Secrets 設定或 Token 權限。錯誤細節：{e}")

with col_reload:
    if st.button("🔄重新讀取資料", use_container_width=True):
        load_excel.clear() 
        if 'trans_df' in st.session_state:
            del st.session_state['trans_df']
        if 'members_df' in st.session_state:
            del st.session_state['members_df']
        st.session_state.last_updated = datetime.now(TW_TZ).strftime('%Y-%m-%d %H:%M:%S')
        st.rerun()

# 執行讀取資料
try:
    df_trans, df_members = load_excel()
except Exception as e:
    st.error("讀取 Excel 失敗，請確認檔案與分頁是否存在。")
    st.stop()

# ==========================================
# 3. 建立網頁分頁
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(["💰收支概況", "📊記帳與管理", "🏸季繳追蹤", "ℹ️後台資訊"])


# ------------------------------------------
# 分頁 1：💰收支概況
# ------------------------------------------
with tab1:
    col_text, col_line, col_charts = st.columns([1, 0.1, 2])
    
    # --- 幕後財務數據計算 ---
    if '類別' in current_trans.columns and '金額' in current_trans.columns:
        total_income = current_trans[current_trans['類別'] == '收入']['金額'].sum()
        total_expense = current_trans[current_trans['類別'] == '支出']['金額'].sum()
        net_balance = total_income - total_expense

        # 場地預留金 (僅幕後扣除)
        if '項目' in current_trans.columns:
            venue_expenses = current_trans[(current_trans['類別'] == '支出') & (current_trans['項目'] == '場地費')]['金額']
            weekly_venue_fee = venue_expenses.mode()[0] if not venue_expenses.empty else 760
        else:
            weekly_venue_fee = 380 * 2
        venue_reserve = weekly_venue_fee * 2 

        # 季繳預留金
        total_remaining = st.session_state.members_df['剩餘次數'].sum()
        cost_per_time = 220
        quarterly_reserve = total_remaining * cost_per_time

        # 可挪用金額 = 總結餘 - 季繳預留 - 場地預留(幕後)
        usable_balance = net_balance - quarterly_reserve - venue_reserve

    # --- 左半邊：文字數據 ---
    with col_text:
        st.subheader("💰總財務概況")
        
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric("💵總結餘金額", f"${net_balance:,.0f}")

        with col_m2:
            st.metric("💳季繳預留金", f"${quarterly_reserve:,.0f}")
            
        st.write("") 
        st.success(f"### 🛡️可挪用金額:\n### **${usable_balance:,.0f}**")
        
        # 顯示最後更新時間（位於可挪用金額框正下方）
        last_sync_time = st.session_state.get('last_updated', '尚未更新')
        st.caption(f"⏱️ 更新時間：{last_sync_time}")

    # --- 中間：垂直分隔線 ---
    with col_line:
        st.markdown(
            '<div style="width: 2px; height: 400px; background-color: #e0e0e0; margin: 0 auto; margin-top: 50px;"></div>', 
            unsafe_allow_html=True
        )

    # --- 右半邊：長條圖 ---
    with col_charts:
        st.subheader("📊收支圖表分析")
        if '類別' in current_trans.columns and '金額' in current_trans.columns and '日期' in current_trans.columns:
            plot_df = current_trans.dropna(subset=['日期']).copy()
            plot_df['月份'] = pd.to_datetime(plot_df['日期']).dt.strftime('%Y-%m')
            monthly_summary = plot_df.groupby(['月份', '類別'])['金額'].sum().reset_index()
            fig_bar = px.bar(monthly_summary, x='月份', y='金額', color='類別', barmode='group', title="每月收支變化",
                            color_discrete_map={'收入':'#28a745', '支出':'#dc3545'})
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
                    updated_trans = pd.concat([st.session_state.trans_df, new_record], ignore_index=True)
                    st.session_state.trans_df = get_sorted_trans_df(updated_trans)
                    st.success(f"✅已成功記帳：{t_type} - {t_item} ${t_amount}")
                    st.rerun()

    # 右側：收支總表與編輯
    with col_right:
        col_title, col_dl = st.columns([7, 3])
        with col_title:
            st.subheader("📝收支總表")
        with col_dl:
            dl_output = io.BytesIO()
            today_str = datetime.now(TW_TZ).strftime('%Y%m%d')
            
            # 準備匯出的純日期且排序好的 Excel 資料
            export_trans_df = get_sorted_trans_df(st.session_state.trans_df)
            export_members_df = get_clean_members_df(st.session_state.members_df)
            
            with pd.ExcelWriter(dl_output, engine='openpyxl') as writer:
                export_trans_df.to_excel(writer, sheet_name=f'收支明細_{today_str}', index=False)
                export_members_df.to_excel(writer, sheet_name=f'季繳名單_{today_str}', index=False)
            
            st.download_button(
                label="📥下載Excel",
                data=dl_output.getvalue(),
                file_name=f"哈蜜瓜收支表_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_tab2"
            )

        if "trans_key_version" not in st.session_state:
            st.session_state.trans_key_version = 0
        
        # 1. 取得排序後的資料
        display_df = get_sorted_trans_df(st.session_state.trans_df)

        # 2. 強制把「日期」轉為 dt.date 物件，防範 data_editor 報錯
        if '日期' in display_df.columns:
            display_df['日期'] = pd.to_datetime(display_df['日期'], errors='coerce').dt.date

        # 3. 處理空值避免顯示 None
        for col in display_df.select_dtypes(include=['object']).columns:
            display_df[col] = display_df[col].fillna("")
            
        # 4. 插入刪除欄位
        display_df.insert(0, "🗑️刪除", False)
        del_col = display_df.columns[0]
            
        # 5. 安全地建立 column_config
        cfg = {
            del_col: st.column_config.CheckboxColumn("刪除", default=False, width="small"),
            "類別": st.column_config.SelectboxColumn("類別", options=["收入", "支出"], required=True),
            "項目": st.column_config.SelectboxColumn("項目", options=item_options),
            "金額": st.column_config.NumberColumn("金額", format="%d"),
            "經手人": st.column_config.SelectboxColumn("經手人", options=handler_options),
            "備註": st.column_config.TextColumn("備註")
        }
        if "日期" in display_df.columns:
            cfg["日期"] = st.column_config.DateColumn("日期", format="YYYY-MM-DD")

        edited_trans = st.data_editor(
            display_df, 
            use_container_width=True, 
            num_rows="fixed",
            hide_index=True,
            key=f"trans_editor_{st.session_state.trans_key_version}",
            column_config=cfg
        )   
        
        if edited_trans[del_col].any():
            st.warning("⚠️發現已勾選的項目，確定要刪除嗎？")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("❎確認刪除", key="confirm_del_trans", use_container_width=True):
                    filtered_df = edited_trans[~edited_trans[del_col]].iloc[:, 1:].reset_index(drop=True)
                    st.session_state.trans_df = get_sorted_trans_df(filtered_df)
                    st.session_state.trans_key_version += 1 
                    st.rerun()

            with col_btn2:
                if st.button("🔙取消刪除", key="cancel_del_trans", use_container_width=True):
                    st.session_state.trans_key_version += 1
                    st.rerun()
        else:
            orig_check = display_df.iloc[:, 1:]
            edit_check = edited_trans.iloc[:, 1:]
            if not edit_check.equals(orig_check):
                st.session_state.trans_df = get_sorted_trans_df(edit_check)
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
        col_title, col_dl = st.columns([7, 3])
        with col_title:
            st.subheader("📋季繳追蹤清單")
        with col_dl:
            dl_output = io.BytesIO()
            today_str = datetime.now(TW_TZ).strftime('%Y%m%d')
            
            export_trans_df = get_sorted_trans_df(st.session_state.trans_df)
            export_members_df = get_clean_members_df(st.session_state.members_df)
            
            with pd.ExcelWriter(dl_output, engine='openpyxl') as writer:
                export_trans_df.to_excel(writer, sheet_name=f'收支明細_{today_str}', index=False)
                export_members_df.to_excel(writer, sheet_name=f'季繳名單_{today_str}', index=False)
                
            st.download_button(
                label="📥下載Excel",
                data=dl_output.getvalue(),
                file_name=f"哈蜜瓜收支表_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_tab3"
            )

        if "members_key_version" not in st.session_state:
            st.session_state.members_key_version = 0
        
        def update_status(times):
            if times <= 0: return "⛔已結束"
            elif times <= 2: return "⚠️剩餘2次，可提醒繳費"
            else: return "🟢進行中"

        display_df = st.session_state.members_df.copy()
        display_df['狀態提醒'] = display_df['剩餘次數'].apply(update_status)
        
        # 安全轉為 dt.date
        for date_col in ['繳費日期', '季繳開始日期']:
            if date_col in display_df.columns:
                display_df[date_col] = pd.to_datetime(display_df[date_col], errors='coerce').dt.date

        display_df.insert(0, "🗑️刪除", False)
        del_col_mem = display_df.columns[0]

        mem_cfg = {
            del_col_mem: st.column_config.CheckboxColumn("刪除", default=False, width="small"),
            "剩餘次數": st.column_config.NumberColumn("剩餘次數", format="%d")
        }
        if "繳費日期" in display_df.columns:
            mem_cfg["繳費日期"] = st.column_config.DateColumn("繳費日期", format="YYYY-MM-DD")
        if "季繳開始日期" in display_df.columns:
            mem_cfg["季繳開始日期"] = st.column_config.DateColumn("季繳開始日期", format="YYYY-MM-DD")

        edited_members = st.data_editor(
            display_df,
            use_container_width=True,
            num_rows="fixed",
            hide_index=True,
            key=f"members_editor_{st.session_state.members_key_version}",
            column_config=mem_cfg
        )

        if edited_members[del_col_mem].any():
            st.warning("⚠️發現已勾選的項目，確定要刪除嗎？")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("❎確認刪除", key="confirm_del_mem", use_container_width=True):
                    st.session_state.members_df = edited_members[~edited_members[del_col_mem]][['姓名', '繳費日期', '季繳開始日期', '剩餘次數']].reset_index(drop=True)
                    st.session_state.members_key_version += 1
                    st.rerun()
            with col_btn2:
                if st.button("🔙取消刪除", key="cancel_del_mem", use_container_width=True):
                    st.session_state.members_key_version += 1
                    st.rerun()
        else:
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
    st.markdown("💡 **注意：如以下內容需修改，請通知團主**")
    st.write("") 
    
    # --- 第一列 (2個欄位) ---
    row1_col1, row1_col2 = st.columns(2)
    
    with row1_col1:
        st.markdown("""
        <div style="background-color: #e3f2fd; padding: 20px; border-radius: 12px; min-height: 190px; color: #1f1f1f;">
            <h3 style="margin-top: 0; color: #0d47a1;">💰收費標準</h3>
            <ul style="margin-bottom: 0; line-height: 1.8;">
                <li><b>季繳球友：</b>預繳10次共2200元</li>
                <li><b>臨打球友：</b>每次250元、只打1小時每次150元</li>
                <li><b>羽毛球費：</b>每桶610元</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with row1_col2:
        st.markdown("""
        <div style="background-color: #fff3e0; padding: 20px; border-radius: 12px; min-height: 190px; color: #1f1f1f;">
            <h3 style="margin-top: 0; color: #e65100;">🎉球團福利</h3>
            <ul style="margin-bottom: 0; line-height: 1.8;">
                <li><b>當月壽星優惠：</b>只要為打過一次的球友，便可享一次臨打免費+飲料任挑</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    st.write("") 
    
    # --- 第二列 (2個欄位) ---
    row2_col1, row2_col2 = st.columns(2)
    
    with row2_col1:
        st.markdown("""
        <div style="background-color: #e8f5e9; padding: 20px; border-radius: 12px; min-height: 190px; color: #1f1f1f;">
            <h3 style="margin-top: 0; color: #1b5e20;">🧑‍💼預估設定</h3>
            <ul style="margin-bottom: 0; line-height: 1.8;">
                <li><b>總結餘金額：</b>收入-支出</li>
                <li><b>季繳預留金：</b>季繳球友剩餘總次數 × 220元</li>
                <li><b>可挪用金額：</b>總結餘金額-季繳預留金-預留2周場地費</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with row2_col2:
        st.markdown("""
        <div style="background-color: #f3e5f5; padding: 20px; border-radius: 12px; min-height: 190px; color: #1f1f1f;">
            <h3 style="margin-top: 0; color: #4a148c;">📈統計結果【2026.4~2026.9】</h3>
            <ul style="margin-bottom: 0; line-height: 1.8;">
                <li>每次開團平均人次：12.43人</li>
                <li>每次開團場地平均使用時數：3.78hrs</li>
                <li>每次開團分攤時數：0.314hr/人場</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
