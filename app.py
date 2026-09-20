import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import math
from ortools.sat.python import cp_model
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="飯食服事自動排班系統", page_icon="🍞", layout="wide")

st.title("🍞 飯食服事自動排班系統")
st.write("歡迎使用！請上傳從 Google 表單下載的原始 CSV 或是 Excel 檔案，並在左側調整學期條件與人員離台/請假/次數微調設定。")

# 1. 檔案上傳
uploaded_file = st.sidebar.file_uploader("📂 上傳表單檔案 (CSV 或 Excel)", type=["csv", "xlsx", "xls"])

# 2. 側邊欄條件設定 (GUI 化)
st.sidebar.header("⚙️ 1. 學期與基本設定")

today = datetime.now().date()
start_date = st.sidebar.date_input("學期開始日期", today)
end_date = st.sidebar.date_input("學期結束日期", today + timedelta(days=90))

weekday_options = {"週一": 0, "週二": 1, "週三": 2, "週四": 3, "週五": 4}
selected_weekday_names = st.sidebar.multiselect("🗓️ 選擇每週服事日期", options=list(weekday_options.keys()), default=[])
selected_weekdays = [weekday_options[name] for name in selected_weekday_names]

date_range_days = (end_date - start_date).days if end_date >= start_date else 0
selected_holidays = st.sidebar.multiselect(
    "國定假日 / 不排班日期", 
    options=[start_date + timedelta(days=i) for i in range(date_range_days + 1)],
    default=[],
    format_func=lambda d: d.strftime("%Y/%m/%d (%a)")
)

st.sidebar.subheader("👥 人力需求設定")
before_count = st.sidebar.number_input("飯前人數 (固定同性別)", min_value=1, max_value=4, value=2)
after_min = st.sidebar.number_input("飯後最少人數", min_value=1, max_value=8, value=4)
after_max = st.sidebar.number_input("飯後最多人數", min_value=1, max_value=8, value=5)

leave_dates_gui = {}
join_dates_gui = {}
custom_adjustments_gui = {}

df_raw = None

if uploaded_file is not None:
    file_name = uploaded_file.name
    try:
        if file_name.endswith('.csv'):
            try:
                df_raw = pd.read_csv(uploaded_file, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, encoding='cp950')
        else:
            df_raw = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ 讀取檔案失敗：{e}。請確認檔案格式是否正確。")

if df_raw is not None:
    # 找到正確的姓名欄位名稱
    name_col = [c for c in df_raw.columns if "姓名" in str(c) or "Name" in str(c)]
    name_col_name = name_col[0] if name_col else df_raw.columns[0]
    
    # 清除姓名前後空格
    df_raw[name_col_name] = df_raw[name_col_name].astype(str).str.strip()
    members_list = [m for m in df_raw[name_col_name].dropna().unique().tolist() if m and m != "nan"]
    
    st.sidebar.markdown("---")
    st.sidebar.header("✈️ 2. 特定人員出勤限制")

    selected_leave_members = st.sidebar.multiselect("選擇離台成員", options=members_list, default=[])
    for m in selected_leave_members:
        l_date = st.sidebar.date_input(f"【{m}】最後服事/離台日期", value=end_date, key=f"leave_{m}")
        leave_dates_gui[m] = l_date

    selected_join_members = st.sidebar.multiselect("選擇延後加入成員", options=members_list, default=[])
    for m in selected_join_members:
        j_date = st.sidebar.date_input(f"【{m}】開始可服事日期", value=start_date, key=f"join_{m}")
        join_dates_gui[m] = j_date

    # 特定人員服事次數微調
    st.sidebar.markdown("---")
    st.sidebar.header("⚖️ 3. 特定人員服事次數微調")
    selected_adj_members = st.sidebar.multiselect("選擇需微調服事次數的成員", options=members_list, default=[])
    
    for m in selected_adj_members:
        st.sidebar.write(f"**👤 【{m}】次數調整**")
        b_adj = st.sidebar.number_input(f"【{m}】飯前次數增減 (負數代表減少)", min_value=-10, max_value=10, value=0, key=f"badj_{m}")
        a_adj = st.sidebar.number_input(f"【{m}】飯後次數增減 (負數代表減少)", min_value=-10, max_value=10, value=0, key=f"aadj_{m}")
        if b_adj != 0 or a_adj != 0:
            custom_adjustments_gui[m] = {"飯前": b_adj, "飯後": a_adj}

# 核心排班函數
def run_scheduler(df, start_date, end_date, active_weekdays, holidays_list, b_count, a_min, a_max, leave_map, join_map, custom_adjustments):
    holidays_dt = [datetime.combine(h, datetime.min.time()) for h in holidays_list]
    leave_dates = {k: datetime.combine(v, datetime.min.time()) for k, v in leave_map.items()}
    join_dates = {k: datetime.combine(v, datetime.min.time()) for k, v in join_map.items()}

    dates = []
    curr = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.min.time())
    
    while curr <= end_dt:
        if curr.weekday() in active_weekdays and curr not in holidays_dt:
            dates.append(curr)
        curr += timedelta(days=1)

    name_col = [c for c in df.columns if "姓名" in str(c) or "Name" in str(c)][0]
    gender_col = [c for c in df.columns if "弟兄" in str(c) or "姊妹" in str(c) or "Br." in str(c)][0]

    members = [str(m).strip() for m in df[name_col].dropna().unique().tolist() if str(m).strip() and str(m).strip() != "nan"]
    gender_map = dict(zip(df[name_col].astype(str).str.strip(), df[gender_col]))
    
    # 廣泛相容星期與時段關鍵字
    weekday_keys = {
        0: ["週一", "星期一", "禮拜一", "mon"],
        1: ["週二", "星期二", "禮拜二", "tue"],
        2: ["週三", "星期三", "禮拜三", "wed"],
        3: ["週四", "星期四", "禮拜四", "thu"],
        4: ["週五", "星期五", "禮拜五", "fri"]
    }
    shift_keys = {
        "飯前": ["飯前", "餐前", "before"],
        "飯後": ["飯後", "餐後", "after"]
    }

    total_dates = len(dates)
    total_members = len(members) if len(members) > 0 else 1
    avg_before = (total_dates * b_count) / total_members
    avg_after = (total_dates * ((a_min + a_max) / 2)) / total_members

    avail_map = {}
    pref_map = {}

    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        if not name or name == "nan": continue
        avail_map[name] = {}
        pref_map[name] = {"飯前": [], "飯後": []}
        
        for d in dates:
            w_idx = d.weekday()
            w_keywords = weekday_keys[w_idx]
            
            for s in ["飯前", "飯後"]:
                s_keywords = shift_keys[s]
                
                matched_col = None
                for col in df.columns:
                    col_clean = str(col).replace("\n", "").replace(" ", "").lower()
                    
                    # 嚴格剔除偏好欄位
                    if "偏好" in col_clean or "preference" in col_clean:
                        continue
                    
                    if any(w_k in col_clean for w_k in w_keywords) and any(s_k in col_clean for s_k in s_keywords):
                        matched_col = col
                        break
                
                if matched_col:
                    val = str(row[matched_col]).strip().lower().replace("\n", "").replace(" ", "")
                    
                    # 擴充否決關鍵字，優先攔截
                    no_keywords = ["不", "否", "no", "not", "0", "無法", "忙碌", "x", "難", "拒"]
                    yes_keywords = ["可以", "可", "1", "yes", "available", "ok", "v", "圈"]
                    
                    if any(k in val for k in no_keywords):
                        avail_map[name][(d, s)] = 0
                    elif any(k in val for k in yes_keywords):
                        avail_map[name][(d, s)] = 1
                    else:
                        avail_map[name][(d, s)] = 0
                else:
                    avail_map[name][(d, s)] = 0

        pref_b_col = [c for c in df.columns if ("偏好" in str(c) or "preference" in str(c).lower()) and ("飯前" in str(c) or "餐前" in str(c))]
        pref_a_col = [c for c in df.columns if ("偏好" in str(c) or "preference" in str(c).lower()) and ("飯後" in str(c) or "餐後" in str(c))]
        if pref_b_col and not pd.isna(row[pref_b_col[0]]):
            pref_map[name]["飯前"] = str(row[pref_b_col[0]])
        if pref_a_col and not pd.isna(row[pref_a_col[0]]):
            pref_map[name]["飯後"] = str(row[pref_a_col[0]])

    model = cp_model.CpModel()
    shifts = ["飯前", "飯後"]
    x = {}
    for p in members:
        for d in dates:
            for s in shifts:
                x[p, d, s] = model.NewBoolVar(f'x_{p}_{d.strftime("%Y%m%d")}_{s}')

    # 硬性絕不排班約束 (Not Available)
    for d in dates:
        for p in members:
            if avail_map.get(p, {}).get((d, "飯前"), 1) == 0:
                model.Add(x[p, d, "飯前"] == 0)
            if avail_map.get(p, {}).get((d, "飯後"), 1) == 0:
                model.Add(x[p, d, "飯後"] == 0)

        males_before = [x[p, d, "飯前"] for p in members if "弟兄" in str(gender_map.get(p, ""))]
        females_before = [x[p, d, "飯前"] for p in members if "姊妹" in str(gender_map.get(p, ""))]
        
        model.Add(sum(x[p, d, "飯前"] for p in members) == b_count)
        b_all_male = model.NewBoolVar(f'before_male_{d.strftime("%Y%m%d")}')
        model.Add(sum(males_before) == b_count).OnlyEnforceIf(b_all_male)
        model.Add(sum(females_before) == 0).OnlyEnforceIf(b_all_male)
        model.Add(sum(females_before) == b_count).OnlyEnforceIf(b_all_male.Not())
        model.Add(sum(males_before) == 0).OnlyEnforceIf(b_all_male.Not())

        model.Add(sum(x[p, d, "飯後"] for p in members) >= a_min)
        model.Add(sum(x[p, d, "飯後"] for p in members) <= a_max)

        # 🟢 新增：飯後任一性別人數不可剛好為 1 人 (允許 0 人，或至少 2 人以上)
        males_after = [x[p, d, "飯後"] for p in members if "弟兄" in str(gender_map.get(p, ""))]
        females_after = [x[p, d, "飯後"] for p in members if "姊妹" in str(gender_map.get(p, ""))]
        model.Add(sum(males_after) != 1)
        model.Add(sum(females_after) != 1)

        for p in members:
            model.Add(x[p, d, "飯前"] + x[p, d, "飯後"] <= 1)

    for p in members:
        for d in dates:
            if p in join_dates and d < join_dates[p]:
                model.Add(x[p, d, "飯前"] == 0)
                model.Add(x[p, d, "飯後"] == 0)
            if p in leave_dates and d > leave_dates[p]:
                model.Add(x[p, d, "飯前"] == 0)
                model.Add(x[p, d, "飯後"] == 0)

    # 精準次數微調與公平性計算
    for p, adj in custom_adjustments.items():
        if p in members:
            b_adj = adj.get("飯前", 0)
            a_adj = adj.get("飯後", 0)
            
            if b_adj != 0:
                target_b = max(0, math.floor(avg_before + b_adj))
                model.Add(sum(x[p, d, "飯前"] for d in dates) <= target_b)
                
            if a_adj != 0:
                target_a = max(0, math.floor(avg_after + a_adj))
                model.Add(sum(x[p, d, "飯後"] for d in dates) <= target_a)

    adjusted_scores = []
    for p in members:
        actual_score = sum(x[p, d, "飯前"] * 2 + x[p, d, "飯後"] * 1 for d in dates)
        
        expected_deduction = 0
        if p in custom_adjustments:
            b_adj = custom_adjustments[p].get("飯前", 0)
            a_adj = custom_adjustments[p].get("飯後", 0)
            expected_deduction = (abs(b_adj) * 2 + abs(a_adj) * 1) if (b_adj < 0 or a_adj < 0) else 0

        virtual_score = actual_score + expected_deduction
        adjusted_scores.append(virtual_score)

    max_s, min_s = model.NewIntVar(0, 100, 'max_s'), model.NewIntVar(0, 100, 'min_s')
    model.AddMaxEquality(max_s, adjusted_scores)
    model.AddMinEquality(min_s, adjusted_scores)

    pref_score_terms = []
    weekday_str_map = {0: "週一", 1: "週二", 2: "週三", 3: "週四", 4: "週五"}
    for p in members:
        for d in dates:
            w_str = weekday_str_map[d.weekday()]
            if w_str in pref_map[p]["飯前"]: pref_score_terms.append(x[p, d, "飯前"] * 3)
            if w_str in pref_map[p]["飯後"]: pref_score_terms.append(x[p, d, "飯後"] * 2)

    model.Minimize((max_s - min_s) * 100 - sum(pref_score_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 15.0
    status = solver.Solve(model)

    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "排班總表"
        
        font_bold = Font(name="微軟正黑體", size=10, bold=True)
        font_regular = Font(name="微軟正黑體", size=10)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Side(style='thin', color='000000')
        border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
        
        fill_header = PatternFill("solid", fgColor="D9EA88")
        fill_week = PatternFill("solid", fgColor="C9DAF8")
        fill_date = PatternFill("solid", fgColor="D9EAD3")
        fill_before = PatternFill("solid", fgColor="FCE5CD")
        fill_after = PatternFill("solid", fgColor="FFF2CC")
        
        active_w_names = [name for name, w_idx in weekday_options.items() if w_idx in active_weekdays]
        header = ["週別", "時段"] + active_w_names
        ws.append(header)
        for col in range(1, len(header) + 1):
            cell = ws.cell(1, col)
            cell.fill, cell.font, cell.alignment, cell.border = fill_header, font_bold, align_center, border_all
            
        grouped_dates = {}
        for d in dates:
            w_num = d.isocalendar()[1]
            if w_num not in grouped_dates: grouped_dates[w_num] = {}
            grouped_dates[w_num][d.weekday()] = d
            
        curr_row = 2
        week_count = 1
        for w_num, d_dict in grouped_dates.items():
            ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row+2, end_column=1)
            w_cell = ws.cell(curr_row, 1, f"第 {week_count} 週")
            w_cell.fill, w_cell.font, w_cell.alignment = fill_week, font_bold, align_center
            
            ws.cell(curr_row, 2, "日期").fill = fill_date
            for idx, w_idx in enumerate(active_weekdays):
                dt = d_dict.get(w_idx)
                ws.cell(curr_row, 3 + idx, dt.strftime("%m月%d日") if dt else "-").fill = fill_week
                
            ws.cell(curr_row+1, 2, "飯前").fill = fill_before
            for idx, w_idx in enumerate(active_weekdays):
                dt = d_dict.get(w_idx)
                names = "、".join([p for p in members if dt and solver.Value(x[p, dt, "飯前"]) == 1])
                ws.cell(curr_row+1, 3 + idx, names)
                
            ws.cell(curr_row+2, 2, "飯後").fill = fill_after
            for idx, w_idx in enumerate(active_weekdays):
                dt = d_dict.get(w_idx)
                names = "、".join([p for p in members if dt and solver.Value(x[p, dt, "飯後"]) == 1])
                ws.cell(curr_row+2, 3 + idx, names)
            
            for r in range(curr_row, curr_row+3):
                for c in range(1, len(header) + 1):
                    cell = ws.cell(r, c)
                    cell.font = font_bold if c <= 2 or r == curr_row else font_regular
                    cell.alignment, cell.border = align_center, border_all
                    
            curr_row += 3
            week_count += 1
            
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 10
        for col_idx in range(3, len(header) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 30

        ws_stats = wb.create_sheet(title="個人權重統計")
        ws_stats.append(["姓名", "身份", "飯前次數 (2分)", "飯後次數 (1分)", "總權重積分", "偏好滿足次數"])
        for col in range(1, 7):
            ws_stats.cell(1, col).fill, ws_stats.cell(1, col).font = fill_header, font_bold
            
        stats_data = []
        for p in members:
            b_cnt = sum(solver.Value(x[p, d, "飯前"]) for d in dates)
            a_cnt = sum(solver.Value(x[p, d, "飯後"]) for d in dates)
            tot = b_cnt * 2 + a_cnt * 1
            pref_hit = 0
            weekday_str_map = {0: "週一", 1: "週二", 2: "週三", 3: "週四", 4: "週五"}
            for d in dates:
                w_str = weekday_str_map[d.weekday()]
                if solver.Value(x[p, d, "飯前"]) == 1 and w_str in pref_map[p]["飯前"]: pref_hit += 1
                if solver.Value(x[p, d, "飯後"]) == 1 and w_str in pref_map[p]["飯後"]: pref_hit += 1
            stats_data.append([p, gender_map.get(p, ""), b_cnt, a_cnt, tot, pref_hit])
            
        for r in sorted(stats_data, key=lambda x: x[4], reverse=True):
            ws_stats.append(r)
            
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        return buffer
    else:
        return None

# 主畫面展示
if df_raw is not None:
    if not selected_weekdays:
        st.error("⚠️ 請至少在左側邊欄選擇一個「每週服事日期」！")
    else:
        st.success(f"✅ 成功讀取表單！共 {len(members_list)} 位成員數據。")
        
        st.write("### 📌 當前特殊排班設定概覽")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info("**🛫 提前離台名單：**\n" + ("\n".join([f"- {k}: {v}" for k, v in leave_dates_gui.items()]) if leave_dates_gui else "無"))
        with col2:
            st.info("**🛬 延後加入名單：**\n" + ("\n".join([f"- {k}: {v}" for k, v in join_dates_gui.items()]) if join_dates_gui else "無"))
        with col3:
            adj_info = [f"- {k}: 飯前({v['飯前']:+d}), 飯後({v['飯後']:+d})" for k, v in custom_adjustments_gui.items()]
            st.info("**⚖️ 次數微調名單：**\n" + ("\n".join(adj_info) if adj_info else "無"))
        
        if st.button("🚀 開始自動排班", type="primary"):
            with st.spinner("演算法正在計算最佳且公平的排班組合..."):
                excel_data = run_scheduler(
                    df_raw, start_date, end_date, selected_weekdays, selected_holidays, 
                    before_count, after_min, after_max, 
                    leave_dates_gui, join_dates_gui, custom_adjustments_gui
                )
                
                if excel_data:
                    st.balloons()
                    st.success("🎉 排班完成！請下載 Excel：")
                    st.download_button(
                        label="📥 下載彩色版排班 Excel 檔案",
                        data=excel_data,
                        file_name="學期飯食服事排班結果_彩色版.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 無法找到符合限制條件的排班組合，請嘗試調整次數限制或人數設定。")
else:
    st.info("👈 請先於左側邊欄上傳 Google 表單檔案 (CSV 或 Excel) 以開始排班。")
