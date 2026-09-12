import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
from ortools.sat.python import cp_model
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="飯食服事自動排班系統", page_icon="🍞", layout="wide")

st.title("🍞 飯食服事自動排班系統")
st.write("歡迎使用！請上傳從 Google 表單下載的原始 CSV 檔案，並在左側調整學期條件與人員離台/請假設定。")

# 1. 檔案上傳
uploaded_file = st.sidebar.file_uploader("📂 上傳 Google 表單 CSV 檔案", type=["csv"])

# 2. 側邊欄條件設定 (GUI 化)
st.sidebar.header("⚙️ 1. 學期與基本設定")

start_date = st.sidebar.date_input("學期開始日期", datetime(2025, 9, 15))
end_date = st.sidebar.date_input("學期結束日期", datetime(2025, 12, 18))

# 國定假日設定
default_holidays = [datetime(2025, 9, 29).date(), datetime(2025, 10, 6).date(), datetime(2025, 10, 10).date()]
selected_holidays = st.sidebar.multiselect(
    "國定假日 / 不排班日期", 
    options=[start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)],
    default=[d for d in default_holidays if start_date <= d <= end_date],
    format_func=lambda d: d.strftime("%Y/%m/%d (%a)")
)

# 人力需求設定
st.sidebar.subheader("👥 人力需求設定")
before_count = st.sidebar.number_input("飯前人數 (固定同性別)", min_value=1, max_value=4, value=2)
after_min = st.sidebar.number_input("飯後最少人數", min_value=1, max_value=8, value=4)
after_max = st.sidebar.number_input("飯後最多人數", min_value=1, max_value=8, value=5)

# --- 人員特殊出勤 GUI 設定區 ---
leave_dates_gui = {}
join_dates_gui = {}

if uploaded_file is not None:
    df_raw = pd.read_csv(uploaded_file)
    members_list = df_raw['姓名 Name'].dropna().unique().tolist()
    
    st.sidebar.markdown("---")
    st.sidebar.header("✈️ 2. 特定人員出勤限制")
    
    default_leaves = {
        "陳佳穎 Elizabeth": datetime(2025, 11, 25).date(),
        "Mia Dumoran": datetime(2025, 11, 27).date(),
        "關佳恩 (Hannah Kwan)": datetime(2025, 12, 1).date()
    }
    default_joins = {
        "Teosaner Yutanesy Iman": datetime(2025, 9, 22).date()
    }

    st.sidebar.subheader("🛫 提前離台人員設定")
    selected_leave_members = st.sidebar.multiselect(
        "選擇離台成員",
        options=members_list,
        default=[m for m in default_leaves.keys() if m in members_list]
    )
    for m in selected_leave_members:
        d_val = default_leaves.get(m, end_date)
        l_date = st.sidebar.date_input(f"【{m}】最後服事/離台日期", value=d_val, key=f"leave_{m}")
        leave_dates_gui[m] = l_date

    st.sidebar.subheader("🛬 延後加入人員設定")
    selected_join_members = st.sidebar.multiselect(
        "選擇延後加入成員",
        options=members_list,
        default=[m for m in default_joins.keys() if m in members_list]
    )
    for m in selected_join_members:
        d_val = default_joins.get(m, start_date)
        j_date = st.sidebar.date_input(f"【{m}】開始可服事日期", value=d_val, key=f"join_{m}")
        join_dates_gui[m] = j_date

# 核心排班函數
def run_scheduler(df, start_date, end_date, holidays_list, b_count, a_min, a_max, leave_map, join_map):
    holidays_dt = [datetime.combine(h, datetime.min.time()) for h in holidays_list]
    leave_dates = {k: datetime.combine(v, datetime.min.time()) for k, v in leave_map.items()}
    join_dates = {k: datetime.combine(v, datetime.min.time()) for k, v in join_map.items()}

    dates = []
    curr = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date, datetime.min.time())
    
    while curr <= end_dt:
        if curr.weekday() in [0, 2, 3] and curr not in holidays_dt:
            dates.append(curr)
        curr += timedelta(days=1)

    members = df['姓名 Name'].dropna().unique().tolist()
    gender_map = dict(zip(df['姓名 Name'], df['弟兄／姊妹 ( Br. / Sr. )']))

    weekday_map = {0: "週一", 2: "週三", 3: "週四"}

    avail_map = {}
    pref_map = {}

    for _, row in df.iterrows():
        name = row['姓名 Name']
        if pd.isna(name): continue
        avail_map[name] = {}
        pref_map[name] = {"飯前": [], "飯後": []}
        
        for d in dates:
            w_str = weekday_map[d.weekday()]
            for s in ["飯前", "飯後"]:
                col_name = [c for c in df.columns if w_str in c and s in c]
                if col_name:
                    val = str(row[col_name[0]])
                    is_available = "1" in val or "可以" in val or "Available" in val
                    avail_map[name][(d, s)] = 1 if is_available else 0
                else:
                    avail_map[name][(d, s)] = 1

        pref_b_col = [c for c in df.columns if "偏好的飯前" in c]
        pref_a_col = [c for c in df.columns if "偏好的飯後" in c]
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

    scores = [sum(x[p, d, "飯前"] * 2 + x[p, d, "飯後"] * 1 for d in dates) for p in members]
    max_s, min_s = model.NewIntVar(0, 100, 'max_s'), model.NewIntVar(0, 100, 'min_s')
    model.AddMaxEquality(max_s, scores)
    model.AddMinEquality(min_s, scores)

    pref_score_terms = []
    for p in members:
        for d in dates:
            w_str = weekday_map[d.weekday()]
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
        
        ws.append(["週別", "時段", "週一", "週三", "週四"])
        for col in range(1, 6):
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
            m_date, w_date, th_date = d_dict.get(0), d_dict.get(2), d_dict.get(3)
            
            ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row+2, end_column=1)
            w_cell = ws.cell(curr_row, 1, f"第 {week_count} 週")
            w_cell.fill, w_cell.font, w_cell.alignment = fill_week, font_bold, align_center
            
            ws.cell(curr_row, 2, "日期").fill = fill_date
            ws.cell(curr_row, 3, m_date.strftime("%m月%d日") if m_date else "-").fill = fill_week
            ws.cell(curr_row, 4, w_date.strftime("%m月%d日") if w_date else "-").fill = fill_week
            ws.cell(curr_row, 5, th_date.strftime("%m月%d日") if th_date else "-").fill = fill_week
            
            ws.cell(curr_row+1, 2, "飯前").fill = fill_before
            ws.cell(curr_row+1, 3, "、".join([p for p in members if m_date and solver.Value(x[p, m_date, "飯前"]) == 1]))
            ws.cell(curr_row+1, 4, "、".join([p for p in members if w_date and solver.Value(x[p, w_date, "飯前"]) == 1]))
            ws.cell(curr_row+1, 5, "、".join([p for p in members if th_date and solver.Value(x[p, th_date, "飯前"]) == 1]))
            
            ws.cell(curr_row+2, 2, "飯後").fill = fill_after
            ws.cell(curr_row+2, 3, "、".join([p for p in members if m_date and solver.Value(x[p, m_date, "飯後"]) == 1]))
            ws.cell(curr_row+2, 4, "、".join([p for p in members if w_date and solver.Value(x[p, w_date, "飯後"]) == 1]))
            ws.cell(curr_row+2, 5, "、".join([p for p in members if th_date and solver.Value(x[p, th_date, "飯後"]) == 1]))
            
            for r in range(curr_row, curr_row+3):
                for c in range(1, 6):
                    cell = ws.cell(r, c)
                    cell.font = font_bold if c <= 2 or r == curr_row else font_regular
                    cell.alignment, cell.border = align_center, border_all
                    
            curr_row += 3
            week_count += 1
            
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 10
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 30
        ws.column_dimensions['E'].width = 30

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
            for d in dates:
                w_str = weekday_map[d.weekday()]
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
if uploaded_file is not None:
    st.success(f"✅ 成功讀取表單！共 {len(df_raw)} 位成員數據。")
    
    st.write("### 📌 當前人員離台 / 延後加入設定概覽")
    col1, col2 = st.columns(2)
    with col1:
        st.info("**🛫 提前離台名單：**\n" + ("\n".join([f"- {k}: {v}" for k, v in leave_dates_gui.items()]) if leave_dates_gui else "無"))
    with col2:
        st.info("**🛬 延後加入名單：**\n" + ("\n".join([f"- {k}: {v}" for k, v in join_dates_gui.items()]) if join_dates_gui else "無"))
    
    if st.button("🚀 開始自動排班", type="primary"):
        with st.spinner("演算法正在計算最佳且公平的排班組合..."):
            excel_data = run_scheduler(
                df_raw, start_date, end_date, selected_holidays, 
                before_count, after_min, after_max, 
                leave_dates_gui, join_dates_gui
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
                st.error("❌ 無法找到符合限制條件的排班組合，請嘗試調整離台日期或人數限制。")
else:
    st.info("👈 請先於左側邊欄上傳 Google 表單 CSV 檔案以開始排班。")
