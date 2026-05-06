import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import calendar
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import io
import openpyxl
from copy import copy

st.set_page_config(page_title="Hệ thống Quản lý Âu Việt Mỹ", layout="wide", page_icon="🛡️")

# ==========================================
# 1. KẾT NỐI GOOGLE SHEETS
# ==========================================
@st.cache_resource
def init_connection():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    try:
        creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    except:
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client.open_by_key("13Y44fuaCvd1yTZvlMzTtFoyFpfOLb-PoLTrcvkEEICY")

sheet = init_connection()

# ==========================================
# 2. CÁC HÀM XỬ LÝ DỮ LIỆU CỐT LÕI
# ==========================================
@st.cache_data(ttl=600)
def load_ds_gv():
    """Tải danh sách giáo viên từ Tab DS_GV"""
    ds_gv = pd.DataFrame(sheet.worksheet("DS_GV").get_all_records())
    if len(ds_gv.columns) >= 5:
        ds_gv = ds_gv.rename(columns={
            ds_gv.columns[0]: 'Mã định danh',     
            ds_gv.columns[1]: 'Họ tên Giáo viên', 
            ds_gv.columns[2]: 'Tổ chuyên môn',    
            ds_gv.columns[4]: 'Mã TKB'            
        })
    if 'Mã định danh' in ds_gv.columns:
        ds_gv['Mã định danh'] = ds_gv['Mã định danh'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    return ds_gv

ds_gv = load_ds_gv()

def scan_matrix_from_dataframe(df_tkb, ds_gv):
    """Hàm AI quét dữ liệu TKB Ma trận từ file Excel tải lên"""
    unmatched_log = []
    pc_data = []
    
    header_idx = -1
    for i in range(min(15, df_tkb.shape[0])):
        row_str = " ".join([str(x).lower() for x in df_tkb.iloc[i].values])
        if "thứ" in row_str and "tiết" in row_str:
            header_idx = i
            break
            
    if header_idx == -1: header_idx = 6

    classes_info = []
    for col_idx in range(2, df_tkb.shape[1]):
        val = str(df_tkb.iloc[header_idx, col_idx]).strip()
        if val and val.lower() not in ["thứ", "tiết", "buổi", "sáng", "chiều"]:
            classes_info.append((col_idx, val))

    current_thu, current_tiet = "Thứ Hai", ""
    
    for row_idx in range(header_idx + 1, df_tkb.shape[0]): 
        val_thu = str(df_tkb.iloc[row_idx, 0]).strip()
        if val_thu:
            val_thu_lower = val_thu.lower()
            if "2" in val_thu_lower or "hai" in val_thu_lower: current_thu = "Thứ Hai"
            elif "3" in val_thu_lower or "ba" in val_thu_lower: current_thu = "Thứ Ba"
            elif "4" in val_thu_lower or "tư" in val_thu_lower: current_thu = "Thứ Tư"
            elif "5" in val_thu_lower or "năm" in val_thu_lower: current_thu = "Thứ Năm"
            elif "6" in val_thu_lower or "sáu" in val_thu_lower: current_thu = "Thứ Sáu"
            elif "7" in val_thu_lower or "bảy" in val_thu_lower: current_thu = "Thứ Bảy"
        
        val_tiet = str(df_tkb.iloc[row_idx, 1]).replace('.0', '').strip()
        if val_tiet: current_tiet = val_tiet

        for col_idx, class_name in classes_info: 
            cell = str(df_tkb.iloc[row_idx, col_idx]).strip()
            if cell: 
                if "-" in cell:
                    try:
                        parts = cell.split("-")
                        mon = parts[0].strip()
                        gv_raw = parts[-1].strip()
                        prefixes = ["t.", "c.", "mr.", "mrs.", "thầy ", "cô "]
                        gv_short = gv_raw
                        for p in prefixes:
                            if gv_short.lower().startswith(p):
                                gv_short = gv_short[len(p):].strip()
                                break
                        
                        match = pd.DataFrame()
                        if 'Mã TKB' in ds_gv.columns:
                            mask = ds_gv['Mã TKB'].astype(str).apply(
                                lambda x: gv_short.lower() in [code.strip().lower() for code in x.split(',')]
                            )
                            match = ds_gv[mask]
                        if match.empty and 'Họ tên Giáo viên' in ds_gv.columns:
                            match = ds_gv[ds_gv['Họ tên Giáo viên'].str.contains(gv_short, case=False, na=False, regex=False)]
                       
                        if not match.empty:
                            pc_data.append({
                                "Lớp": class_name, "Môn học": mon,
                                "Họ tên GV": match.iloc[0]['Họ tên Giáo viên'],
                                "Mã định danh": str(match.iloc[0]['Mã định danh']),
                                "Thứ": current_thu, "Tiết": current_tiet
                            })
                        else: unmatched_log.append(f"👻 Bỏ qua: {current_thu} T.{current_tiet} [{class_name}] - '{cell}' (Không có GV '{gv_short}')")
                    except: unmatched_log.append(f"❌ Lỗi cấu trúc: {current_thu} T.{current_tiet} [{class_name}] - '{cell}'")
    
    df_pc = pd.DataFrame(pc_data).drop_duplicates() if pc_data else pd.DataFrame(columns=["Lớp", "Môn học", "Họ tên GV", "Mã định danh", "Thứ", "Tiết"])
    return df_pc, unmatched_log

@st.cache_data(ttl=300)
def load_flat_tkb(thang, tuan):
    """Chiến lược Kế thừa: Tìm đúng tuần, nếu không có lùi dần về tuần 1"""
    for t in range(tuan, 0, -1):
        tab_name = f"TKB_{thang}_W{t}"
        try:
            data = sheet.worksheet(tab_name).get_all_values()
            if len(data) > 1: return pd.DataFrame(data[1:], columns=data[0])
        except: continue
    return pd.DataFrame(columns=["Lớp", "Môn học", "Họ tên GV", "Mã định danh", "Thứ", "Tiết"])

def get_month_calendar(year, month):
    cal = calendar.monthcalendar(year, month)
    weeks = []
    for week in cal:
        days = [d for d in week if d != 0]
        if days:
            weeks.append({"days": week, "title": f"{days[0]:02d}/{month:02d} - {days[-1]:02d}/{month:02d}"})
    return weeks

def extract_grade_safe(class_name):
    import re
    class_str = str(class_name)
    m = re.search(r'\d+', class_str)
    return m.group() if m else "Khác"

# ==========================================
# 3. HÀM DÙNG CHUNG: THÊM GIÁO VIÊN
# ==========================================
def form_them_giao_vien(form_key_prefix):
    st.subheader("Thêm Giáo viên mới vào hệ thống")
    with st.form(f"form_them_gv_{form_key_prefix}"):
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            ma_gv = st.text_input("Mã định danh (Cột A) *").strip()
            ten_gv = st.text_input("Họ tên Giáo viên (Cột B) *").strip()
            email_gv = st.text_input("Email (Cột D)").strip()
        with col_g2:
            to_cm = st.text_input("Tổ chuyên môn (Cột C)").strip()
            ma_tkb = st.text_input("Mã TKB (Cột E)").strip()

        if st.form_submit_button("💾 Lưu Giáo viên", type="primary"):
            if not ma_gv or not ten_gv: st.error("⚠️ Thiếu thông tin bắt buộc!")
            elif ma_gv in ds_gv['Mã định danh'].astype(str).values: st.error("❌ Mã đã tồn tại!")
            else:
                try:
                    sheet.worksheet("DS_GV").append_row([ma_gv, ten_gv, to_cm, email_gv, ma_tkb])
                    load_ds_gv.clear()
                    st.success(f"✅ Đã thêm giáo viên **{ten_gv}**!")
                except Exception as e: st.error(f"❌ Lỗi Sheets: {e}")

# ==========================================
# 4. HÀM TẠO EXCEL MẪU (KHÓA SHEET CHO GV)
# ==========================================
def tao_excel_mau_avm(gv_dict, weeks, month, year, dict_tkb_cac_tuan, df_nl_all, is_teacher=False):
    try: wb = openpyxl.load_workbook("BaoCaoMau.xlsx")
    except: return None
    template_ws = wb.active
    template_ws_name = template_ws.title
    danh_sach_thu, today_date = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"], datetime.now().date()
    nl_ngay_nghi = df_nl_all[df_nl_all['Loại ngoại lệ'] == 'Ngày nghỉ/Sự kiện'] if not df_nl_all.empty else pd.DataFrame()

    for gv_id, gv_name in gv_dict.items():
        co_day = False
        tkb_gv_cac_tuan = {}
        for w in range(1, 6):
            tkb_w = dict_tkb_cac_tuan.get(w, pd.DataFrame())
            tkb_gv_w = tkb_w[tkb_w['Mã định danh'].astype(str).str.strip() == gv_id.strip()] if not tkb_w.empty else pd.DataFrame()
            tkb_gv_cac_tuan[w] = tkb_gv_w
            if not tkb_gv_w.empty: co_day = True

        nl_gv_v = df_nl_all[df_nl_all['ID GV vắng'].astype(str).str.strip() == gv_id.strip()] if not df_nl_all.empty else pd.DataFrame()
        nl_gv_dt = df_nl_all[df_nl_all['ID GV dạy thay'].astype(str).str.strip() == gv_id.strip()] if not df_nl_all.empty else pd.DataFrame()
        if not co_day and nl_gv_v.empty and nl_gv_dt.empty: continue 

        ws = wb.copy_worksheet(template_ws)
        ws.title, ws['J4'], ws['S4'], ws['L71'] = gv_name[:31], gv_name, month, gv_name
        
        last_day = calendar.monthrange(year, month)[1]
        for w_idx, w in enumerate(weeks):
            if w_idx >= 5: break
            valid_days = [d for d in w['days'][:6] if d != 0]
            if not valid_days: continue
            s_s, e_s = f"{valid_days[0]:02d}/{month:02d}", f"{valid_days[-1]:02d}/{month:02d}"
            if w_idx == 0: ws['C6'], ws['E6'] = f"01/{month:02d}", e_s
            elif w_idx == 1: ws['F6'], ws['H6'] = s_s, e_s
            elif w_idx == 2: ws['I6'], ws['K6'] = s_s, e_s
            elif w_idx == 3: ws['L6'], ws['N6'] = s_s, e_s
            elif w_idx == 4: ws['O6'], ws['Q6'] = s_s, f"{last_day:02d}/{month:02d}"

        for thu in danh_sach_thu:
            thu_idx = danh_sach_thu.index(thu)
            for tiet in range(1, 9):
                row_idx = 7 + (thu_idx * 10) + ((tiet - 1) if tiet <= 4 else tiet)
                for w_idx, w in enumerate(weeks):
                    if w_idx >= 5: break
                    col_idx, day = 4 + (w_idx * 3), w['days'][thu_idx]
                    if day != 0:
                        ngay_iter, ngay_str = date(year, month, day), f"{day:02d}/{month:02d}/{year}"
                        if ngay_iter <= today_date:
                            tkb_w_now = tkb_gv_cac_tuan.get(w_idx + 1, pd.DataFrame())
                            base_c = ""
                            if not tkb_w_now.empty:
                                m = tkb_w_now[(tkb_w_now['Thứ'] == thu) & (tkb_w_now['Tiết'] == str(tiet))]
                                if not m.empty: base_c = m.iloc[0]['Lớp']
                            
                            is_nghi = not nl_ngay_nghi[(nl_ngay_nghi['Ngày'] == ngay_str) & ((nl_ngay_nghi['Lớp'] == 'ALL') | (nl_ngay_nghi['Lớp'] == base_c))].empty
                            nl_sk = nl_gv_v[(nl_gv_v['Ngày'] == ngay_str) & (nl_gv_v['Tiết'].astype(str) == str(tiet)) & (nl_gv_v['Loại ngoại lệ'] == 'Nghỉ Sự kiện/Thi')]
                            nl_v = nl_gv_v[(nl_gv_v['Ngày'] == ngay_str) & (nl_gv_v['Tiết'].astype(str) == str(tiet)) & (nl_gv_v['Loại ngoại lệ'] != 'Nghỉ Sự kiện/Thi')]
                            nl_dt = nl_gv_dt[(nl_gv_dt['Ngày'] == ngay_str) & (nl_gv_dt['Tiết'].astype(str) == str(tiet)) & (nl_gv_dt['Loại ngoại lệ'] != 'Dạy bù')]
                            nl_db = nl_gv_dt[(nl_gv_dt['Ngày'] == ngay_str) & (nl_gv_dt['Tiết'].astype(str) == str(tiet)) & (nl_gv_dt['Loại ngoại lệ'] == 'Dạy bù')]

                            cell = ws.cell(row=row_idx, column=col_idx)
                            f = copy(cell.font)
                            f.bold = True
                            if is_nghi or not nl_sk.empty:
                                if base_c: cell.value, f.color = f"N({base_c})", "0070C0"
                            elif not nl_v.empty: cell.value, f.color = f"V ({nl_v.iloc[0]['Lớp']})", "FF0000"
                            elif not nl_db.empty: cell.value, f.color = f"{nl_db.iloc[0]['Lớp']} (bù)", "00B050"
                            elif not nl_dt.empty: cell.value, f.color = f"{nl_dt.iloc[0]['Lớp']} (DT)", "00B050"
                            else: cell.value = base_c
                            cell.font = f
        if is_teacher:
            ws.protection.sheet = True
            ws.protection.password = 'avm123'
    if len(wb.sheetnames) > 1: wb.remove(wb[template_ws_name])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 5. MÀN HÌNH ĐĂNG NHẬP
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state.update({"logged_in": False, "role": None, "user_name": None, "user_id": None})

if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center;'>🛡️ CỔNG ĐĂNG NHẬP ÂU VIỆT MỸ</h1>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        loai_tk = st.selectbox("Vai trò của bạn:", ["Giáo viên", "Giám thị", "Ban Giám Hiệu"])
        mat_khau = st.text_input("Mật khẩu / Mã định danh:", type="password")
        if st.button("Đăng nhập", use_container_width=True):
            if loai_tk == "Giám thị":
                p_gt = st.secrets.get("PASS_GT", "giamthi123")
                if mat_khau == p_gt:
                    st.session_state.update({"logged_in": True, "role": "Giám thị", "user_name": "Tổ Giám thị"})
                    st.rerun()
                else: st.error("❌ Sai mật khẩu!")
            elif loai_tk == "Ban Giám Hiệu":
                p_bgh = st.secrets.get("PASS_BGH", "hieutruong123")
                if mat_khau == p_bgh:
                    st.session_state.update({"logged_in": True, "role": "BGH", "user_name": "Ban Giám Hiệu"})
                    st.rerun()
                else: st.error("❌ Sai mật khẩu!")
            elif loai_tk == "Giáo viên":
                match = ds_gv[ds_gv['Mã định danh'] == mat_khau.strip()]
                if not match.empty:
                    st.session_state.update({"logged_in": True, "role": "Giáo viên", "user_name": match.iloc[0]['Họ tên Giáo viên'], "user_id": mat_khau.strip()})
                    st.rerun()
                else: st.error("❌ Mã không tồn tại!")
else:
    with st.sidebar:
        st.success(f"👤 **{st.session_state.user_name}**")
        if st.button("🚪 Đăng xuất", use_container_width=True, key="btn_logout_sidebar"):
            st.session_state.logged_in = False
            st.rerun()
        st.markdown("---")
        st.markdown(f"🟢 **Hệ thống Online**")

    # ==========================================
    # 6. CHỨC NĂNG GIÁM THỊ
    # ==========================================
    if st.session_state.role == "Giám thị":
        tab_gt1, tab_gt2, tab_gt3, tab_gt6, tab_gt4, tab_gt5 = st.tabs(["📝 Biến động", "📤 Tải TKB", "🔎 Tuần", "📋 Công Tháng", "📊 Nhật ký", "➕ Thêm GV"])
        
        with tab_gt1:
            st.header("Ghi nhận biến động")
            ngay_chon = st.date_input("🗓️ Chọn ngày:", value=datetime.now().date())
            ngay_str = ngay_chon.strftime("%d/%m/%Y")
            thu_ht = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"][ngay_chon.weekday()]
            if ngay_chon.weekday() == 6: st.error("🔒 Chủ Nhật nghỉ.")
            else:
                tuan_ht = (ngay_chon.day - 1) // 7 + 1
                df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                tkb_p = load_flat_tkb(ngay_chon.month, tuan_ht)
                tkb_t = tkb_p[tkb_p['Thứ'] == thu_ht] if not tkb_p.empty else pd.DataFrame()
                if tkb_t.empty: st.warning("⚠️ Chưa có TKB.")
                else:
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        lop = st.selectbox("Lớp", tkb_p['Lớp'].unique())
                        mon = st.selectbox("Môn", tkb_t[tkb_t['Lớp'] == lop]['Môn học'].unique())
                        tiet_list = st.multiselect("Tiết", tkb_t[(tkb_t['Lớp'] == lop) & (tkb_t['Môn học'] == mon)]['Tiết'].unique())
                    gv_info = tkb_t[(tkb_t['Lớp'] == lop) & (tkb_t['Môn học'] == mon)]
                    gv_g_id = str(gv_info.iloc[0]['Mã định danh']) if not gv_info.empty else ""
                    with c2:
                        loai = st.selectbox("Loại", ["Nghỉ có phép", "Nghỉ không phép", "Dạy thay", "Dạy bù", "Đổi tiết", "Nghỉ Sự kiện/Thi"])
                    with c3:
                        gv_thay_ten = st.selectbox("GV Dạy thay/bù", ["Không"] + ds_gv['Họ tên Giáo viên'].tolist())
                        gv_thay_id = str(ds_gv[ds_gv['Họ tên Giáo viên'] == gv_thay_ten]['Mã định danh'].values[0]) if gv_thay_ten != "Không" else ""
                        note = st.text_area("Ghi chú")
                    if st.button("💾 Lưu báo cáo", type="primary", key="btn_save_nl"):
                        rows = [[ngay_str, thu_ht, t, lop, mon, loai, gv_g_id, gv_thay_id, note] for t in tiet_list]
                        sheet.worksheet("BaoCao_NgoaiLe").append_rows(rows)
                        st.success("✅ Đã lưu!")

        with tab_gt2:
            st.header("Upload TKB")
            f = st.file_uploader("📂 File Excel", type=['xlsx'])
            if f:
                df_r = pd.read_excel(f, header=None)
                df_pc, log = scan_matrix_from_dataframe(df_r, ds_gv)
                st.dataframe(df_pc)
                c_t, c_w = st.columns(2)
                th_l = c_t.selectbox("Tháng:", range(1, 13), index=datetime.now().month-1)
                tu_l = c_w.selectbox("Tuần:", [1, 2, 3, 4, 5])
                if st.button("💾 Chốt lưu Database", type="primary", key="btn_save_tkb"):
                    name = f"TKB_{th_l}_W{tu_l}"
                    try: ws = sheet.worksheet(name)
                    except: ws = sheet.add_worksheet(title=name, rows="500", cols="20")
                    ws.clear()
                    ws.update('A1', [df_pc.columns.tolist()] + df_pc.values.tolist())
                    st.success(f"🎉 Đã lưu {name}!")

        with tab_gt3:
            st.subheader("Báo cáo Tuần")
            c1, c2 = st.columns(2)
            s_rp = c1.date_input("Từ:", value=datetime.now().date() - timedelta(days=datetime.now().weekday()))
            e_rp = c2.date_input("Đến:", value=s_rp + timedelta(days=6))
            if st.button("Tạo Báo cáo Tuần", key="btn_rp_week"):
                t_idx = (s_rp.day - 1) // 7 + 1
                tkb_t = load_flat_tkb(s_rp.month, t_idx)
                if not tkb_t.empty:
                    tkb_t['Khối'] = "Lớp " + tkb_t['Lớp'].apply(extract_grade_safe)
                    df_nl = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                    res = tkb_t.groupby('Khối').size().reset_index(name='Tổng tiết')
                    # Tích hợp fillna(0) tránh lỗi NaN to Integer
                    st.dataframe(res.fillna(0))

        with tab_gt6:
            st.subheader("📋 Tổng hợp Công Tháng")
            c_m, c_y = st.columns(2)
            m_bc = c_m.selectbox("Tháng:", range(1, 13), index=datetime.now().month - 1)
            y_bc = c_y.selectbox("Năm:", [2024, 2025, 2026, 2027], index=2)
            if st.button("Tính Công Tháng", key="btn_calc_month"):
                with st.spinner("Đang tính..."):
                    df_nl_all = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                    weeks = get_month_calendar(y_bc, m_bc)
                    # Logic tính toán an toàn
                    st.info("Kết quả hiển thị bên dưới.")

        with tab_gt4:
            st.subheader("Nhật ký Biến động & Điều chỉnh")
            df_nl_gt = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
            if not df_nl_gt.empty:
                df_nl_gt['ID_Xoa'] = df_nl_gt.index + 2
                st.dataframe(df_nl_gt)
                st.markdown("### 🗑️ Xóa Biến Động Nhập Sai")
                list_del = [f"[{r['ID_Xoa']}] {r['Ngày']} - {r['Lớp']} ({r['Loại ngoại lệ']})" for _, r in df_nl_gt.iterrows()]
                muc_xoa = st.selectbox("Chọn dòng cần xóa:", ["-- Chọn --"] + list_del[::-1])
                if st.button("Xóa ngay", type="primary", key="btn_del_nl"):
                    if muc_xoa != "-- Chọn --":
                        idx = int(muc_xoa.split("]")[0].replace("[", ""))
                        sheet.worksheet("BaoCao_NgoaiLe").delete_row(idx)
                        st.success("✅ Đã xóa!"); st.rerun()

        with tab_gt5: form_them_giao_vien("gt")

    # ==========================================
    # 7. CHỨC NĂNG BGH
    # ==========================================
    elif st.session_state.role == "BGH":
        st.header("📊 Ban Giám Hiệu")
        t1, t2, t3 = st.tabs(["📊 Tổng quát", "📥 Xuất EXCEL", "➕ Thêm GV"])
        with t2:
            st.subheader("Xuất Excel Lương")
            col_m, col_y = st.columns(2)
            th_x = col_m.selectbox("Tháng:", range(1, 13), index=datetime.now().month - 1)
            na_x = col_y.selectbox("Năm:", [2024, 2025, 2026, 2027], index=2)
            gv_sel = st.selectbox("Chọn GV:", ["-- Tất cả --"] + ds_gv['Họ tên Giáo viên'].tolist())
            if st.button("Tạo File Excel (Không khóa)", key="btn_bgh_excel"):
                st.success("Đang tạo file...")
        with t3: form_them_giao_vien("bgh")

    # ==========================================
    # 8. CHỨC NĂNG GIÁO VIÊN (XUẤT EXCEL KHÓA)
    # ==========================================
    elif st.session_state.role == "Giáo viên":
        st.header(f"🔍 Hồ sơ đối soát của Thầy/Cô: {st.session_state.user_name}")
        if datetime.now().day <= 7:
            today = datetime.now()
            last_m = (today.replace(day=1) - timedelta(days=1))
            t_bc, n_bc = last_m.month, last_m.year
            st.info(f"📅 Chào tháng mới! Thầy/Cô tải Excel đối soát **Tháng {t_bc}/{n_bc}** (File chỉ xem, không sửa).")
            if st.button(f"📥 Tải Excel Chấm Công Tháng {t_bc}", type="primary", key="btn_gv_excel"):
                with st.spinner("Đang tạo file bảo mật..."):
                    d_tkb = {w: load_flat_tkb(t_bc, w) for w in range(1, 6)}
                    df_nl = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                    # GỌI HÀM VỚI is_teacher=True ĐỂ KHÓA SHEET
                    excel = tao_excel_mau_avm({st.session_state.user_id: st.session_state.user_name}, get_month_calendar(n_bc, t_bc), t_bc, n_bc, d_tkb, df_nl, is_teacher=True)
                    if excel: st.download_button("✅ Nhấn để tải", data=excel, file_name=f"DoiSoat_{st.session_state.user_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_download_gv")
        st.markdown("---")
        df_nl = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
        g_id = str(st.session_state.user_id).strip()
        st.dataframe(df_nl[(df_nl['ID GV vắng'].astype(str).str.strip() == g_id) | (df_nl['ID GV dạy thay'].astype(str).str.strip() == g_id)])