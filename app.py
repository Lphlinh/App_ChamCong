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
# 2. HÀM QUÉT TKB MA TRẬN (Fix lỗi đọc thiếu ngày Thứ Hai + Nhận diện Lớp động)
# ==========================================
@st.cache_data(ttl=600) 
def load_master_data(selected_month=None):
    ds_gv = pd.DataFrame(sheet.worksheet("DS_GV").get_all_records())
    
    # Ép tên cột theo vị trí để chống lỗi Unicode từ GSheets
    if len(ds_gv.columns) >= 5:
        ds_gv = ds_gv.rename(columns={
            ds_gv.columns[0]: 'Mã định danh',     
            ds_gv.columns[1]: 'Họ tên Giáo viên', 
            ds_gv.columns[2]: 'Tổ chuyên môn',    
            ds_gv.columns[4]: 'Mã TKB'            
        })

    if 'Mã định danh' in ds_gv.columns:
        ds_gv['Mã định danh'] = ds_gv['Mã định danh'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    unmatched_log = []
    target_ws_name = f"TKB_PhanCong_{selected_month}" if selected_month else "TKB_PhanCong"

    try:
        try:
            tkb_raw = sheet.worksheet(target_ws_name).get_all_values()
        except:
            tkb_raw = sheet.worksheet("TKB_PhanCong").get_all_values()
            
        df_tkb = pd.DataFrame(tkb_raw)
        
        # --- ĐIỂM NÂNG CẤP: Tự động lấy tên Lớp từ Dòng 1 (Index 0) ---
        classes = []
        for col_idx in range(2, df_tkb.shape[1]):
            classes.append(str(df_tkb.iloc[0, col_idx]).strip())
            
        pc_data = []
        current_thu, current_tiet = "Thứ Hai", ""
        
        # --- ĐIỂM NÂNG CẤP: Bắt đầu quét từ Dòng 2 (Index 1) thay vì Dòng 8 ---
        for row_idx in range(1, df_tkb.shape[0]): 
            val_thu = str(df_tkb.iloc[row_idx, 0]).strip().lower() if df_tkb.shape[1] > 0 else ""
            if val_thu:
                if "2" in val_thu or "hai" in val_thu: current_thu = "Thứ Hai"
                elif "3" in val_thu or "ba" in val_thu: current_thu = "Thứ Ba"
                elif "4" in val_thu or "tư" in val_thu: current_thu = "Thứ Tư"
                elif "5" in val_thu or "năm" in val_thu: current_thu = "Thứ Năm"
                elif "6" in val_thu or "sáu" in val_thu: current_thu = "Thứ Sáu"
                elif "7" in val_thu or "bảy" in val_thu: current_thu = "Thứ Bảy"
            
            val_tiet = str(df_tkb.iloc[row_idx, 1]).replace('.0', '').strip() if df_tkb.shape[1] > 1 else ""
            if val_tiet: current_tiet = val_tiet
  
            for col_idx, class_name in enumerate(classes, start=2): 
                if col_idx < df_tkb.shape[1]:
                    cell = str(df_tkb.iloc[row_idx, col_idx]).strip()
                    if cell: 
                        if "-" in cell:
                            try:
                                parts = cell.split("-")
                                if len(parts) >= 2:
                                    mon = parts[0].strip()
                                    gv_raw = parts[-1].strip() # Lấy phần tử sau cùng cho an toàn
                                    gv_raw_lower = gv_raw.lower()
                                    
                                    if gv_raw_lower.startswith(("t.", "c.", "mr.", "mrs.")):
                                        if gv_raw_lower.startswith("t.") or gv_raw_lower.startswith("c."):
                                            gv_short = gv_raw[2:].strip()
                                        elif gv_raw_lower.startswith("mr."):
                                            gv_short = gv_raw[3:].strip()
                                        elif gv_raw_lower.startswith("mrs."):
                                            gv_short = gv_raw[4:].strip()
                                        
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
                                        else:
                                            unmatched_log.append(f"⚠️ {current_thu} - Tiết {current_tiet} - Lớp {class_name}: Không tìm thấy GV '{gv_short}'")
                                    else:
                                        unmatched_log.append(f"👻 Bỏ qua: {current_thu} - Tiết {current_tiet} - Lớp {class_name} ('{cell}' thiếu T./C./Mr./Mrs.)")
                            except Exception as e:
                                unmatched_log.append(f"❌ {current_thu} - Tiết {current_tiet} - Lớp {class_name}: Lỗi cấu trúc '{cell}'")
                        else:
                            unmatched_log.append(f"👻 Bỏ qua: {current_thu} - Tiết {current_tiet} - Lớp {class_name} ('{cell}' thiếu gạch nối)")
        
        pc_chuyenmon = pd.DataFrame(pc_data).drop_duplicates() if pc_data else pd.DataFrame(columns=["Lớp", "Môn học", "Họ tên GV", "Mã định danh", "Thứ", "Tiết"])
    except Exception as e:
        pc_chuyenmon = pd.DataFrame(columns=["Lớp", "Môn học", "Họ tên GV", "Mã định danh", "Thứ", "Tiết"])
  
    return ds_gv, pc_chuyenmon, unmatched_log

ds_gv, pc_chuyenmon, unmatched_log = load_master_data()

def get_month_calendar(year, month):
    cal = calendar.monthcalendar(year, month)
    weeks = []
    for week in cal:
        days = [d for d in week if d != 0]
        if days:
            start_date = f"{days[0]:02d}/{month:02d}"
            end_date = f"{days[-1]:02d}/{month:02d}"
            weeks.append({"days": week, "title": f"{start_date} - {end_date}"})
    return weeks

# ==========================================
# 3. MÀN HÌNH ĐĂNG NHẬP
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
                try: pass_gt = st.secrets["PASS_GT"]
                except: pass_gt = "giamthi123"
                if mat_khau == pass_gt:
                    st.session_state.update({"logged_in": True, "role": "Giám thị", "user_name": "Tổ Giám thị"})
                    st.rerun()
                else: st.error("❌ Sai mật khẩu Giám thị!")
            elif loai_tk == "Ban Giám Hiệu":
                try: pass_bgh = st.secrets["PASS_BGH"]
                except: pass_bgh = "hieutruong123"
                if mat_khau == pass_bgh:
                    st.session_state.update({"logged_in": True, "role": "BGH", "user_name": "Ban Giám Hiệu"})
                    st.rerun()
                else: st.error("❌ Sai mật khẩu Ban Giám Hiệu!")
            elif loai_tk == "Giáo viên":
                gv_match = ds_gv[ds_gv['Mã định danh'] == mat_khau.strip()]
                if not gv_match.empty:
                    st.session_state.update({"logged_in": True, "role": "Giáo viên", 
                                             "user_name": gv_match.iloc[0]['Họ tên Giáo viên'], "user_id": mat_khau.strip()})
                    st.rerun()
                else: st.error("❌ Mã định danh không tồn tại trong hệ thống!")
else:
    with st.sidebar:
        st.success(f"👤 **{st.session_state.user_name}**")
        if st.button("🚪 Đăng xuất", use_container_width=True):
            st.session_state.logged_in = False
            st.rerun()
  
        st.markdown("---")
        st.markdown(f"🟢 **Hệ thống Online**<br><small>Lần kết nối: {datetime.now().strftime('%H:%M:%S %d/%m')}</small>", unsafe_allow_html=True)

    # ==========================================
    # 4. CHỨC NĂNG GIÁM THỊ
    # ==========================================
    if st.session_state.role == "Giám thị":
        tab_gt1, tab_gt2 = st.tabs(["📝 Ghi nhận biến động", "🔎 Báo cáo Cuối tuần"])
        
        with tab_gt1:
            st.header("Ghi nhận sự cố")
            col_date, _ = st.columns([1, 2])
            with col_date:
                ngay_chon = st.date_input("🗓️ Chọn ngày ghi nhận:", value=datetime.now().date())
                ngay_str = ngay_chon.strftime("%d/%m/%Y")
            thu_hien_tai = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"][ngay_chon.weekday()]
            
            if ngay_chon.weekday() == 6:
                st.error("🔒 HỆ THỐNG ĐÃ KHÓA SỔ. Hôm nay là Chủ Nhật, không thể cập nhật dữ liệu.")
            else:
                with st.spinner("Đang tải dữ liệu Thời khóa biểu..."):
                    df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                    df_today = df_ngoai_le[df_ngoai_le['Ngày'] == ngay_str] if not df_ngoai_le.empty else pd.DataFrame()
                    tkb_today = pc_chuyenmon[pc_chuyenmon['Thứ'] == thu_hien_tai]

                if pc_chuyenmon.empty:
                    st.error("❌ App chưa nhận diện được dữ liệu từ TKB. Vui lòng báo BGH kiểm tra lại file Google Sheets.")
                else:
                    st.markdown("---")
                    col1, col2, col3 = st.columns(3)
                    classes = ["KHTN", "KHXH", "11A", "11C", "10A", "10C", "9A", "8A", "7A", "6A"]
                    with col1:
                        lop = st.selectbox("Lớp", classes)
                        mon_hople = tkb_today[tkb_today['Lớp'] == lop]['Môn học'].dropna().unique().tolist()
                    if not mon_hople:
                        st.warning(f"📭 Vào {thu_hien_tai} ({ngay_str}), lớp {lop} KHÔNG CÓ lịch học theo TKB.")
                    else:
                        with col1:
                            mon = st.selectbox("Môn", mon_hople)
                            tiet_hop_le = sorted([str(t) for t in tkb_today[(tkb_today['Lớp'] == lop) & (tkb_today['Môn học'] == mon)]['Tiết'].dropna().unique().tolist()])
                            tiet_list = st.multiselect("Chọn Tiết", options=tiet_hop_le, default=tiet_hop_le)

                        gv_info = tkb_today[(tkb_today['Lớp'] == lop) & (tkb_today['Môn học'] == mon)]
                        gv_goc_ten = gv_info.iloc[0]['Họ tên GV'] if not gv_info.empty else "N/A"
                        gv_goc_id = str(gv_info.iloc[0]['Mã định danh']) if not gv_info.empty else ""
                        with col2:
                            st.info(f"GV Phụ trách: **{gv_goc_ten}**")
                            loai = st.selectbox("Loại", ["Nghỉ có phép", "Nghỉ không phép", "Dạy thay", "Đổi tiết", "Nghỉ Sự kiện/Thi"])
                        
                        gv_ban_list = []
                        for t in tiet_list:
                            gv_ban_list.extend(tkb_today[tkb_today['Tiết'] == str(t)]['Mã định danh'].astype(str).tolist())
                            if not df_today.empty:
                                ca_nay = df_today[df_today['Tiết'].astype(str) == str(t)]
                                gv_ban_list.extend(ca_nay['ID GV vắng'].astype(str).tolist())
                                gv_ban_list.extend(ca_nay['ID GV dạy thay'].astype(str).tolist())
                        
                        if gv_goc_id: gv_ban_list.append(gv_goc_id) 
                        gv_ban_list = list(set([x for x in gv_ban_list if x != ""]))
                        df_gv_ranh = ds_gv[~ds_gv['Mã định danh'].astype(str).isin(gv_ban_list)]
                        danh_sach_day_thay = ["Không"] + df_gv_ranh['Họ tên Giáo viên'].tolist()

                        with col3:
                            gv_thay_ten = st.selectbox("GV Dạy thay (Hệ thống ẩn GV đang bận)", danh_sach_day_thay)
                            gv_thay_id = str(ds_gv[ds_gv['Họ tên Giáo viên'] == gv_thay_ten]['Mã định danh'].values[0]) if gv_thay_ten != "Không" else ""
                            note = st.text_area("Ghi chú")

                        if st.button("💾 Lưu báo cáo", type="primary"):
                            if len(tiet_list) == 0:
                                st.warning("⚠️ Vui lòng chọn ít nhất 1 tiết học trước khi lưu!")
                            else:
                                with st.spinner("Đang ghi nhận dữ liệu..."):
                                    rows_to_add = [[ngay_str, thu_hien_tai, t, lop, mon, loai, gv_goc_id, gv_thay_id, note] for t in tiet_list]
                                    sheet.worksheet("BaoCao_NgoaiLe").append_rows(rows_to_add)
                                    st.success(f"✅ Đã ghi nhận thành công cho ngày {ngay_str}!")

                    st.markdown("---")
                    with st.expander("🏖️ Khai báo Ngày Nghỉ / Sự kiện (Không tính vắng)", expanded=False):
                        st.info("Dùng để ghi nhận ngày nghỉ lễ, ngày thi... Các tiết thuộc phạm vi này sẽ được in **N(Lớp)** trên Excel.")
                        col_n1, col_n2 = st.columns(2)
                        pham_vi = col_n1.selectbox("Phạm vi nghỉ:", ["Toàn trường", "Chọn lớp cụ thể"])
                        lop_nghi = "ALL"
                        if pham_vi == "Chọn lớp cụ thể":
                            lop_nghi = col_n1.selectbox("Chọn Lớp nghỉ:", classes, key="lop_nghi_cb")
                        
                        ly_do = col_n2.text_input("Ghi chú (Tên ngày lễ, sự kiện...):")
                        if st.button("💾 Lưu Ngày Nghỉ", type="primary", key="btn_luu_nghi"):
                            row_meta = [ngay_str, thu_hien_tai, "ALL", lop_nghi, "ALL", "Ngày nghỉ/Sự kiện", "ALL", "", ly_do]
                            sheet.worksheet("BaoCao_NgoaiLe").append_rows([row_meta])
                            st.success(f"✅ Đã lưu ngày {ngay_str} là Ngày nghỉ ({pham_vi})!")

        with tab_gt2:
            st.subheader("Báo cáo Kiểm dò chéo Sổ đầu bài")
            col_d1, col_d2 = st.columns(2)
            with col_d1: start_rp = st.date_input("Từ ngày:", value=datetime.now().date() - timedelta(days=datetime.now().weekday()))
            with col_d2: end_rp = st.date_input("Đến ngày:", value=start_rp + timedelta(days=6))
            
            if st.button("Tạo Báo cáo Tuần", type="primary"):
                with st.spinner("Đang tính toán số liệu tuần..."):
                    if pc_chuyenmon.empty:
                        st.error("❌ Hệ thống chưa nhận diện được dữ liệu từ Thời khóa biểu gốc.")
                    else:
                        df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
                        
                        tkb_tuan = pc_chuyenmon.copy()
                        tkb_tuan['Khối'] = "Lớp " + tkb_tuan['Lớp'].astype(str).str.extract(r'^(\d+)')[0].fillna("Khác")
                        rp_tkb = tkb_tuan.groupby('Khối').size().reset_index(name='Tổng TKB phải dạy')
                        
                        if not df_ngoai_le.empty:
                            df_ngoai_le['Ngày chuẩn'] = pd.to_datetime(df_ngoai_le['Ngày'], format='%d/%m/%Y', errors='coerce')
                            mask_rp = (df_ngoai_le['Ngày chuẩn'].dt.date >= start_rp) & (df_ngoai_le['Ngày chuẩn'].dt.date <= end_rp)
                            df_rp = df_ngoai_le.loc[mask_rp].copy()
                        else:
                            df_rp = pd.DataFrame()
                        
                        if not df_rp.empty:
                            df_rp['Khối'] = "Lớp " + df_rp['Lớp'].astype(str).str.extract(r'^(\d+)')[0].fillna("Khác")
                            mask_su_co = ~df_rp['Loại ngoại lệ'].isin(['Ngày nghỉ/Sự kiện', 'Nghỉ Sự kiện/Thi'])
                            df_rp_loi = df_rp[mask_su_co]
                            
                            rp_vang = df_rp_loi[df_rp_loi['ID GV vắng'].astype(str).str.strip() != ""].groupby('Khối').size().reset_index(name='Số tiết Nghỉ (Vắng)')
                            rp_thay = df_rp_loi[df_rp_loi['ID GV dạy thay'].astype(str).str.strip() != ""].groupby('Khối').size().reset_index(name='Số tiết Dạy thay')
                        else:
                            rp_vang = pd.DataFrame(columns=['Khối', 'Số tiết Nghỉ (Vắng)'])
                            rp_thay = pd.DataFrame(columns=['Khối', 'Số tiết Dạy thay'])
                        
                        rp_final = pd.merge(rp_tkb, rp_vang, on='Khối', how='left').fillna(0)
                        rp_final = pd.merge(rp_final, rp_thay, on='Khối', how='left').fillna(0)

                        cols_to_fix = ['Tổng TKB phải dạy', 'Số tiết Nghỉ (Vắng)', 'Số tiết Dạy thay']
                        for col in cols_to_fix:
                            if col not in rp_final.columns:
                                rp_final[col] = 0
                            rp_final[col] = pd.to_numeric(rp_final[col], errors='coerce').fillna(0)
                        
                        rp_final['Tổng Thực Dạy'] = rp_final['Tổng TKB phải dạy'] - rp_final['Số tiết Nghỉ (Vắng)'] + rp_final['Số tiết Dạy thay']
                        
                        tong_cong = rp_final[['Tổng TKB phải dạy', 'Số tiết Nghỉ (Vắng)', 'Số tiết Dạy thay', 'Tổng Thực Dạy']].sum()
                        rp_final.loc['TOÀN TRƯỜNG'] = tong_cong
                        rp_final.at['TOÀN TRƯỜNG', 'Khối'] = "TOÀN TRƯỜNG"
                        
                        for col in ['Tổng TKB phải dạy', 'Số tiết Nghỉ (Vắng)', 'Số tiết Dạy thay', 'Tổng Thực Dạy']:
                            rp_final[col] = rp_final[col].astype(int)

                        st.dataframe(rp_final, use_container_width=True)
                        st.success("✅ Đã tạo xong báo cáo tuần!")

    # ==========================================
    # 5. CHỨC NĂNG GIÁO VIÊN
    # ==========================================
    elif st.session_state.role == "Giáo viên":
        st.header(f"🔍 Hồ sơ đối soát của Thầy/Cô: {st.session_state.user_name}")
        with st.spinner("Đang truy xuất dữ liệu cá nhân..."):
            df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
            if not df_ngoai_le.empty:
                gv_id_str = str(st.session_state.user_id).strip()
                df_vang = df_ngoai_le[df_ngoai_le['ID GV vắng'].astype(str).str.strip() == gv_id_str].copy()
                if not df_vang.empty: 
                    df_vang['Vai trò'] = df_vang['Loại ngoại lệ'].apply(lambda x: "Sự kiện/Thi (Không trừ)" if x == "Nghỉ Sự kiện/Thi" else "Vắng mặt (-)")
                df_thay = df_ngoai_le[df_ngoai_le['ID GV dạy thay'].astype(str).str.strip() == gv_id_str].copy()
                if not df_thay.empty: df_thay['Vai trò'] = "Dạy thay (+)"
                
                df_ketqua = pd.concat([df_vang, df_thay])
                if not df_ketqua.empty:
                    st.dataframe(df_ketqua[['Ngày', 'Thứ', 'Tiết', 'Lớp', 'Môn', 'Vai trò', 'Loại ngoại lệ']], use_container_width=True)
                else: st.info("🎉 Tuyệt vời! Thầy/Cô đảm bảo 100% công giảng dạy.")
            else: st.info("Hệ thống hiện chưa có dữ liệu biến động nào.")

    # ==========================================
    # 6. CHỨC NĂNG BGH (DASHBOARD & EXCEL TEMPLATE)
    # ==========================================
    elif st.session_state.role == "BGH":
        st.header("📊 Bảng điều khiển dành cho Ban Giám Hiệu")
        
        with st.spinner("Đang tải dữ liệu toàn trường..."):
            df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
            if df_ngoai_le.empty:
                min_date, max_date = datetime.now().date(), datetime.now().date()
                df_filtered = pd.DataFrame(columns=['Ngày', 'Thứ', 'Tiết', 'Lớp', 'Môn', 'Loại ngoại lệ', 'ID GV vắng', 'ID GV dạy thay', 'Ghi chú'])
            else:
                df_ngoai_le['Ngày chuẩn'] = pd.to_datetime(df_ngoai_le['Ngày'], format='%d/%m/%Y', errors='coerce')
                min_date = df_ngoai_le['Ngày chuẩn'].min().date()
                max_date = df_ngoai_le['Ngày chuẩn'].max().date()
                date_range = st.date_input("🗓️ Chọn khoảng thời gian xem báo cáo:", value=(min_date, max_date))
                
                if isinstance(date_range, tuple) and len(date_range) == 2: start_date, end_date = date_range
                elif isinstance(date_range, tuple) and len(date_range) == 1: start_date = end_date = date_range[0]
                else: start_date = end_date = date_range
                    
                mask = (df_ngoai_le['Ngày chuẩn'].dt.date >= start_date) & (df_ngoai_le['Ngày chuẩn'].dt.date <= end_date)
                df_filtered = df_ngoai_le.loc[mask].copy()

            dict_gv_ten = pd.Series(ds_gv['Họ tên Giáo viên'].values, index=ds_gv['Mã định danh'].astype(str).str.strip()).to_dict()
            dict_gv_to = pd.Series(ds_gv['Tổ chuyên môn'].values, index=ds_gv['Mã định danh'].astype(str).str.strip()).to_dict()
            
            if not df_filtered.empty:
                id_vang_clean = df_filtered['ID GV vắng'].astype(str).str.strip()
                id_thay_clean = df_filtered['ID GV dạy thay'].astype(str).str.strip()
                
                df_filtered['Giáo viên Vắng'] = id_vang_clean.map(dict_gv_ten).fillna("Không rõ")
                df_filtered['Tổ Vắng'] = id_vang_clean.map(dict_gv_to).fillna("Không rõ")
                df_filtered['Giáo viên Dạy thay'] = id_thay_clean.map(dict_gv_ten).fillna("Không có")
                df_filtered['Tổ Dạy thay'] = id_thay_clean.map(dict_gv_to).fillna("Không có")

        tab1, tab2, tab3, tab4 = st.tabs(["📊 Tổng quát", "🏢 Theo Tổ", "📥 Xuất EXCEL (Chấm Công)", "🗂️ TKB Đã Quét"])
        
        with tab1:
            st.subheader("Nhật ký Biến động Tổng quát")
            mask_metrics = ~df_filtered['Loại ngoại lệ'].isin(['Ngày nghỉ/Sự kiện', 'Nghỉ Sự kiện/Thi'])
            df_metrics = df_filtered[mask_metrics]
            
            tong_su_co = len(df_metrics)
            so_ca_day_thay = len(df_metrics[df_metrics['ID GV dạy thay'].astype(str).str.strip() != '']) if not df_metrics.empty else 0
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Tổng tiết báo vắng", tong_su_co)
            col2.metric("Số tiết đã Dạy thay", so_ca_day_thay, delta_color="normal")
            col3.metric("Số tiết Lớp tự học", tong_su_co - so_ca_day_thay, delta_color="inverse")
            if not df_filtered.empty:
                st.dataframe(df_filtered[['Ngày', 'Thứ', 'Tiết', 'Lớp', 'Môn', 'Loại ngoại lệ', 'Giáo viên Vắng', 'Giáo viên Dạy thay', 'Ghi chú']], use_container_width=True)
            else: st.info("🎉 Trường đang hoạt động ổn định, chưa phát sinh ca vắng/dạy thay nào.")
        
        with tab2:
            st.subheader("Thống kê theo Tổ Chuyên môn")
            if not df_filtered.empty:
                df_to = df_filtered[~df_filtered['Loại ngoại lệ'].isin(['Ngày nghỉ/Sự kiện', 'Nghỉ Sự kiện/Thi'])]
                s_vang = df_to[df_to['Tổ Vắng'] != "Không rõ"].groupby('Tổ Vắng').size().rename('Số tiết Vắng')
                s_thay = df_to[df_to['Tổ Dạy thay'] != "Không có"].groupby('Tổ Dạy thay').size().rename('Số tiết Dạy thay')
                
                df_to_tonghop = pd.concat([s_vang, s_thay], axis=1).fillna(0).astype(int)
                df_to_tonghop.index.name = 'Tổ chuyên môn'
                df_to_tonghop = df_to_tonghop.reset_index()
                
                st.dataframe(df_to_tonghop[['Tổ chuyên môn', 'Số tiết Vắng', 'Số tiết Dạy thay']], use_container_width=True)
            else: st.info("📭 Không có dữ liệu để thống kê theo tổ.")

        with tab4:
            st.subheader("Dữ liệu Thời Khóa Biểu đang áp dụng")
            if unmatched_log:
                with st.expander("⚠️ Cảnh báo: Có lỗi trong quá trình tự động quét TKB", expanded=True):
                    st.error("Các ô TKB sau đây bị sai cú pháp hoặc ký hiệu Giáo viên không khớp với DS_GV:")
                    for log in unmatched_log: st.write(log)
            
            st.markdown("---")
            st.subheader("Lưu trữ TKB theo tháng")
            thang_luu = st.selectbox("Chọn tháng muốn lưu TKB hiện tại:", range(1, 13), index=datetime.now().month-1)
            if st.button(f"📤 Lưu dữ liệu vào tab TKB_PhanCong_{thang_luu}", type="primary"):
                with st.spinner(f"Đang tạo/cập nhật tab TKB_PhanCong_{thang_luu}..."):
                    try:
                        tab_name = f"TKB_PhanCong_{thang_luu}"
                        try:
                            ws_target = sheet.worksheet(tab_name)
                        except:
                            ws_target = sheet.add_worksheet(title=tab_name, rows="100", cols="20")
                        
                        ws_target.clear()
                        df_up = pc_chuyenmon.astype(str).fillna("")
                        data = [df_up.columns.tolist()] + df_up.values.tolist()
                        ws_target.update('A1', data)
                        st.success(f"✅ Đã lưu trữ TKB tháng {thang_luu} thành công!")
                    except Exception as e:
                        st.error(f"❌ Lỗi lưu trữ: {e}")
                        
            st.markdown("---")
            if pc_chuyenmon.empty: st.warning("⚠️ Hệ thống chưa nhận diện được dữ liệu TKB hợp lệ.")
            else: st.dataframe(pc_chuyenmon, use_container_width=True)

        with tab3:
            st.subheader("Tạo Bảng Chấm Công Lương (Sử dụng Template Mẫu)")
            col_m, col_y = st.columns(2)
            with col_m: thang_xuat = st.selectbox("Chọn Tháng:", range(1, 13), index=datetime.now().month - 1)
            with col_y: nam_xuat = st.selectbox("Chọn Năm:", [2024, 2025, 2026, 2027], index=2)
            
            _, tkb_thang_chon, _ = load_master_data(thang_xuat)
            
            weeks = get_month_calendar(nam_xuat, thang_xuat)
            ds_gv['HienThi_BGH'] = ds_gv['Họ tên Giáo viên'] + " - ID: " + ds_gv['Mã định danh'].astype(str)
            gv_chon = st.selectbox("Chọn Giáo viên để xuất Excel:", ["-- Chọn Giáo viên --"] + ds_gv['HienThi_BGH'].tolist())
            
            def tao_excel_mau_avm(gv_dict, weeks, month, year, df_tkb_all, df_nl_all):
                try: wb = openpyxl.load_workbook("BaoCaoMau.xlsx")
                except FileNotFoundError: return None

                template_ws = wb.active
                template_ws_name = template_ws.title
                danh_sach_thu = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy"]
                today_date = datetime.now().date()
                nl_ngay_nghi = df_nl_all[df_nl_all['Loại ngoại lệ'] == 'Ngày nghỉ/Sự kiện']

                for gv_id, gv_name in gv_dict.items():
                    tkb_gv = df_tkb_all[df_tkb_all['Mã định danh'].astype(str).str.strip() == gv_id.strip()]
                    nl_gv_v = df_nl_all[df_nl_all['ID GV vắng'].astype(str).str.strip() == gv_id.strip()]
                    nl_gv_dt = df_nl_all[df_nl_all['ID GV dạy thay'].astype(str).str.strip() == gv_id.strip()]

                    if tkb_gv.empty and nl_gv_v.empty and nl_gv_dt.empty: continue 

                    ws = wb.copy_worksheet(template_ws)
                    ws.title = gv_name[:31]
                    ws['J4'], ws['S4'], ws['L71'] = gv_name, month, gv_name

                    last_day_of_month = calendar.monthrange(year, month)[1]
                    for w_idx, w in enumerate(weeks):
                        if w_idx >= 5: break 
                        valid_days = [d for d in w['days'][:6] if d != 0] 
                        if not valid_days: continue
                        
                        start_str = f"{valid_days[0]:02d}/{month:02d}"
                        end_str = f"{valid_days[-1]:02d}/{month:02d}"

                        if w_idx == 0: ws['C6'], ws['E6'] = f"01/{month:02d}", end_str           
                        elif w_idx == 1: ws['F6'], ws['H6'] = start_str, end_str           
                        elif w_idx == 2: ws['I6'], ws['K6'] = start_str, end_str           
                        elif w_idx == 3: ws['L6'], ws['N6'] = start_str, end_str           
                        elif w_idx == 4: ws['O6'], ws['Q6'] = start_str, f"{last_day_of_month:02d}/{month:02d}" 

                    for thu in danh_sach_thu:
                        thu_idx = danh_sach_thu.index(thu)
                        tkb_thu = tkb_gv[tkb_gv['Thứ'] == thu]

                        for tiet in range(1, 9): 
                            row_idx = 7 + (thu_idx * 10) + ((tiet - 1) if tiet <= 4 else tiet)
                            base_class = ""
                            tkb_match = tkb_thu[tkb_thu['Tiết'] == str(tiet)]
                            if not tkb_match.empty: base_class = tkb_match.iloc[0]['Lớp']

                            for w_idx, w in enumerate(weeks):
                                if w_idx >= 5: break
                                col_idx = 4 + (w_idx * 3) 
                                day = w['days'][thu_idx]

                                if day != 0:
                                    ngay_iter = date(year, month, day)
                                    if ngay_iter > today_date:
                                        ws.cell(row=row_idx, column=col_idx, value="")
                                    else:
                                        ngay_str = f"{day:02d}/{month:02d}/{year}"
                                        
                                        is_ngay_nghi = False
                                        if not nl_ngay_nghi.empty:
                                            match_nghi = nl_ngay_nghi[(nl_ngay_nghi['Ngày'] == ngay_str) & 
                                                                      ((nl_ngay_nghi['Lớp'] == 'ALL') | (nl_ngay_nghi['Lớp'] == base_class))]
                                            if not match_nghi.empty: is_ngay_nghi = True
                                        
                                        nl_sk_match = nl_gv_v[(nl_gv_v['Ngày'] == ngay_str) & (nl_gv_v['Tiết'].astype(str) == str(tiet)) & (nl_gv_v['Loại ngoại lệ'] == 'Nghỉ Sự kiện/Thi')]
                                        nl_v_match = nl_gv_v[(nl_gv_v['Ngày'] == ngay_str) & (nl_gv_v['Tiết'].astype(str) == str(tiet)) & (nl_gv_v['Loại ngoại lệ'] != 'Nghỉ Sự kiện/Thi')]
                                        nl_dt_match = nl_gv_dt[(nl_gv_dt['Ngày'] == ngay_str) & (nl_gv_dt['Tiết'].astype(str) == str(tiet))]
                                        
                                        target_cell = ws.cell(row=row_idx, column=col_idx)
                                        
                                        new_font = copy(target_cell.font)
                                        new_font.bold = True
                                        new_align = copy(target_cell.alignment)
                                        new_align.wrap_text, new_align.shrink_to_fit = False, False
                                        
                                        if is_ngay_nghi or not nl_sk_match.empty:
                                            if base_class:
                                                target_cell.value, new_font.color = f"N({base_class})", "0070C0"
                                                target_cell.font, target_cell.alignment = new_font, new_align
                                        elif not nl_v_match.empty:
                                            target_cell.value, new_font.color = f"{nl_v_match.iloc[0]['Lớp']} (V)", "FF0000"
                                            target_cell.font, target_cell.alignment = new_font, new_align
                                        elif not nl_dt_match.empty:
                                            target_cell.value, new_font.color = f"{nl_dt_match.iloc[0]['Lớp']} (DT)", "00B050"
                                            target_cell.font, target_cell.alignment = new_font, new_align
                                        else:
                                            if base_class: target_cell.value = base_class
                                else:
                                    ws.cell(row=row_idx, column=col_idx, value="")

                if len(wb.sheetnames) > 1: wb.remove(wb[template_ws_name])
                else:
                    template_ws.title, template_ws['A1'] = "KhongCoDuLieu", "Giáo viên không có phân công."

                output = io.BytesIO()
                wb.save(output)
                return output.getvalue()

            st.markdown("---")
            if gv_chon == "-- Chọn Giáo viên --":
                st.info("👆 Chọn Giáo viên để tải Bảng Chấm Công cá nhân, hoặc tải Bảng Chấm Công Toàn Trường bên dưới.")
                
            col_ex1, col_ex2 = st.columns(2)
            df_nl_full = df_ngoai_le.copy() if not df_ngoai_le.empty else pd.DataFrame(columns=['Ngày', 'Thứ', 'Tiết', 'Lớp', 'Môn', 'Loại ngoại lệ', 'ID GV vắng', 'ID GV dạy thay', 'Ghi chú'])
            
            with col_ex1:
                if gv_chon != "-- Chọn Giáo viên --":
                    if st.button(f"📥 Tải Excel CÁ NHÂN ({gv_chon.split(' - ')[0]})", type="primary"):
                        with st.spinner("Đang tạo Excel..."):
                            gv_id_str = gv_chon.split(" - ID: ")[-1].strip()
                            gv_name_str = gv_chon.split(" - ID: ")[0].strip()
                            excel_data = tao_excel_mau_avm({gv_id_str: gv_name_str}, weeks, thang_xuat, nam_xuat, tkb_thang_chon, df_nl_full)
                            if excel_data: st.download_button("✅ Tải File", data=excel_data, file_name=f"ChamCong_{gv_name_str}_{thang_xuat}.xlsx")
            
            with col_ex2:
                if st.button("📥 Tải Excel TOÀN TRƯỜNG", type="primary"):
                    with st.spinner("Đang tổng hợp..."):
                        gv_dict_all = {str(row['Mã định danh']): row['Họ tên Giáo viên'] for _, row in ds_gv.iterrows()}
                        excel_data_all = tao_excel_mau_avm(gv_dict_all, weeks, thang_xuat, nam_xuat, tkb_thang_chon, df_nl_full)
                        if excel_data_all: st.download_button("✅ Tải File", data=excel_data_all, file_name=f"ChamCong_ToanTruong_T{thang_xuat}.xlsx")