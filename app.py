import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Hệ thống Quản lý Âu Việt Mỹ", layout="wide", page_icon="🛡️")

# --- KẾT NỐI GOOGLE SHEETS (BẢN CLOUD BẢO MẬT) ---
@st.cache_resource
def init_connection():
    import json
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    # Lấy chìa khóa từ Két sắt bí mật của Streamlit thay vì file cứng
    creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    # Dán lại ID file Sheets của Thầy vào đây
    spreadsheet_id = "13Y44fuaCvd1yTZvlMzTtFoyFpfOLb-PoLTrcvkEEICY" 
    
    sheet = client.open_by_key(spreadsheet_id)
    return sheet

sheet = init_connection()

# --- HÀM ĐỌC DỮ LIỆU TỪ SHEETS ---
def load_master_data():
    ds_gv = pd.DataFrame(sheet.worksheet("DS_GV").get_all_records())
    pc_chuyenmon = pd.DataFrame(sheet.worksheet("PC_ChuyenMon").get_all_records())
    return ds_gv, pc_chuyenmon

ds_gv, pc_chuyenmon = load_master_data()

st.title("🛡️ Hệ thống Quản lý Chấm công & Quỹ lương (BẢN CHÍNH THỨC)")

menu = ["1. Giám thị Chấm công", "2. Giáo viên Đối soát", "3. BGH Xem Báo cáo"]
choice = st.sidebar.selectbox("VAI TRÒ NGƯỜI DÙNG", menu)

# ==========================================
# VAI TRÒ 1: GIÁM THỊ NHẬP LIỆU (Giữ nguyên như bản tối ưu hôm qua)
# ==========================================
if choice == "1. Giám thị Chấm công":
    st.header("📝 Ghi nhận biến động (Ngoại lệ)")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        lop = st.selectbox("Chọn Lớp", pc_chuyenmon['Lớp'].unique())
        danh_sach_mon = pc_chuyenmon[pc_chuyenmon['Lớp'] == lop]['Môn học'].unique()
        mon = st.selectbox("Chọn Môn", danh_sach_mon)
        tiet_list = st.multiselect("Chọn các tiết học", [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], default=[1])
    
    gv_info = pc_chuyenmon[(pc_chuyenmon['Lớp'] == lop) & (pc_chuyenmon['Môn học'] == mon)]
    gv_goc_ten = gv_info['Họ tên GV'].values[0] if not gv_info.empty else ""
    gv_goc_id = gv_info['Mã định danh'].values[0] if not gv_info.empty else ""
    
    with col2:
        st.info(f"GV Phụ trách: **{gv_goc_ten}** (ID: {gv_goc_id})")
        loai = st.selectbox("Loại biến động", ["Nghỉ có phép", "Nghỉ không phép", "Dạy thay", "Đổi tiết"])
    
    with col3:
        gv_thay_ten = st.selectbox("GV Dạy thay (nếu có)", ["Không"] + ds_gv['Họ tên Giáo viên'].tolist())
        note = st.text_area("Ghi chú")
        
        gv_thay_id = ""
        if gv_thay_ten != "Không":
            gv_thay_id = ds_gv[ds_gv['Họ tên Giáo viên'] == gv_thay_ten]['Mã định danh'].values[0]

    if st.button("💾 Lưu báo cáo lên Hệ thống", type="primary"):
        if len(tiet_list) == 0:
            st.warning("⚠️ Thầy/Cô vui lòng chọn ít nhất 1 tiết học trước khi lưu!")
        else:
            with st.spinner("Đang đồng bộ dữ liệu..."):
                now = datetime.now()
                ngay = now.strftime("%d/%m/%Y")
                weekdays = ["Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"]
                thu = weekdays[now.weekday()]
                
                rows_to_add = []
                for tiet in tiet_list:
                    row_data = [ngay, thu, tiet, lop, mon, loai, str(gv_goc_id), str(gv_thay_id), note]
                    rows_to_add.append(row_data)
                
                sheet.worksheet("BaoCao_NgoaiLe").append_rows(rows_to_add)
                tiet_hien_thi = ", ".join(map(str, sorted(tiet_list)))
                st.success(f"✅ Đã tách và ghi nhận thành công {len(tiet_list)} tiết (Tiết: {tiet_hien_thi}) vào Database!")

# ==========================================
# VAI TRÒ 2: GIÁO VIÊN ĐỐI SOÁT (Nâng cấp chọn Tên)
# ==========================================
elif choice == "2. Giáo viên Đối soát":
    st.header("🔍 Tra cứu lịch sử cá nhân")
    
    # Tạo danh sách hiển thị: "Tên Giáo viên - ID"
    ds_gv['HienThi'] = ds_gv['Họ tên Giáo viên'] + " - ID: " + ds_gv['Mã định danh'].astype(str)
    danh_sach_chon = ["-- Vui lòng chọn Tên của Thầy/Cô --"] + ds_gv['HienThi'].tolist()
    
    selected_gv = st.selectbox("Tìm kiếm Tên Giáo viên:", danh_sach_chon)
    
    if selected_gv != "-- Vui lòng chọn Tên của Thầy/Cô --":
        # Tách lấy cái ID nằm ở cuối chuỗi để đi tìm dữ liệu
        gv_id_str = selected_gv.split(" - ID: ")[-1].strip()
        
        with st.spinner("Đang truy xuất dữ liệu từ máy chủ..."):
            df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
            
            if not df_ngoai_le.empty:
                df_ngoai_le['ID GV vắng'] = df_ngoai_le['ID GV vắng'].astype(str)
                df_ngoai_le['ID GV dạy thay'] = df_ngoai_le['ID GV dạy thay'].astype(str)
                
                df_vang = df_ngoai_le[df_ngoai_le['ID GV vắng'] == gv_id_str].copy()
                if not df_vang.empty:
                    df_vang['Vai trò'] = "Vắng mặt (-)"
                    
                df_thay = df_ngoai_le[df_ngoai_le['ID GV dạy thay'] == gv_id_str].copy()
                if not df_thay.empty:
                    df_thay['Vai trò'] = "Dạy thay (+)"
                
                df_ketqua = pd.concat([df_vang, df_thay])
                
                if not df_ketqua.empty:
                    st.success(f"Tìm thấy {len(df_ketqua)} bản ghi liên quan.")
                    st.dataframe(df_ketqua[['Ngày', 'Thứ', 'Tiết', 'Lớp', 'Môn', 'Vai trò', 'Loại ngoại lệ', 'Ghi chú']], use_container_width=True)
                else:
                    st.info("🎉 Tuyệt vời! Tháng này Thầy/Cô đảm bảo 100% công giảng dạy, không có biến động.")
            else:
                st.info("Hệ thống hiện chưa có dữ liệu ngoại lệ nào.")

# ==========================================
# VAI TRÒ 3: BGH XEM BÁO CÁO (Dashboard 3 Tầng + Lọc Thời gian + Drill-down)
# ==========================================
elif choice == "3. BGH Xem Báo cáo":
    st.header("📊 Bảng điều khiển dành cho Ban Giám Hiệu")
    
    with st.spinner("Đang tổng hợp số liệu toàn trường..."):
        df_ngoai_le = pd.DataFrame(sheet.worksheet("BaoCao_NgoaiLe").get_all_records())
        
        if df_ngoai_le.empty:
            st.info("Trường đang hoạt động ổn định, chưa phát sinh ca vắng/dạy thay nào.")
        else:
            # --- 1. XÂY DỰNG BỘ LỌC THỜI GIAN ---
            df_ngoai_le['Ngày chuẩn'] = pd.to_datetime(df_ngoai_le['Ngày'], format='%d/%m/%Y', errors='coerce')
            min_date = df_ngoai_le['Ngày chuẩn'].min().date()
            max_date = df_ngoai_le['Ngày chuẩn'].max().date()
            
            st.markdown("### 🗓️ BỘ LỌC THỜI GIAN")
            date_range = st.date_input("Chọn khoảng thời gian cần xem báo cáo:", value=(min_date, max_date))
            
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_date, end_date = date_range
            elif isinstance(date_range, tuple) and len(date_range) == 1:
                start_date = end_date = date_range[0]
            else:
                start_date = end_date = date_range
                
            mask = (df_ngoai_le['Ngày chuẩn'].dt.date >= start_date) & (df_ngoai_le['Ngày chuẩn'].dt.date <= end_date)
            df_filtered = df_ngoai_le.loc[mask].copy()
            
            st.markdown("---")
            
            if df_filtered.empty:
                st.warning("📭 Không có dữ liệu ngoại lệ nào trong khoảng thời gian Thầy/Cô vừa chọn.")
            else:
                # Dịch ID sang Tên và Tổ
                dict_gv_ten = pd.Series(ds_gv['Họ tên Giáo viên'].values, index=ds_gv['Mã định danh'].astype(str)).to_dict()
                dict_gv_to = pd.Series(ds_gv['Tổ chuyên môn'].values, index=ds_gv['Mã định danh'].astype(str)).to_dict()
                
                df_filtered['Giáo viên Vắng'] = df_filtered['ID GV vắng'].astype(str).map(dict_gv_ten).fillna("Không rõ")
                df_filtered['Tổ Vắng'] = df_filtered['ID GV vắng'].astype(str).map(dict_gv_to).fillna("Không rõ")
                df_filtered['Giáo viên Dạy thay'] = df_filtered['ID GV dạy thay'].astype(str).map(dict_gv_ten).fillna("Không có")
                df_filtered['Tổ Dạy thay'] = df_filtered['ID GV dạy thay'].astype(str).map(dict_gv_to).fillna("Không có")

                # --- 2. TẠO 3 TAB HIỂN THỊ CHUYÊN SÂU ---
                tab1, tab2, tab3 = st.tabs(["📊 1. TỔNG QUÁT", "🏢 2. THEO TỔ CHUYÊN MÔN", "👩‍🏫 3. THEO GIÁO VIÊN"])
                
                # TAB 1: TỔNG QUÁT TOÀN TRƯỜNG
                with tab1:
                    tong_su_co = len(df_filtered)
                    so_ca_day_thay = len(df_filtered[df_filtered['ID GV dạy thay'] != ''])
                    so_ca_bo_trong = tong_su_co - so_ca_day_thay
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Tổng số tiết báo vắng", tong_su_co, "tiết")
                    col2.metric("Số tiết đã có người Dạy thay", so_ca_day_thay, "đảm bảo tiến độ", delta_color="normal")
                    col3.metric("Số tiết Lớp tự học (Trống)", so_ca_bo_trong, "cần lưu ý", delta_color="inverse")
                    
                    st.subheader("Nhật ký biến động chi tiết")
                    st.dataframe(df_filtered[['Ngày', 'Tiết', 'Lớp', 'Môn', 'Giáo viên Vắng', 'Loại ngoại lệ', 'Giáo viên Dạy thay', 'Ghi chú']], use_container_width=True)
                
                # TAB 2: THỐNG KÊ & TRA CỨU THEO TỔ
                with tab2:
                    st.info("Bảng tổng hợp: Đánh giá nhanh tình hình nhân sự các Tổ.")
                    df_to_vang = df_filtered.groupby('Tổ Vắng').size().reset_index(name='Số tiết Vắng')
                    df_to_thay = df_filtered[df_filtered['Tổ Dạy thay'] != "Không có"].groupby('Tổ Dạy thay').size().reset_index(name='Số tiết Dạy thay')
                    
                    df_to_tonghop = pd.merge(df_to_vang, df_to_thay, left_on='Tổ Vắng', right_on='Tổ Dạy thay', how='outer').fillna(0)
                    df_to_tonghop['Tổ chuyên môn'] = df_to_tonghop['Tổ Vắng'].combine_first(df_to_tonghop['Tổ Dạy thay'])
                    df_to_tonghop = df_to_tonghop[['Tổ chuyên môn', 'Số tiết Vắng', 'Số tiết Dạy thay']]
                    df_to_tonghop['Số tiết Vắng'] = df_to_tonghop['Số tiết Vắng'].astype(int)
                    df_to_tonghop['Số tiết Dạy thay'] = df_to_tonghop['Số tiết Dạy thay'].astype(int)
                    
                    st.dataframe(df_to_tonghop, use_container_width=True)
                    
                    # TÍNH NĂNG MỚI: DRILL-DOWN THEO TỔ
                    st.markdown("---")
                    st.markdown("#### 🔎 Tra cứu chi tiết sự cố theo Tổ")
                    danh_sach_to = sorted(list(set(df_filtered['Tổ Vắng'].dropna().tolist() + df_filtered['Tổ Dạy thay'].dropna().tolist())))
                    danh_sach_to = [t for t in danh_sach_to if t not in ["Không có", "Không rõ"]]
                    
                    to_duoc_chon = st.selectbox("Chọn Tổ chuyên môn để xem chi tiết:", ["-- Vui lòng chọn Tổ --"] + danh_sach_to)
                    if to_duoc_chon != "-- Vui lòng chọn Tổ --":
                        df_chi_tiet_to = df_filtered[(df_filtered['Tổ Vắng'] == to_duoc_chon) | (df_filtered['Tổ Dạy thay'] == to_duoc_chon)]
                        st.dataframe(df_chi_tiet_to[['Ngày', 'Tiết', 'Lớp', 'Môn', 'Giáo viên Vắng', 'Loại ngoại lệ', 'Giáo viên Dạy thay', 'Ghi chú']], use_container_width=True)
                
                # TAB 3: THỐNG KÊ & TRA CỨU THEO GIÁO VIÊN
                with tab3:
                    st.info("Bảng tổng hợp: Căn cứ đối soát Quỹ lương cuối tháng.")
                    df_gv_vang = df_filtered.groupby('Giáo viên Vắng').size().reset_index(name='Tổng tiết Vắng (-)')
                    df_gv_thay = df_filtered[df_filtered['Giáo viên Dạy thay'] != "Không có"].groupby('Giáo viên Dạy thay').size().reset_index(name='Tổng tiết Dạy thay (+)')
                    
                    df_gv_tonghop = pd.merge(df_gv_vang, df_gv_thay, left_on='Giáo viên Vắng', right_on='Giáo viên Dạy thay', how='outer').fillna(0)
                    df_gv_tonghop['Họ tên Giáo viên'] = df_gv_tonghop['Giáo viên Vắng'].combine_first(df_gv_tonghop['Giáo viên Dạy thay'])
                    df_gv_tonghop = df_gv_tonghop[['Họ tên Giáo viên', 'Tổng tiết Vắng (-)', 'Tổng tiết Dạy thay (+)']]
                    df_gv_tonghop['Tổng tiết Vắng (-)'] = df_gv_tonghop['Tổng tiết Vắng (-)'].astype(int)
                    df_gv_tonghop['Tổng tiết Dạy thay (+)'] = df_gv_tonghop['Tổng tiết Dạy thay (+)'].astype(int)
                    
                    st.dataframe(df_gv_tonghop, use_container_width=True)
                    
                    # TÍNH NĂNG MỚI: DRILL-DOWN THEO GIÁO VIÊN (Bao gồm Tổng hợp Khối)
                    st.markdown("---")
                    st.markdown("#### 🔎 Tra cứu & Tổng hợp Lương theo Giáo viên")
                    
                    ds_gv['HienThi_BGH'] = ds_gv['Họ tên Giáo viên'] + " - ID: " + ds_gv['Mã định danh'].astype(str)
                    danh_sach_chon_bgh = ["-- Vui lòng chọn Tên/Mã định danh Giáo viên --"] + ds_gv['HienThi_BGH'].tolist()
                    
                    gv_duoc_chon = st.selectbox("Tìm kiếm Tên hoặc Mã định danh:", danh_sach_chon_bgh)
                    
                    if gv_duoc_chon != "-- Vui lòng chọn Tên/Mã định danh Giáo viên --":
                        gv_id_str = gv_duoc_chon.split(" - ID: ")[-1].strip()
                        
                        df_vang = df_filtered[df_filtered['ID GV vắng'].astype(str) == gv_id_str].copy()
                        if not df_vang.empty: df_vang['Vai trò'] = "Vắng mặt (-)"
                            
                        df_thay = df_filtered[df_filtered['ID GV dạy thay'].astype(str) == gv_id_str].copy()
                        if not df_thay.empty: df_thay['Vai trò'] = "Dạy thay (+)"
                        
                        df_chi_tiet_gv = pd.concat([df_vang, df_thay])
                        
                        if not df_chi_tiet_gv.empty:
                            
                            # =========================================================
                            # BẢNG 1: TỔNG HỢP THEO KHỐI (Giống ảnh 2 của Thầy)
                            # =========================================================
                            st.markdown("##### 💰 Bảng tổng hợp số tiết theo Khối (Căn cứ nhân đơn giá)")
                            
                            # Thuật toán tự bóc tách số đầu tiên của Lớp để tìm Khối (VD: 10A1 -> L10)
                            df_chi_tiet_gv['Khối'] = "L" + df_chi_tiet_gv['Lớp'].astype(str).str.extract(r'^(\d+)')[0].fillna("Khác")
                            
                            # Xoay bảng (Pivot) để các Khối nằm ngang
                            pt_khoi = pd.pivot_table(df_chi_tiet_gv, values='Tiết', index='Vai trò', columns='Khối', aggfunc='count', fill_value=0)
                            
                            # Thêm cột TỔNG ở đầu tiên
                            pt_khoi.insert(0, 'TỔNG CỘNG', pt_khoi.sum(axis=1))
                            
                            st.dataframe(pt_khoi, use_container_width=True)
                            
                            # =========================================================
                            # BẢNG 2: CHI TIẾT TỪNG SỰ CỐ (Giống ảnh 1 của Thầy)
                            # =========================================================
                            st.markdown("##### 📝 Lịch sử chi tiết")
                            st.dataframe(df_chi_tiet_gv[['Ngày', 'Tiết', 'Lớp', 'Môn', 'Vai trò', 'Loại ngoại lệ', 'Ghi chú']], use_container_width=True)
                            
                        else:
                            st.success("Giáo viên này không có ca vắng/dạy thay nào trong khoảng thời gian đã chọn.")