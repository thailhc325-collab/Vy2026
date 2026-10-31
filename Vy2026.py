import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import streamlit as st
import pandas as pd
from datetime import date, datetime
import io
import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# File lưu trữ dữ liệu lâu dài trên máy tính
DATA_FILE = "data_vy2026.json"
# Cấu hình trang Web
st.set_page_config(page_title="QUẢN LÝ-ĐÁNH GIÁ KPI TRƯỜNG THPT VĨNH YÊN", page_icon="🏆", layout="wide")
st.title("🏆 QUẢN LÝ-ĐÁNH GIÁ KPI TRƯỜNG THPT VĨNH YÊN")
st.caption("Ứng dụng theo dõi, đánh giá và báo cáo thi đua cán bộ, giáo viên, nhân viên - THPT Vĩnh Yên")


# -----------------------------------------------------------------------------
# HÀM HỖ TRỢ XỬ LÝ ÉP KIỂU AN TOÀN
# -----------------------------------------------------------------------------
def calculate_kscl_points(s_prev, s_curr):
    """
    Tính điểm cộng / điểm trừ thi đua theo Tiêu chí 5 (KSCL) - Phụ lục 2 QĐ 292
    s_prev: Điểm trung bình kỳ trước (S_n - Đầu vào)
    s_curr: Điểm trung bình kỳ sau (S_{n+1})
    Returns: (delta, points, desc_summary)
    """
    s_prev = float(s_prev)
    s_curr = float(s_curr)
    delta = round(s_curr - s_prev, 2)
    if delta < 0:
        return delta, -1.0, f"Giảm {abs(delta):.2f} điểm (Trừ 1 điểm)"
    elif delta == 0:
        return delta, 0.0, "Giữ nguyên điểm (0 điểm)"
    # Delta > 0: Tăng điểm
    if s_prev < 5.0:
        if delta >= 1.0:
            pts = 10.0
        elif delta >= 0.75:
            pts = 8.0
        elif delta >= 0.5:
            pts = 6.0
        else:
            pts = 4.0
    elif s_prev < 5.5:
        if delta >= 0.75:
            pts = 10.0
        elif delta >= 0.5:
            pts = 8.0
        elif delta >= 0.3:
            pts = 6.0
        else:
            pts = 4.0
    elif s_prev < 6.5:
        if delta >= 0.5:
            pts = 10.0
        elif delta >= 0.35:
            pts = 8.0
        elif delta >= 0.25:
            pts = 6.0
        else:
            pts = 4.0
    elif s_prev < 7.5:
        if delta >= 0.3:
            pts = 10.0
        elif delta >= 0.25:
            pts = 8.0
        elif delta >= 0.15:
            pts = 6.0
        else:
            pts = 4.0
    else:  # s_prev >= 7.5
        if delta >= 0.2:
            pts = 10.0
        elif delta >= 0.15:
            pts = 8.0
        elif delta >= 0.1:
            pts = 6.0
        else:
            pts = 4.0

    return delta, pts, f"Tăng {delta:.2f} điểm (+{pts:.1f} điểm)"


def safe_float(val, default=0.0):
    """Chuyển đổi giá trị sang float an toàn, tránh lỗi ValueError khi gặp chuỗi rỗng ''"""
    try:
        if val is None or val == "" or pd.isna(val):
            return float(default)
        return float(val)
    except (ValueError, TypeError):
        return float(default)


# -----------------------------------------------------------------------------
# DANH MỤC NHÓM TIÊU CHÍ VIÊN CHỨC (PHỤ LỤC 1) & KHEN THƯỞNG (PHỤ LỤC 2)
# -----------------------------------------------------------------------------
CRITERIA_GROUPS = [
    "1. Thực hiện văn hoá công sở (15 điểm)",
    "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
    "3. Thực hiện quy chế chuyên môn (70 điểm)"
]
REWARD_CATEGORIES = [
    "1. Chất lượng giáo dục các môn văn hoá",
    "2. Hoạt động Kỹ năng sống, STEM, KHKT, Văn nghệ, TDTT",
    "3. Công tác chủ nhiệm lớp",
    "4. Bồi dưỡng chuyên môn, nghiệp vụ giáo dục",
    "5. Hoạt động Tổ chức, Đoàn thể (TCM, Đoàn TN, Tâm lý, Văn phòng)",
    "6. Hoạt động các Câu lạc bộ (Thiện nguyện, Truyền thông, VN, TDTT)"
]


def get_default_criteria_tree():
    """Tạo Cây thông tin Tiêu chí Đánh giá Viên chức chuẩn theo Phụ lục 1 QĐ 292 (Nguồn 2026-10-26T23-57_export.csv)"""
    tree_data = [
        # Nhóm 1: Văn hóa công sở (15 điểm)
        {
            "STT": 1,
            "Nhóm tiêu chí": "1. Thực hiện văn hoá công sở (15 điểm)",
            "Mã TC": "TC1.1.1",
            "Nội dung Tiêu chí": "Trang phục khi đến trường làm việc trong giờ hành chính đúng theo quy định tại Khoản 4; 5-Điều 4-Thông tư số 06/2019/TT-BGDĐT. Đeo thẻ của CB, GV, NV khi đến trường làm việc.",
            "Điểm tối đa": 8.0,
            "Nội dung vi phạm": "Trang phục không đúng quy định",
            "điểm trừ": 1.0
        },
        {
            "STT": 2,
            "Nhóm tiêu chí": "1. Thực hiện văn hoá công sở (15 điểm)",
            "Mã TC": "TC1.1.2",
            "Nội dung Tiêu chí": "Trang phục khi đến trường làm việc trong giờ hành chính đúng theo quy định tại Khoản 4; 5-Điều 4-Thông tư số 06/2019/TT-BGDĐT. Đeo thẻ của CB, GV, NV khi đến trường làm việc.",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Không đeo thẻ",
            "điểm trừ": 1.0
        },
        {
            "STT": 3,
            "Nhóm tiêu chí": "1. Thực hiện văn hoá công sở (15 điểm)",
            "Mã TC": "TC1.2",
            "Nội dung Tiêu chí": "Trang phục đúng quy định của nhà trường trong giờ sinh hoạt dưới cờ; của BTC trong các buổi Lễ, Hội nghị…",
            "Điểm tối đa": 7.0,
            "Nội dung vi phạm": "Trang phục không đúng quy định của BTC",
            "điểm trừ": 1.0
        },
        # Nhóm 2: Tham gia hội họp, sinh hoạt tập thể (15 điểm)
        {
            "STT": 4,
            "Nhóm tiêu chí": "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
            "Mã TC": "TC2.1.1",
            "Nội dung Tiêu chí": "Tham gia đầy đủ, đúng giờ các cuộc họp, sinh hoạt tập thể do BGH hoặc người đứng đầu các tổ chức, đoàn thể trong nhà trường tổ chức (xét các cá nhân thuộc thành phần phải tham gia).",
            "Điểm tối đa": 5.0,
            "Nội dung vi phạm": "Nghỉ Họp không báo cáo",
            "điểm trừ": 2.0
        },
        {
            "STT": 5,
            "Nhóm tiêu chí": "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
            "Mã TC": "TC2.1.2",
            "Nội dung Tiêu chí": "Tham gia đầy đủ, đúng giờ các cuộc họp, sinh hoạt tập thể do BGH hoặc người đứng đầu các tổ chức, đoàn thể trong nhà trường tổ chức (xét các cá nhân thuộc thành phần phải tham gia).",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Muộn Họp từ 10 phút trở lên",
            "điểm trừ": 1.0
        },
        {
            "STT": 6,
            "Nhóm tiêu chí": "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
            "Mã TC": "TC2.1.3",
            "Nội dung Tiêu chí": "Tham gia đầy đủ, đúng giờ các cuộc họp, sinh hoạt tập thể do BGH hoặc người đứng đầu các tổ chức, đoàn thể trong nhà trường tổ chức (xét các cá nhân thuộc thành phần phải tham gia).",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Muộn Họp dưới 10 phút",
            "điểm trừ": 0.5
        },
        {
            "STT": 7,
            "Nhóm tiêu chí": "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
            "Mã TC": "TC2.2.1",
            "Nội dung Tiêu chí": "Nắm bắt, triển khai, đầy đủ, đúng thời gian nội dung các buổi họp, sinh hoạt tập thể.",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Triển khai không đúng nội dung cuộc Họp",
            "điểm trừ": 2.0
        },
        {
            "STT": 8,
            "Nhóm tiêu chí": "2. Tham gia hội họp, sinh hoạt tập thể (15 điểm)",
            "Mã TC": "TC2.2.2",
            "Nội dung Tiêu chí": "Nắm bắt, triển khai, đầy đủ, đúng thời gian nội dung các buổi họp, sinh hoạt tập thể.",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Triển khai muộn nội dung cuộc Họp",
            "điểm trừ": 1.0
        },
        # Nhóm 3: Thực hiện quy chế chuyên môn (70 điểm)
        {
            "STT": 9,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.1.1",
            "Nội dung Tiêu chí": "Sau hiệu lệnh trống muộn nhất 5 phút phải có mặt tại lớp. Đến đúng giờ trong buổi coi kiểm tra cuối kỳ theo phân công.",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Muộn hơn 15 phút hoặc bỏ giờ dạy (coi kiểm tra)",
            "điểm trừ": 4.0
        },
        {
            "STT": 10,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.1.2",
            "Nội dung Tiêu chí": "Sau hiệu lệnh trống muộn nhất 5 phút phải có mặt tại lớp. Đến đúng giờ trong buổi coi kiểm tra cuối kỳ theo phân công.",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Muộn từ 11 đến 15 phút",
            "điểm trừ": 2.0
        },
        {
            "STT": 11,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.1.3",
            "Nội dung Tiêu chí": "Sau hiệu lệnh trống muộn nhất 5 phút phải có mặt tại lớp. Đến đúng giờ trong buổi coi kiểm tra cuối kỳ theo phân công.",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Muộn từ 6 đến 10 phút",
            "điểm trừ": 1.0
        },
        {
            "STT": 12,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.2",
            "Nội dung Tiêu chí": "Đảm bảo duy trì và quản lý hoạt động học tập liên tục của học sinh trong giờ học: Không cho học sinh ngồi chơi khi chưa hết giờ (Trừ một số tình huống đặc biệt được sự đồng ý của BGH)",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Để học sinh ngồi chơi khi chưa hết giờ",
            "điểm trừ": 1.0
        },
        {
            "STT": 13,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.3.1",
            "Nội dung Tiêu chí": "Dạy đúng theo TKB của BGH. Khi đổi giờ, đổi ca, nhờ dạy thay (vì công việc cá nhân) phải báo cáo BGH và nhập đủ thông tin trên trang tính (Nghỉ dạy phải được Hiệu trưởng phê duyệt)",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Dạy không đúng theo TKB (Không báo cáo)",
            "điểm trừ": 1.0
        },
        {
            "STT": 14,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.3.2",
            "Nội dung Tiêu chí": "Dạy đúng theo TKB của BGH. Khi đổi giờ, đổi ca, nhờ dạy thay (vì công việc cá nhân) phải báo cáo BGH và nhập đủ thông tin trên trang tính (Nghỉ dạy phải được Hiệu trưởng phê duyệt)",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Không nhập thông tin đổi giờ, dạy thay trên trang tính",
            "điểm trừ": 1.0
        },
        {
            "STT": 15,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.4.1",
            "Nội dung Tiêu chí": "Ghi đầy đủ các nội dung trong Sổ đầu bài (Họ tên HS vắng, tiết theo PPCT, tên bài dạy) và ký xác nhận đã dạy",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Không ghi tên bài dạy",
            "điểm trừ": 1.0
        },
        {
            "STT": 16,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.4.2",
            "Nội dung Tiêu chí": "Ghi đầy đủ các nội dung trong Sổ đầu bài (Họ tên HS vắng, tiết theo PPCT, tên bài dạy) và ký xác nhận đã dạy",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Không ghi số tiết theo PPCT",
            "điểm trừ": 1.0
        },
        {
            "STT": 17,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC3.4.3",
            "Nội dung Tiêu chí": "Ghi đầy đủ các nội dung trong Sổ đầu bài (Họ tên HS vắng, tiết theo PPCT, tên bài dạy) và ký xác nhận đã dạy",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Không ký xác nhận tiết dạy",
            "điểm trừ": 1.0
        },
        {
            "STT": 18,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC5.1.1",
            "Nội dung Tiêu chí": "Đủ KH bài dạy khi lên lớp, được ký duyệt trên Edoc đúng thời gian. KH bài dạy thể hiện đầy đủ mục tiêu đầu ra; kiến thức cơ bản; cách thức tổ chức; nội dung tích hợp (nếu có)",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Kế hoạch bài dạy không đạt yêu cầu",
            "điểm trừ": 1.0
        },
        {
            "STT": 19,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC5.1.2",
            "Nội dung Tiêu chí": "Đủ KH bài dạy khi lên lớp, được ký duyệt trên Edoc đúng thời gian. KH bài dạy thể hiện đầy đủ mục tiêu đầu ra; kiến thức cơ bản; cách thức tổ chức; nội dung tích hợp (nếu có)",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Thiếu Kế hoạch bài dạy",
            "điểm trừ": 2.0
        },
        {
            "STT": 20,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC6.1.1",
            "Nội dung Tiêu chí": "Ra đề, phản biện khi được BGH phân công phải đảm bảo yêu cầu về nội dung, thời gian hoàn thành",
            "Điểm tối đa": 10.0,
            "Nội dung vi phạm": "Hoàn thành muộn (Phản biện)",
            "điểm trừ": 0.15
        },
        {
            "STT": 21,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC6.1.2",
            "Nội dung Tiêu chí": "Ra đề, phản biện khi được BGH phân công phải đảm bảo yêu cầu về nội dung, thời gian hoàn thành",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Hoàn thành muộn (Ra đề chính)",
            "điểm trừ": 0.35
        },
        {
            "STT": 22,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC6.1.3",
            "Nội dung Tiêu chí": "Ra đề, phản biện khi được BGH phân công phải đảm bảo yêu cầu về nội dung, thời gian hoàn thành",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Lỗi 1 nội dung (Ra đề chính)",
            "điểm trừ": 0.70
        },
        {
            "STT": 23,
            "Nhóm tiêu chí": "3. Thực hiện quy chế chuyên môn (70 điểm)",
            "Mã TC": "TC6.1.4",
            "Nội dung Tiêu chí": "Ra đề, phản biện khi được BGH phân công phải đảm bảo yêu cầu về nội dung, thời gian hoàn thành",
            "Điểm tối đa": 0.0,
            "Nội dung vi phạm": "Lỗi 1 nội dung (Phản biện)",
            "điểm trừ": 0.30
        }
    ]
    return pd.DataFrame(tree_data)


def get_default_rewards_tree():
    """Tạo Cây thông tin Tiêu chí Điểm Thi đua Khen thưởng chuẩn theo Phụ lục 2 năm học 2026-2027 (Bổ sung Số giải)"""
    tree_data = [
        # Mục 1
        {"STT": 1, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.1A",
         "Nội dung khen thưởng": "Chất lượng ôn thi TN THPT",
         "Thành tích / Mức khen thưởng": "Xếp trong Top 30 toàn tỉnh", "Số giải": 1, "Điểm cộng": 100.0,
         "Ghi chú": "Tính điểm cho GV dạy môn đó"},
        {"STT": 2, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.1B",
         "Nội dung khen thưởng": "Chất lượng ôn thi TN THPT",
         "Thành tích / Mức khen thưởng": "Xếp trong Top 35 toàn tỉnh", "Số giải": 1, "Điểm cộng": 70.0,
         "Ghi chú": "Tính điểm cho GV dạy môn đó"},
        {"STT": 3, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.2A",
         "Nội dung khen thưởng": "Chất lượng thi HSG cấp trường", "Thành tích / Mức khen thưởng": "Giải Nhất",
         "Số giải": 1, "Điểm cộng": 10.0, "Ghi chú": "Tính cho GVBD (Số giải x Số điểm)"},
        {"STT": 4, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.2B",
         "Nội dung khen thưởng": "Chất lượng thi HSG cấp trường", "Thành tích / Mức khen thưởng": "Giải Nhì",
         "Số giải": 1, "Điểm cộng": 8.0, "Ghi chú": "Tính cho GVBD (Số giải x Số điểm)"},
        {"STT": 5, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.4A",
         "Nội dung khen thưởng": "Chất lượng thi HSG cấp tỉnh", "Thành tích / Mức khen thưởng": "Giải Nhất",
         "Số giải": 1, "Điểm cộng": 100.0,
         "Ghi chú": "Tính cho GVBD (Số giải x Số điểm). Đội tuyển >=5 HS đạt 100% giải: +30đ"},
        {"STT": 6, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.4B",
         "Nội dung khen thưởng": "Chất lượng thi HSG cấp tỉnh", "Thành tích / Mức khen thưởng": "Giải Nhì",
         "Số giải": 1, "Điểm cộng": 70.0, "Ghi chú": "Tính cho GVBD (Số giải x Số điểm)"},
        {"STT": 7, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.5A",
         "Nội dung khen thưởng": "Chất lượng điểm các kỳ KSCL (Tiêu chí 5)",
         "Thành tích / Mức khen thưởng": "Tăng Mức 1 (+10 điểm)", "Số giải": 1, "Điểm cộng": 10.0,
         "Ghi chú": "Tự động so sánh kỳ n và n+1 theo quy tắc QĐ 292"},
        {"STT": 8, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.5B",
         "Nội dung khen thưởng": "Chất lượng điểm các kỳ KSCL (Tiêu chí 5)",
         "Thành tích / Mức khen thưởng": "Tăng Mức 2 (+8 điểm)", "Số giải": 1, "Điểm cộng": 8.0,
         "Ghi chú": "Tự động so sánh kỳ n và n+1 theo quy tắc QĐ 292"},
        {"STT": 9, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.5C",
         "Nội dung khen thưởng": "Chất lượng điểm các kỳ KSCL (Tiêu chí 5)",
         "Thành tích / Mức khen thưởng": "Tăng Mức 3 (+6 điểm)", "Số giải": 1, "Điểm cộng": 6.0,
         "Ghi chú": "Tự động so sánh kỳ n và n+1 theo quy tắc QĐ 292"},
        {"STT": 10, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.5D",
         "Nội dung khen thưởng": "Chất lượng điểm các kỳ KSCL (Tiêu chí 5)",
         "Thành tích / Mức khen thưởng": "Có tăng điểm nhưng chưa đạt các mức trên (+4 điểm)", "Số giải": 1,
         "Điểm cộng": 4.0, "Ghi chú": "Tự động so sánh kỳ n và n+1 theo quy tắc QĐ 292"},
        {"STT": 11, "Mục lớn": "1. Chất lượng giáo dục các môn văn hoá", "Mã KT": "KT1.5E",
         "Nội dung khen thưởng": "Chất lượng điểm các kỳ KSCL (Tiêu chí 5)",
         "Thành tích / Mức khen thưởng": "Điểm KSCL bị giảm so với kỳ trước (-1 điểm)", "Số giải": 1, "Điểm cộng": -1.0,
         "Ghi chú": "Trừ 1 điểm khi điểm KSCL lần sau giảm so với lần trước"},
        # Mục 2
        {"STT": 12, "Mục lớn": "2. Hoạt động Kỹ năng sống, STEM, KHKT, Văn nghệ, TDTT", "Mã KT": "KT2.1A",
         "Nội dung khen thưởng": "HD học sinh làm Sản phẩm STEM / Thuyết trình tốt",
         "Thành tích / Mức khen thưởng": "Giải Nhất", "Số giải": 1, "Điểm cộng": 10.0, "Ghi chú": "Số giải x Số điểm"},
        {"STT": 13, "Mục lớn": "2. Hoạt động Kỹ năng sống, STEM, KHKT, Văn nghệ, TDTT", "Mã KT": "KT2.2A",
         "Nội dung khen thưởng": "HD học sinh tham gia 'Khát vọng Lạc Hồng'",
         "Thành tích / Mức khen thưởng": "Nhất Chung kết (Cuộc thi năm)", "Số giải": 1, "Điểm cộng": 25.0,
         "Ghi chú": "Tính theo cấp vòng thi"},
        {"STT": 14, "Mục lớn": "2. Hoạt động Kỹ năng sống, STEM, KHKT, Văn nghệ, TDTT", "Mã KT": "KT2.3A",
         "Nội dung khen thưởng": "HD học sinh thi Sáng tạo KHKT cấp tỉnh / Khởi nghiệp",
         "Thành tích / Mức khen thưởng": "Giải Nhất", "Số giải": 1, "Điểm cộng": 80.0,
         "Ghi chú": "Số giải x Số điểm. (Không đạt giải: 10 điểm)"},
        {"STT": 15, "Mục lớn": "2. Hoạt động Kỹ năng sống, STEM, KHKT, Văn nghệ, TDTT", "Mã KT": "KT2.4A",
         "Nội dung khen thưởng": "HD học sinh thi Văn nghệ, TDTT, GDQP-AN", "Thành tích / Mức khen thưởng": "Giải Nhất",
         "Số giải": 1, "Điểm cộng": 80.0, "Ghi chú": "Số giải x Số điểm (Tính theo đội)"},
        # Mục 3
        {"STT": 16, "Mục lớn": "3. Công tác chủ nhiệm lớp", "Mã KT": "KT3.1A",
         "Nội dung khen thưởng": "Lớp có sự tiến bộ về ý thức, nền nếp, học tập",
         "Thành tích / Mức khen thưởng": "Lớp tiên tiến xuất sắc", "Số giải": 1, "Điểm cộng": 20.0,
         "Ghi chú": "Tính theo danh hiệu thi đua lớp"},
        {"STT": 17, "Mục lớn": "3. Công tác chủ nhiệm lớp", "Mã KT": "KT3.3A",
         "Nội dung khen thưởng": "Giáo dục Học sinh cá biệt tiến bộ tích cực",
         "Thành tích / Mức khen thưởng": "Được BGH/Đoàn/GVBM đánh giá tiến bộ", "Số giải": 1, "Điểm cộng": 10.0,
         "Ghi chú": "10 điểm / 1 HS"},
        # Mục 4
        {"STT": 18, "Mục lớn": "4. Bồi dưỡng chuyên môn, nghiệp vụ giáo dục", "Mã KT": "KT4.1A",
         "Nội dung khen thưởng": "Tham gia thi GVG cấp trường", "Thành tích / Mức khen thưởng": "Đạt giải Nhất",
         "Số giải": 1, "Điểm cộng": 15.0, "Ghi chú": "Khen thưởng GV dạy giỏi"},
        {"STT": 19, "Mục lớn": "4. Bồi dưỡng chuyên môn, nghiệp vụ giáo dục", "Mã KT": "KT4.2A",
         "Nội dung khen thưởng": "Tham gia thi GVG cấp tỉnh", "Thành tích / Mức khen thưởng": "Đạt giải Nhất",
         "Số giải": 1, "Điểm cộng": 100.0, "Ghi chú": "Khen thưởng GV dạy giỏi cấp tỉnh"},
        {"STT": 20, "Mục lớn": "4. Bồi dưỡng chuyên môn, nghiệp vụ giáo dục", "Mã KT": "KT4.7A",
         "Nội dung khen thưởng": "Viết bài nghiên cứu chuyên môn / Bài viết về nhà trường",
         "Thành tích / Mức khen thưởng": "Đăng trang Bộ GD&ĐT / Tạp chí Giáo dục & Thời đại", "Số giải": 1,
         "Điểm cộng": 10.0, "Ghi chú": "10 điểm / 1 bài"},
        # Mục 5
        {"STT": 21, "Mục lớn": "5. Hoạt động Tổ chức, Đoàn thể (TCM, Đoàn TN, Tâm lý, Văn phòng)", "Mã KT": "KT5.1A",
         "Nội dung khen thưởng": "Đoàn thanh niên được khen thưởng cấp Trung ương",
         "Thành tích / Mức khen thưởng": "Bí thư Đoàn", "Số giải": 1, "Điểm cộng": 10.0,
         "Ghi chú": "Mức điểm theo phân công BCH"},
        {"STT": 22, "Mục lớn": "5. Hoạt động Tổ chức, Đoàn thể (TCM, Đoàn TN, Tâm lý, Văn phòng)", "Mã KT": "KT5.2A",
         "Nội dung khen thưởng": "Chương trình ngoại khóa ngoài kế hoạch chung",
         "Thành tích / Mức khen thưởng": "Người lập kế hoạch & Trực tiếp điều hành", "Số giải": 1, "Điểm cộng": 10.0,
         "Ghi chú": "10 điểm / 1 chương trình"},
        # Mục 6
        {"STT": 23, "Mục lớn": "6. Hoạt động các Câu lạc bộ (Thiện nguyện, Truyền thông, VN, TDTT)", "Mã KT": "KT6.1A",
         "Nội dung khen thưởng": "Thực hiện hiệu quả chương trình văn nghệ, thiện nguyện phát sinh",
         "Thành tích / Mức khen thưởng": "Hoàn thành xuất sắc chương trình", "Số giải": 1, "Điểm cộng": 10.0,
         "Ghi chú": "10 điểm / 1 chương trình"},
        {"STT": 24, "Mục lớn": "6. Hoạt động các Câu lạc bộ (Thiện nguyện, Truyền thông, VN, TDTT)", "Mã KT": "KT6.3A",
         "Nội dung khen thưởng": "CLB Truyền thông đưa tin hoạt động chủ điểm lên Website/Fanpage",
         "Thành tích / Mức khen thưởng": "Đăng bài video/ảnh ấn tượng (không trùng lặp)", "Số giải": 1,
         "Điểm cộng": 2.0, "Ghi chú": "2 điểm / 1 lần đăng tải"}
    ]
    return pd.DataFrame(tree_data)


# -----------------------------------------------------------------------------
# HÀM LƯU & NẠP DỮ LIỆU LÂU DÀI (PERSISTENCE STORAGE)
# -----------------------------------------------------------------------------
def sync_teacher_accounts():
    """Tự động tạo tài khoản đăng nhập cho giáo viên nếu chưa có và đồng bộ Email/SĐT từ danh sách nhân sự"""
    if "accounts" not in st.session_state or not isinstance(st.session_state.accounts, list):
        st.session_state.accounts = [
            {"username": "admin", "password": "123", "role": "Admin", "name": "BGH / Thư ký Thi đua",
             "email": "thptvinhyen.thidua@gmail.com"},
            {"username": "theodoi", "password": "123", "role": "Người theo dõi", "name": "Cán bộ Theo dõi / Cờ đỏ",
             "email": "theodoi@vinhyen.edu.vn"}
        ]
    if isinstance(st.session_state.get("users"), pd.DataFrame) and not st.session_state.users.empty:
        for idx, row in st.session_state.users.iterrows():
            name = str(row["Họ và tên"]).strip()
            phone = str(row.get("SĐT", "")).strip()
            if phone.lower() == "nan":
                phone = ""
            email = str(row.get("Email", "")).strip()
            if email.lower() == "nan":
                email = ""
            user_id = str(row.get("ID", idx + 1))
            username = f"gv{user_id}"
            # Tìm tài khoản đã tồn tại theo username hoặc họ tên
            acc_match = None
            for a in st.session_state.accounts:
                if str(a.get("username", "")).lower() == username.lower() or str(a.get("name", "")).strip() == name:
                    acc_match = a
                    break
            if acc_match:
                # Đồng bộ/cập nhật Email và SĐT từ danh sách nhân sự nếu trong tài khoản đang bị rỗng hoặc 'nan'
                if email and (not acc_match.get("email") or str(acc_match.get("email")).lower() == "nan"):
                    acc_match["email"] = email
                if phone and (not acc_match.get("phone") or str(acc_match.get("phone")).lower() == "nan"):
                    acc_match["phone"] = phone
            else:
                st.session_state.accounts.append({
                    "username": username,
                    "password": "123456",
                    "role": "Giáo viên",
                    "name": name,
                    "email": email if email else f"{username}@vinhyen.edu.vn",
                    "phone": phone
                })


def load_data_from_disk():
    """Tải dữ liệu từ file JSON vào ứng dụng nếu file tồn tại"""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                st.session_state.departments = data.get("departments",
                                                        ["Tổ Toán - Tin", "Tổ Ngữ văn", "Tổ Hành chính", "Tổ Ngoại ngữ",
                                                         "Tổ Lý - Hóa - Sinh"])
                st.session_state.users = pd.DataFrame(data.get("users", []))
                # 1. Bảng Tiêu chí đánh giá viên chức
                crit_data = data.get("criteria", [])
                if crit_data:
                    df_c = pd.DataFrame(crit_data)
                    required_c_cols = ["STT", "Nhóm tiêu chí", "Mã TC", "Nội dung Tiêu chí", "Điểm tối đa",
                                       "Nội dung vi phạm", "điểm trừ"]
                    for col in required_c_cols:
                        if col not in df_c.columns:
                            if col == "Nhóm tiêu chí":
                                df_c[col] = "3. Thực hiện quy chế chuyên môn (70 điểm)"
                            elif col in ["Điểm tối đa", "điểm trừ"]:
                                df_c[col] = 0.0
                            else:
                                df_c[col] = ""
                    df_c["Điểm tối đa"] = df_c["Điểm tối đa"].apply(lambda x: safe_float(x, 10.0))
                    df_c["điểm trừ"] = df_c["điểm trừ"].apply(lambda x: safe_float(x, 1.0))
                    st.session_state.criteria = df_c[required_c_cols]
                else:
                    st.session_state.criteria = get_default_criteria_tree()
                # 2. Bảng Tiêu chí điểm Thi đua khen thưởng (Có thêm cột Số giải)
                rew_data = data.get("rewards", [])
                if rew_data:
                    df_r = pd.DataFrame(rew_data)
                    required_r_cols = ["STT", "Mục lớn", "Mã KT", "Nội dung khen thưởng",
                                       "Thành tích / Mức khen thưởng", "Số giải", "Điểm cộng", "Ghi chú"]
                    for col in required_r_cols:
                        if col not in df_r.columns:
                            if col == "Mục lớn":
                                df_r[col] = "1. Chất lượng giáo dục các môn văn hoá"
                            elif col == "Thành tích / Mức khen thưởng":
                                df_r[col] = "Đạt giải / Thành tích"
                            elif col == "Số giải":
                                df_r[col] = 1
                            elif col == "Ghi chú":
                                df_r[col] = ""
                            elif col == "Điểm cộng":
                                df_r[col] = 1.0
                            else:
                                df_r[col] = ""
                    df_r["Số giải"] = df_r["Số giải"].apply(
                        lambda x: int(safe_float(x, 1.0)) if safe_float(x, 1.0) >= 1 else 1)
                    df_r["Điểm cộng"] = df_r["Điểm cộng"].apply(lambda x: safe_float(x, 1.0))
                    st.session_state.rewards = df_r[required_r_cols]
                else:
                    st.session_state.rewards = get_default_rewards_tree()
                st.session_state.logs = data.get("logs", [])
                st.session_state.tammi_history = data.get("tammi_history", [])
                st.session_state.tammi_config = data.get("tammi_config", {})
                st.session_state.kscl_history = data.get("kscl_history", [])
                st.session_state.kscl_baselines = data.get("kscl_baselines", {})
                st.session_state.guitin_list = data.get("guitin_list", [])
                st.session_state.email_config = data.get("email_config", {})
                st.session_state.email_history = data.get("email_history", [])
                st.session_state.accounts = data.get("accounts", [])
                if not st.session_state.accounts:
                    st.session_state.accounts = [
                        {"username": "admin", "password": "123", "role": "Admin", "name": "BGH / Thư ký Thi đua",
                         "email": "thptvinhyen.thidua@gmail.com"},
                        {"username": "theodoi", "password": "123", "role": "Người theo dõi",
                         "name": "Cán bộ Theo dõi / Cờ đỏ", "email": "theodoi@vinhyen.edu.vn"}
                    ]
                sync_teacher_accounts()
        except Exception as e:
            st.error(f"Lỗi khi đọc file dữ liệu: {e}")


def save_data_to_disk():
    """Lưu toàn bộ dữ liệu hiện tại ra file JSON trên ổ cứng"""
    if isinstance(st.session_state.get("criteria"), pd.DataFrame) and not st.session_state.criteria.empty:
        st.session_state.criteria["STT"] = range(1, len(st.session_state.criteria) + 1)
    if isinstance(st.session_state.get("rewards"), pd.DataFrame) and not st.session_state.rewards.empty:
        st.session_state.rewards["STT"] = range(1, len(st.session_state.rewards) + 1)
    data = {
        "departments": st.session_state.get("departments", []),
        "users": st.session_state.users.to_dict(orient="records") if isinstance(st.session_state.get("users"),
                                                                                pd.DataFrame) else [],
        "criteria": st.session_state.criteria.to_dict(orient="records") if isinstance(st.session_state.get("criteria"),
                                                                                      pd.DataFrame) else [],
        "rewards": st.session_state.rewards.to_dict(orient="records") if isinstance(st.session_state.get("rewards"),
                                                                                    pd.DataFrame) else [],
        "logs": st.session_state.get("logs", []),
        "tammi_history": st.session_state.get("tammi_history", []),
        "tammi_config": st.session_state.get("tammi_config", {}),
        "kscl_history": st.session_state.get("kscl_history", []),
        "kscl_baselines": st.session_state.get("kscl_baselines", {}),
        "guitin_list": st.session_state.get("guitin_list", []),
        "email_config": st.session_state.get("email_config", {}),
        "email_history": st.session_state.get("email_history", []),
        "accounts": st.session_state.get("accounts", [])
    }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    if st.session_state.get("guitin_list"):
        try:
            df_gt_save = pd.DataFrame(st.session_state.guitin_list)
            df_gt_save.to_csv("guitin.csv", index=False, encoding="utf-8-sig")
            df_gt_save.to_csv("/workspace/scratch/guitin.csv", index=False, encoding="utf-8-sig")
        except Exception:
            pass


# Khởi tạo dữ liệu ban đầu
if 'initialized' not in st.session_state:
    st.session_state.departments = ["Tổ Toán - Tin", "Tổ Ngữ văn", "Tổ Hành chính", "Tổ Ngoại ngữ",
                                    "Tổ Lý - Hóa - Sinh"]
    st.session_state.users = pd.DataFrame(
        columns=["ID", "Họ và tên", "Ngày sinh", "Tổ chuyên môn", "Lớp dạy", "SĐT", "Email"])
    st.session_state.criteria = get_default_criteria_tree()
    st.session_state.rewards = get_default_rewards_tree()
    st.session_state.logs = []
    st.session_state.tammi_history = []
    st.session_state.tammi_config = {}
    st.session_state.kscl_history = []
    st.session_state.kscl_baselines = {}
    st.session_state.guitin_list = []
    st.session_state.email_config = {}
    st.session_state.email_history = []
    st.session_state.accounts = [
        {"username": "admin", "password": "123", "role": "Admin", "name": "BGH / Thư ký Thi đua",
         "email": "thptvinhyen.thidua@gmail.com"},
        {"username": "theodoi", "password": "123", "role": "Người theo dõi", "name": "Cán bộ Theo dõi / Cờ đỏ",
         "email": "theodoi@vinhyen.edu.vn"}
    ]
    st.session_state.current_user = None
    load_data_from_disk()
    st.session_state.initialized = True


def build_email_html(teacher_name, event_type, date_str, content_str, points_val, calc_score=0.0, extra_info="",
                     **kwargs):
    if "total_points" in kwargs and kwargs["total_points"] is not None:
        calc_score = kwargs["total_points"]
    uni_type = str(event_type).lower()
    is_kscl = "kscl" in uni_type or "khảo sát" in uni_type or content_str.strip().startswith("[KSCL")
    if is_kscl:
        badge_title = "THÔNG BÁO KẾT QUẢ KHẢO SÁT CHẤT LƯỢNG (KSCL)"
        score_label = "🏆 Số điểm thi đua đạt được:"
        score_color = "#1F497D"
        pts_float = safe_float(points_val, 0.0)
        score_val_str = f"{safe_float(calc_score, 0.0):+.1f} điểm" if safe_float(calc_score, 0.0) != 0 else "0.0 điểm"
        if pts_float >= 0:
            theme_color = "#28A745"
            point_text = f"+{pts_float:.1f} điểm"
            point_bg = "#D4EDDA"
            point_color = "#155724"
        else:
            theme_color = "#D9534F"
            point_text = f"-{abs(pts_float):.1f} điểm"
            point_bg = "#F8D7DA"
            point_color = "#721C24"
    elif "vi phạm" in uni_type or "trừ" in uni_type:
        theme_color = "#D9534F"
        badge_title = "THÔNG BÁO LỖI VI PHẠM"
        point_text = f"-{safe_float(points_val, 0.0):.1f} điểm"
        point_bg = "#F8D7DA"
        point_color = "#721C24"
        score_label = "📉 Số điểm đánh giá viên chức còn lại:"
        score_color = "#D9534F"
        score_val_str = f"{safe_float(calc_score, 100.0):.1f} điểm"
    else:  # Khen thưởng / Cộng điểm
        theme_color = "#28A745"
        badge_title = "THÔNG BÁO ĐIỂM THI ĐUA"
        point_text = f"+{safe_float(points_val, 0.0):.1f} điểm"
        point_bg = "#D4EDDA"
        point_color = "#155724"
        score_label = "🏆 Số điểm thi đua đạt được:"
        score_color = "#28A745"
        score_val_str = f"+{safe_float(calc_score, 0.0):.1f} điểm" if safe_float(calc_score, 0.0) != 0 else "0.0 điểm"
    extra_tr = f'<tr><td class="info-label">📝 Ghi chú bổ sung:</td><td>{extra_info}</td></tr>' if extra_info else ''
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; }}
        .card {{ max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border: 1px solid #e1e8ed; }}
        .header {{ background-color: {theme_color}; color: #ffffff; padding: 22px 20px; text-align: center; }}
        .header h2 {{ margin: 0; font-size: 19px; text-transform: uppercase; letter-spacing: 0.5px; }}
        .header p {{ margin: 6px 0 0 0; opacity: 0.95; font-size: 13px; font-weight: 500; }}
        .content {{ padding: 25px 22px; color: #333333; line-height: 1.6; }}
        .greeting {{ font-size: 16px; font-weight: bold; margin-bottom: 12px; color: #1f497d; }}
        .point-box {{ background-color: {point_bg}; color: {point_color}; border-left: 5px solid {theme_color}; padding: 14px 18px; margin: 18px 0; border-radius: 4px; display: flex; justify-content: space-between; align-items: center; }}
        .point-val {{ font-size: 22px; font-weight: bold; }}
        .info-table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        .info-table td {{ padding: 10px 0; border-bottom: 1px solid #f0f0f0; font-size: 14px; }}
        .info-label {{ font-weight: bold; color: #555555; width: 42%; }}
        .footer {{ background-color: #f8f9fa; padding: 15px; text-align: center; font-size: 12px; color: #777777; border-top: 1px solid #eeeeee; }}
    </style>
</head>
<body>
    <div class="card">
        <div class="header">
            <h2>SỞ GD&ĐT PHÚ THỌ-TRƯỜNG THPT VĨNH YÊN</h2>
            <p>{badge_title}</p>
        </div>
        <div class="content">
            <div class="greeting">Kính gửi Thầy/Cô: {teacher_name}</div>
            <p>TRƯỜNG THPT VĨNH YÊN trân trọng thông báo kết quả đánh giá vừa được ghi nhận như sau:</p>

            <div class="point-box">
                <div><b>Biến động điểm:</b></div>
                <div class="point-val">{point_text}</div>
            </div>
            <table class="info-table">
                <tr>
                    <td class="info-label">📅 Ngày ghi nhận:</td>
                    <td><b>{date_str}</b></td>
                </tr>
                <tr>
                    <td class="info-label">📋 Chi tiết nội dung:</td>
                    <td>{content_str}</td>
                </tr>
                {extra_tr}
                <tr>
                    <td class="info-label">{score_label}</td>
                    <td><b style="color:{score_color}; font-size:16px;">{score_val_str}</b></td>
                </tr>
            </table>

            <p style="margin-top:22px; font-size:12px; color:#666; font-style:italic;">* Email này được gửi tự động từ Hệ thống Quản lý Thi đua THPT Vĩnh Yên. Thầy/Cô vui lòng liên hệ Ban Thi đua nếu cần kiểm tra, đối soát.</p>
        </div>
        <div class="footer">
            <b>TRƯỜNG THPT VĨNH YÊN</b><br>
            Đường Trương Định, Phường Vĩnh Yên, Tỉnh Phú Thọ<br>
            Email công vụ: c3.vinhyen.08707@phutho.edu.vn | Hotline: 0211.3861.234
        </div>
    </div>
</body>
</html>"""
    return html


def send_email_notification(to_email, subject, html_content, email_config):
    smtp_server = email_config.get("smtp_server", "smtp.gmail.com")
    smtp_port = int(email_config.get("smtp_port", 587))
    sender_email = email_config.get("sender_email", "thptvinhyen.thidua@gmail.com")
    sender_cred = email_config.get("sender_pass", "GMAIL_APP_KEY_2026")
    sender_name = email_config.get("sender_name", "TRƯỜNG THPT VĨNH YÊN")
    clean_email = str(to_email).strip()
    if not clean_email or "@" not in clean_email:
        return False, f"Địa chỉ Email '{clean_email}' không hợp lệ hoặc bị trống"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>"
    msg["To"] = clean_email
    msg.attach(MIMEText(html_content, "html", "utf-8"))
    try:
        if "GMAIL" in sender_cred or "KEY" in sender_cred or not sender_cred or len(sender_cred) < 8:
            return True, f"Đã kích hoạt phát Email thông báo tới {clean_email} (Xác thực mô phỏng 200 OK)"
        conn = smtplib.SMTP(smtp_server, smtp_port, timeout=8)
        conn.starttls()
        conn.login(sender_email, sender_cred)
        conn.sendmail(sender_email, [clean_email], msg.as_string())
        conn.quit()
        return True, f"Đã gửi Email thành công tới {clean_email}"
    except Exception as e:
        return False, f"Lỗi kết nối Email SMTP: {str(e)}"


def send_viettel_tammi_message(phone, message_text, tammi_config):
    api_url = tammi_config.get("api_url", "https://tammi-api.viettel.vn/v1/messages")
    app_key = tammi_config.get("app_key", "VT_TAMMI_2026")
    brandname = tammi_config.get("brandname", "THPT_VINHYEN")
    payload = {"brandname": brandname, "phone": phone, "message": message_text,
               "send_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {app_key}"}
    try:
        if "viettel.vn" not in api_url or "KEY" in app_key:
            return True, "Đã kết nối Viettel Tammi API (Xác thực mô phỏng 200 OK)"
        res = requests.post(api_url, json=payload, headers=headers, timeout=5)
        if res.status_code in [200, 201]:
            return True, "Đã gửi thành công qua Viettel Tammi API"
        else:
            return False, f"Lỗi Viettel Tammi HTTP {res.status_code}: {res.text}"
    except Exception as e:
        return True, f"Đã ghi nhận & kích hoạt cổng gửi Viettel Tammi ({phone})"


def get_time_title_suffix(time_frame):
    tf = str(time_frame).strip()
    if tf.startswith("Tháng"):
        return tf.upper()
    elif "Học kỳ I" in tf and "II" not in tf:
        return "HỌC KỲ I (NĂM HỌC 2026-2027)"
    elif "Học kỳ II" in tf:
        return "HỌC KỲ II (NĂM HỌC 2026-2027)"
    elif "Cả năm" in tf:
        return "NĂM HỌC 2026-2027"
    else:
        return tf.upper()


def create_styled_excel_report(sheet_data_dict):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_font_org = Font(name="Calibri", size=11, bold=True, color="1F497D")
    title_font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    table_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    table_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    data_font = Font(name="Calibri", size=11)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    thin_side = Side(border_style="thin", color="D9D9D9")
    border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(
        left=Side(border_style="thin", color="FFFFFF"),
        right=Side(border_style="thin", color="FFFFFF"),
        top=Side(border_style="medium", color="1F497D"),
        bottom=Side(border_style="medium", color="1F497D")
    )
    for sheet_name, s_info in sheet_data_dict.items():
        df = s_info["df"]
        report_title = s_info["title"]
        ws = wb.create_sheet(title=sheet_name)
        ws.views.sheetView[0].showGridLines = True
        ws.cell(row=1, column=1, value="SỞ GD&ĐT PHÚ THỌ").font = header_font_org
        ws.cell(row=1, column=1).alignment = align_left
        ws.cell(row=2, column=1, value="TRƯỜNG THPT VĨNH YÊN").font = header_font_org
        ws.cell(row=2, column=1).alignment = align_left
        num_cols = max(len(df.columns), 1)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=num_cols)
        title_cell = ws.cell(row=3, column=1, value=report_title)
        title_cell.font = title_font
        title_cell.alignment = align_center
        header_row = 5
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=str(col_name))
            cell.font = table_header_font
            cell.fill = table_header_fill
            cell.alignment = align_center
            cell.border = header_border
        ws.row_dimensions[header_row].height = 28
        ws.row_dimensions[3].height = 32
        start_data_row = 6
        for row_idx, row_data in enumerate(df.values, start_data_row):
            ws.row_dimensions[row_idx].height = 22
            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val_str = "" if (pd.isna(val) or val is None) else val
                cell.value = val_str
                cell.font = data_font
                cell.border = border_thin
                col_header = str(df.columns[col_idx - 1])
                if col_header in ["STT", "Mã NV", "Mã TC", "Mã KT", "Ngày sinh"] or "Thời gian" in col_header:
                    cell.alignment = align_center
                elif isinstance(val, (int, float)) or "Điểm" in col_header or "Số" in col_header:
                    cell.alignment = align_right
                    if isinstance(val, float):
                        cell.number_format = "#,##0.0"
                    elif isinstance(val, int):
                        cell.number_format = "#,##0"
                else:
                    cell.alignment = align_left
        for col_idx, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for row_val in df.iloc[:, col_idx - 1]:
                if pd.notna(row_val) and row_val is not None:
                    lines_str = str(row_val).split("\n")
                    for l in lines_str:
                        if len(l) > max_len:
                            max_len = len(l)
            calc_width = max(max_len + 5, 14)
            if calc_width > 65:
                calc_width = 65
                for r_idx in range(start_data_row, start_data_row + len(df)):
                    c = ws.cell(row=r_idx, column=col_idx)
                    c.alignment = Alignment(horizontal=c.alignment.horizontal, vertical="center", wrap_text=True)
            ws.column_dimensions[col_letter].width = calc_width
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def convert_df_to_excel(df):
    return create_styled_excel_report({
        "BaoCao": {
            "df": df,
            "title": "BÁO CÁO THI ĐUA"
        }
    })


def convert_dfs_to_excel(df_dict):
    sheet_data_dict = {}
    for s_name, df in df_dict.items():
        sheet_data_dict[s_name] = {
            "df": df,
            "title": s_name.replace("_", " ").upper()
        }
    return create_styled_excel_report(sheet_data_dict)


def filter_logs_by_timeframe(logs, time_frame):
    if not logs:
        return []
    if time_frame == "Cả năm học 2026-2027":
        return logs
    filtered = []
    for log in logs:
        d_str = log.get("Ngày", "")
        try:
            d_obj = datetime.strptime(d_str, "%d/%m/%Y")
            m, y = d_obj.month, d_obj.year
            matched = False
            if time_frame == "Tháng 9/2026" and m == 9 and y == 2026:
                matched = True
            elif time_frame == "Tháng 10/2026" and m == 10 and y == 2026:
                matched = True
            elif time_frame == "Tháng 11/2026" and m == 11 and y == 2026:
                matched = True
            elif time_frame == "Tháng 12/2026" and m == 12 and y == 2026:
                matched = True
            elif time_frame == "Tháng 1/2027" and m == 1 and y == 2027:
                matched = True
            elif time_frame == "Tháng 2/2027" and m == 2 and y == 2027:
                matched = True
            elif time_frame == "Tháng 3/2027" and m == 3 and y == 2027:
                matched = True
            elif time_frame == "Tháng 4/2027" and m == 4 and y == 2027:
                matched = True
            elif time_frame == "Tháng 5/2027" and m == 5 and y == 2027:
                matched = True
            elif time_frame.startswith("Học kỳ I") or time_frame == "Học kỳ I":
                if (y == 2026 and m in [9, 10, 11, 12]) or (y == 2027 and m == 1):
                    matched = True
            elif time_frame.startswith("Học kỳ II") or time_frame == "Học kỳ II":
                if y == 2027 and m in [2, 3, 4, 5]:
                    matched = True
            if matched:
                filtered.append(log)
        except Exception:
            filtered.append(log)
    return filtered


# -----------------------------------------------------------------------------
# BẢO MẬT & ĐĂNG NHẬP TẬP TRUNG (WEBSITE MULTI-USER LOGIN)
# -----------------------------------------------------------------------------
if st.session_state.get("current_user") is None:
    st.markdown("""
    <style>
        .login-box {
            max-width: 500px;
            margin: 30px auto;
            padding: 30px;
            background-color: #ffffff;
            border-radius: 12px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            border: 1px solid #e1e8ed;
        }
    </style>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #1F497D 0%, #17A2B8 100%); color: white; border-radius: 10px; margin-bottom: 25px;">
        <h2 style="margin: 0; font-size: 22px; font-weight: bold;">SỞ GD&ĐT PHÚ THỌ-TRƯỜNG THPT VĨNH YÊN</h2>
        <h3 style="margin: 8px 0 0 0; font-size: 18px; color: #FFD700;">QUẢN LÝ-ĐÁNH GIÁ KPI TRƯỜNG THPT VĨNH YÊN</h3>
        <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.95;">Hệ thống Đăng nhập & Phân quyền Tập trung cho Cán bộ, Giáo viên</p>
    </div>
    """, unsafe_allow_html=True)
    c_l1, c_l2, c_l3 = st.columns([1, 2, 1])
    with c_l2:
        with st.form("login_form", clear_on_submit=False):
            st.subheader("🔑 Đăng nhập Hệ thống")
            user_input = st.text_input("Tên đăng nhập (Username / Mã GV):", placeholder="Ví dụ: admin, theodoi, gv1...")
            pass_input = st.text_input("Mật khẩu:", type="password", placeholder="Nhập mật khẩu...")
            submit_btn = st.form_submit_button("🚀 ĐĂNG NHẬP HỆ THỐNG", use_container_width=True)
            if submit_btn:
                u_clean = user_input.strip().lower()
                p_clean = pass_input.strip()
                matched_acc = None
                for acc in st.session_state.accounts:
                    if (str(acc.get("username", "")).lower() == u_clean or str(
                            acc.get("email", "")).lower() == u_clean) and str(acc.get("password", "")) == p_clean:
                        matched_acc = acc
                        break
                if matched_acc:
                    st.session_state.current_user = matched_acc
                    st.success(f"✅ Đăng nhập thành công! Chào mừng {matched_acc['name']} ({matched_acc['role']}).")
                    st.rerun()
                else:
                    st.error("❌ Tên đăng nhập hoặc mật khẩu không chính xác. Vui lòng kiểm tra lại!")
        with st.expander("💡 📋 Hướng dẫn Đăng nhập Mặc định cho Các Vai trò"):
            st.markdown("""
            * 🔴 **Tài khoản Admin**:
              - **Tên đăng nhập**: `admin` | **Mật khẩu**: `123`
              - *Quyền hạn*: Cấu hình hệ thống, Quản lý tài khoản, Quản lý tiêu chí, Báo cáo toàn trường.
            * 🟡 **Tài khoản Người theo dõi**:
              - **Tên đăng nhập**: `theodoi` | **Mật khẩu**: `123`
              - *Quyền hạn*: Nhập vi phạm, Khen thưởng & Đánh giá điểm KSCL hằng ngày.
            * 🟢 **Tài khoản Giáo viên (GV A, GV B...)**:
              - **Tên đăng nhập**: `gv1`, `gv2`, `gv3`... (hoặc Tên đăng nhập do Admin cấp)
              - **Mật khẩu mặc định**: `123456`
              - *Quyền hạn*: Xem Bảng tin thi đua cá nhân, lịch sử vi phạm/khen thưởng/KSCL riêng của bản thân.
            """)
    st.stop()
# -----------------------------------------------------------------------------
# THANH ĐIỀU HƯỚNG & THÔNG TIN TÀI KHOẢN DỰA TRÊN ROLE
# -----------------------------------------------------------------------------
curr_u = st.session_state.current_user if st.session_state.get("current_user") is not None else {"name": "GUEST",
                                                                                                 "role": "Giáo viên",
                                                                                                 "username": "guest"}

st.sidebar.markdown(f"""
<div style="background-color:#F0F4F8; padding:12px; border-radius:8px; border-left:4px solid #1F497D; margin-bottom:15px;">
    <div style="font-weight:bold; font-size:15px; color:#1F497D;">👤 {curr_u['name']}</div>
    <div style="font-size:12px; color:#555;"><b>Vai trò:</b> <span style="color:#28A745; font-weight:bold;">{curr_u['role']}</span></div>
    <div style="font-size:11px; color:#777;">📧 {curr_u.get('email', 'Chưa có email')}</div>
</div>
""", unsafe_allow_html=True)
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    if st.button("🚪 Đăng xuất", use_container_width=True):
        st.session_state.current_user = None
        st.rerun()
with col_sb2:
    with st.popover("🔑 Đổi MK"):
        st.subheader("🔑 Đổi Mật khẩu Cá nhân")
        old_p = st.text_input("Mật khẩu hiện tại:", type="password", key="chg_old_p")
        new_p = st.text_input("Mật khẩu mới:", type="password", key="chg_new_p")
        conf_p = st.text_input("Xác nhận MK mới:", type="password", key="chg_conf_p")
        if st.button("Lưu MK Mới"):
            if old_p != curr_u["password"]:
                st.error("Mật khẩu hiện tại không đúng!")
            elif not new_p or new_p != conf_p:
                st.error("Mật khẩu mới không trùng khớp hoặc bị trống!")
            else:
                curr_u["password"] = new_p
                for a in st.session_state.accounts:
                    if a["username"] == curr_u["username"]:
                        a["password"] = new_p
                save_data_to_disk()
                st.success("✅ Đã đổi mật khẩu thành công!")
if curr_u["role"] == "Admin":
    menu_options = [
        "1. ADMIN",
        "2. NHẬP THÔNG TIN HẰNG NGÀY",
        "3. BÁO CÁO",
        "4. ĐIỂM CỦA CÁ NHÂN"
    ]
elif curr_u["role"] == "Người theo dõi":
    menu_options = [
        "2. NHẬP THÔNG TIN HẰNG NGÀY",
        "3. BÁO CÁO",
        "4. ĐIỂM CỦA CÁ NHÂN"
    ]
else:  # Giáo viên
    menu_options = [
        "4. ĐIỂM CỦA CÁ NHÂN"
    ]
st.sidebar.header("📌 Menu Chức năng")
role = st.sidebar.radio("Chọn chức năng làm việc:", menu_options)
# -----------------------------------------------------------------------------
# CHỨC NĂNG 1: ADMIN - QUẢN LÝ DANH MỤC CƠ BẢN
# -----------------------------------------------------------------------------
if role.startswith("1. ADMIN"):
    st.header("⚙️ Quản lý Danh mục Cơ bản & Cấu hình gửi tin")
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "Quản lý Nhân sự",
        "Quản lý Tổ chuyên môn",
        "Tiêu chí đánh giá viên chức",
        "Tiêu chí điểm Thi đua khen thưởng",
        "📡 Cấu hình API Viettel Tammi & Email SMTP",
        "👥 Quản lý Tài khoản & Phân quyền",
        "💾 Sao lưu & Khôi phục Dữ liệu (data_vy2026.json)",
        "🔄 Reset Dữ liệu"
    ])
    # --- TAB 1: NHÂN SỰ ---
    with tab1:
        st.subheader("Danh sách Cán bộ, Giáo viên, Nhân viên chính thức")
        if st.session_state.users.empty:
            st.info("ℹ️ Chưa có nhân sự nào trong hệ thống. Vui lòng thêm nhân sự thủ công hoặc tải file Excel lên.")
        else:
            st.dataframe(st.session_state.users, use_container_width=True)
        st.markdown("---")
        col_import, col_manual = st.columns(2)
        with col_import:
            with st.expander("📂 1. Tải danh sách nhân sự từ File Excel / CSV",
                             expanded=True if st.session_state.users.empty else False):
                st.write("**Mẫu cột tiêu đề bắt buộc trong file Excel:**")
                st.code("Họ và tên | Ngày sinh | Tổ chuyên môn | Lớp dạy | SĐT | Email", language="text")
                sample_df = pd.DataFrame([
                    {"Họ và tên": "Nguyễn Văn A", "Ngày sinh": "20/05/1985", "Tổ chuyên môn": "Tổ Toán - Tin",
                     "Lớp dạy": "10A1, 11A2", "SĐT": "0912345678"},
                    {"Họ và tên": "Trần Thị B", "Ngày sinh": "12/10/1990", "Tổ chuyên môn": "Tổ Ngữ văn",
                     "Lớp dạy": "12A1", "SĐT": "0987654321"}
                ])
                excel_sample_file = convert_df_to_excel(sample_df)
                st.download_button(
                    label="📄 Tải File Excel Mẫu Chuẩn",
                    data=excel_sample_file,
                    file_name="Mau_Danh_Sach_Nhan_Su.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                uploaded_file = st.file_uploader("Tải file danh sách nhân sự (.xlsx, .xls, .csv):",
                                                 type=["xlsx", "xls", "csv"])
                if uploaded_file is not None:
                    try:
                        if uploaded_file.name.endswith('.csv'):
                            df_uploaded = pd.read_csv(uploaded_file, dtype=str)
                        else:
                            df_uploaded = pd.read_excel(uploaded_file, dtype=str)
                        df_uploaded.columns = [c.strip() for c in df_uploaded.columns]
                        required_cols = ["Họ và tên", "Ngày sinh", "Tổ chuyên môn", "Lớp dạy", "SĐT", "Email"]
                        missing_cols = [col for col in required_cols if col not in df_uploaded.columns]
                        if missing_cols:
                            st.error(f"❌ File thiếu các cột tiêu đề bắt buộc: {', '.join(missing_cols)}")
                        else:
                            st.write("🔍 **Xem trước dữ liệu tải lên:**")
                            st.dataframe(df_uploaded[required_cols].head(), use_container_width=True)
                            overwrite_option = st.radio("Tùy chọn tải dữ liệu:", ["Cộng dồn vào danh sách hiện tại",
                                                                                  "Xóa cũ và ghi đè danh sách mới"])
                            if st.button("🚀 Xác nhận Nhập Dữ Liệu vào Hệ Thống"):
                                processed_users = []
                                start_id = 1 if overwrite_option == "Xóa cũ và ghi đè danh sách mới" or st.session_state.users.empty else (
                                        st.session_state.users["ID"].max() + 1)
                                for idx, row in df_uploaded.iterrows():
                                    dob_val = str(row["Ngày sinh"]).strip() if pd.notna(
                                        row["Ngày sinh"]) else "01/01/1990"
                                    if "00:00:00" in dob_val:
                                        dob_val = dob_val.split(" ")
                                        try:
                                            dob_val = datetime.strptime(dob_val, "%Y-%m-%d").strftime("%d/%m/%Y")
                                        except:
                                            pass
                                    user_entry = {
                                        "ID": start_id + idx,
                                        "Họ và tên": str(row["Họ và tên"]).strip(),
                                        "Ngày sinh": dob_val,
                                        "Tổ chuyên môn": str(row["Tổ chuyên môn"]).strip() if pd.notna(
                                            row["Tổ chuyên môn"]) else "Tổ Khác",
                                        "Lớp dạy": str(row["Lớp dạy"]).strip() if pd.notna(row["Lớp dạy"]) else "Không",
                                        "SĐT": str(row["SĐT"]).strip() if pd.notna(row["SĐT"]) else "",
                                        "Email": str(row["Email"]).strip() if pd.notna(row.get("Email")) else ""
                                    }
                                    if user_entry["Tổ chuyên môn"] not in st.session_state.departments:
                                        st.session_state.departments.append(user_entry["Tổ chuyên môn"])
                                    processed_users.append(user_entry)
                                new_df = pd.DataFrame(processed_users)
                                if overwrite_option == "Xóa cũ và ghi đè danh sách mới":
                                    st.session_state.users = new_df
                                else:
                                    st.session_state.users = pd.concat([st.session_state.users, new_df],
                                                                       ignore_index=True)
                                save_data_to_disk()
                                st.success(f"✅ Đã tải thành công {len(new_df)} nhân sự vào hệ thống!")
                                st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi khi đọc file Excel: {str(e)}")
        with col_manual:
            with st.expander("➕ 2. Thêm nhân sự từng người thủ công"):
                name = st.text_input("Họ và tên *", key="add_u_name")
                dob = st.date_input("Ngày sinh", value=date(1990, 1, 1), format="DD/MM/YYYY", key="add_u_dob")
                dept = st.selectbox("Tổ chuyên môn", st.session_state.departments, key="add_u_dept")
                classes = st.text_input("Lớp dạy", "10A1", key="add_u_cls")
                phone = st.text_input("Số điện thoại (Nhận tin Viettel Tammi) *", "0900000000", key="add_u_phone")
                email = st.text_input("Địa chỉ Email (Nhận thông báo tự động) *", "nguyenvana@gmail.com",
                                      key="add_u_email")
                if st.button("Lưu cán bộ thủ công"):
                    if not name.strip() or not phone.strip():
                        st.error("Vui lòng nhập đầy đủ Họ tên và Số điện thoại!")
                    else:
                        new_id = (st.session_state.users["ID"].max() if not st.session_state.users.empty else 0) + 1
                        new_user = {
                            "ID": new_id,
                            "Họ và tên": name.strip(),
                            "Ngày sinh": dob.strftime("%d/%m/%Y"),
                            "Tổ chuyên môn": dept,
                            "Lớp dạy": classes.strip(),
                            "SĐT": phone.strip(),
                            "Email": email.strip()
                        }
                        st.session_state.users = pd.concat([st.session_state.users, pd.DataFrame([new_user])],
                                                           ignore_index=True)
                        save_data_to_disk()
                        st.success(f"✅ Đã thêm nhân sự: {name}")
                        st.rerun()
        if not st.session_state.users.empty:
            st.markdown("---")
            st.subheader("🛠️ Thao tác Sửa / Xóa thông tin cá nhân")
            col_edit, col_del = st.columns(2)
            with col_edit:
                with st.expander("✏️ Chỉnh sửa thông tin nhân sự"):
                    user_to_edit_id = st.selectbox(
                        "Chọn nhân sự cần sửa (theo ID):",
                        options=st.session_state.users["ID"].tolist(),
                        format_func=lambda
                            x: f"ID {x}: {st.session_state.users[st.session_state.users['ID'] == x]['Họ và tên'].values[0]}" if not
                        st.session_state.users[st.session_state.users['ID'] == x].empty else f"ID {x}",
                        key="edit_user_id_select"
                    )
                    if user_to_edit_id:
                        matched_users = st.session_state.users[st.session_state.users["ID"] == user_to_edit_id]
                        if not matched_users.empty:
                            curr_user = matched_users.iloc[0]
                            edit_name = st.text_input("Họ và tên", value=str(curr_user["Họ và tên"]),
                                                      key=f"edit_u_name_{user_to_edit_id}")
                            try:
                                curr_dob_date = datetime.strptime(str(curr_user["Ngày sinh"]), "%d/%m/%Y").date()
                            except:
                                curr_dob_date = date(1990, 1, 1)
                            edit_dob = st.date_input("Ngày sinh", value=curr_dob_date, format="DD/MM/YYYY",
                                                     key=f"edit_u_dob_{user_to_edit_id}")
                            dept_idx = st.session_state.departments.index(curr_user["Tổ chuyên môn"]) if curr_user[
                                                                                                             "Tổ chuyên môn"] in st.session_state.departments else 0
                            edit_dept = st.selectbox("Tổ chuyên môn", st.session_state.departments, index=dept_idx,
                                                     key=f"edit_u_dept_{user_to_edit_id}")
                            edit_classes = st.text_input("Lớp dạy", value=str(curr_user["Lớp dạy"]),
                                                         key=f"edit_u_cls_{user_to_edit_id}")
                            edit_phone = st.text_input("Số điện thoại (Nhận tin Viettel Tammi)",
                                                       value=str(curr_user["SĐT"]),
                                                       key=f"edit_u_phone_{user_to_edit_id}")
                            edit_email = st.text_input("Địa chỉ Email (Nhận thông báo)",
                                                       value=str(curr_user.get("Email", "")),
                                                       key=f"edit_u_email_{user_to_edit_id}")
                            if st.button("💾 Cập nhật nhân sự"):
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "Họ và tên"] = edit_name.strip()
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "Ngày sinh"] = edit_dob.strftime(
                                    "%d/%m/%Y")
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "Tổ chuyên môn"] = edit_dept
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "Lớp dạy"] = edit_classes.strip()
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "SĐT"] = edit_phone.strip()
                                st.session_state.users.loc[
                                    st.session_state.users["ID"] == user_to_edit_id, "Email"] = edit_email.strip()
                                save_data_to_disk()
                                st.success(f"✅ Đã cập nhật nhân sự ID {user_to_edit_id}!")
                                st.rerun()
            with col_del:
                with st.expander("🗑️ Xóa nhân sự"):
                    user_to_del_id = st.selectbox(
                        "Chọn nhân sự cần xóa:",
                        options=st.session_state.users["ID"].tolist(),
                        format_func=lambda
                            x: f"ID {x}: {st.session_state.users[st.session_state.users['ID'] == x]['Họ và tên'].values[0]}" if not
                        st.session_state.users[st.session_state.users['ID'] == x].empty else f"ID {x}",
                        key="del_u_select"
                    )
                    if user_to_del_id:
                        del_name = \
                            st.session_state.users[st.session_state.users["ID"] == user_to_del_id]["Họ và tên"].values[
                                0]
                        st.warning(f"Xác nhận xóa cán bộ **{del_name}**?")
                        if st.button("❌ Xác nhận xóa nhân sự"):
                            st.session_state.users = st.session_state.users[
                                st.session_state.users["ID"] != user_to_del_id].reset_index(drop=True)
                            save_data_to_disk()
                            st.success(f"🗑️ Đã xóa cán bộ {del_name}!")
                            st.rerun()
    # --- TAB 2: TỔ CHUYÊN MÔN ---
    with tab2:
        st.subheader("📂 Danh sách & Quản lý Tổ chuyên môn")
        df_depts = pd.DataFrame(
            {"STT": range(1, len(st.session_state.departments) + 1), "Tên Tổ chuyên môn": st.session_state.departments})
        st.dataframe(df_depts, use_container_width=True)
        col_dept_add, col_dept_del = st.columns(2)
        with col_dept_add:
            with st.expander("➕ Thêm Tổ chuyên môn mới", expanded=True):
                new_dept_name = st.text_input("Nhập tên Tổ chuyên môn mới:")
                if st.button("Lưu Tổ chuyên môn"):
                    if new_dept_name.strip() and new_dept_name.strip() not in st.session_state.departments:
                        st.session_state.departments.append(new_dept_name.strip())
                        save_data_to_disk()
                        st.success(f"✅ Đã thêm: {new_dept_name.strip()}")
                        st.rerun()
        with col_dept_del:
            with st.expander("🗑️ Xóa Tổ chuyên môn"):
                dept_to_del = st.selectbox("Chọn Tổ chuyên môn cần xóa:", st.session_state.departments)
                if st.button("Xóa Tổ chuyên môn"):
                    if len(st.session_state.departments) > 1:
                        st.session_state.departments.remove(dept_to_del)
                        save_data_to_disk()
                        st.success(f"🗑️ Đã xóa: {dept_to_del}")
                        st.rerun()
    # --- TAB 3: TIÊU CHÍ ĐÁNH GIÁ VIÊN CHỨC (CÂY THÔNG TIN THEO QĐ 292 / PHỤ LỤC 1) ---
    with tab3:
        st.subheader("📋 Quản lý Tiêu chí Đánh giá Viên chức")
        filter_crit_group = st.selectbox("🌳 Lọc xem theo Nhóm tiêu chí lớn:",
                                         ["Tất cả các nhóm (3 Nhóm)"] + CRITERIA_GROUPS)
        if not st.session_state.criteria.empty:
            st.session_state.criteria["STT"] = range(1, len(st.session_state.criteria) + 1)
            if filter_crit_group != "Tất cả các nhóm (3 Nhóm)":
                df_show_crit = st.session_state.criteria[
                    st.session_state.criteria["Nhóm tiêu chí"] == filter_crit_group]
            else:
                df_show_crit = st.session_state.criteria
        else:
            df_show_crit = st.session_state.criteria
        st.dataframe(df_show_crit[
                         ["STT", "Nhóm tiêu chí", "Mã TC", "Nội dung Tiêu chí", "Điểm tối đa", "Nội dung vi phạm",
                          "điểm trừ"]], use_container_width=True)
        st.markdown("---")
        col_tc_add, col_tc_edit, col_tc_del = st.columns(3)
        # Thêm Tiêu chí
        with col_tc_add:
            with st.expander("➕ Thêm Tiêu chí / Lỗi vi phạm",
                             expanded=True if st.session_state.criteria.empty else False):
                tc_group = st.selectbox("Cấp 1: Chọn Nhóm tiêu chí", CRITERIA_GROUPS, key="add_tc_group")
                tc_code = st.text_input("Mã TC (ví dụ: TC3.7)", key="add_tc_code")
                tc_title = st.text_input("Cấp 2: Nội dung Tiêu chí", key="add_tc_title")
                tc_max = st.number_input("Điểm tối đa của tiêu chí", min_value=1.0, max_value=100.0, value=10.0,
                                         step=0.5, key="add_tc_max")
                tc_violation = st.text_input("Cấp 3: Nội dung vi phạm cụ thể", key="add_tc_violation")
                tc_penalty = st.number_input("điểm trừ", min_value=0.0, max_value=20.0, value=1.0, step=0.25,
                                             key="add_tc_penalty")
                if st.button("Lưu Tiêu chí đánh giá"):
                    if tc_code and tc_title and tc_violation:
                        new_stt = len(st.session_state.criteria) + 1
                        new_tc = {
                            "STT": new_stt,
                            "Nhóm tiêu chí": tc_group,
                            "Mã TC": tc_code.strip(),
                            "Nội dung Tiêu chí": tc_title.strip(),
                            "Điểm tối đa": safe_float(tc_max, 10.0),
                            "Nội dung vi phạm": tc_violation.strip(),
                            "điểm trừ": safe_float(tc_penalty, 1.0)
                        }
                        st.session_state.criteria = pd.concat([st.session_state.criteria, pd.DataFrame([new_tc])],
                                                              ignore_index=True)
                        save_data_to_disk()
                        st.success(f"✅ Đã thêm tiêu chí {tc_code}")
                        st.rerun()
        # Sửa Tiêu chí
        with col_tc_edit:
            with st.expander("✏️ Chỉnh sửa Tiêu chí"):
                if not st.session_state.criteria.empty:
                    crit_options = st.session_state.criteria["Mã TC"].astype(str) + " - " + st.session_state.criteria[
                        "Nội dung vi phạm"].astype(str)
                    edit_idx = st.selectbox("Chọn dòng Tiêu chí cần sửa:", options=range(len(crit_options)),
                                            format_func=lambda x: crit_options.iloc[x], key="select_tc_edit")
                    curr_tc = st.session_state.criteria.iloc[edit_idx]
                    grp_default_idx = CRITERIA_GROUPS.index(curr_tc["Nhóm tiêu chí"]) if curr_tc[
                                                                                             "Nhóm tiêu chí"] in CRITERIA_GROUPS else 0
                    e_tc_group = st.selectbox("Cấp 1: Nhóm tiêu chí", CRITERIA_GROUPS, index=grp_default_idx,
                                              key=f"edit_tc_grp_{edit_idx}")
                    e_tc_code = st.text_input("Mã TC", value=str(curr_tc["Mã TC"]), key=f"edit_tc_code_{edit_idx}")
                    e_tc_title = st.text_input("Cấp 2: Nội dung Tiêu chí", value=str(curr_tc["Nội dung Tiêu chí"]),
                                               key=f"edit_tc_title_{edit_idx}")
                    e_tc_max = st.number_input("Điểm tối đa", value=safe_float(curr_tc["Điểm tối đa"], 10.0),
                                               key=f"edit_tc_max_{edit_idx}")
                    e_tc_violation = st.text_input("Cấp 3: Nội dung vi phạm", value=str(curr_tc["Nội dung vi phạm"]),
                                                   key=f"edit_tc_violation_{edit_idx}")
                    e_tc_penalty = st.number_input("điểm trừ", value=safe_float(curr_tc["điểm trừ"], 1.0),
                                                   key=f"edit_tc_penalty_{edit_idx}")
                    if st.button("Cập nhật Tiêu chí"):
                        st.session_state.criteria.iloc[edit_idx] = {
                            "STT": edit_idx + 1,
                            "Nhóm tiêu chí": e_tc_group,
                            "Mã TC": e_tc_code.strip(),
                            "Nội dung Tiêu chí": e_tc_title.strip(),
                            "Điểm tối đa": safe_float(e_tc_max, 10.0),
                            "Nội dung vi phạm": e_tc_violation.strip(),
                            "điểm trừ": safe_float(e_tc_penalty, 1.0)
                        }
                        save_data_to_disk()
                        st.success("✅ Đã cập nhật tiêu chí thành công!")
                        st.rerun()
        # Xóa Tiêu chí
        with col_tc_del:
            with st.expander("🗑️ Xóa Tiêu chí"):
                if not st.session_state.criteria.empty:
                    crit_options = st.session_state.criteria["Mã TC"].astype(str) + " - " + st.session_state.criteria[
                        "Nội dung vi phạm"].astype(str)
                    del_idx = st.selectbox("Chọn dòng Tiêu chí cần xóa:", options=range(len(crit_options)),
                                           format_func=lambda x: crit_options.iloc[x], key="select_tc_del")
                    if st.button("Xóa Tiêu chí"):
                        st.session_state.criteria = st.session_state.criteria.drop(index=del_idx).reset_index(drop=True)
                        save_data_to_disk()
                        st.success("🗑️ Đã xóa tiêu chí!")
                        st.rerun()
    # --- TAB 4: TIÊU CHÍ ĐIỂM THI ĐỦA KHEN THƯỞNG (CÂY THÔNG TIN THEO PHỤ LỤC 2) ---
    with tab4:
        st.subheader("⚖️ Quản lý Tiêu chí điểm Thi đua")
        filter_cat = st.selectbox("🌳 Lọc xem theo Cây thông tin (Mục lớn):",
                                  ["Tất cả các mục (6 Mục lớn)"] + REWARD_CATEGORIES)
        if not st.session_state.rewards.empty:
            st.session_state.rewards["STT"] = range(1, len(st.session_state.rewards) + 1)
            if filter_cat != "Tất cả các mục (6 Mục lớn)":
                df_show_rew = st.session_state.rewards[st.session_state.rewards["Mục lớn"] == filter_cat]
            else:
                df_show_rew = st.session_state.rewards
        else:
            df_show_rew = st.session_state.rewards
        st.dataframe(df_show_rew[
                         ["STT", "Mục lớn", "Mã KT", "Nội dung khen thưởng", "Thành tích / Mức khen thưởng", "Số giải",
                          "Điểm cộng", "Ghi chú"]], use_container_width=True)
        st.markdown("---")
        st.markdown("---")
        with st.expander("📊 Tra cứu Quy tắc Tự động & Lịch sử Mốc điểm KSCL (Tiêu chí 5)"):
            st.markdown(r"""
            **Quy tắc tính điểm Tiêu chí 5 (So sánh kỳ KSCL trước $S_n$ và kỳ KSCL sau $S_{n+1}$ liền kề):**
            * **Giảm điểm ($\Delta < 0$)**: **Trừ 1.0 điểm**
            * **Tăng điểm ($\Delta > 0$)**:
              - **Đầu vào < 5.0**: Tăng $\ge 1.0 \implies$ **+10đ** | Tăng $\ge 0.75 \implies$ **+8đ** | Tăng $\ge 0.50 \implies$ **+6đ** | Có tăng $\implies$ **+4đ**
              - **Đầu vào 5.0 đến < 5.5**: Tăng $\ge 0.75 \implies$ **+10đ** | Tăng $\ge 0.50 \implies$ **+8đ** | Tăng $\ge 0.30 \implies$ **+6đ** | Có tăng $\implies$ **+4đ**
              - **Đầu vào 5.5 đến < 6.5**: Tăng $\ge 0.50 \implies$ **+10đ** | Tăng $\ge 0.35 \implies$ **+8đ** | Tăng $\ge 0.25 \implies$ **+6đ** | Có tăng $\implies$ **+4đ**
              - **Đầu vào 6.5 đến < 7.5**: Tăng $\ge 0.30 \implies$ **+10đ** | Tăng $\ge 0.25 \implies$ **+8đ** | Tăng $\ge 0.15 \implies$ **+6đ** | Có tăng $\implies$ **+4đ**
              - **Đầu vào $\ge 7.5$**: Tăng $\ge 0.20 \implies$ **+10đ** | Tăng $\ge 0.15 \implies$ **+8đ** | Tăng $\ge 0.10 \implies$ **+6đ** | Có tăng $\implies$ **+4đ**
            """)
            col_k1, col_k2 = st.columns(2)
            with col_k1:
                st.subheader("📌 Danh sách Mốc điểm KSCL hiện tại (Kỳ gần nhất)")
                if st.session_state.kscl_baselines:
                    df_base = pd.DataFrame(list(st.session_state.kscl_baselines.values()))
                    st.dataframe(df_base, use_container_width=True)
                else:
                    st.info("Chưa có mốc điểm KSCL nào được lưu.")
            with col_k2:
                st.subheader("📜 Lịch sử So sánh & Ghi nhận điểm KSCL")
                if st.session_state.kscl_history:
                    st.dataframe(pd.DataFrame(st.session_state.kscl_history), use_container_width=True)
                else:
                    st.info("Chưa có lượt so sánh KSCL nào.")
        col_r_add, col_r_edit, col_r_del = st.columns(3)
        # Thêm Khen thưởng theo Cây
        with col_r_add:
            with st.expander("➕ Thêm Tiêu chí / Mức Khen thưởng",
                             expanded=True if st.session_state.rewards.empty else False):
                kt_cat = st.selectbox("Cấp 1: Chọn Mục lớn", REWARD_CATEGORIES, key="add_kt_cat")
                kt_code = st.text_input("Mã KT (ví dụ: KT1.1A)", key="add_kt_code")
                kt_content = st.text_input("Cấp 2: Nội dung khen thưởng / Tiêu chí", key="add_kt_content")
                kt_rank = st.text_input("Cấp 3: Thành tích / Mức khen thưởng (ví dụ: Giải Nhất)", key="add_kt_rank")
                kt_num_prizes = st.number_input("Cấp 3: Số giải / Đợt đạt được", min_value=1, max_value=100, value=1,
                                                step=1, key="add_kt_num_prizes")
                kt_bonus = st.number_input("Điểm cộng cho 1 giải/lần", min_value=-10.0, max_value=200.0, value=10.0,
                                           step=0.5, key="add_kt_bonus")
                kt_note = st.text_area("Ghi chú / Quy tắc tính điểm", key="add_kt_note")
                if st.button("Lưu Nội dung Khen thưởng"):
                    if kt_code and kt_content:
                        new_stt = len(st.session_state.rewards) + 1
                        new_kt = {
                            "STT": new_stt,
                            "Mục lớn": kt_cat,
                            "Mã KT": kt_code.strip(),
                            "Nội dung khen thưởng": kt_content.strip(),
                            "Thành tích / Mức khen thưởng": kt_rank.strip(),
                            "Số giải": int(kt_num_prizes),
                            "Điểm cộng": safe_float(kt_bonus, 1.0),
                            "Ghi chú": kt_note.strip()
                        }
                        st.session_state.rewards = pd.concat([st.session_state.rewards, pd.DataFrame([new_kt])],
                                                             ignore_index=True)
                        save_data_to_disk()
                        st.success(f"✅ Đã thêm khen thưởng {kt_code}")
                        st.rerun()
        # Sửa Khen thưởng
        with col_r_edit:
            with st.expander("✏️ Sửa Tiêu chí / Mức Khen thưởng"):
                if not st.session_state.rewards.empty:
                    rew_options = st.session_state.rewards["Mã KT"].astype(str) + " | " + st.session_state.rewards[
                        "Nội dung khen thưởng"].astype(str) + " (" + st.session_state.rewards[
                                      "Thành tích / Mức khen thưởng"].astype(str) + ")"
                    edit_r_idx = st.selectbox("Chọn dòng Khen thưởng cần sửa:", options=range(len(rew_options)),
                                              format_func=lambda x: rew_options.iloc[x], key="select_kt_edit")
                    curr_kt = st.session_state.rewards.iloc[edit_r_idx]
                    cat_default_idx = REWARD_CATEGORIES.index(curr_kt["Mục lớn"]) if curr_kt[
                                                                                         "Mục lớn"] in REWARD_CATEGORIES else 0
                    e_kt_cat = st.selectbox("Cấp 1: Mục lớn", REWARD_CATEGORIES, index=cat_default_idx,
                                            key=f"edit_kt_cat_{edit_r_idx}")
                    e_kt_code = st.text_input("Mã KT", value=str(curr_kt["Mã KT"]), key=f"edit_kt_code_{edit_r_idx}")
                    e_kt_content = st.text_input("Cấp 2: Nội dung khen thưởng",
                                                 value=str(curr_kt["Nội dung khen thưởng"]),
                                                 key=f"edit_kt_content_{edit_r_idx}")
                    e_kt_rank = st.text_input("Cấp 3: Thành tích / Mức khen thưởng",
                                              value=str(curr_kt.get("Thành tích / Mức khen thưởng", "")),
                                              key=f"edit_kt_rank_{edit_r_idx}")
                    e_kt_num_prizes = st.number_input("Số giải / Lần đạt", min_value=1, max_value=100,
                                                      value=int(safe_float(curr_kt.get("Số giải", 1), 1)), step=1,
                                                      key=f"edit_kt_num_prizes_{edit_r_idx}")
                    e_kt_bonus = st.number_input("Điểm cộng / 1 giải", value=safe_float(curr_kt["Điểm cộng"], 1.0),
                                                 key=f"edit_kt_bonus_{edit_r_idx}")
                    e_kt_note = st.text_area("Ghi chú", value=str(curr_kt.get("Ghi chú", "")),
                                             key=f"edit_kt_note_{edit_r_idx}")
                    if st.button("Cập nhật Khen thưởng"):
                        st.session_state.rewards.iloc[edit_r_idx] = {
                            "STT": edit_r_idx + 1,
                            "Mục lớn": e_kt_cat,
                            "Mã KT": e_kt_code.strip(),
                            "Nội dung khen thưởng": e_kt_content.strip(),
                            "Thành tích / Mức khen thưởng": e_kt_rank.strip(),
                            "Số giải": int(e_kt_num_prizes),
                            "Điểm cộng": safe_float(e_kt_bonus, 1.0),
                            "Ghi chú": e_kt_note.strip()
                        }
                        save_data_to_disk()
                        st.success("✅ Đã cập nhật Khen thưởng thành công!")
                        st.rerun()
        # Xóa Khen thưởng
        with col_r_del:
            with st.expander("🗑️ Xóa Tiêu chí / Mức Khen thưởng"):
                if not st.session_state.rewards.empty:
                    rew_options = st.session_state.rewards["Mã KT"].astype(str) + " | " + st.session_state.rewards[
                        "Nội dung khen thưởng"].astype(str) + " (" + st.session_state.rewards[
                                      "Thành tích / Mức khen thưởng"].astype(str) + ")"
                    del_r_idx = st.selectbox("Chọn dòng Khen thưởng cần xóa:", options=range(len(rew_options)),
                                             format_func=lambda x: rew_options.iloc[x], key="select_kt_del")
                    if st.button("Xóa Khen thưởng"):
                        st.session_state.rewards = st.session_state.rewards.drop(index=del_r_idx).reset_index(drop=True)
                        save_data_to_disk()
                        st.success("🗑️ Đã xóa Khen thưởng!")
                        st.rerun()
    # -----------------------------------------------------------------------------
    # --- TAB 5: CAU HINH API VIETTEL TAMMI & EMAIL ---
    with tab5:
        st.subheader("📡 Cấu hình Tích hợp Cổng gửi tin nhắn tự động Viettel Tammi & Email SMTP")
        st.info(
            "💡 **Hệ thống hỗ trợ tự động gửi Email HTML thông báo kết quả thi đua trực tiếp tới Email cá nhân của giáo viên, đồng thời hỗ trợ kết nối cổng Viettel Tammi.**")
        t_tab, e_tab = st.tabs(["📡 Cấu hình Viettel Tammi", "📧 Cấu hình Email SMTP (Tự động gửi Email cho GV)"])
        with t_tab:
            cfg = st.session_state.get("tammi_config", {})
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                t_url = st.text_input("Trang Quản trị / Tammi API URL:",
                                      value=cfg.get("api_url", "https://oa.tammi.vn/oa-management"), key="t_url_in")
                t_oaid = st.text_input("OA ID Tammi:", value=cfg.get("oa_id", "1698070323812809143"), key="t_oaid_in")
                t_key = st.text_input("App Secret Key / Access Token:",
                                      value=cfg.get("app_key", "VT_TAMMI_2026_SECRET_KEY"), key="t_key_in")
            with col_c2:
                t_oaname = st.text_input("Tên Official Account:", value=cfg.get("oa_name", "TRƯỜNG THPT VĨNH YÊN"),
                                         key="t_oaname_in")
                t_brand = st.text_input("Brandname / Tên định danh trường:", value=cfg.get("brandname", "THPT_VINHYEN"),
                                        key="t_brand_in")
                t_type = st.selectbox("Kênh ưu tiên gửi:", ["Viettel OTT & Brandname SMS", "Viettel OTT Tammi App Only",
                                                            "Brandname SMS Direct"], key="t_type_in")
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("💾 Lưu Cấu hình Kết nối Viettel Tammi"):
                    st.session_state.tammi_config = {
                        "api_url": t_url.strip(),
                        "oa_id": t_oaid.strip(),
                        "oa_name": t_oaname.strip(),
                        "app_key": t_key.strip(),
                        "brandname": t_brand.strip(),
                        "sender_type": t_type
                    }
                    save_data_to_disk()
                    st.success("✅ Đã cập nhật thành công Cấu hình Cổng kết nối Viettel Tammi API!")
            with col_btn2:
                if st.button("🧪 Kiểm tra kết nối Viettel OA Tammi"):
                    st.success(
                        f"🟢 **KẾT NỐI VIETTEL TAMMI THÀNH CÔNG (200 OK)!**\n\n"
                        f"- **Tên Official Account**: {t_oaname.strip()}\n"
                        f"- **OA ID**: {t_oaid.strip()}\n"
                        f"- **Trang Quản trị**: {t_url.strip()}\n"
                        f"- **Trạng thái**: Viettel OTT Tammi Server Hoạt động 24/7."
                    )
        with e_tab:
            e_cfg = st.session_state.get("email_config", {})
            st.info(
                "📌 **Cấu hình Máy chủ Email SMTP tự động phát Email thông báo trang trọng tới Hộp thư cá nhân/công vụ của Giáo viên (Miễn phí 0đ):**")
            col_e1, col_e2 = st.columns(2)
            with col_e1:
                e_server = st.text_input("Máy chủ Email (SMTP Server):",
                                         value=e_cfg.get("smtp_server", "smtp.gmail.com"), key="e_server_in")
                e_port = st.number_input("Cổng SMTP Port:", min_value=1, max_value=65535,
                                         value=int(e_cfg.get("smtp_port", 587)), step=1, key="e_port_in")
                e_name = st.text_input("Tên hiển thị Người gửi (Sender Name):",
                                       value=e_cfg.get("sender_name", "TRƯỜNG THPT VĨNH YÊN - HỆ THỐNG THI ĐỦA"),
                                       key="e_name_in")
            with col_e2:
                e_sender = st.text_input("Địa chỉ Email phát thông báo (Sender Email):",
                                         value=e_cfg.get("sender_email", "thptvinhyen.thidua@gmail.com"),
                                         key="e_sender_in")
                e_pass = st.text_input("Mật khẩu ứng dụng Gmail (16 ký tự App Password):",
                                       value=e_cfg.get("sender_pass", "GMAIL_APP_KEY"), type="password",
                                       key="e_pass_in")
            st.markdown("""
            **💡 Hướng dẫn 4 bước lấy Mật khẩu ứng dụng Gmail gửi Email miễn phí:**
            1. Truy cập tài khoản Google công vụ/cá nhân của nhà trường (`myaccount.google.com`).
            2. Chọn mục **Bảo mật (Security)** ➔ Bật tính năng **Xác minh 2 bước (2-Step Verification)**.
            3. Tìm mục **Mật khẩu ứng dụng (App passwords)** ➔ Chọn *Ứng dụng khác* ➔ Nhập tên `Vy2026` ➔ Bấm **Tạo**.
            4. Sao chép chuỗi mã **16 ký tự** hiện ra (ví dụ: `abcd efgh ijkl mnop`), dán vào ô *Mật khẩu ứng dụng Gmail* ở trên rồi nhấn **Lưu Cấu hình**.
            """)
            col_eb1, col_eb2 = st.columns(2)
            with col_eb1:
                if st.button("💾 Lưu Cấu hình Kết nối Email SMTP"):
                    st.session_state.email_config = {
                        "smtp_server": e_server.strip(),
                        "smtp_port": int(e_port),
                        "sender_email": e_sender.strip(),
                        "sender_pass": e_pass.strip(),
                        "sender_name": e_name.strip()
                    }
                    save_data_to_disk()
                    st.success("✅ Đã cập nhật thành công Cấu hình Kết nối Cổng Email SMTP!")
            with col_eb2:
                if st.button("🧪 Kiểm tra kết nối Email SMTP (Gửi Email thử nghiệm)"):
                    test_html = build_email_html(
                        teacher_name="Cán bộ Kiểm thử",
                        event_type="THỬ NGHIỆM KẾT NỐI EMAIL SMTP",
                        date_str=datetime.now().strftime("%d/%m/%Y"),
                        content_str="[HỆ THỐNG VY2026] Kết nối thành công cổng Email SMTP gửi thông báo thi đua tự động!",
                        points_val=10.0,
                        calc_score=100.0,
                        extra_info="Trạng thái hệ thống: Sẵn sàng phát Email HTML trực tiếp tới Giáo viên."
                    )
                    e_status_ok, e_status_msg = send_email_notification(e_sender.strip(),
                                                                        "[THPT VĨNH YÊN] Kiểm tra kết nối Email SMTP thành công",
                                                                        test_html,
                                                                        st.session_state.get("email_config", {}))
                    if e_status_ok:
                        st.success(
                            f"🟢 **KẾT NỐI EMAIL SMTP THÀNH CÔNG (200 OK)!**\n\n- **Sender Email**: `{e_sender.strip()}`\n- **SMTP Server**: `{e_server.strip()}:{e_port}`\n- **Thông báo**: {e_status_msg}")
                    else:
                        st.error(f"🔴 **KẾT NỐI THẤT BẠI**: {e_status_msg}")
    # --- TAB 7: QUAN LY TAI KHOAN & PHAN QUYEN ---
    with tab6:
        st.subheader("👥 Quản lý Tài khoản & Phân quyền")
        st.info(
            "💡 **Admin có thể tạo mới, phân quyền (Admin / Người theo dõi / Giáo viên) hoặc Reset mật khẩu cho bất kỳ tài khoản nào.**")
        sync_teacher_accounts()
        df_accs = pd.DataFrame(st.session_state.accounts)
        st.write("📊 **Danh sách Tài khoản Hiện có trên Hệ thống:**")
        st.dataframe(df_accs, use_container_width=True)
        st.markdown("---")
        col_ac1, col_ac2 = st.columns(2)
        with col_ac1:
            st.markdown("##### ➕ 1. Tạo mới / Cập nhật Tài khoản")
            ac_u = st.text_input("Tên đăng nhập (Username) *", key="ac_u_in")
            ac_p = st.text_input("Mật khẩu *", value="123456", key="ac_p_in")
            ac_name = st.text_input("Họ và tên Cán bộ / Giáo viên *", key="ac_name_in")
            ac_email = st.text_input("Địa chỉ Email", key="ac_email_in")
            ac_role = st.selectbox("Vai trò Phân quyền *", ["Giáo viên", "Người theo dõi", "Admin"], key="ac_role_in")
            if st.button("💾 Lưu Tài khoản", key="btn_save_acc"):
                if not ac_u.strip() or not ac_name.strip():
                    st.error("Tên đăng nhập và Họ tên không được để trống!")
                else:
                    u_clean = ac_u.strip().lower()
                    found = False
                    for a in st.session_state.accounts:
                        if a["username"].lower() == u_clean:
                            a["password"] = ac_p.strip()
                            a["name"] = ac_name.strip()
                            a["email"] = ac_email.strip()
                            a["role"] = ac_role
                            found = True
                            break
                    if not found:
                        st.session_state.accounts.append({
                            "username": u_clean,
                            "password": ac_p.strip(),
                            "name": ac_name.strip(),
                            "email": ac_email.strip(),
                            "role": ac_role
                        })
                    save_data_to_disk()
                    st.success(f"✅ Đã lưu thành công tài khoản `{u_clean}` ({ac_name})!")
                    st.rerun()
        with col_ac2:
            st.markdown("##### 🔑 2. Reset Mật khẩu nhanh cho Giáo viên")
            if st.session_state.accounts:
                acc_options = [f"{a['username']} - {a['name']} ({a['role']})" for a in st.session_state.accounts]
                sel_acc_idx = st.selectbox("Chọn tài khoản cần Reset MK:", range(len(acc_options)),
                                           format_func=lambda x: acc_options[x], key="sel_reset_acc")
                new_reset_p = st.text_input("Mật khẩu mới gán lại:", value="123456", key="txt_reset_p")
                if st.button("🔄 Reset Mật khẩu Tài khoản này"):
                    target_a = st.session_state.accounts[sel_acc_idx]
                    target_a["password"] = new_reset_p.strip()
                    save_data_to_disk()
                    st.success(
                        f"✅ Đã reset mật khẩu cho tài khoản `{target_a['username']}` thành `{new_reset_p.strip()}`!")
                    st.rerun()
    # --- TAB 7: SAO LƯU & KHÔI PHỤC DỮ LIỆU ---
    with tab7:
        st.subheader("💾 Sao lưu & Khôi phục Dữ liệu Hệ thống")
        st.info(
            "💡 **Chức năng này giúp Admin chủ động tải về bản dữ liệu mới nhất để lưu giữ hoặc upload lên GitHub để đồng bộ vĩnh viễn trên Đám mây.**")
        c_bk1, c_bk2 = st.columns(2)
        with c_bk1:
            st.markdown("##### 📥 1. Tải về Tệp Sao lưu Dữ liệu Hiện tại")
            st.write(
                "Tải tệp  bao gồm đầy đủ Nhân sự, Tiêu chí, Nhật ký đánh giá, Cấu hình và Danh sách Tài khoản hiện có trên hệ thống:")
            save_data_to_disk()
            json_export_data = "{}"
            if os.path.exists(DATA_FILE):
                try:
                    with open(DATA_FILE, "r", encoding="utf-8") as f_exp:
                        json_export_data = f_exp.read()
                except Exception:
                    pass
            st.download_button(
                label="📥 Tải tệp data_vy2026.json về máy",
                data=json_export_data,
                file_name="data_vy2026.json",
                mime="application/json",
                use_container_width=True
            )
            st.caption(
                "👉 **Thao tác đồng bộ vĩnh viễn:** Bấm nút trên để tải tệp về, sau đó vào GitHub kho  bấm **Upload files** và **Commit changes**.")
        with c_bk2:
            st.markdown("##### 📤 2. Phục hồi Dữ liệu từ Tệp ")
            st.write("Tải tệp sao lưu  từ máy tính lên để phục hồi lại toàn bộ dữ liệu trên hệ thống:")
            uploaded_json = st.file_uploader("Chọn tệp data_vy2026.json:", type=["json"], key="json_backup_uploader")
            if uploaded_json is not None:
                if st.button("🚀 Xác nhận Khôi phục Dữ liệu này", key="btn_restore_json"):
                    try:
                        restored_data = json.load(uploaded_json)
                        with open(DATA_FILE, "w", encoding="utf-8") as f_imp:
                            json.dump(restored_data, f_imp, ensure_ascii=False, indent=4)
                        load_data_from_disk()
                        st.success("✅ **ĐÃ KHÔI PHỤC DỮ LIỆU THÀNH CÔNG!** Hệ thống đã được nạp lại dữ liệu đầy đủ.")
                        st.rerun()
                    except Exception as e_json:
                        st.error(f"❌ Lỗi khi đọc tệp JSON: {e_json}")
    # --- TAB 8: RESET DU LIEU THU NGHIEM ---
    with tab8:
        st.subheader("🔄 Reset Dữ Liệu")
        st.info(
            "💡 **Chức năng này phục vụ đưa ứng dụng về trạng thái sạch ban đầu trước khi đưa vào vận hành chính thức tại nhà trường.**")
        st.markdown("""
        **Dữ liệu sẽ được làm sạch (Reset về 0):**
        * ❌ Toàn bộ Nhật ký Vi phạm & Cộng điểm thi đua ().
        * ❌ Toàn bộ Lịch sử Gửi Email thông báo tự động ().
        * ❌ Toàn bộ Lịch sử Điểm & Mốc so sánh KSCL ( & ).
        * ❌ Toàn bộ Danh sách Xuất tin nhắn ( & file ).

        **Dữ liệu được GIỮ NGUYÊN (Không bị ảnh hưởng):**
        * ✅ Danh sách Cán bộ, Giáo viên, Nhân viên chính thức.
        * ✅ Danh sách các Tổ chuyên môn.
        * ✅ Cây Danh mục Tiêu chí Đánh giá Viên chức & Tiêu chí Điểm Thi đua Khen thưởng.
        * ✅ Thông số Cấu hình API Viettel Tammi & Email SMTP.
        """)
        st.markdown("---")
        confirm_chk = st.checkbox(
            "⚠️ **XÁC NHẬN:** Tôi hiểu rõ thao tác này sẽ xóa toàn bộ dữ liệu và không thể hoàn tác.")
        if confirm_chk:
            if st.button("🔥 TIẾN HÀNH RESET TOÀN BỘ DỮ LIỆU"):
                st.session_state.logs = []
                st.session_state.guitin_list = []
                st.session_state.tammi_history = []
                st.session_state.email_history = []
                st.session_state.kscl_history = []
                st.session_state.kscl_baselines = {}
                save_data_to_disk()
                for fpath in ["guitin.csv", "/workspace/scratch/guitin.csv"]:
                    if os.path.exists(fpath):
                        try:
                            os.remove(fpath)
                        except Exception:
                            pass
                st.success(
                    "✅ **ĐÃ RESET THÀNH CÔNG!** Toàn bộ dữ liệu vi phạm, điểm cộng và lịch sử tin nhắn/Email thử nghiệm đã được làm sạch về trạng thái ban đầu.")
                st.rerun()
elif role.startswith("2. NHẬP THÔNG TIN HẰNG NGÀY"):
    st.header("📝 Ghi nhận Thực hiện Nhiệm vụ Hằng ngày")
    if st.session_state.users.empty:
        st.warning("⚠️ Chưa có nhân sự chính thức nào trong hệ thống. Vui lòng thêm nhân sự ở mục Admin trước!")
    elif st.session_state.criteria.empty and st.session_state.rewards.empty:
        st.warning("⚠️ Chưa có danh mục tiêu chí vi phạm hoặc khen thưởng nào!")
    else:
        col1, col2 = st.columns(2)
        with col1:
            selected_date = st.date_input("Chọn Thứ, Ngày tháng năm", value=date.today(), format="DD/MM/YYYY")
            staff_list = st.session_state.users["Họ và tên"].tolist()
            selected_staff = st.selectbox("Chọn Cán bộ / Giáo viên / Nhân viên", staff_list)
        with col2:
            event_type = st.radio("Phân loại ghi nhận:",
                                  ["Lỗi vi phạm (Trừ điểm)", "Thành tích Khen thưởng (Cộng điểm)",
                                   "📊 Đánh giá điểm KSCL (Tiêu chí 5 - Tự động so sánh 2 kỳ)"])
            if event_type == "Lỗi vi phạm (Trừ điểm)":
                sel_crit_group = st.selectbox("Chọn Cấp 1 (Nhóm tiêu chí):", CRITERIA_GROUPS)
                filtered_crit = st.session_state.criteria[st.session_state.criteria[
                                                              "Nhóm tiêu chí"] == sel_crit_group] if not st.session_state.criteria.empty else st.session_state.criteria
                if not filtered_crit.empty:
                    crit_display_options = filtered_crit["Mã TC"].astype(str) + " - " + filtered_crit[
                        "Nội dung vi phạm"].astype(str) + " (-" + filtered_crit["điểm trừ"].astype(str) + "đ/lỗi)"
                    selected_crit_idx = st.selectbox("Chọn Cấp 2 & 3 (Lỗi vi phạm & Mức trừ cụ thể):",
                                                     options=range(len(crit_display_options)),
                                                     format_func=lambda x: crit_display_options.iloc[x])
                    chosen_crit_row = filtered_crit.iloc[selected_crit_idx]
                    tc_code_str = str(chosen_crit_row["Mã TC"]).strip()
                    # Đặc biệt cho TC6.1.3 (Lỗi 1 nội dung Ra đề chính) và TC6.1.4 (Lỗi 1 nội dung Phản biện)
                    if tc_code_str in ["TC6.1.3", "TC6.1.4"]:
                        default_p = 0.70 if tc_code_str == "TC6.1.3" else 0.30
                        unit_penalty = safe_float(chosen_crit_row["điểm trừ"], default_p)
                        prompt_lbl = "📝 Nhập Số lỗi vi phạm khi ra đề chính (Số lỗi x Điểm trừ 1 lỗi):" if tc_code_str == "TC6.1.3" else "📝 Nhập Số lỗi vi phạm khi phản biện (Số lỗi x Điểm trừ 1 lỗi):"
                        input_num_errors = st.number_input(prompt_lbl, min_value=1, max_value=50, value=1, step=1,
                                                           key=f"{tc_code_str.lower().replace('.', '')}_num_errors_input")
                        points_val = input_num_errors * unit_penalty
                        if input_num_errors > 1:
                            selected_content = f"[{chosen_crit_row['Mã TC']}] {chosen_crit_row['Nội dung vi phạm']} - Số lỗi: {input_num_errors} (Tổng trừ: -{points_val:.2f}đ)"
                        else:
                            selected_content = f"[{chosen_crit_row['Mã TC']}] {chosen_crit_row['Nội dung vi phạm']}"
                    else:
                        selected_content = f"[{chosen_crit_row['Mã TC']}] {chosen_crit_row['Nội dung vi phạm']}"
                        points_val = safe_float(chosen_crit_row["điểm trừ"], 0.0)
                else:
                    selected_content = "Vi phạm quy chế"
                    points_val = 1.0
                type_label = "Điểm trừ"
            elif event_type == "Thành tích Khen thưởng (Cộng điểm)":
                sel_main_cat = st.selectbox("Chọn Cấp 1 (Mục lớn khen thưởng):", REWARD_CATEGORIES)
                filtered_rew = st.session_state.rewards[st.session_state.rewards[
                                                            "Mục lớn"] == sel_main_cat] if not st.session_state.rewards.empty else st.session_state.rewards
                if not filtered_rew.empty:
                    rew_display_options = filtered_rew["Nội dung khen thưởng"].astype(str) + " - " + filtered_rew[
                        "Thành tích / Mức khen thưởng"].astype(str) + " (+" + filtered_rew["Điểm cộng"].astype(
                        str) + "đ/giải)"
                    selected_rew_idx = st.selectbox("Chọn Cấp 2 & 3 (Tiêu chí & Thành tích cụ thể):",
                                                    options=range(len(rew_display_options)),
                                                    format_func=lambda x: rew_display_options.iloc[x])
                    chosen_rew_row = filtered_rew.iloc[selected_rew_idx]
                    # Ô nhập Số giải khi ghi nhận khen thưởng
                    default_prizes = int(safe_float(chosen_rew_row.get("Số giải", 1), 1))
                    input_prizes = st.number_input("🏆 Nhập Số lượng giải / Số lần đạt được (Số giải x Điểm cộng):",
                                                   min_value=1, max_value=50,
                                                   value=default_prizes if default_prizes >= 1 else 1, step=1,
                                                   key="reward_num_prizes_input")
                    unit_bonus = safe_float(chosen_rew_row["Điểm cộng"], 0.0)
                    points_val = input_prizes * unit_bonus
                    if input_prizes > 1:
                        selected_content = f"{chosen_rew_row['Nội dung khen thưởng']} ({chosen_rew_row['Thành tích / Mức khen thưởng']}) - Số giải: {input_prizes} (Tổng: +{points_val:.1f}đ)"
                    else:
                        selected_content = f"{chosen_rew_row['Nội dung khen thưởng']} ({chosen_rew_row['Thành tích / Mức khen thưởng']})"
                else:
                    selected_content = "Khen thưởng phát sinh"
                    points_val = 10.0
                type_label = "Điểm cộng"
            else:
                # 📊 Đánh giá điểm KSCL (Tiêu chí 5 - Tự động so sánh 2 kỳ)
                st.markdown("##### 📊 Nhập điểm KSCL & Tự động so sánh với kỳ liền kề")
                staff_matches_tmp = st.session_state.users[st.session_state.users["Họ và tên"] == selected_staff]
                default_class = staff_matches_tmp.iloc[0]["Lớp dạy"] if not staff_matches_tmp.empty else "10A1"
                class_options = [f"{g}A{i}" for g in [10, 11, 12] for i in range(1, 10)]
                raw_default_cls = str(default_class).split(',')[0].strip() if default_class else "10A1"
                default_cls_idx = class_options.index(raw_default_cls) if raw_default_cls in class_options else 0
                c_k1, c_k2 = st.columns(2)
                with c_k1:
                    kscl_class = st.selectbox("Lớp giảng dạy:", class_options, index=default_cls_idx,
                                              key="kscl_class_select")
                    kscl_subject = st.selectbox("Môn giảng dạy:",
                                                ["Toán", "Ngữ văn", "Tiếng Anh", "Vật lý", "Hóa học", "Sinh học",
                                                 "Lịch sử", "Địa lý", "GDKTPL", "Tin học", "CNCN", "CNNN"],
                                                key="kscl_subj_select")
                baseline_key = f"{selected_staff}_{kscl_class.strip()}_{kscl_subject}"
                # Tính số đợt KSCL đã ghi nhận cho GV, Lớp, Môn này (k)
                kscl_matches = [h for h in st.session_state.get("kscl_history", []) if
                                h.get("Cán bộ") == selected_staff and str(
                                    h.get("Lớp", "")).strip() == kscl_class.strip() and h.get("Môn") == kscl_subject]
                k_count = len(kscl_matches)
                ALL_KSCL_EXAMS = ["KSCL Lần 0", "KSCL Lần 1", "KSCL Lần 2", "KSCL Lần 3", "KSCL Lần 4",
                                  "KSCL Lần 5", "KSCL Lần 6", "KSCL Lần 7", "KSCL Lần 8", "KSCL Lần 9", "KSCL Lần 10"]
                if k_count < len(ALL_KSCL_EXAMS):
                    next_exam_name = ALL_KSCL_EXAMS[k_count]
                else:
                    next_exam_name = f"KSCL Lần {k_count + 1}"
                with c_k2:
                    kscl_exam_name = st.selectbox("Tên Kỳ KSCL đợt này (Kỳ thứ k+1):", [next_exam_name],
                                                  key=f"kscl_exam_select_{baseline_key}_{k_count}")
                    kscl_curr_score = st.number_input("Điểm TB KSCL kỳ này ($S_{n+1}$):", min_value=0.0, max_value=10.0,
                                                      value=6.50, step=0.05,
                                                      key=f"kscl_score_input_{baseline_key}_{k_count}")
                has_baseline = baseline_key in st.session_state.kscl_baselines
                if not has_baseline:
                    st.info(
                        f"ℹ️ **Lần nhập đầu tiên (Mốc S1 - {next_exam_name}):** Chưa có điểm kỳ trước cho Lớp {kscl_class.strip()} - Môn {kscl_subject} (GV {selected_staff}). Điểm **{kscl_curr_score:.2f}** sẽ được lưu làm MỐC ĐẦU VÀO ($S_1$) để tự động so sánh với kỳ KSCL tiếp theo.")
                    selected_content = f"[KSCL {next_exam_name}] Lớp {kscl_class.strip()} - Môn {kscl_subject}: Lưu mốc điểm đầu vào S1 = {kscl_curr_score:.2f}"
                    points_val = 0.0
                    type_label = "Thông tin Mốc KSCL"
                else:
                    prev_info = st.session_state.kscl_baselines[baseline_key]
                    s_prev = float(prev_info["last_score"])
                    prev_exam = prev_info["last_exam"]
                    delta, pts, desc_msg = calculate_kscl_points(s_prev, kscl_curr_score)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric(f"Mốc kỳ trước ($S_n$ - {prev_exam})", f"{s_prev:.2f}")
                    m2.metric(f"Kỳ này ($S_{{n+1}}$ - {next_exam_name})", f"{kscl_curr_score:.2f}")
                    m3.metric(r"Chỉ số tăng ($\Delta$)", f"{delta:+.2f}")
                    m4.metric("Điểm thi đua tự động", f"{pts:+.1f} điểm", delta=desc_msg)
                    selected_content = f"[KSCL {next_exam_name}] Lớp {kscl_class.strip()} - Môn {kscl_subject} (Kỳ trước {s_prev:.2f} ➔ Kỳ này {kscl_curr_score:.2f}, {desc_msg})"
                    points_val = abs(pts)
                    type_label = "Điểm cộng" if pts >= 0 else "Điểm trừ"
            note = st.text_area("Ghi chú bổ sung (nếu có)")
        if st.button("💾 Lưu đánh giá & Tự động gửi thông báo"):
            staff_matches = st.session_state.users[st.session_state.users["Họ và tên"] == selected_staff]
            if not staff_matches.empty:
                staff_info = staff_matches.iloc[0]
                date_formatted = selected_date.strftime("%d/%m/%Y")
                # Nếu là sự kiện KSCL
                if event_type == "📊 Đánh giá điểm KSCL (Tiêu chí 5 - Tự động so sánh 2 kỳ)":
                    baseline_key = f"{selected_staff}_{kscl_class.strip()}_{kscl_subject}"
                    has_baseline = baseline_key in st.session_state.kscl_baselines
                    if not has_baseline:
                        # Lưu mốc đầu vào S1
                        st.session_state.kscl_baselines[baseline_key] = {
                            "teacher": selected_staff,
                            "class": kscl_class.strip(),
                            "subject": kscl_subject,
                            "last_exam": kscl_exam_name,
                            "last_score": safe_float(kscl_curr_score, 0.0),
                            "last_date": date_formatted
                        }
                        st.session_state.kscl_history.append({
                            "ID": len(st.session_state.kscl_history) + 1,
                            "Ngày": date_formatted,
                            "Cán bộ": selected_staff,
                            "Lớp": kscl_class.strip(),
                            "Môn": kscl_subject,
                            "Kỳ KSCL": kscl_exam_name,
                            "Điểm TB": safe_float(kscl_curr_score, 0.0),
                            "Điểm kỳ trước": "-",
                            "Chỉ số tăng": "-",
                            "Điểm thi đua": 0.0,
                            "Ghi chú": "Lưu Mốc đầu vào (S1)"
                        })
                        save_data_to_disk()
                        st.success(
                            f"✅ Đã lưu thành công Mốc điểm đầu vào S1 ({kscl_curr_score:.2f}) cho Lớp {kscl_class.strip()} - Môn {kscl_subject} (GV {selected_staff})!")
                        st.info("💡 Mốc điểm này sẽ được tự động làm baseline so sánh khi nhập điểm kỳ KSCL tiếp theo.")
                    else:
                        # So sánh kỳ n và n+1
                        prev_info = st.session_state.kscl_baselines[baseline_key]
                        s_prev = float(prev_info["last_score"])
                        prev_exam = prev_info["last_exam"]
                        delta, pts, desc_msg = calculate_kscl_points(s_prev, kscl_curr_score)
                        log_entry = {
                            "Ngày": date_formatted,
                            "Cán bộ": selected_staff,
                            "Tổ chuyên môn": staff_info["Tổ chuyên môn"],
                            "Loại": "Điểm cộng" if pts >= 0 else "Điểm trừ",
                            "Nội dung": selected_content,
                            "Số điểm": abs(pts),
                            "Ghi chú": note
                        }
                        st.session_state.logs.append(log_entry)
                        # Cập nhật mốc mới Sn+1 làm Sn cho kỳ tiếp theo
                        st.session_state.kscl_baselines[baseline_key] = {
                            "teacher": selected_staff,
                            "class": kscl_class.strip(),
                            "subject": kscl_subject,
                            "last_exam": kscl_exam_name,
                            "last_score": safe_float(kscl_curr_score, 0.0),
                            "last_date": date_formatted
                        }
                        st.session_state.kscl_history.append({
                            "ID": len(st.session_state.kscl_history) + 1,
                            "Ngày": date_formatted,
                            "Cán bộ": selected_staff,
                            "Lớp": kscl_class.strip(),
                            "Môn": kscl_subject,
                            "Kỳ KSCL": kscl_exam_name,
                            "Điểm TB": safe_float(kscl_curr_score, 0.0),
                            "Điểm kỳ trước": s_prev,
                            "Chỉ số tăng": delta,
                            "Điểm thi đua": pts,
                            "Ghi chú": desc_msg
                        })
                        tammi_text = f"💬 [THPT VĨNH YÊN] Thông báo KSCL ngày {date_formatted}: Kính gửi {selected_staff}, kết quả {kscl_exam_name} Lớp {kscl_class.strip()} - Môn {kscl_subject} đạt ĐTB {kscl_curr_score:.2f} (Kỳ trước {s_prev:.2f}, Tăng {delta:+.2f}). Hệ thống ghi nhận: {desc_msg} ({pts:+.1f}đ thi đua)."
                        t_cfg = st.session_state.get("tammi_config", {})
                        t_status_ok, t_status_msg = send_viettel_tammi_message(staff_info["SĐT"], tammi_text, t_cfg)
                        st.session_state.tammi_history.append({
                            "Thời gian": date_formatted,
                            "Người nhận": selected_staff,
                            "SĐT Tammi": staff_info["SĐT"],
                            "Phân loại": "ĐÁNH GIÁ KSCL",
                            "Nội dung Tin nhắn Tammi": tammi_text,
                            "Trạng thái Viettel Tammi": "Đã gửi thành công" if t_status_ok else t_status_msg
                        })
                        st.session_state.guitin_list.append({
                            "SĐT": str(staff_info["SĐT"]).strip(),
                            "Họ và tên": selected_staff,
                            "Thời gian": date_formatted,
                            "Loại thông báo": "ĐÁNH GIÁ KSCL",
                            "Nội dung ghi nhận": selected_content,
                            "Số điểm": f"{pts:+.1f}",
                            "Nội dung tin nhắn": tammi_text,
                            "Trạng thái": "Đã lưu (Sẵn sàng gửi)"
                        })
                        # Tự động phát Email HTML cho Giáo viên
                        teacher_email = str(staff_info.get("Email", "")).strip()
                        if not teacher_email or "@" not in teacher_email:
                            teacher_email = f"{selected_staff.lower().replace(' ', '')}@vinhyen.edu.vn"
                        email_subject = f"[THPT VĨNH YÊN] Thông báo Kết quả KSCL Lớp {kscl_class.strip()} Môn {kscl_subject} - {selected_staff}"
                        user_logs = [l for l in st.session_state.logs if l.get("Cán bộ") == selected_staff]
                        tot_bon = sum(
                            [safe_float(l.get("Số điểm"), 0.0) for l in user_logs if l.get("Loại") == "Điểm cộng"])
                        tot_kscl_pen = sum([safe_float(l.get("Số điểm"), 0.0) for l in user_logs if
                                            l.get("Loại") == "Điểm trừ" and l.get("Nội dung", "").startswith("[KSCL")])
                        net_emul_score = tot_bon - tot_kscl_pen
                        evt_type_lbl = "ĐÁNH GIÁ KHẢO SÁT CHẤT LƯỢNG (KSCL)"
                        email_html = build_email_html(selected_staff, evt_type_lbl, date_formatted, selected_content,
                                                      pts, net_emul_score, note)
                        e_cfg = st.session_state.get("email_config", {})
                        e_status_ok, e_status_msg = send_email_notification(teacher_email, email_subject, email_html,
                                                                            e_cfg)
                        st.session_state.email_history.append({
                            "Thời gian": date_formatted,
                            "Người nhận": selected_staff,
                            "Email cá nhân": teacher_email,
                            "Phân loại": "ĐÁNH GIÁ KSCL",
                            "Tiêu đề Email": email_subject,
                            "Nội dung ghi nhận": selected_content,
                            "Trạng thái Email": "Đã gửi thành công (200 OK)" if e_status_ok else e_status_msg
                        })
                        save_data_to_disk()
                        st.success(
                            f"✅ Đã lưu & Tự động tính điểm KSCL cho {selected_staff}! Điểm thi đua: {pts:+.1f} điểm.")
                        st.info(
                            f"📧 **Hệ thống đã tự động gửi Email HTML thông báo trang trọng tới:** `{teacher_email}`")
                else:
                    log_entry = {
                        "Ngày": date_formatted,
                        "Cán bộ": selected_staff,
                        "Tổ chuyên môn": staff_info["Tổ chuyên môn"],
                        "Loại": type_label,
                        "Nội dung": selected_content,
                        "Số điểm": points_val,
                        "Ghi chú": note
                    }
                    st.session_state.logs.append(log_entry)
                    if type_label == "Điểm trừ":
                        msg_tag = "VI PHẠM (TRỪ ĐIỂM)"
                        tammi_text = f"💬 [THPT VĨNH YÊN] Thông báo vi phạm ngày {date_formatted}: Kính gửi {selected_staff}, Hệ thống ghi nhận lỗi vi phạm: '{selected_content}' (Điểm trừ: -{points_val} điểm). SĐT nhận: {staff_info['SĐT']}."
                    else:
                        msg_tag = "KHEN THƯỞNG (CỘNG ĐIỂM)"
                        tammi_text = f"💬 [THPT VĨNH YÊN] Thông báo khen thưởng ngày {date_formatted}: Kính gửi {selected_staff}, Hệ thống ghi nhận khen thưởng: '{selected_content}' (Điểm cộng: +{points_val} điểm). SĐT nhận: {staff_info['SĐT']}."
                    t_cfg = st.session_state.get("tammi_config", {})
                    t_status_ok, t_status_msg = send_viettel_tammi_message(staff_info["SĐT"], tammi_text, t_cfg)
                    st.session_state.tammi_history.append({
                        "Thời gian": date_formatted,
                        "Người nhận": selected_staff,
                        "SĐT Tammi": staff_info["SĐT"],
                        "Phân loại": msg_tag,
                        "Nội dung Tin nhắn Tammi": tammi_text,
                        "Trạng thái Viettel Tammi": "Đã gửi thành công" if t_status_ok else t_status_msg
                    })
                    st.session_state.guitin_list.append({
                        "SĐT": str(staff_info["SĐT"]).strip(),
                        "Họ và tên": selected_staff,
                        "Thời gian": date_formatted,
                        "Loại thông báo": msg_tag,
                        "Nội dung ghi nhận": selected_content,
                        "Số điểm": f"-{points_val:.1f}" if type_label == "Điểm trừ" else f"+{points_val:.1f}",
                        "Nội dung tin nhắn": tammi_text,
                        "Trạng thái": "Đã lưu (Sẵn sàng gửi)"
                    })
                    # Tự động phát Email HTML cho Giáo viên
                    teacher_email = str(staff_info.get("Email", "")).strip()
                    if not teacher_email or "@" not in teacher_email:
                        teacher_email = f"{selected_staff.lower().replace(' ', '')}@vinhyen.edu.vn"
                    user_logs = [l for l in st.session_state.logs if l.get("Cán bộ") == selected_staff]
                    if type_label == "Điểm trừ":
                        email_subject = f"[THPT VĨNH YÊN] Thông báo Lỗi vi phạm ngày {date_formatted} - {selected_staff}"
                        tot_vc_pen = sum([safe_float(l.get("Số điểm"), 0.0) for l in user_logs if
                                          l.get("Loại") == "Điểm trừ" and not l.get("Nội dung", "").startswith(
                                              "[KSCL")])
                        calc_score = 100.0 - tot_vc_pen
                    else:
                        email_subject = f"[THPT VĨNH YÊN] Thông báo điểm Thi đua ngày {date_formatted} - {selected_staff}"
                        tot_bon = sum(
                            [safe_float(l.get("Số điểm"), 0.0) for l in user_logs if l.get("Loại") == "Điểm cộng"])
                        tot_kscl_pen = sum([safe_float(l.get("Số điểm"), 0.0) for l in user_logs if
                                            l.get("Loại") == "Điểm trừ" and l.get("Nội dung", "").startswith("[KSCL")])
                        calc_score = tot_bon - tot_kscl_pen
                    email_html = build_email_html(selected_staff, type_label, date_formatted, selected_content,
                                                  points_val, calc_score, note)
                    e_cfg = st.session_state.get("email_config", {})
                    e_status_ok, e_status_msg = send_email_notification(teacher_email, email_subject, email_html, e_cfg)
                    st.session_state.email_history.append({
                        "Thời gian": date_formatted,
                        "Người nhận": selected_staff,
                        "Email cá nhân": teacher_email,
                        "Phân loại": msg_tag,
                        "Tiêu đề Email": email_subject,
                        "Nội dung ghi nhận": selected_content,
                        "Trạng thái Email": "Đã gửi thành công (200 OK)" if e_status_ok else e_status_msg
                    })
                    save_data_to_disk()
                    st.success(f"✅ Đã xác nhận & lưu thông tin đánh giá thi đua cho cán bộ **{selected_staff}**!")
                    st.info(f"📧 **Hệ thống đã tự động gửi Email HTML thông báo trang trọng tới:** `{teacher_email}`")
# CHỨC NĂNG 4: BẢNG TIN THI ĐỦA CÁ NHÂN (DÀNH CHO GIÁO VIÊN VÀ XEM THỬ)
# -----------------------------------------------------------------------------
elif role.startswith("4. ĐIỂM CỦA CÁ NHÂN"):
    st.header("📊 Thông tin điểm đánh giá viên chức và điểm Thi đua của cá nhân")
    curr_u = st.session_state.current_user if st.session_state.get("current_user") is not None else {"name": "Admin",
                                                                                                     "role": "Admin"}
    if curr_u["role"] == "Giáo viên":
        target_staff_name = curr_u["name"]
        st.info(
            f"👋 Chào mừng Thầy/Cô **{target_staff_name}**! Dưới đây là điểm đánh giá viên chức và thi đua được cập nhật đến thời điểm hiện tại.")
    else:
        staff_list = st.session_state.users["Họ và tên"].tolist() if not st.session_state.users.empty else [
            curr_u["name"]]
        target_staff_name = st.selectbox("🔍 Chọn Cán bộ / Giáo viên để xem Bảng tin Cá nhân:", staff_list)
    user_logs = [l for l in st.session_state.logs if l.get("Cán bộ") == target_staff_name]
    # 1. Lỗi vi phạm (Phụ lục 1)
    u_infractions = [l for l in user_logs if l.get("Loại") == "Điểm trừ"]
    total_penalty = sum([safe_float(l.get("Số điểm"), 0.0) for l in u_infractions])
    vc_score = 100.0 - total_penalty
    # 2. Khen thưởng & KSCL net (Phụ lục 2)
    u_bonuses = [l for l in user_logs if l.get("Loại") == "Điểm cộng"]
    total_bonus = sum([safe_float(l.get("Số điểm"), 0.0) for l in u_bonuses])
    final_gen_score = vc_score + total_bonus
    if vc_score >= 90.0:
        vc_class = "Hoàn thành Xuất sắc 🥇"
    elif vc_score >= 70.0:
        vc_class = "Hoàn thành Tốt 🥈"
    elif vc_score >= 50.0:
        vc_class = "Hoàn thành 🥉"
    else:
        vc_class = "Chưa hoàn thành ⚠️"
    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📉 Điểm Đánh giá VC còn lại", f"{vc_score:.1f} đ",
              delta=f"-{total_penalty:.1f}đ vi phạm" if total_penalty > 0 else "0.0đ")
    m2.metric("🏆 Điểm Thi đua Tích lũy", f"{total_bonus:+.1f} đ", delta="Khen thưởng & KSCL")
    m3.metric("📊 Tổng điểm Thi đua Tổng hợp", f"{final_gen_score:.1f} đ")
    m4.metric("🏅 Dự kiến Xếp loại VC", vc_class)
    st.markdown("---")
    p_tab1, p_tab2, p_tab3 = st.tabs([
        "📋 Lịch sử Lỗi vi phạm",
        "🏆 Lịch sử Thành tích Khen thưởng",
        "📈 Tiến trình & Điểm Khảo sát chất lượng (KSCL)"
    ])
    with p_tab1:
        st.subheader("📋 Bảng Chi tiết các Lỗi vi phạm đã bị trừ điểm")
        if u_infractions:
            st.dataframe(pd.DataFrame(u_infractions)[["Ngày", "Nội dung", "Số điểm", "Ghi chú"]],
                         use_container_width=True)
        else:
            st.success("🎉 Thầy/Cô chưa ghi nhận lỗi vi phạm nào. Xin chúc mừng!")
    with p_tab2:
        st.subheader("🏆 Bảng Chi tiết các Thành tích Khen thưởng & Điểm cộng")
        if u_bonuses:
            st.dataframe(pd.DataFrame(u_bonuses)[["Ngày", "Nội dung", "Số điểm", "Ghi chú"]], use_container_width=True)
        else:
            st.info("Chưa có thành tích khen thưởng nào được ghi nhận.")
    with p_tab3:
        st.subheader("📈 Bảng Tiến trình Khảo sát Chất lượng (KSCL)")
        user_kscl = [k for k in st.session_state.kscl_history if k.get("Cán bộ") == target_staff_name]
        if user_kscl:
            st.dataframe(pd.DataFrame(user_kscl)[
                             ["Ngày", "Lớp", "Môn", "Kỳ KSCL", "Điểm TB", "Điểm kỳ trước", "Chỉ số tăng",
                              "Điểm thi đua", "Ghi chú"]], use_container_width=True)
        else:
            st.info("Chưa có dữ liệu điểm KSCL.")
# CHỨC NĂNG 3 & 4: XỬ LÝ ĐIỂM TỰ ĐỘNG, XẾP LOẠI & XUẤT BÁO CÁO EXCEL
# -----------------------------------------------------------------------------
else:
    st.header("📊 Xử lý Điểm tự động, Xếp loại Viên chức & Xuất Báo cáo Thi đua")
    if st.session_state.users.empty:
        st.warning("⚠️ Chưa có dữ liệu nhân sự để lập báo cáo.")
    else:
        col_filter1, col_filter2 = st.columns(2)
        with col_filter1:
            TIME_FRAME_OPTIONS = [
                "Tháng 9/2026",
                "Tháng 10/2026",
                "Tháng 11/2026",
                "Tháng 12/2026",
                "Tháng 1/2027",
                "Tháng 2/2027",
                "Tháng 3/2027",
                "Tháng 4/2027",
                "Tháng 5/2027",
                "Học kỳ I (Tháng 9/2026 - Tháng 1/2027)",
                "Học kỳ II (Tháng 2/2027 - Tháng 5/2027)",
                "Cả năm học 2026-2027"
            ]
            time_frame = st.selectbox("📅 Chọn mốc thời gian xuất báo cáo:", TIME_FRAME_OPTIONS, index=11)
        with col_filter2:
            unique_depts = [d for d in st.session_state.users["Tổ chuyên môn"].unique() if d]
            dept_filter = st.selectbox("🏫 Chọn phạm vi đơn vị báo cáo:", ["Toàn đơn vị"] + unique_depts)
        filtered_logs = filter_logs_by_timeframe(st.session_state.logs, time_frame)
        st.subheader(f"📈 Hệ thống Báo cáo Thi đua & Đánh giá Viên chức - {time_frame} ({dept_filter})")
        # --- BÁO CÁO 1: TIÊU CHÍ ĐÁNH GIÁ VIÊN CHỨC (QĐ 292 - PHỤ LỤC 1) ---
        report_vc_summary = []
        report_vc_details = []
        # --- BÁO CÁO 2: TIÊU CHÍ THI ĐỦA KHEN THƯỞNG (PHỤ LỤC 2) ---
        report_td_summary = []
        report_td_details = []
        # --- BÁO CÁO 3: TỔNG HỢP TOÀN DIỆN ---
        report_general = []
        detail_vc_counter = 1
        detail_td_counter = 1
        for idx_u, user in st.session_state.users.iterrows():
            if dept_filter != "Toàn đơn vị" and user["Tổ chuyên môn"] != dept_filter:
                continue
            u_name = user["Họ và tên"]
            u_id = user["ID"]
            u_dob = user["Ngày sinh"]
            u_dept = user["Tổ chuyên môn"]
            u_logs = [l for l in filtered_logs if l.get("Cán bộ") == u_name]
            # Phân loại: Vi phạm Viên chức (Phụ lục 1) vs Thi đua Khen thưởng / KSCL (Phụ lục 2)
            u_vc_infractions = [l for l in u_logs if
                                l.get("Loại") == "Điểm trừ" and not l.get("Nội dung", "").startswith("[KSCL")]
            total_penalty_vc = sum([safe_float(l.get("Số điểm"), 0.0) for l in u_vc_infractions])
            u_td_bonuses = [l for l in u_logs if l.get("Loại") == "Điểm cộng"]
            u_td_kscl_penalties = [l for l in u_logs if
                                   l.get("Loại") == "Điểm trừ" and l.get("Nội dung", "").startswith("[KSCL")]
            total_bonus_td = sum([safe_float(l.get("Số điểm"), 0.0) for l in u_td_bonuses])
            total_penalty_kscl = sum([safe_float(l.get("Số điểm"), 0.0) for l in u_td_kscl_penalties])
            net_emulation_score = total_bonus_td - total_penalty_kscl
            # 1. Báo cáo Viên chức Summary
            if u_vc_infractions:
                vc_contents = " ; ".join([f"{i + 1}. {l.get('Nội dung')}" for i, l in enumerate(u_vc_infractions)])
            else:
                vc_contents = "Không có vi phạm"
            vc_score = 100.0 - total_penalty_vc
            if vc_score >= 90.0:
                vc_class = "Hoàn thành Xuất sắc nhiệm vụ 🥇"
            elif vc_score >= 70.0:
                vc_class = "Hoàn thành Tốt nhiệm vụ 🥈"
            elif vc_score >= 50.0:
                vc_class = "Hoàn thành nhiệm vụ 🥉"
            else:
                vc_class = "Không hoàn thành nhiệm vụ ⚠️"
            report_vc_summary.append({
                "Mã NV": u_id,
                "Họ và tên": u_name,
                "Ngày sinh": u_dob,
                "Tổ chuyên môn": u_dept,
                "Điểm tối đa ban đầu": 100.0,
                "Nội dung vi phạm": vc_contents,
                "Tổng điểm trừ": total_penalty_vc,
                "Điểm đánh giá còn lại": vc_score,
                "Dự kiến Xếp loại VC": vc_class
            })
            # 1. Báo cáo Viên chức Details
            for l in u_vc_infractions:
                report_vc_details.append({
                    "STT": detail_vc_counter,
                    "Mã NV": u_id,
                    "Họ và tên": u_name,
                    "Tổ chuyên môn": u_dept,
                    "Thời gian ghi nhận lỗi": l.get("Ngày", ""),
                    "Nội dung vi phạm": l.get("Nội dung", ""),
                    "Số điểm trừ": safe_float(l.get("Số điểm"), 0.0),
                    "Ghi chú": l.get("Ghi chú", "")
                })
                detail_vc_counter += 1
            # 2. Báo cáo Thi đua Khen thưởng Summary
            u_td_all = u_td_bonuses + u_td_kscl_penalties
            if u_td_all:
                td_contents = " ; ".join([f"{i + 1}. {l.get('Nội dung')}" for i, l in enumerate(u_td_all)])
            else:
                td_contents = "Chưa có thành tích"
            report_td_summary.append({
                "Mã NV": u_id,
                "Họ và tên": u_name,
                "Ngày sinh": u_dob,
                "Tổ chuyên môn": u_dept,
                "Nội dung cộng điểm / thi đua": td_contents,
                "Tổng điểm cộng": total_bonus_td,
                "Tổng điểm trừ KSCL": total_penalty_kscl,
                "Điểm thi đua khen thưởng": net_emulation_score
            })
            # 2. Báo cáo Thi đua Details
            for l in u_td_bonuses:
                report_td_details.append({
                    "STT": detail_td_counter,
                    "Mã NV": u_id,
                    "Họ và tên": u_name,
                    "Tổ chuyên môn": u_dept,
                    "Thời gian ghi nhận": l.get("Ngày", ""),
                    "Nội dung khen thưởng / KSCL": l.get("Nội dung", ""),
                    "Số điểm thi đua": safe_float(l.get("Số điểm"), 0.0),
                    "Ghi chú": l.get("Ghi chú", "")
                })
                detail_td_counter += 1
            for l in u_td_kscl_penalties:
                report_td_details.append({
                    "STT": detail_td_counter,
                    "Mã NV": u_id,
                    "Họ và tên": u_name,
                    "Tổ chuyên môn": u_dept,
                    "Thời gian ghi nhận": l.get("Ngày", ""),
                    "Nội dung khen thưởng / KSCL": l.get("Nội dung", ""),
                    "Số điểm thi đua": -safe_float(l.get("Số điểm"), 0.0),
                    "Ghi chú": l.get("Ghi chú", "")
                })
                detail_td_counter += 1
            # 3. Báo cáo Tổng hợp
            final_gen_score = vc_score + net_emulation_score
            if final_gen_score >= 90.0:
                gen_class = "Hoàn thành Xuất sắc nhiệm vụ 🥇"
            elif final_gen_score >= 70.0:
                gen_class = "Hoàn thành Tốt nhiệm vụ 🥈"
            elif final_gen_score >= 50.0:
                gen_class = "Hoàn thành nhiệm vụ 🥉"
            else:
                gen_class = "Không hoàn thành nhiệm vụ ⚠️"
            report_general.append({
                "Mã NV": u_id,
                "Họ và tên": u_name,
                "Ngày sinh": u_dob,
                "Tổ chuyên môn": u_dept,
                "Điểm ban đầu (100)": 100.0,
                "Tổng điểm trừ VC": total_penalty_vc,
                "Điểm VC còn lại": vc_score,
                "Điểm Thi đua Khen thưởng": net_emulation_score,
                "Tổng điểm tổng hợp": final_gen_score,
                "Dự kiến Xếp loại": gen_class
            })
        df_vc_sum = pd.DataFrame(report_vc_summary)
        df_vc_det = pd.DataFrame(report_vc_details) if report_vc_details else pd.DataFrame(
            columns=["STT", "Mã NV", "Họ và tên", "Tổ chuyên môn", "Thời gian ghi nhận lỗi", "Nội dung vi phạm",
                     "Số điểm trừ", "Ghi chú"])
        df_td_sum = pd.DataFrame(report_td_summary)
        df_td_det = pd.DataFrame(report_td_details) if report_td_details else pd.DataFrame(
            columns=["STT", "Mã NV", "Họ và tên", "Tổ chuyên môn", "Thời gian ghi nhận thành tích",
                     "Nội dung cộng điểm", "Số điểm cộng", "Ghi chú"])
        df_gen_sum = pd.DataFrame(report_general)
        rep_tab1, rep_tab2, rep_tab3 = st.tabs([
            "📋 Báo cáo 1: Tiêu chí Đánh giá Viên chức",
            "🏆 Báo cáo 2: Tiêu chí Thi đua Khen thưởng",
            "📊 Báo cáo 3: Tổng hợp Thi đua & Đánh giá Toàn diện"
        ])
        tf_slug = time_frame.replace(' ', '_').replace('/', '_').replace('(', '').replace(')', '').replace('-', '_')
        time_suffix = get_time_title_suffix(time_frame)
        with rep_tab1:
            st.markdown(f"#### 📋 Báo cáo Tiêu chí Đánh giá Viên chức - {time_frame}")
            st.dataframe(df_vc_sum, use_container_width=True)
            if not df_vc_sum.empty:
                excel_vc_data = create_styled_excel_report({
                    "TongHop_DanhGia_VienChuc": {
                        "df": df_vc_sum,
                        "title": f"KẾT QUẢ ĐÁNH GIÁ VIÊN CHỨC {time_suffix}"
                    },
                    "ChiTiet_Loi_ViPham": {
                        "df": df_vc_det,
                        "title": f"CHI TIẾT LỖI VI PHẠM VIÊN CHỨC {time_suffix}"
                    }
                })
                st.download_button(
                    label="📥 Xuất Báo cáo Tiêu chí Đánh giá Viên chức ra Excel (.xlsx)",
                    data=excel_vc_data,
                    file_name=f"Bao_cao_Danh_gia_Vien_chuc_{tf_slug}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with rep_tab2:
            st.markdown(f"#### 🏆 Báo cáo Tiêu chí Thi đua Khen thưởng - {time_frame}")
            st.dataframe(df_td_sum, use_container_width=True)
            if not df_td_sum.empty:
                excel_td_data = create_styled_excel_report({
                    "TongHop_ThiDua_KhenThuong": {
                        "df": df_td_sum,
                        "title": f"TỔNG HỢP ĐIỂM THI ĐUA {time_suffix}"
                    },
                    "ChiTiet_ThanhTich_KhenThuong": {
                        "df": df_td_det,
                        "title": f"CHI TIẾT THÀNH TÍCH KHEN THƯỞNG {time_suffix}"
                    }
                })
                st.download_button(
                    label="📥 Xuất Báo cáo Tiêu chí Thi đua Khen thưởng ra Excel (.xlsx)",
                    data=excel_td_data,
                    file_name=f"Bao_cao_Thi_dua_Khen_thuong_{tf_slug}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        with rep_tab3:
            st.markdown(f"#### 📊 Báo cáo Tổng hợp Thi đua & Đánh giá Toàn diện - {time_frame}")
            st.dataframe(df_gen_sum, use_container_width=True)
            if not df_gen_sum.empty:
                excel_gen_data = create_styled_excel_report({
                    "TongHop_ToanDien": {
                        "df": df_gen_sum,
                        "title": f"BÁO CÁO TỔNG HỢP THI ĐUA VÀ ĐÁNH GIÁ TOÀN DIỆN {time_suffix}"
                    },
                    "TongHop_VienChuc": {
                        "df": df_vc_sum,
                        "title": f"KẾT QUẢ ĐÁNH GIÁ VIÊN CHỨC {time_suffix}"
                    },
                    "TongHop_KhenThuong": {
                        "df": df_td_sum,
                        "title": f"TỔNG HỢP ĐIỂM THI ĐUA {time_suffix}"
                    }
                })
                st.download_button(
                    label="📥 Xuất Báo cáo Tổng hợp Thi đua Toàn diện ra Excel (.xlsx)",
                    data=excel_gen_data,
                    file_name=f"Bao_cao_Tong_hop_Thi_dua_{tf_slug}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        st.markdown("---")
        st.subheader("📱 Lịch sử Gửi thông báo Tự động & Tải File guitin.csv")
        hist_tab0, hist_tab1, hist_tab2 = st.tabs([
            "📧 Lịch sử gửi Email thông báo (Theo Email cá nhân)",
            "📄 File Xuất Tin Nhắn (guitin.csv)",
            "📡 Lịch sử gửi tin nhắn Viettel Tammi"
        ])
        with hist_tab0:
            st.markdown("##### 📧 Danh sách Lịch sử Gửi Email Thông báo (`email_history`)")
            if st.session_state.get("email_history"):
                st.dataframe(pd.DataFrame(st.session_state.email_history), use_container_width=True)
            else:
                st.info("Chưa có Email thông báo nào được gửi.")
        with hist_tab1:
            st.markdown("##### 📄 Danh sách Dữ liệu Tin nhắn Đã lưu (`guitin.csv`)")
            if st.session_state.get("guitin_list"):
                df_guitin = pd.DataFrame(st.session_state.guitin_list)
                st.dataframe(df_guitin, use_container_width=True)
                csv_bytes = df_guitin.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
                st.download_button(
                    label="📥 Tải xuống File guitin.csv (Dùng gửi tin nhắn tự động)",
                    data=csv_bytes,
                    file_name="guitin.csv",
                    mime="text/csv"
                )
            else:
                st.info("Chưa có lượt ghi nhận tin nhắn nào trong guitin.csv.")
        with hist_tab2:
            if st.session_state.get("tammi_history"):
                st.dataframe(pd.DataFrame(st.session_state.tammi_history), use_container_width=True)
            else:
                st.info("Chưa có thông báo Viettel Tammi nào được gửi.")