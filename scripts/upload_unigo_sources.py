import subprocess
import sys
from pathlib import Path

NOTEBOOK_ID = "f6c754f2-0291-40f3-bdc8-8b8cd37ef396"

SOURCES_TO_UPLOAD = [
    r"D:\UNIGO\Phân phối chương trình\Robotics\KHUNG CHƯƠNG TRÌNH ROBOTICS TIỂU HỌC & THCS UNIGO.pdf",
    r"D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\3456-VV_huong_dan_trien_khai_Khung_nang_luc_so_cho_HS_885ca.pdf",
    r"D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\cong-van-5512-bgddt-2020_d8bd32d0a4.pdf",
    r"D:\UNIGO\Hệ thống mẫu văn bản\Công_văn_quy_định\3439_QD_BGDDT_signed_c93bf_c86ef.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\2. Chuẩn đầu ra ICT - Tiểu học.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\3. Khung chương trình + PPCT ICT - Tiểu học\Khung chương trình ICT - Lớp 1.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\3. Khung chương trình + PPCT ICT - Tiểu học\Khung chương trình ICT - Lớp 2.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\3. Khung chương trình + PPCT ICT - Tiểu học\Khung chương trình ICT - Lớp 3.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\3. Khung chương trình + PPCT ICT - Tiểu học\Khung chương trình ICT - Lớp 4.pdf",
    r"D:\UNIGO\Phân phối chương trình\Môn ICT\3. Khung chương trình + PPCT ICT - Tiểu học\Khung chương trình ICT - Lớp 5.pdf",
]

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    print(f"Bắt đầu tải {len(SOURCES_TO_UPLOAD)} tài liệu vào Notebook: {NOTEBOOK_ID}")
    
    for i, file_path in enumerate(SOURCES_TO_UPLOAD, start=1):
        p = Path(file_path)
        if not p.exists():
            print(f"[{i}/{len(SOURCES_TO_UPLOAD)}] ⚠️ Không tìm thấy file: {file_path}")
            continue
            
        print(f"\n[{i}/{len(SOURCES_TO_UPLOAD)}] Đang tải: {p.name} ({p.stat().st_size / 1024 / 1024:.2f} MB)...")
        cmd = ["nlm", "source", "add", NOTEBOOK_ID, "--file", str(p), "--wait", "-p", "tk2"]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
        if res.returncode == 0:
            print(f"  ✓ Thành công: {res.stdout.strip()}")
        else:
            print(f"  ✗ Thất bại: {res.stderr.strip() or res.stdout.strip()}")

    print("\nHoàn tất nạp tài liệu!")

if __name__ == "__main__":
    main()
