import os
import glob
import re
import xlrd
import openpyxl

def normalize_name(name):
    if not name:
        return ""
    # remove extra spaces and lowercase
    return re.sub(r'\s+', ' ', str(name).strip()).lower()

# Mapping between classes and source NhanXet files
class_sources = {
    '1a1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (1).xls', 'tin': None},
    '1c1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (4).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (2).xls'},
    '2a1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (5).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (3).xls'},
    '2c1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (7).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (6).xls'},
    '3a1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (9).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (8).xls'},
    '3c1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (11).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (10).xls'},
    '4c1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (13).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (12).xls'},
    '5c1': {'rob': 'NhanXetMonHocCuaGiaoVien.xls (14).xls', 'tin': 'NhanXetMonHocCuaGiaoVien.xls (15).xls'},
}

base_dir = r'D:\UNIGO\Kiemtra_danhgia\Thi_GK1\TH'

def load_source_comments(filename):
    if not filename:
        return {}
    path = os.path.join(base_dir, filename)
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    hdr_row = 5
    res = {}
    for r in range(hdr_row + 1, sh.nrows):
        ho = str(sh.cell_value(r, 2)).strip()
        ten = str(sh.cell_value(r, 3)).strip()
        full_name = f"{ho} {ten}".strip()
        # Col 7 if 8 cols, col 8 if 9 cols
        c_nx = 7 if sh.ncols == 8 else 8
        nx = str(sh.cell_value(r, c_nx)).strip()
        res[normalize_name(full_name)] = nx
    return res

def process_all_classes():
    summary_results = []
    
    for cname, srcs in sorted(class_sources.items()):
        dest_filename = f'filemau_danhgia_dinhkyc1_tt27_2020_{cname}_GK1.xlsx'
        dest_path = os.path.join(base_dir, dest_filename)
        
        if not os.path.exists(dest_path):
            print(f"Error: Not found {dest_path}")
            continue
            
        rob_comments = load_source_comments(srcs['rob'])
        tin_comments = load_source_comments(srcs['tin'])
        
        wb = openpyxl.load_workbook(dest_path)
        ws = wb['TH-CN (Tin học)']
        
        count_updated = 0
        students_info = []
        
        for r in range(2, ws.max_row + 1):
            stt = ws.cell(r, 1).value
            sid = ws.cell(r, 2).value
            name = ws.cell(r, 3).value
            dob = ws.cell(r, 4).value
            
            if not name:
                continue
                
            norm_name = normalize_name(name)
            t_nx = tin_comments.get(norm_name, '')
            r_nx = rob_comments.get(norm_name, '')
            
            # Format combined comment
            if t_nx and r_nx:
                combined_nx = f"Tin học: {t_nx} Robotic: {r_nx}"
            elif t_nx:
                combined_nx = f"Tin học: {t_nx}"
            elif r_nx:
                combined_nx = f"Robotic: {r_nx}"
            else:
                combined_nx = ""
                print(f"  [WARNING] No comment found for student: {name} in class {cname.upper()}")
            
            # Column 5: Mức đạt được -> để trống theo yêu cầu người dùng
            ws.cell(r, 5).value = None
            
            # Column 6: Điểm KTĐK -> để trống
            ws.cell(r, 6).value = None
            
            # Column 7: Mã nhận xét -> để trống
            ws.cell(r, 7).value = None
            
            # Column 8: Nội dung nhận xét
            ws.cell(r, 8).value = combined_nx
            
            # Column 9: Thời điểm đánh giá -> 'GK1'
            ws.cell(r, 9).value = 'GK1'
            
            count_updated += 1
            students_info.append({
                'stt': stt,
                'name': name,
                'len_nx': len(combined_nx),
                'sample': combined_nx[:60] + '...' if len(combined_nx) > 60 else combined_nx
            })
            
        wb.save(dest_path)
        summary_results.append({
            'class': cname.upper(),
            'file': dest_filename,
            'count': count_updated,
            'students': students_info
        })
        print(f"Successfully updated {count_updated} students in {dest_filename}")

    return summary_results

if __name__ == '__main__':
    results = process_all_classes()
    print("\nAll classes processed successfully.")
