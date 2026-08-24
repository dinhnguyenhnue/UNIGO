import shutil, os, sys
sys.stdout.reconfigure(encoding='utf-8')

ART_DIR = r"C:\Users\bmngu\.gemini\antigravity-ide\brain\aa8b9142-1b13-4668-be83-50a1e29bc7f9"
IMG5_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images"
IMG7_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images"

files_map = {
    os.path.join(ART_DIR, "lop5_bai3_travel_situations_1787588464189.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai3_travel.jpg"),
    os.path.join(ART_DIR, "lop5_bai3_search_magic_1787588509257.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai3_search.jpg"),
    os.path.join(ART_DIR, "lop5_bai4_folder_tree_1787588487450.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai4_tree.jpg"),
    os.path.join(ART_DIR, "lop7_bai3_file_data_hub_1787588553175.jpg"): os.path.join(IMG7_DIR, "ai_lop7_bai3_data_hub.jpg"),
    os.path.join(ART_DIR, "lop7_bai4_social_networks_1787588576615.jpg"): os.path.join(IMG7_DIR, "ai_lop7_bai4_social_networks.jpg"),
}

for src, dst in files_map.items():
    if os.path.exists(src):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        print(f"Copied {os.path.basename(src)} -> {dst}")
    else:
        print(f"Source not found: {src}")
