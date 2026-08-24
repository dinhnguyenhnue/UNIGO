import shutil, os, sys
sys.stdout.reconfigure(encoding='utf-8')

ART_DIR = r"C:\Users\bmngu\.gemini\antigravity-ide\brain\aa8b9142-1b13-4668-be83-50a1e29bc7f9"
IMG5_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_5\Tuần_04\images"
IMG7_DIR = r"D:\UNIGO\KHBD_Tin_học\Lớp_7\Tuần_04\images"

files_map = {
    os.path.join(ART_DIR, "ai_lop5_bai4_4steps_flow_1787589343562.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai4_4steps_flow.jpg"),
    os.path.join(ART_DIR, "ai_lop5_bai3_search_flow_1787589371749.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai3_search_flow.jpg"),
    os.path.join(ART_DIR, "ai_lop5_bai4_hoctap_folders_1787589400284.jpg"): os.path.join(IMG5_DIR, "ai_lop5_bai4_hoctap_folders.jpg"),
    os.path.join(ART_DIR, "ai_lop7_bai3_shortcuts_1787589423570.jpg"): os.path.join(IMG7_DIR, "ai_lop7_bai3_shortcuts.jpg"),
    os.path.join(ART_DIR, "ai_lop7_bai4_5k_safety_1787589447074.jpg"): os.path.join(IMG7_DIR, "ai_lop7_bai4_5k_safety.jpg"),
}

for src, dst in files_map.items():
    if os.path.exists(src):
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copy2(src, dst)
        print(f"Copied {os.path.basename(src)} -> {dst}")
    else:
        print(f"Source not found: {src}")
