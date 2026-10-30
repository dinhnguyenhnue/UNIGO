import sys
from pathlib import Path
from notebooklm_tools.cli.main import cli_main

robotics_id = "34985e71-e053-47ca-a169-8125e774f8d0"
cong_nghe_id = "1ab9d9b3-8942-4df0-b838-b9dd142e4e04"

tasks = [
    # Robotics
    (robotics_id, r"D:\UNIGO\SGK\Robotics\Giáo trình - OLLO Initiate.pdf"),
    (robotics_id, r"D:\UNIGO\SGK\Robotics\Giáo trình - OLLO Kinder.pdf"),
    # Cong nghe
    (cong_nghe_id, r"D:\UNIGO\KHBD_Cong_nghe\SGK\SGK Công nghệ 4 KNTT.pdf"),
    (cong_nghe_id, r"D:\UNIGO\KHBD_Cong_nghe\SGK\SGK công nghệ 5 KNTT.pdf"),
    (cong_nghe_id, r"D:\UNIGO\KHBD_Cong_nghe\SGK\SGK Công nghệ 6 KNTT.pdf"),
    (cong_nghe_id, r"D:\UNIGO\KHBD_Cong_nghe\SGK\SGK Công nghệ 7 KNTT.pdf"),
    (cong_nghe_id, r"D:\UNIGO\KHBD_Cong_nghe\SGK\SGK CÔNG NGHỆ 8 KNTT.pdf"),
]

for nb_id, file_path in tasks:
    p = Path(file_path)
    if not p.exists():
        print(f"[ERROR] File not found: {file_path}")
        continue
    print(f"\n==========================================")
    print(f"Uploading: {p.name} ({p.stat().st_size / (1024*1024):.1f} MB) -> {nb_id}")
    print(f"==========================================")
    sys.argv = ['nlm', 'source', 'add', nb_id, '--file', str(p), '--wait']
    try:
        cli_main()
    except SystemExit as e:
        if e.code != 0:
            print(f"CLI exited with code {e.code}")
    except Exception as e:
        print(f"Error uploading {p.name}: {e}")

print("\n>>> ALL UPLOADS COMPLETED! <<<")
