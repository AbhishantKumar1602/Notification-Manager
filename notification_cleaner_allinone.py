
import os
import sys
import sqlite3
from pathlib import Path
from datetime import datetime, timedelta, timezone
import shutil
import json
import time
import tkinter as tk
import subprocess
from tkinter import messagebox, simpledialog
from bs4 import BeautifulSoup

# ================= CONFIG =================
# Dynamic paths
WPN_DB_PATH = os.path.join( 
    os.environ["LOCALAPPDATA"],
    "Microsoft",
    "Windows",
    "Notifications",
    "wpndatabase.db"
)
TEMP_DB = os.path.join(os.environ.get("TEMP", r"C:\\temp"), "wpn_copy.db")
OUTPUT_JSON = "1_notification_classified.json"

# Handle PyInstaller bundle path for dismiss_notif.exe logic
if getattr(sys, 'frozen', False):
    # If run as an EXE, expect the helper to be in the same temp folder
    DISMISS_EXE = os.path.join(sys._MEIPASS, "dismiss_notif.exe")
else:
    # If run as a script, look in the build folder
    base_path = os.path.dirname(os.path.abspath(__file__))
    DISMISS_EXE = os.path.join(base_path, "dismiss_notif", "bin", "Release", "net8.0-windows10.0.19041.0", "win-x64", "publish", "dismiss_notif.exe")

DELAY_BETWEEN_DISMISS = 0.1  # seconds
# =========================================

def load_target_titles():
    """Reads target titles from a text file in the same directory."""
    titles = []
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    file_path = os.path.join(base_path, "target_titles.txt")
    
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            titles = [line.strip() for line in f if line.strip()]
    return titles

def load_target_apps():
    """Reads target app names from a text file in the same directory."""
    apps = []
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    file_path = os.path.join(base_path, "target_apps.txt")
    
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            apps = [line.strip() for line in f if line.strip()]
    return apps

def filetime_to_datetime(ft):
    if not ft:
        return None
    return datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ft // 10)

def decode_payload(blob):
    if not blob:
        return None
    for enc in ("utf-16le", "utf-8"):
        try:
            decoded = blob.decode(enc)
            if "<toast" in decoded:
                return decoded
        except Exception:
            pass
    return None

def parse_payload(xml):
    if not xml:
        return "", "", "N/A"
    soup = BeautifulSoup(xml, "xml")
    texts = soup.find_all("text")
    title = texts[0].text.strip() if len(texts) > 0 else ""
    subtitle = texts[1].text.strip() if len(texts) > 1 else ""
    attribution = soup.find("text", {"placement": "attribution"})
    attribution = attribution.text.strip() if attribution else "N/A"
    return title, subtitle, attribution

def copy_wpn_db():
    Path(os.path.dirname(TEMP_DB)).mkdir(parents=True, exist_ok=True)
    shutil.copy2(WPN_DB_PATH, TEMP_DB)
    wal = WPN_DB_PATH + "-wal"
    if os.path.exists(wal):
        shutil.copy2(wal, TEMP_DB + "-wal")

def export_and_classify_notifications():
    copy_wpn_db()
    conn = sqlite3.connect(TEMP_DB)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT Id, Payload, ArrivalTime, ExpiryTime "
        "FROM Notification WHERE Type='toast' ORDER BY ArrivalTime DESC"
    )
    now = datetime.now(timezone.utc)
    wanted = []
    unwanted = []
    
    target_titles = load_target_titles()
    target_apps = load_target_apps()
    
    for notif_id, payload, arrival_ft, expiry_ft in cursor.fetchall():
        arrival = filetime_to_datetime(arrival_ft)
        expiry = filetime_to_datetime(expiry_ft)
        if not arrival or not expiry:
            continue
        if not (arrival <= now <= expiry):
            continue
        payload_xml = decode_payload(payload)
        title, subtitle, attribution = parse_payload(payload_xml)
        entry = {
            "Id": notif_id,
            "Title": title,
            "Subtitle": subtitle,
            "Attribution": attribution,
            "ArrivalTime": arrival.isoformat(),
            "ExpiryTime": expiry.isoformat()
        }
        
        is_wanted = False
        if title in target_titles:
            is_wanted = True
        else:
            for app in target_apps:
                if (app.lower() in title.lower() or 
                    app.lower() in attribution.lower() or 
                    app.lower() in subtitle.lower() or 
                    (payload_xml and app.lower() in payload_xml.lower())):
                    is_wanted = True
                    break
        
        if is_wanted:
            wanted.append(entry)
        else:
            unwanted.append(entry)
    final_json = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "wanted": wanted,
        "unwanted": unwanted
    }
    Path(os.path.dirname(OUTPUT_JSON)).mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(final_json, f, indent=4, ensure_ascii=False)
    conn.close()
    print("✅ Notification classification complete")
    print(f"   Wanted   : {len(wanted)}")
    print(f"   Unwanted : {len(unwanted)}")
    print(f"📄 Output → {OUTPUT_JSON}")

def load_classified_notifications():
    if not os.path.exists(OUTPUT_JSON):
        raise FileNotFoundError(f"JSON not found: {OUTPUT_JSON}")
    with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
        return json.load(f)

def dismiss_unwanted(unwanted_list):
    if not unwanted_list:
        print("ℹ️ No unwanted notifications to dismiss")
        return
    print(f"🧹 Dismissing {len(unwanted_list)} unwanted notifications...\n")
    for notif in unwanted_list:
        title = notif.get("Title", "").strip()
        notif_id = notif.get("Id")
        if not title:
            continue
        print(f"❌ Dismissing: {title}  (ID: {notif_id})")
        try:
            result = subprocess.run(
                [DISMISS_EXE, title],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode != 0:
                print(f"⚠️ Failed to dismiss '{title}': {result.stderr.strip()}")
        except Exception as e:
            print(f"⚠️ Exception while dismissing '{title}': {e}")
        time.sleep(DELAY_BETWEEN_DISMISS)

def main():
    # Ensure config files exist
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
        
    for filename in ["target_titles.txt", "target_apps.txt"]:
        fpath = os.path.join(base_path, filename)
        if not os.path.exists(fpath):
            with open(fpath, "w", encoding="utf-8") as f:
                f.write("") # Create empty file

    # Initialize UI (Hidden root window)
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True) # Make sure popups appear on top

    # 1. Configuration Verification Popup
    msg = (
        "Configuration Check:\n\n"
        "1. Have you added your wanted App Names to 'target_apps.txt'?\n"
        "2. Have you added your wanted Titles to 'target_titles.txt'?\n\n"
        "Click Yes to start cleaning. Click No to exit and edit files."
    )
    if not messagebox.askyesno("Notification Cleaner", msg):
        messagebox.showinfo("Exiting", "Please update the text files in the folder and restart the app.")
        root.destroy()
        return

    # 2. Interval Input Popup
    run_interval = simpledialog.askinteger(
        "Settings", 
        "Enter run interval in seconds:", 
        initialvalue=15, 
        minvalue=1,
        parent=root
    )
    if run_interval is None:
        run_interval = 15 # Default if cancelled
    
    root.destroy() # Cleanup UI

    print(f"🟢 Notification Auto Controller started (Interval: {run_interval}s)")

    while True:
        print("\n▶ Running: Extract_Notification_From_Database")
        export_and_classify_notifications()
        print("\n▶ Running: Remove_Unwanted_Notifications")
        data = load_classified_notifications()
        wanted = data.get("wanted", [])
        unwanted = data.get("unwanted", [])
        print(f"✅ Wanted notifications   : {len(wanted)} (left untouched)")
        print(f"❌ Unwanted notifications : {len(unwanted)}\n")
        dismiss_unwanted(unwanted)
        print("\n🎯 Done. Wanted notifications preserved.")
        print(f"\n⏳ Waiting {run_interval} seconds...\n")
        time.sleep(run_interval)

if __name__ == "__main__":
    main()
