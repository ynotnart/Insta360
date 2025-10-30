
# AutoInsta_v1_3_EXE
# This project aims to automate the insta360 conversion process from insta raw files to mp4 and gpx, and apply it to batch raw files. With the goal 
# of running on other windows machines to automate conversion processes.
# 
# Created: 7/23/25
# Source: ChatGPT
#
# python -m PyInstaller --onefile --windowed "C:\Users\ynotn\OneDrive\Desktop\Desktop_HP\desktop\Projects\Visual_Studio\Code_Bank\360_Drive\AutoInsta_v1_3_EXE.py"
# Look for file in "dist"
# 
# 7/23/25: modified color plus to 100 and bitrate to 49. Halted development of AutoInsta_v1_2.py in further this branch, 'EXE'.
# 7/17/25: sometimes last insta file is not moved into converted folder. This is the patch! Bitrate update to "50". Actually is '49'
# 7/17/25: updated bitrate pointer coords.
# 

import os
import time
import shutil
import subprocess
import pyautogui
import pygetwindow as gw
import keyboard
import threading
import sys
import tkinter as tk
from tkinter import filedialog

pyautogui.FAILSAFE = True

# === EXPORT OPTION COORDINATES (UPDATE THESE BASED ON YOUR SCREEN) ===
EXPORT_COORDS = {
    "StabilizationType": (1498, 155),
    "DirectionLock": (1587, 286),
    "Lens guard": (1494, 250),
    "Off": (1587, 383),
    "Color opt": (1495, 334),
    "ColorPlusbutton": (1861, 235),
    "ColorPlusslider": (1831, 282),
    "Audio": (1495, 433),
    "Noise_Red": (1589, 284),
    "Export": (1712, 35),
    "360": (1030, 212),
    "bitrate": (930, 500),
    "GPX": (882, 765),
    "Start_Export": (1165, 955)
}

def choose_folder(prompt):
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title=prompt)
    return folder

# Set these dynamically later
INPUT_FOLDER = OUTPUT_FOLDER = MP4_OUTPUT_FOLDER = None

def esc_listener():
    keyboard.wait('esc')
    print("\n🚨 ESC key pressed. Exiting gracefully...")
    close_insta360()
    sys.exit(0)

def apply_main_settings():
    print("🎛️ Applying main video settings (outside export)...")
    for key in ["StabilizationType", "DirectionLock", "Lens guard", "Off",
                "Color opt", "ColorPlusbutton", "ColorPlusslider",
                "Audio", "Noise_Red"]:
        if key in EXPORT_COORDS:
            click_if_not_yellow(EXPORT_COORDS[key])

def open_file_with_insta360(file_path):
    print(f"Opening file with default associated app: {file_path}")
    os.startfile(file_path)
    time.sleep(30)
    try:
        for w in gw.getWindowsWithTitle("Insta360"):
            if not w.isMaximized:
                print("🖥️ Maximizing Insta360 Studio window...")
                w.maximize()
                break
    except Exception as e:
        print(f"⚠️ Could not maximize window: {e}")

def close_insta360():
    print("❌ Closing Insta360 Studio...")
    try:
        os.system('taskkill /IM "Insta360 Studio.exe" /F')
    except Exception as e:
        print(f"⚠️ Could not kill Insta360 Studio process: {e}")
    time.sleep(2)

def export_file():
    print("🎞️ Opening Export dialog...")
    pyautogui.hotkey('ctrl', 'e')
    time.sleep(2)
    print("🛠️ Setting export options inside dialog...")
    for key in ["Export", "360", "bitrate", "GPX", "Start_Export"]:
        if key not in EXPORT_COORDS:
            continue
        coord = EXPORT_COORDS[key]
        if key == "Start_Export":
            pyautogui.click(coord)
            time.sleep(0.5)
        else:
            click_if_not_yellow(coord)
    print("✅ Starting export...")
    pyautogui.press('enter')
    time.sleep(5)

def is_file_stable(file_path, wait_time=3):
    try:
        size1 = os.path.getsize(file_path)
        time.sleep(wait_time)
        size2 = os.path.getsize(file_path)
        return size1 == size2
    except FileNotFoundError:
        return False

def wait_for_export(mp4_name, timeout=1800):
    print(f"⏳ Waiting for export of {mp4_name}...")
    full_path = os.path.join(MP4_OUTPUT_FOLDER, mp4_name)
    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(full_path) and is_file_stable(full_path):
            print("✅ MP4 export complete and stable.")
            return True
        time.sleep(2)
    print("❌ Timeout waiting for export.")
    return False

def process_insta_file(insta_file):
    file_path = os.path.join(INPUT_FOLDER, insta_file)
    mp4_file = insta_file.replace(".insv", ".mp4")
    mp4_path = os.path.join(MP4_OUTPUT_FOLDER, mp4_file)
    output_path = os.path.join(OUTPUT_FOLDER, insta_file)

    open_file_with_insta360(file_path)
    apply_main_settings()
    export_file()

    if not wait_for_export(mp4_file):
        print(f"⚠️ Skipping {insta_file} due to timeout.")
        close_insta360()
        return

    close_insta360()
    print("🕒 Waiting for file locks to release...")
    time.sleep(5)  # Give Windows/Explorer time to release file handle

    print(f"📁 Moving {insta_file} to converted folder...")
    for attempt in range(15):
        try:
            if not os.path.exists(file_path):
                print("✅ File already moved.")
                return
            shutil.move(file_path, output_path)
            print(f"✅ Successfully moved {insta_file}.")
            return
        except PermissionError:
            print(f"⏳ File in use. Retrying move ({attempt + 1}/15)...")
            time.sleep(2)
        except Exception as e:
            print(f"⚠️ Unexpected error during move: {e}")
            time.sleep(2)

    print(f"❌ Failed to move {insta_file} after 15 attempts.")


def is_yellow(pixel):
    r, g, b = pixel
    return r > 200 and g > 200 and b < 100

def click_if_not_yellow(coord):
    pixel = pyautogui.pixel(coord[0], coord[1])
    if not is_yellow(pixel):
        pyautogui.click(coord)
        time.sleep(0.5)
    else:
        print(f"Skipping click at {coord}, already yellow.")

def main():
    global INPUT_FOLDER, OUTPUT_FOLDER, MP4_OUTPUT_FOLDER

    print("📂 Please select required folders:")
    INPUT_FOLDER = choose_folder("Select INPUT folder containing .insv files")
    OUTPUT_FOLDER = choose_folder("Select OUTPUT folder for converted .insv files")
    MP4_OUTPUT_FOLDER = choose_folder("Select folder where exported MP4s will appear")

    if not INPUT_FOLDER or not OUTPUT_FOLDER or not MP4_OUTPUT_FOLDER:
        print("❌ One or more folders not selected. Exiting.")
        return

    print(f"\nINPUT: {INPUT_FOLDER}\nOUTPUT: {OUTPUT_FOLDER}\nMP4 OUT: {MP4_OUTPUT_FOLDER}\n")

    insv_files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".insv")]
    if not insv_files:
        print("❗ No Insta360 files found.")
        return

    for f in insv_files:
        process_insta_file(f)
        print("⏳ Waiting 5 seconds before next file...\n")
        time.sleep(5)

    print("🎉 All files processed.")

if __name__ == "__main__":
    esc_thread = threading.Thread(target=esc_listener, daemon=True)
    esc_thread.start()
    print("🚀 Press ESC at any time to cancel the script.\n")

    try:
        main()
    except pyautogui.FailSafeException:
        print("🚨 Script stopped by user (mouse moved to corner).")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

    input("📥 Press Enter to exit...")
