# -*- coding: utf-8 -*-
"""Cài font SVN-Gilroy (bundled trong skill) theo per-user — KHÔNG cần quyền admin.

Chạy:  python3 install_fonts.py
- macOS   : copy vào ~/Library/Fonts            (app tự nhận, mở lại app đang chạy)
- Linux   : copy vào ~/.local/share/fonts rồi fc-cache
- Windows : copy vào %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts + đăng ký registry HKCU
            (cần: pip install fonttools)
"""
import os, sys, glob, shutil, platform, subprocess
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

_HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.normpath(os.path.join(_HERE, "..", "assets", "fonts"))


def _font_files():
    files = []
    for pat in ("*.otf", "*.ttf", "*.OTF", "*.TTF"):
        files += glob.glob(os.path.join(FONT_DIR, pat))
    return sorted(set(files))


def _copy_all(dest_dir, files):
    os.makedirs(dest_dir, exist_ok=True)
    n_new = n_have = 0
    for f in files:
        dest = os.path.join(dest_dir, os.path.basename(f))
        if os.path.exists(dest):
            n_have += 1
        else:
            shutil.copyfile(f, dest); n_new += 1
    return n_new, n_have


def install_mac(files):
    dest = os.path.expanduser("~/Library/Fonts")
    n_new, n_have = _copy_all(dest, files)
    print(f"macOS: đã chép {n_new} font mới, {n_have} đã có sẵn → {dest}")
    print("Mở lại app (PowerPoint/Keynote/LibreOffice) để nhận font 'SVN-Gilroy'.")


def install_linux(files):
    dest = os.path.expanduser("~/.local/share/fonts/SVN-Gilroy")
    n_new, n_have = _copy_all(dest, files)
    print(f"Linux: đã chép {n_new} font mới, {n_have} đã có sẵn → {dest}")
    try:
        subprocess.run(["fc-cache", "-f"], check=False)
    except FileNotFoundError:
        print("(không có fc-cache, bỏ qua cập nhật cache)")


def install_windows(files):
    import winreg
    from fontTools.ttLib import TTFont
    dest_dir = os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Windows', 'Fonts')
    os.makedirs(dest_dir, exist_ok=True)
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
    n_ok = 0
    for f in files:
        try:
            nm = TTFont(f)['name']
            full = (nm.getDebugName(4) or (nm.getDebugName(1) + " " + (nm.getDebugName(2) or ""))).strip()
            ext = os.path.splitext(f)[1].lower()
            suffix = " (OpenType)" if ext == '.otf' else " (TrueType)"
            dest = os.path.join(dest_dir, full.replace(' ', '') + ext)
            if os.path.exists(dest):
                status = "đã có sẵn"
            else:
                shutil.copyfile(f, dest); status = "đã cài"
            winreg.SetValueEx(key, full + suffix, 0, winreg.REG_SZ, dest)
            print(f"{status}: {full}{suffix}"); n_ok += 1
        except Exception as e:
            print("Lỗi:", os.path.basename(f), "->", e)
    winreg.CloseKey(key)
    print(f"\nHoàn tất: {n_ok}/{len(files)} font. Mở lại PowerPoint để dùng 'SVN-Gilroy'.")


def main():
    files = _font_files()
    if not files:
        print("Không có file font trong", FONT_DIR)
        print("Bộ SVN-Gilroy (16 file .otf/.ttf) có BẢN QUYỀN nên bản công khai không kèm. Cách lấy:")
        print("  1. Xin bộ font nội bộ từ phòng Marketing (Drive/Zalo), chép 16 file vào thư mục trên, chạy lại script này.")
        print("  2. Hoặc máy đã cài SVN-Gilroy sẵn thì không cần làm gì — module dựng slide gọi font theo tên.")
        print("  3. Không có font: PowerPoint sẽ hiện font dự phòng; đổi FONT trong brand_pptx.py sang Be Vietnam Pro / Montserrat.")
        return
    sysname = platform.system()
    if sysname == "Darwin":
        install_mac(files)
    elif sysname == "Windows":
        install_windows(files)
    else:
        install_linux(files)


if __name__ == "__main__":
    main()
