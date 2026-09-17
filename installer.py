from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
import sys
from pathlib import Path


APP_NAME="三会文件生成系统"
PUBLISHER="北京雷石天地电子技术股份有限公司"


def message(text,title=APP_NAME,flags=0x40):
    return ctypes.windll.user32.MessageBoxW(None,text,title,flags)


def make_shortcut(shortcut,target,working_dir,icon):
    shortcut.parent.mkdir(parents=True,exist_ok=True)
    ps=("$w=New-Object -ComObject WScript.Shell;"
        f"$s=$w.CreateShortcut('{str(shortcut).replace("'","''")}');"
        f"$s.TargetPath='{str(target).replace("'","''")}';"
        f"$s.WorkingDirectory='{str(working_dir).replace("'","''")}';"
        f"$s.IconLocation='{str(icon).replace("'","''")},0';$s.Save()")
    subprocess.run(["powershell.exe","-NoProfile","-ExecutionPolicy","Bypass","-Command",ps],check=True,creationflags=0x08000000)


def main():
    silent="--silent" in sys.argv
    payload=Path(getattr(sys,"_MEIPASS",Path(__file__).parent))/"payload"/"ThreeMeetingsApp.exe"
    if not payload.is_file():
        if not silent: message("安装包不完整，未找到主程序。",flags=0x10)
        return 1
    local=Path(os.environ.get("LOCALAPPDATA",Path.home()/"AppData"/"Local"))
    roaming=Path(os.environ.get("APPDATA",Path.home()/"AppData"/"Roaming"))
    desktop=Path(os.environ.get("USERPROFILE",Path.home()))/"Desktop"
    app_dir=local/"雷石股份"/APP_NAME
    data_dir=local/"雷石股份"/f"{APP_NAME}数据"
    target=app_dir/f"{APP_NAME}.exe"
    try:
        app_dir.mkdir(parents=True,exist_ok=True);data_dir.mkdir(parents=True,exist_ok=True)
        shutil.copy2(payload,target)
        make_shortcut(desktop/f"{APP_NAME}.lnk",target,app_dir,target)
        make_shortcut(roaming/"Microsoft"/"Windows"/"Start Menu"/"Programs"/"雷石股份"/f"{APP_NAME}.lnk",target,app_dir,target)
    except Exception as exc:
        if not silent: message(f"安装失败：{exc}",flags=0x10)
        return 1
    if not silent: message(f"安装完成。\n\n桌面和开始菜单已创建“{APP_NAME}”快捷方式。\n会议资料将保存在：\n{data_dir}")
    subprocess.Popen([str(target)],cwd=str(app_dir),creationflags=0x08000000)
    return 0


if __name__=="__main__":
    raise SystemExit(main())
