"""Play a notification sound (and show a Windows toast with a short message).

    python scripts/notify.py ["message"] [--sound alarm|ok|chime] [--no-toast]

The sound comes first from the Windows system sounds (asterisk / exclamation / a
C: scale chime), so nothing has to be installed. The toast uses the Windows
notification centre through PowerShell; if that is unavailable the sound still plays.
"""
from __future__ import annotations

import argparse
import subprocess
import sys


def play(sound: str) -> None:
    import winsound

    if sound == "chime":
        for freq in (660, 880, 1100):
            winsound.Beep(freq, 160)
    elif sound == "alarm":
        for _ in range(3):
            winsound.MessageBeep(winsound.MB_ICONHAND)
            winsound.Beep(1000, 250)
    else:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)
        winsound.Beep(880, 200)
        winsound.Beep(1175, 300)


def toast(message: str) -> bool:
    ps = r'''
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$template = @"
<toast><visual><binding template="ToastGeneric"><text>Claude Code</text><text>MESSAGE</text></binding></visual><audio src="ms-winsoundevent:Notification.Default"/></toast>
"@
$template = $template.Replace("MESSAGE", [System.Security.SecurityElement]::Escape($env:NOTIFY_MESSAGE))
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml($template)
$toast = New-Object Windows.UI.Notifications.ToastNotification $xml
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Claude Code").Show($toast)
'''
    try:
        import os

        env = dict(os.environ, NOTIFY_MESSAGE=message)
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], env=env, capture_output=True, text=True, timeout=20)
        return r.returncode == 0
    except Exception:
        return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("message", nargs="?", default="Done.")
    ap.add_argument("--sound", choices=["ok", "alarm", "chime"], default="ok")
    ap.add_argument("--no-toast", action="store_true")
    args = ap.parse_args()
    play(args.sound)
    shown = False if args.no_toast else toast(args.message)
    print(f"sound played ({args.sound}); toast {'shown' if shown else 'not shown'}: {args.message}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
