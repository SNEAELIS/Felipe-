"""
monitor.py

Watch a folder for newly created files and show a Windows notification.

Usage:
  python monitor.py "C:\path\to\watch"
  # or run with pythonw to avoid console window:
  pythonw monitor.py "C:\path\to\watch"

This script is designed to run without admin privileges. To run at login, create a shortcut
to pythonw.exe with this script path in your Startup folder (shell:startup).
"""

import sys
import time
import threading
import argparse

from pathlib import Path
from collections import deque, defaultdict

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler, DirCreatedEvent, FileCreatedEvent
except Exception:
    print("Missing dependency 'watchdog'. Install via: pip install --user watchdog", file=sys.stderr)
    raise

# win10toast shows Windows toast notifications from user session
try:
    from win10toast import ToastNotifier
    _TOASTER = ToastNotifier()
    _HAS_TOAST = True
except Exception:
    _TOASTER = None
    _HAS_TOAST = False


class NewFileHandler(FileSystemEventHandler):
    def __init__(self, notify_func, debounce_seconds=1.0):
        super().__init__()
        self.notify = notify_func
        self.debounce = debounce_seconds

        self._last = {}
        self._lock = threading.Lock()


    def _should_precess(self, path):
        now = time.time()
        with self._lock:
            last = self._last.get(path)
            if last and (now-last) < self.debounce:
                return False
        self._last[path] = now

        if len(self._last) > 50:
            cutoff = now - (self.debounce * 10)
            for key, value in list(self._last.items()):
                if value < cutoff:
                    del self._last[key]

        return True


    def on_created(self, event):
        if event.is_directory:
            return

        full = getattr(event, 'src_path', None) or getattr(event, 'pathname', None) or ''

        if not full:
            return
        if not self._should_precess(full):
            return

        self.notify(full)



def show_notification_win(name, fullpath):
    title = "file added"
    msg = name

    if _HAS_TOAST and _TOASTER:
        try:
            _TOASTER.show_toast(title, msg, duration=10, threaded=True)
            return
        except Exception:
            pass

    print(f'{time.strftime('%Y-%m-%d %H:%M:%S')} - {title}: {fullpath}')


def main():
    def notify(fullpath):
        name = Path(fullpath).name
        threading.Thread(target=show_notification_win, args=(name, fullpath), daemon=True).start()

    p = argparse.ArgumentParser(description='Watch folder for new files and show notifications.')
    p.add_argument('path', help='Folder to watch')
    p.add_argument('--debounce',
                   type=float,
                   default=1.0,
                   help='Debounce seconds for duplicate events'
                   )
    args = p.parse_args()

    target = Path(args.path).expanduser()
    if not target.exists() or not target.is_dir():
        print("Path does not exist or is not a directory:", target, file=sys.stderr)
        sys.exit(2)

    handler = NewFileHandler(notify, debounce_seconds=args.debounce)
    observer = Observer()
    observer.schedule(handler, str(target), recursive=True)
    observer.start()
    print(f"Watching folder: {target}  (Press Ctrl+F2 to stop)")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Stopping...")
    finally:
        observer.stop()
        observer.join()


if __name__ == "__main__":
    main()
