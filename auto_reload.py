import subprocess
import sys
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os

# Set your main script names here
WEB_SCRIPT = 'run_web.py'
DESKTOP_SCRIPT = 'attendance.py'

class RestartHandler(FileSystemEventHandler):
    def __init__(self, restart_callback):
        self.restart_callback = restart_callback

    def on_any_event(self, event):
        if event.is_directory:
            return
        if event.src_path.endswith('.py'):
            print(f"Detected change in {event.src_path}, restarting...")
            self.restart_callback()

class AppReloader:
    def __init__(self, script):
        self.script = script
        self.process = None
        self.start_app()

    def start_app(self):
        if self.process:
            self.process.terminate()
            self.process.wait()
        print(f"Starting {self.script}...")
        self.process = subprocess.Popen([sys.executable, self.script])

    def restart(self):
        self.start_app()

    def stop(self):
        if self.process:
            self.process.terminate()
            self.process.wait()

if __name__ == "__main__":
    # Choose which app to reload: 'web' or 'desktop'
    mode = sys.argv[1] if len(sys.argv) > 1 else 'web'
    script = WEB_SCRIPT if mode == 'web' else DESKTOP_SCRIPT
    reloader = AppReloader(script)
    event_handler = RestartHandler(reloader.restart)
    observer = Observer()
    observer.schedule(event_handler, path='.', recursive=True)
    observer.start()
    print(f"Watching for changes. Running {script}. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
        reloader.stop()
    observer.join()
