import shutil
import time
from pathlib import Path
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# Target root path setup
PRIMARY_DIR = Path(r"P:\\")
FALLBACK_DIR = Path(r"\\192.168.1.4\B1_Shr")

BASE_DIR = PRIMARY_DIR if PRIMARY_DIR.exists() else FALLBACK_DIR

# Monitor P:\Inbox and move incoming scans directly to AUGUST archive
WATCH_DIR = BASE_DIR / "Inbox"
ARCHIVE_DIR = (
    BASE_DIR
    / "Attachments"
    / "Acknowledged Deliveries"
    / "2026"
    / "AUGUST"
)

# Auto-create Inbox folder if not present
WATCH_DIR.mkdir(parents=True, exist_ok=True)
ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)


class DeliveryHandler(FileSystemEventHandler):

    def on_created(self, event):
        if event.is_directory or not event.src_path.endswith(".pdf"):
            return

        pdf_path = Path(event.src_path)
        print(f"[DETECTED] New delivery note: {pdf_path.name}")

        time.sleep(2)  # Wait for file transfer to complete

        try:
            target_path = ARCHIVE_DIR / pdf_path.name
            shutil.move(str(pdf_path), str(target_path))
            print(f"[SUCCESS] Moved {pdf_path.name} -> {ARCHIVE_DIR}")
        except Exception as e:
            print(f"[ERROR] Failed to process {pdf_path.name}: {e}")


if __name__ == "__main__":
    event_handler = DeliveryHandler()
    observer = Observer()
    observer.schedule(event_handler, str(WATCH_DIR), recursive=False)
    observer.start()

    print(f"Monitoring folder: {WATCH_DIR}")
    print(f"Archive destination: {ARCHIVE_DIR}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()