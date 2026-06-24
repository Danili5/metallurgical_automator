import os
import sys
import time
import pandas as pd
from pathlib import Path
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

load_dotenv()

DATA = os.getenv("DATA")
FORMS = os.getenv("FORMS")

class Main(FileSystemEventHandler):
    def on_created(self, event):    
        if event.is_directory:
            return
        
        file = Path(event.src_path)

        if file.name.startswith("~$") or file.suffix not in [".csv", ".is_tens"]:
            return

        time.sleep(0.5)

        file_number = file.stem

        test_results_file = None
        test_meta_file = None

        if file.suffix == ".csv":
            test_results_file = file
            test_meta_file = file.with_suffix(".is_tens")
        else:
            test_meta_file = file
            test_results_file = file.with_suffix(".csv")

        if not test_results_file.exists() or not test_meta_file.exists():
            return

        if file.suffix != ".csv":
            return

        try:
            print(test_results_file, test_meta_file, flush = True)
        except PermissionError:
            print("Permission Error retrying again", flush = True)
        except Exception:
            pritn("Processing Error", flush = True)

observer = Observer()

observer.schedule(Main(), DATA)
observer.start()

try:
    if observer.is_alive():
        print("started", flush = True)

    while observer.is_alive():
        observer.join(1)
finally:
    observer.stop()
    observer.join()