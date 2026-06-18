from pathlib import Path
from docx import Document
from docx.shared import Pt
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from dotenv import load_dotenv
import pandas as pd
import os

load_dotenv()

TEMPLATE_PATH = os.getenv("TEMPLATE_DIRECTORY")
REPORT_PATH = os.getenv("REPORT_DIRECTORY")
DATA_PATH = os.getenv("DATA_DIRECTORY")

class Main(FileSystemEventHandler):
    def on_created(self, event):
        try:
            template = Document(list(Path(TEMPLATE_PATH).glob("*.docx"))[0])
            data = {f"+{key} {sub_key}+": sub_value for key, value in pd.read_csv(max([file for file in Path(DATA_PATH).glob("*.csv") if file.is_file()], key = lambda file: file.stat().st_ctime), skiprows = [0,2], index_col = 0).astype(str).to_dict().items() for sub_key, sub_value in value.items()}

            for row in template.tables[1].rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:        
                        for key, value in data.items():
                            if key in paragraph.text:
                                try:
                                    value = f"{int(value):,}"
                                except ValueError:
                                    pass

                                new_text = paragraph.text.replace(key, value)
                                paragraph.text = ""
                                paragraph.text = new_text
                                for run in paragraph.runs:
                                    run.font.name = "Bookman Old Style"
                                    run.font.size = Pt(10)

                        if "+" in paragraph.text:
                            paragraph.text = ""

            template.save(str(Path(REPORT_PATH)) + r"\weld_report.docx")
        except Exception as exception:
            print(exception)

main = Main()
observer = Observer()
observer.schedule(main, Path(DATA_PATH))
observer.start()

try:
    while observer.is_alive():
        observer.join(1)
finally:
    observer.stop()
    observer.join()