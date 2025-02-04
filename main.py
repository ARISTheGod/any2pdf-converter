import logging
import os
import time
from functools import wraps
from pathlib import Path
from typing import List

import PyPDF2
import aspose.slides as slides
import aspose.words as words
import img2pdf
from dotenv import load_dotenv
from fpdf import FPDF  # For converting text files
from fpdf.enums import XPos, YPos  # Enums for new cell parameters

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def log_execution_time(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        logging.info(f"{func.__name__} executed in {end_time - start_time:.2f} seconds")
        return result
    return wrapper


# Custom subclass so that every added page automatically draws the code box
class CodePDF(FPDF):
    def header(self):
        # For pages after the first, automatically draw the code box from top margin.
        # (On the first page we add a header manually before drawing the box.)
        if self.page_no() > 1:
            x0 = 10
            y0 = 10
            box_width = self.w - 20
            box_height = self.h - 20  # leave a 10-unit bottom margin
            self.set_fill_color(240, 240, 240)
            self.rect(x0, y0, box_width, box_height, 'FD')
            self.set_xy(x0 + 2, y0 + 2)


class FileConverter:
    def __init__(self, input_folder: str, output_folder: str):
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)
        self.temp_folder = self.output_folder / "temp"
        self.temp_folder.mkdir(parents=True, exist_ok=True)

    @log_execution_time
    def convert_ppt_to_pdf(self, input_file: Path) -> Path:
        with slides.Presentation(str(input_file)) as presentation:
            output_file = self.temp_folder / f"{input_file.stem}.pdf"
            presentation.save(str(output_file), slides.export.SaveFormat.PDF)
        logging.info(f"Converted {input_file} to PDF")
        return output_file

    @log_execution_time
    def convert_doc_to_pdf(self, input_file: Path) -> Path:
        doc = words.Document(str(input_file))
        output_file = self.temp_folder / f"{input_file.stem}.pdf"
        doc.save(str(output_file), words.SaveFormat.PDF)
        logging.info(f"Converted {input_file} to PDF")
        return output_file

    @log_execution_time
    def convert_image_to_pdf(self, input_file: Path) -> Path:
        output_file = self.temp_folder / f"{input_file.stem}.pdf"
        with open(str(output_file), "wb") as f:
            f.write(img2pdf.convert(str(input_file)))
        logging.info(f"Converted {input_file} to PDF")
        return output_file

    @log_execution_time
    def convert_text_to_pdf(self, input_file: Path) -> Path:
        # Use our custom CodePDF class so that header() is called on every new page.
        pdf = CodePDF()
        pdf.add_page()
        pdf.set_font("Courier", size=10)
        # Write the file path header on the first page.
        pdf.cell(0, 10, text=f"File path: {input_file}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(4)
        # Draw the code box for the first page manually below the header:
        current_y = pdf.get_y()
        x0 = 10
        box_width = pdf.w - 20
        box_height = pdf.h - current_y - 10
        pdf.set_fill_color(240, 240, 240)
        pdf.rect(x0, current_y, box_width, box_height, 'FD')
        pdf.set_xy(x0 + 2, current_y + 2)

        box_width_inner = box_width - 4

        with open(input_file, "r", encoding="utf-8") as file:
            for line in file:
                # Check if the current y-position nears the bottom (i.e. if there's not enough space to print a new line)
                if pdf.get_y() > pdf.h - 15:
                    pdf.add_page()  # on new page, header() gets called automatically
                pdf.multi_cell(box_width_inner, 5, text=line)

        output_file = self.temp_folder / f"{input_file.stem}.pdf"
        pdf.output(str(output_file))
        logging.info(f"Converted {input_file} to PDF")
        return output_file

    @log_execution_time
    def process_files(self) -> List[Path]:
        pdf_files = []
        for file in self.input_folder.iterdir():
            if file.is_file():
                try:
                    ext = file.suffix.lower()
                    if ext in ('.ppt', '.pptx'):
                        pdf_files.append(self.convert_ppt_to_pdf(file))
                    elif ext in ('.doc', '.docx'):
                        pdf_files.append(self.convert_doc_to_pdf(file))
                    elif ext in ('.png', '.jpg', '.jpeg'):
                        pdf_files.append(self.convert_image_to_pdf(file))
                    elif ext == '.pdf':
                        pdf_files.append(file)
                    elif ext in ('.py', '.md'):
                        pdf_files.append(self.convert_text_to_pdf(file))
                    else:
                        logging.warning(f"Unsupported file type: {file}")
                except Exception as e:
                    logging.error(f"Error processing {file}: {str(e)}")
        return pdf_files

    @log_execution_time
    def merge_pdfs(self, pdf_files: List[Path]) -> Path:
        merger = PyPDF2.PdfMerger()
        for pdf in pdf_files:
            merger.append(str(pdf))
        output_file = self.output_folder / "merged_all_files.pdf"
        merger.write(str(output_file))
        merger.close()
        logging.info(f"Merged PDFs saved to {output_file}")
        return output_file

    def cleanup(self) -> None:
        for file in self.temp_folder.iterdir():
            file.unlink()
        self.temp_folder.rmdir()
        logging.info("Temporary files cleaned up")


def main() -> None:
    input_folder = os.getenv('INPUT_FOLDER')
    output_folder = os.getenv('OUTPUT_FOLDER')

    if not input_folder or not output_folder:
        logging.error("INPUT_FOLDER and OUTPUT_FOLDER must be set in the .env file")
        return

    converter = FileConverter(input_folder, output_folder)
    pdf_files = converter.process_files()
    merged_file = converter.merge_pdfs(pdf_files)
    converter.cleanup()
    logging.info(f"All files merged into {merged_file}")


if __name__ == "__main__":
    main()
