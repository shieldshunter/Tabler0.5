import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QListWidget, QCheckBox, QSlider, QComboBox, QAbstractItemView
from PyQt5.QtCore import Qt, QMimeData
from PIL import Image
from img2table.document import Image as Img2TableImage
from img2table.ocr import TesseractOCR
from pdf2image import convert_from_path
import pandas as pd
import pytesseract
import os
import fitz  # PyMuPDF

class PDFTableExtractor(QWidget):
    def __init__(self):
        super().__init__()

        self.layout = QVBoxLayout(self)
        self.load_button = QPushButton("Load PDF", self)
        self.layout.addWidget(self.load_button)

        self.list_widget = QListWidget(self)
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.layout.addWidget(self.list_widget)

        self.implicit_rows = QCheckBox("Implicit Rows", self)
        self.layout.addWidget(self.implicit_rows)

        self.borderless_tables = QCheckBox("Borderless Tables", self)
        self.layout.addWidget(self.borderless_tables)

        self.confidence_slider = QSlider(Qt.Horizontal, self)
        self.confidence_slider.setMinimum(0)
        self.confidence_slider.setMaximum(100)
        self.confidence_slider.setValue(50)
        self.layout.addWidget(self.confidence_slider)

        self.report_type = QComboBox(self)
        self.report_type.addItems(["Extract Tables", "Extract Images"])
        self.layout.addWidget(self.report_type)

        self.extract_button = QPushButton("Start Extraction", self)
        self.layout.addWidget(self.extract_button)

        self.load_button.clicked.connect(self.loadPDF)
        self.extract_button.clicked.connect(self.extract)

        self.fname = None
        self.image_files = []

    def loadPDF(self):
        self.fname = QFileDialog.getOpenFileName(self, 'Open file', '', "Pdf files (*.pdf)")
        if self.fname[0]:
            images = convert_from_path(self.fname[0])
            self.image_files = [f"Page_{i + 1}.png" for i in range(len(images))]
            for i, img in enumerate(images):
                img.save(self.image_files[i])

            self.list_widget.addItems(self.image_files)

    def process_file(self, file_path):
        pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR'  # your tesseract path here
        ocr = TesseractOCR()
        img = Img2TableImage(src=file_path)
        extracted_tables = img.extract_tables(ocr=ocr,
                                              implicit_rows=self.implicit_rows.isChecked(),
                                              borderless_tables=self.borderless_tables.isChecked(),
                                              min_confidence=self.confidence_slider.value())
        if extracted_tables:
            df = extracted_tables[0].df.copy()
            df_original = df.copy()  # make a copy of original dataframe
            df = df.iloc[1:, 1:]  # remove first row and first column
            df.fillna("", inplace=True)
            i = 0
            while i < len(df):
                if df.iloc[i, 1] == "":  # check if second column is empty
                    df.iloc[i - 1] = df.iloc[i - 1].astype(str) + " " + df.iloc[i].astype(str)  # append the whole row to the previous row
                    df = df.drop(df.index[i])  # remove the row
                    df = df.reset_index(drop=True)  # reset the index
                else:
                    i += 1
            return df, df_original
        else:
            print(f"No table found in file: {file_path}")
            return pd.DataFrame(), pd.DataFrame()

    def extract_images(self):
        pdf_file = fitz.open(self.fname[0])
        save_path = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
        for page_index in range(len(pdf_file)):
            page = pdf_file[page_index]
            image_list = page.get_images(full=True)
            for image_index, img in enumerate(image_list):
                xref = img[0]
                base_image = pdf_file.extract_image(xref)
                image_data = base_image["image"]
                image_name = f"{os.path.basename(self.fname[0])}_page{page_index+1}_image{image_index+1}.png"
                with open(os.path.join(save_path, image_name), "wb") as image_file:
                    image_file.write(image_data)

    def extract(self):
        if self.report_type.currentText() == "Extract Tables":
            self.extract_tables()
        elif self.report_type.currentText() == "Extract Images":
            self.extract_images()

    def extract_tables(self):
        selected_items = self.list_widget.selectedItems()
        if selected_items:
            dfs, dfs_original = zip(*[self.process_file(item.text()) for item in selected_items])
            save_path = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
            if save_path:
                output_name = os.path.join(save_path, os.path.basename(self.fname[0]).replace('.pdf', '') + '-output.csv')
                output_name_original = os.path.join(save_path, os.path.basename(self.fname[0]).replace('.pdf', '') + '-original.csv')
                final_df = pd.concat(dfs, axis=0)
                final_df.to_csv(output_name)
                final_df_original = pd.concat(dfs_original, axis=0)
                final_df_original.to_csv(output_name_original)
                print("Table extraction completed successfully.")
        else:
            # If no pages are selected for processing, process the entire PDF as a single image
            dfs, dfs_original = self.process_file(self.fname[0])
            save_path = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
            if save_path:
                output_name = os.path.join(save_path, os.path.basename(self.fname[0]).replace('.pdf', '') + '-output.csv')
                output_name_original = os.path.join(save_path, os.path.basename(self.fname[0]).replace('.pdf', '') + '-original.csv')
                dfs.to_csv(output_name)
                dfs_original.to_csv(output_name_original)
                print("Table extraction completed successfully.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFTableExtractor()
    ex.show()
    sys.exit(app.exec_())
