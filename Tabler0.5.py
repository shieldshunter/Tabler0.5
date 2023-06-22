import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QListWidget, QCheckBox, QSlider, QComboBox, QAbstractItemView, QGraphicsView, QGraphicsScene, QGraphicsRectItem, QFrame
from PyQt5.QtCore import Qt, QMimeData, QRectF
from PyQt5.QtGui import QPixmap, QImage
from PIL import Image
from img2table.document import Image as Img2TableImage
from img2table.ocr import TesseractOCR
from pdf2image import convert_from_path
import pandas as pd
import pytesseract
import os
import fitz  # PyMuPDF
import cv2
import numpy as np


class CropGraphicsView(QGraphicsView):
    def __init__(self):
        super().__init__()
        self.setScene(QGraphicsScene())
        self.rect_item = QGraphicsRectItem()
        self.scene().addItem(self.rect_item)
        self.setFrameStyle(QFrame.NoFrame)

    def set_image(self, img_path):
        self.scene().clear()
        self.scene().addPixmap(QPixmap(img_path))
        self.rect_item = QGraphicsRectItem()
        self.scene().addItem(self.rect_item)

    def mousePressEvent(self, event):
        self.origin = event.pos()
        self.rect_item.setRect(QRectF(self.origin, self.origin))

    def mouseMoveEvent(self, event):
        self.rect_item.setRect(QRectF(self.origin, event.pos()).normalized())

    def mouseReleaseEvent(self, event):
        self.crop_rect = self.rect_item.rect()


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

        self.crop_button = QPushButton("Crop Image", self)
        self.layout.addWidget(self.crop_button)

        self.extract_button = QPushButton("Start Extraction", self)
        self.layout.addWidget(self.extract_button)

        self.load_button.clicked.connect(self.loadPDF)
        self.crop_button.clicked.connect(self.crop_image)
        self.extract_button.clicked.connect(self.extract)

        self.fname = None
        self.image_files = []
        self.crop_view = CropGraphicsView()

    def loadPDF(self):
        self.fname = QFileDialog.getOpenFileName(self, 'Open file', '', "Pdf files (*.pdf)")
        if self.fname[0]:
            images = convert_from_path(self.fname[0])
            self.image_files = [f"{os.path.splitext(self.fname[0])[0]}_Page_{i + 1}.png" for i in range(len(images))]
            for i, img in enumerate(images):
                img.save(self.image_files[i])

            self.list_widget.addItems(self.image_files)

    def crop_image(self):
        selected_files = self.list_widget.selectedItems()
        for file in selected_files:
            self.crop_view.set_image(file.text())
            self.crop_view.show()
            img = cv2.imread(file.text())
            x, y, w, h = map(int, self.crop_view.crop_rect.getCoords())
            crop_img = img[y:y + h, x:x + w]
            cv2.imwrite(file.text(), crop_img)

    def extract(self):
        # Process each image file
        for image_file in self.image_files:
            # Load image
            img = Img2TableImage(image_file)

            # OCR engine
            ocr = TesseractOCR()

            # Extract tables
            extracted_tables = img.extract_tables(ocr=ocr,
                                                  layout_kwargs={'borderless': self.borderless_tables.isChecked()},
                                                  ocr_kwargs={'scale': 1, 'remove_noise': False, 'remove_border': True,
                                                              'binarization': self.confidence_slider.value(),
                                                              'implicit_rows': self.implicit_rows.isChecked()})

            # Save the extracted tables to CSV files
            for i, table in enumerate(extracted_tables):
                output_csv = image_file.replace('.png', f'_table_{i + 1}.csv')
                table.to_csv(output_csv, index=False)

        print("Extraction Completed!")


def main():
    app = QApplication(sys.argv)
    extractor = PDFTableExtractor()
    extractor.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
