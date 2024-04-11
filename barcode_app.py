import sys
import PIL
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QLineEdit, QPushButton
import barcode
from barcode.writer import ImageWriter
import win32print
import win32ui
import os
import sys
import win32con
from PIL import Image, ImageWin

class BarcodeApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Barcode Generator')
        self.layout = QVBoxLayout()

        self.label1 = QLabel('Serial 1:')
        self.field1 = QLineEdit()
        self.layout.addWidget(self.label1)
        self.layout.addWidget(self.field1)

        self.label2 = QLabel('Serial 2:')
        self.field2 = QLineEdit()
        self.layout.addWidget(self.label2)
        self.layout.addWidget(self.field2)

        self.label3 = QLabel('Serial 3:')
        self.field3 = QLineEdit()
        self.layout.addWidget(self.label3)
        self.layout.addWidget(self.field3)

        self.generateButton = QPushButton('Generate Barcodes')
        self.exitButton = QPushButton('Exit App')
        self.generateButton.clicked.connect(self.generateBarcodes)
        self.exitButton.clicked.connect(self.exitApp)
        self.layout.addWidget(self.generateButton)
        self.layout.addWidget(self.exitButton)
        self.setLayout(self.layout)

    def generateBarcodes(self):
        data1 = self.field1.text()
        data2 = self.field2.text()
        data3 = self.field3.text()

        render_options = {
            "module_width": 0.4,
            "font_size": 15,
            "text_distance":7,
            "center_text": True,
            "module_height": 12,
            "write_text": True,
            "quiet_zone": 1,
        }

        for i, data in enumerate([data1, data2, data3], start=1):
            if data == "":
                continue
            barcode_class = barcode.get_barcode_class('code128')
            barcode_instance = barcode_class(data, writer=ImageWriter())
            filename = f'barcode_{i}'

            # Get the directory of the executable
            executable_dir = os.path.dirname(sys.argv[0])
            # Use this directory to construct the path for saving images
            image_dir = os.path.join(executable_dir, "images")
            # Create the images directory if it doesn't exist
            if not os.path.exists(image_dir):
                os.makedirs(image_dir)

            image_path = os.path.join(image_dir, filename)
            barcode_instance.save(image_path, render_options)

            imageFullPath = image_path + '.png';
            to_be_resized = Image.open(imageFullPath)  # open in a PIL Image object
            newSize = (656, 279)
            resized = to_be_resized.resize(newSize, resample=PIL.Image.NEAREST)  # you can choose

            resized.save(imageFullPath)  # save the resized image

            # Print the generated barcode image
            self.printImage(image_path + ".png")

    def exitApp(self):
        sys.exit(app.exec())

    def printImage(self, filename):
        printer_name = win32print.GetDefaultPrinter()
        printer_dc = win32ui.CreateDC()
        printer_dc.CreatePrinterDC(printer_name)

        printer_size = printer_dc.GetDeviceCaps(3), printer_dc.GetDeviceCaps(4)

        image = Image.open(filename)
        image = image.rotate(90, expand=True)  # Rotate the image by 90 degrees
        image = image.convert("1")  # Convert to 1-bit image (black and white)
        image_width, image_height = image.size

        dib = ImageWin.Dib(image)
        dib_width, dib_height = dib.size

        printer_dc.StartDoc(filename)
        printer_dc.StartPage()

        x1 = 0
        y1 = 20
        x2 = x1 + dib_width
        y2 = y1 + dib_height

        dib.draw(printer_dc.GetHandleOutput(), (x1, y1, x2, y2))

        printer_dc.EndPage()
        printer_dc.EndDoc()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec())