import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, 
                             QWidget, QLabel, QGridLayout, QDesktopWidget, QScrollArea)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt

# Fonksiyon modüllerini içe aktar
sys.path.append(os.path.join(os.path.dirname(__file__), 'functions'))
from functions.excel_merger import ExcelMergerWindow
from functions.excel_filter import ExcelFilterWindow
from functions.csv_to_excel import CsvToExcelWindow
from functions.excel_analyzer import ExcelAnalyzerWindow
from functions.excel_cleaner import ExcelCleanerWindow
from functions.excel_comparer import ExcelComparerWindow
from functions.excel_visualizer import ExcelVisualizerWindow
from functions.excel_reporter import ExcelReporterWindow
from functions.excel_macro import ExcelMacroWindow
from functions.excel_validator import ExcelValidatorWindow
from functions.excel_formula import ExcelFormulaWindow

class ExcelHubWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        
    def initUI(self):
        # Ana pencere ayarları
        self.setWindowTitle('Excel İşlem Merkezi')
        self.setMinimumSize(800, 600)
        self.center()
        
        # Ana widget ve scroll area
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Başlık
        title_label = QLabel('Excel İşlem Merkezi')
        title_font = QFont('Arial', 24, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # Alt başlık
        subtitle_label = QLabel('Excel\'i açmadan işlemlerinizi gerçekleştirin')
        subtitle_font = QFont('Arial', 14)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(subtitle_label)
        
        # Scroll Area oluştur
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        
        # Fonksiyon butonları için grid layout
        grid_layout = QGridLayout()
        scroll_layout.addLayout(grid_layout)
        
        # Scroll area'yı ayarla
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)
        
        # Excel Birleştirme butonu
        merge_button = self.create_function_button('Excel Dosyalarını Birleştir', 
                                                 'Excel dosyalarını tek bir dosyada birleştirin')
        merge_button.clicked.connect(self.open_excel_merger)
        grid_layout.addWidget(merge_button, 0, 0)
        
        # Excel Filtreleme butonu
        filter_button = self.create_function_button('Excel Dosyasında Veri Filtreleme', 
                                                 'Excel dosyalarındaki verileri filtreleyip yeni dosya oluşturun')
        filter_button.clicked.connect(self.open_excel_filter)
        grid_layout.addWidget(filter_button, 0, 1)
        
        # CSV'den Excel'e Dönüştürme butonu
        csv_to_excel_button = self.create_function_button('CSV Dosyasını Excel\'e Dönüştür', 
                                                 'CSV dosyalarını Excel formatına dönüştürün')
        csv_to_excel_button.clicked.connect(self.open_csv_to_excel)
        grid_layout.addWidget(csv_to_excel_button, 1, 0)
        
        # Excel Veri Analizi butonu
        excel_analyzer_button = self.create_function_button('Excel Veri Analizi', 
                                                 'Excel dosyalarındaki verileri analiz edin')
        excel_analyzer_button.clicked.connect(self.open_excel_analyzer)
        grid_layout.addWidget(excel_analyzer_button, 1, 1)
        
        # Excel Veri Temizleme butonu
        excel_cleaner_button = self.create_function_button('Excel Veri Temizleme', 
                                                 'Excel dosyalarındaki verileri temizleyin')
        excel_cleaner_button.clicked.connect(self.open_excel_cleaner)
        grid_layout.addWidget(excel_cleaner_button, 2, 0)
        
        # Excel Dosya Karşılaştırma butonu
        excel_comparer_button = self.create_function_button('Excel Dosya Karşılaştırma', 
                                                 'İki Excel dosyasını karşılaştırın')
        excel_comparer_button.clicked.connect(self.open_excel_comparer)
        grid_layout.addWidget(excel_comparer_button, 2, 1)
        
        # Excel Veri Görselleştirme butonu
        excel_visualizer_button = self.create_function_button('Excel Veri Görselleştirme', 
                                                 'Excel verilerinizi grafiklerle görselleştirin')
        excel_visualizer_button.clicked.connect(self.open_excel_visualizer)
        grid_layout.addWidget(excel_visualizer_button, 3, 0)
        
        # Excel Raporlama butonu
        excel_reporter_button = self.create_function_button('Excel Raporlama', 
                                                 'Excel verilerinizden otomatik raporlar oluşturun')
        excel_reporter_button.clicked.connect(self.open_excel_reporter)
        grid_layout.addWidget(excel_reporter_button, 3, 1)
        
        # Excel Makro butonu
        excel_macro_button = self.create_function_button('Excel Makro', 
                                                 'Excel makrolarını yönetin ve çalıştırın')
        excel_macro_button.clicked.connect(self.open_excel_macro)
        grid_layout.addWidget(excel_macro_button, 4, 0)
        
        # Excel Veri Doğrulama butonu
        excel_validator_button = self.create_function_button('Excel Veri Doğrulama', 
                                                 'Excel verilerinizi doğrulayın ve hataları tespit edin')
        excel_validator_button.clicked.connect(self.open_excel_validator)
        grid_layout.addWidget(excel_validator_button, 4, 1)
        
        # Excel Formül Oluşturucu butonu
        excel_formula_button = self.create_function_button('Excel Formül Oluşturucu', 
                                                 'Excel için formüller oluşturun ve düzenleyin')
        excel_formula_button.clicked.connect(self.open_excel_formula)
        grid_layout.addWidget(excel_formula_button, 5, 0)
        
        # Boşluk ekle
        scroll_layout.addStretch()
        
    def create_function_button(self, title, description):
        # Basit bir buton oluştur
        button = QPushButton()
        button.setText(f"{title}\n\n{description}")
        button.setMinimumSize(250, 150)
        button.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 2px solid #c0c0c0;
                border-radius: 10px;
                text-align: center;
                padding: 10px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
                border: 2px solid #a0a0a0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """)
        
        return button
    
    def open_excel_merger(self):
        self.excel_merger_window = ExcelMergerWindow()
        self.excel_merger_window.show()
    
    def open_excel_filter(self):
        self.excel_filter_window = ExcelFilterWindow()
        self.excel_filter_window.show()
        
    def open_csv_to_excel(self):
        self.csv_to_excel_window = CsvToExcelWindow()
        self.csv_to_excel_window.show()
    
    def open_excel_analyzer(self):
        self.excel_analyzer_window = ExcelAnalyzerWindow()
        self.excel_analyzer_window.show()
    
    def open_excel_cleaner(self):
        self.excel_cleaner_window = ExcelCleanerWindow()
        self.excel_cleaner_window.show()
    
    def open_excel_comparer(self):
        self.excel_comparer_window = ExcelComparerWindow()
        self.excel_comparer_window.show()
    
    def open_excel_visualizer(self):
        self.excel_visualizer_window = ExcelVisualizerWindow()
        self.excel_visualizer_window.show()
    
    def open_excel_reporter(self):
        self.excel_reporter_window = ExcelReporterWindow()
        self.excel_reporter_window.show()
        
    def open_excel_macro(self):
        self.excel_macro_window = ExcelMacroWindow()
        self.excel_macro_window.show()
        
    def open_excel_validator(self):
        self.excel_validator_window = ExcelValidatorWindow()
        self.excel_validator_window.show()
        
    def open_excel_formula(self):
        self.excel_formula_window = ExcelFormulaWindow()
        self.excel_formula_window.show()
    
    def center(self):
        # Pencereyi ekranın ortasına konumlandır
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern görünüm için
    window = ExcelHubWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()