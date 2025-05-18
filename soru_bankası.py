
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QRadioButton, QTextEdit,QTableWidget, QTableWidgetItem, QButtonGroup
import xlwings as xw
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtGui import QPainter

# --- Ortak soru listesi ---
soru_listesi = []

class YeniSoruEklePenceresi(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Yeni Soru Ekle')
        self.setGeometry(300, 300, 1000, 400)
        self.initUI()

    def initUI(self):
        ana_layout = QHBoxLayout()

        # Sol taraf - Soru ve cevaplar
        sol_layout = QVBoxLayout()

        soru_label = QLabel('SORU')
        self.soru_edit = QTextEdit()
        sol_layout.addWidget(soru_label)
        sol_layout.addWidget(self.soru_edit)

        self.cevaplar = []
        self.radio_group = QButtonGroup()

        for i in range(5):
            cevap_layout = QHBoxLayout()
            cevap = QLineEdit()
            radio_button = QRadioButton('Doğru Şık')
            self.radio_group.addButton(radio_button, id=i)

            cevap_layout.addWidget(QLabel(f'{i+1}. Yanıt'))
            cevap_layout.addWidget(cevap)
            cevap_layout.addWidget(radio_button)

            sol_layout.addLayout(cevap_layout)
            self.cevaplar.append((cevap, radio_button))

        self.ekle_buton = QPushButton('SORU BANKASINA EKLE')
        self.ekle_buton.clicked.connect(self.soru_ekle)
        sol_layout.addWidget(self.ekle_buton)

        self.buton_2 = QPushButton('SORU BANKASINA EXCEL OLARAK KAYDET.')
        self.buton_2.clicked.connect(self.excel_kaydet)  
        sol_layout.addWidget(self.buton_2)

        ana_layout.addLayout(sol_layout)

        # Sağ taraf - Soru tablosu
        self.soru_tablosu = QTableWidget()
        self.soru_tablosu.setColumnCount(6)
        self.soru_tablosu.setHorizontalHeaderLabels(
            ['Soru', '1. Seçenek', '2. Seçenek', '3. Seçenek', '4. Seçenek', '5. Seçenek']
        )
        ana_layout.addWidget(self.soru_tablosu)

        self.setLayout(ana_layout)

        # SoruSec penceresi örneğini başta None yapıyoruz
        self.soru_sec_penceresi = None

    def soru_ekle(self):
        soru = self.soru_edit.toPlainText()
        secenekler = [cevap_edit.text() for cevap_edit, _ in self.cevaplar]

        dogru_index = self.radio_group.checkedId()

        if soru and secenekler[0] and secenekler[1] and dogru_index != -1:
            row_position = self.soru_tablosu.rowCount()
            self.soru_tablosu.insertRow(row_position)

            self.soru_tablosu.setItem(row_position, 0, QTableWidgetItem(soru))
            for i, secenek in enumerate(secenekler):
                self.soru_tablosu.setItem(row_position, i+1, QTableWidgetItem(secenek))

            # doğru cevap
            dogru_cevap = secenekler[dogru_index]

            # Soru listesine ekle
            soru_listesi.append([soru] + secenekler + [dogru_cevap])

            # Eğer soru_sec_penceresi açıksa onu güncelle
            if self.soru_sec_penceresi is not None:
                self.soru_sec_penceresi.tabloyu_guncelle()

            # Temizle
            self.soru_edit.clear()
            for cevap_edit, _ in self.cevaplar:
                cevap_edit.clear()
            self.radio_group.setExclusive(False)
            for _, radio in self.cevaplar:
                radio.setChecked(False)
            self.radio_group.setExclusive(True)
        else:
            print("En az soru, iki seçenek ve bir doğru cevap seçilmelidir.")

   

    def excel_kaydet(self):  
        rows = self.soru_tablosu.rowCount()
        cols = self.soru_tablosu.columnCount()
    
        data = []
        for row in range(rows):
            row_data = []
            for col in range(cols):
                item = self.soru_tablosu.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
    
        headers = [self.soru_tablosu.horizontalHeaderItem(i).text() for i in range(cols)]
    
        # Excel penceresi açılır
        app = xw.App(visible=True , add_book=False)
        wb = app.books.add()
        ws = wb.sheets[0]
    
        # Başlıkları ve verileri yaz
        ws.range('A1').value = [headers] + data
    

class AnaPencere(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Soru Bankası')
        self.setGeometry(200, 200, 400, 300)
        self.initUI()

    def initUI(self):
        menubar = self.menuBar()
        islem_menu = menubar.addMenu('İşlem')

        yeni_soru_ekle_action = QAction('Yeni Soru Ekle', self)
        yeni_soru_ekle_action.triggered.connect(self.yeni_soru_ekle)
        islem_menu.addAction(yeni_soru_ekle_action)

        soru_sec_action = QAction('Soru Seç', self)
        soru_sec_action.triggered.connect(self.soru_sec)
        islem_menu.addAction(soru_sec_action)

    def yeni_soru_ekle(self):
        self.yeni_soru_penceresi = YeniSoruEklePenceresi()
        self.yeni_soru_penceresi.show()

    def soru_sec(self):
        self.sorusec_penceresi = SoruSec()
        self.sorusec_penceresi.show()
        self.sorusec_penceresi.tabloyu_guncelle()

class SoruSec(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Soru Seç')
        self.setGeometry(300, 300, 1000, 400)
        self.initUI()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()

        self.soru_tablosu = QTableWidget()
        self.soru_tablosu.setColumnCount(7)
        self.soru_tablosu.setHorizontalHeaderLabels(
            ['Soru', '1. Seçenek', '2. Seçenek', '3. Seçenek', '4. Seçenek', '5. Seçenek', 'Cevap']
        )
        layout.addWidget(self.soru_tablosu)

        buton_layout = QHBoxLayout()
        self.yazdir_buton = QPushButton('YAZDIR')
        self.yazdir_buton.clicked.connect(self.yazdir)
        buton_layout.addWidget(self.yazdir_buton)
        layout.addLayout(buton_layout)

        central_widget.setLayout(layout)

    def tabloyu_guncelle(self):
        self.soru_tablosu.setRowCount(0)
        for soru in soru_listesi:
            row_position = self.soru_tablosu.rowCount()
            self.soru_tablosu.insertRow(row_position)
            for sutun, veri in enumerate(soru):
                self.soru_tablosu.setItem(row_position, sutun, QTableWidgetItem(veri))
    
    def yazdir(self):      
     printer = QPrinter(QPrinter.HighResolution)
     dialog = QPrintDialog(printer, self)
     
     if dialog.exec_() == QPrintDialog.Accepted:
         painter = QPainter()
         if painter.begin(printer):
             # Tabloyu bir widget olarak render et
             self.soru_tablosu.render(painter)
             painter.end()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = AnaPencere()
    window.show()
    sys.exit(app.exec_())
