import sys
import numpy as np
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QMessageBox, QLineEdit, QLabel, QGridLayout
from PyQt5.QtCore import QTimer
from datetime import datetime, timedelta
import calendar






class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.folder_path = ""
        self.excel_data = {}

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Selectează un Folder")
        self.setGeometry(100, 100, 800, 600)
        
        layout = QVBoxLayout()

        self.btn_select_folder = QPushButton("Selectează Folder", self)
        self.btn_select_folder.clicked.connect(self.select_folder)
        layout.addWidget(self.btn_select_folder)

        self.setLayout(layout)
        self.show()

    def select_folder(self):
        self.folder_path = QFileDialog.getExistingDirectory(self, "Selectează un folder")
        if self.folder_path:
            self.check_and_create_excel_files()
            self.load_excel_files()
            self.open_main_window()


    def create_ghid(self):
        ghid_path = os.path.join(self.folder_path, "Ghid.txt")
        with open(ghid_path, "w") as f:
            f.write("Acesta este un ghid pentru utilizatorii aplicatiei noastre.\n\n")
            f.write("1. Introducere:\n")
            f.write("\tAceasta secțiune oferă o prezentare generală a aplicației.\n\n")
            f.write("2. Instalare:\n")
            f.write("\tPasul 1: Descărcați aplicația de pe site-ul nostru.\n")
            f.write("\tPasul 2: Rulați programul de instalare și urmați instrucțiunile.\n\n")
            f.write("3. Utilizare:\n")
            f.write("\t\t3.1. Crearea unui cont:\n")
            f.write("\t\t\tPasul 1: Deschideți aplicația și apăsați pe 'Creare cont'.\n")
            f.write("\t\t\tPasul 2: Introduceți informațiile solicitate și apăsați 'Submit'.\n")
            f.write("\t\t3.2. Autentificare:\n")
            f.write("\t\t\tPasul 1: Introduceți numele de utilizator și parola.\n")
            f.write("\t\t\tPasul 2: Apăsați pe 'Autentificare'.\n\n")
            f.write("4. Funcționalități:\n")
            f.write("\tAplicația noastră oferă următoarele funcționalități:\n")
            f.write("\t\t- Gestionare utilizatori\n")
            f.write("\t\t- Generare rapoarte\n")
            f.write("\t\t- Configurări avansate\n\n")
            f.write("5. Întrebări frecvente:\n")
            f.write("\tAceasta secțiune răspunde la întrebările cele mai comune.\n\n")
            f.write("6. Asistență tehnică:\n")
            f.write("\tPentru asistență tehnică, vă rugăm să ne contactați la suport@aplicatie.ro.\n")
            
            
            
            
            
    def check_and_create_excel_files(self):
        file_names = [" - intrari", " - vanzare cu factura", " - vanzari casa de marcat", " - transfer", " - centralizator", " - centralizator SGR"]
        file_paths = [os.path.join(self.folder_path, os.path.basename(self.folder_path) + name + ".xlsx") for name in file_names]

        for file_path in file_paths:
            if not os.path.exists(file_path):
                df = pd.DataFrame()
                df.to_excel(file_path, index=False)
        #self.create_ghid()

    def load_excel_files(self):
        file_names = [" - intrari", " - vanzare cu factura", " - vanzari casa de marcat", " - transfer", " - centralizator", " - centralizator SGR"]
        file_paths = [os.path.join(self.folder_path, os.path.basename(self.folder_path) + name + ".xlsx") for name in file_names]

        self.excel_data = {os.path.basename(file_path): pd.read_excel(file_path) for file_path in file_paths}

    def open_main_window(self):
        self.main_window = MainPage(self.folder_path, self.excel_data)
        self.main_window.show()

class MainPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Pagina Principală" + " - " + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        self.btn_intrari = QPushButton("Intrari", self)
        self.btn_intrari.clicked.connect(self.open_intrari_window)
        layout.addWidget(self.btn_intrari)

        self.btn_vanzare_factura = QPushButton("Vanzare cu Factura", self)
        self.btn_vanzare_factura.clicked.connect(self.open_vanzare_factura_window)
        layout.addWidget(self.btn_vanzare_factura)

        self.btn_vanzare_casa_marcat = QPushButton("Vanzare Casa Marcat", self)
        self.btn_vanzare_casa_marcat.clicked.connect(self.open_vanzare_casa_marcat_window)
        layout.addWidget(self.btn_vanzare_casa_marcat)

        self.btn_transfer_intre_gest = QPushButton("Transfer intre Gest", self)
        self.btn_transfer_intre_gest.clicked.connect(self.open_transfer_intre_gest_window)
        layout.addWidget(self.btn_transfer_intre_gest)

        self.btn_centralizator = QPushButton("Centralizator", self)
        self.btn_centralizator.clicked.connect(self.open_centralizator)
        layout.addWidget(self.btn_centralizator)

        self.btn_centralizator_sgr = QPushButton("Centralizator SGR", self)
        self.btn_centralizator_sgr.clicked.connect(self.open_centralizator_sgr)
        layout.addWidget(self.btn_centralizator_sgr)


        self.setLayout(layout)

    def open_centralizator_sgr(self):
        self.centralizator_sgr_window = CentralizatorSGRPage(self.folder_path, self.excel_data)
        self.centralizator_sgr_window.show()

    def open_centralizator(self):
        self.centralizator_window = CentralizatorPage(self.folder_path, self.excel_data)
        self.centralizator_window.show()


    def refresh_data(self):
        self.check_and_create_excel_files()
        self.load_excel_files()
        QMessageBox.information(self, "Info", "Datele au fost reîncărcate cu succes!")

    def check_and_create_excel_files(self):
        file_names = [" - intrari", " - vanzare cu factura", " - vanzari casa de marcat", " - transfer", " - centralizator", " - centralizator SGR"]
        file_paths = [os.path.join(self.folder_path, os.path.basename(self.folder_path) + name + ".xlsx") for name in file_names]

        for file_path in file_paths:
            if not os.path.exists(file_path):
                df = pd.DataFrame()
                df.to_excel(file_path, index=False)

    def load_excel_files(self):
        file_names = [" - intrari", " - vanzare cu factura", " - vanzari casa de marcat", " - transfer", " - centralizator", " - centralizator SGR"]
        file_paths = [os.path.join(self.folder_path, os.path.basename(self.folder_path) + name + ".xlsx") for name in file_names]

        self.excel_data = {os.path.basename(file_path): pd.read_excel(file_path) for file_path in file_paths}

    def open_intrari_window(self):
        self.intrari_window = IntrariPage(self.folder_path, self.excel_data)
        self.intrari_window.show()

    def open_vanzare_factura_window(self):
        self.vanzare_factura_window = VanzareFacturaPage(self.folder_path, self.excel_data)
        self.vanzare_factura_window.show()

    def open_vanzare_casa_marcat_window(self):
        self.vanzare_casa_marcat_window = VanzareCasaMarcatPage(self.folder_path, self.excel_data)
        self.vanzare_casa_marcat_window.show()

    def open_transfer_intre_gest_window(self):
        self.transfer_intre_gest_window = TransferIntreGestPage(self.folder_path, self.excel_data)
        self.transfer_intre_gest_window.show()








class CentralizatorSGRPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.sold_initial_value = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Centralizator SGR" + " - " + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        self.layout = QGridLayout()
        self.setLayout(self.layout)

        # Buton pentru setare sold initial
        self.btn_sold_initial = QPushButton("Setare Sold Initial SGR", self)
        self.btn_sold_initial.clicked.connect(self.open_sold_initial)
        self.layout.addWidget(self.btn_sold_initial, 0, 0)

        # Buton pentru refresh centralizator
        self.btn_refresh_centralizator = QPushButton("Refresh Centralizator SGR", self)
        self.btn_refresh_centralizator.clicked.connect(self.refresh_centralizator)
        self.layout.addWidget(self.btn_refresh_centralizator, 0, 1)

        # Back button
        self.btn_back = QPushButton("Back", self)
        self.btn_back.clicked.connect(self.close)
        self.layout.addWidget(self.btn_back, 0, 2)

    def open_sold_initial(self):
        self.sold_initial_window = QWidget()
        self.sold_initial_window.setWindowTitle("Setare Sold Initial SGR")
        self.sold_initial_window.setGeometry(150, 150, 400, 200)

        layout = QGridLayout()
        self.sold_initial_window.setLayout(layout)

        layout.addWidget(QLabel("Sold Initial"), 0, 0)
        self.sold_initial_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.sold_initial_input, 0, 1)

        layout.addWidget(QLabel("Luna (1-12)"), 1, 0)
        self.month_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.month_input, 1, 1)

        layout.addWidget(QLabel("Anul"), 2, 0)
        self.year_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.year_input, 2, 1)

        btn_save = QPushButton("Save", self.sold_initial_window)
        btn_save.clicked.connect(self.save_sold_initial)
        layout.addWidget(btn_save, 3, 0)

        btn_cancel = QPushButton("Cancel", self.sold_initial_window)
        btn_cancel.clicked.connect(self.sold_initial_window.close)
        layout.addWidget(btn_cancel, 3, 1)

        self.sold_initial_window.show()

    def save_sold_initial(self):
        try:
            self.sold_initial_value = float(self.sold_initial_input.text())
            self.month = int(self.month_input.text())
            self.year = int(self.year_input.text())
        except ValueError:
            QMessageBox.warning(self, "Eroare", "Vă rugăm să introduceți valori numerice valide pentru sold, lună și an.")
            return

        if self.sold_initial_value < 0:
            QMessageBox.warning(self, "Eroare", "Soldul inițial nu poate fi negativ.")
            return

        centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator SGR.xlsx")

        # Generăm toate zilele lunii respective
        start_date = datetime(self.year, self.month, 1)
        end_date = (start_date + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        dates = pd.date_range(start_date, end_date)

        data = {
            "Data": dates.strftime('%d/%m/%Y'),
            "Intrari SGR": [0] * len(dates),
            "Vanzari SGR Casa Marcat": [0] * len(dates),
            "Vanzari SGR Factura": [0] * len(dates),
            "Transfer SGR": [0] * len(dates),
            "Sold Final": [0] * len(dates),
            "Sold Initial": [self.sold_initial_value] * len(dates),
            "Numele si Anul": f"{os.path.basename(self.folder_path)}"
        }

        df = pd.DataFrame(data)
        writer = pd.ExcelWriter(centralizator_path, engine='xlsxwriter')

        # Salvăm în fișierul Excel
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            }
        )



        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)


        writer.close()

        QMessageBox.information(self, "Succes", f"Fișierul centralizator a fost creat/actualizat")
        self.completeaza_centralizator()
        self.sold_initial_window.close()

    def refresh_centralizator(self):
        # if - centralizator SGR.xlsx is empty then error message
        centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator SGR.xlsx")
        if pd.read_excel(centralizator_path).empty:
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator SGR este gol.")
            return
        self.completeaza_centralizator()

    def completeaza_centralizator(self):
        centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator SGR.xlsx")

        if not os.path.exists(centralizator_path):
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator nu există.")
            return

        try:
            df_centralizator = pd.read_excel(centralizator_path)
        except FileNotFoundError:
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator nu a putut fi citit.")
            return

        # Convertim coloanele la float pentru a evita problemele de compatibilitate
        df_centralizator['Intrari SGR'] = df_centralizator['Intrari SGR'].astype(float)
        df_centralizator['Vanzari SGR Casa Marcat'] = df_centralizator['Vanzari SGR Casa Marcat'].astype(float)
        df_centralizator['Vanzari SGR Factura'] = df_centralizator['Vanzari SGR Factura'].astype(float)
        df_centralizator['Transfer SGR'] = df_centralizator['Transfer SGR'].astype(float)
        df_centralizator['Sold Final'] = df_centralizator['Sold Final'].astype(float)

        # Resetăm coloanele care vor fi recalculată
        df_centralizator['Intrari SGR'] = 0.0
        df_centralizator['Vanzari SGR Casa Marcat'] = 0.0
        df_centralizator['Vanzari SGR Factura'] = 0.0
        df_centralizator['Transfer SGR'] = 0.0
        df_centralizator['Sold Final'] = 0.0

        # Citim fișierele de date
        intrari_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - intrari.xlsx")
        casa_de_marcat_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - vanzari casa de marcat.xlsx")
        facturi_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - vanzare cu factura.xlsx")
        transfer_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - transfer.xlsx")

        try:
            df_intrari = pd.read_excel(intrari_path)
        except FileNotFoundError:
            df_intrari = pd.DataFrame()

        try:
            df_casa_de_marcat = pd.read_excel(casa_de_marcat_path)
        except FileNotFoundError:
            df_casa_de_marcat = pd.DataFrame()

        try:
            df_facturi = pd.read_excel(facturi_path)
        except FileNotFoundError:
            df_facturi = pd.DataFrame()

        try:
            df_transfer = pd.read_excel(transfer_path)
        except FileNotFoundError:
            df_transfer = pd.DataFrame()

        # Extragem anul și luna din prima coloană "Data"
        year = df_centralizator['Data'].iloc[0].split('/')[-1]
        month = df_centralizator['Data'].iloc[0].split('/')[1]
        self.year = int(year)
        self.month = int(month)

        # Procesăm datele pentru fiecare categorie
        if not df_intrari.empty:
            for idx, row in df_intrari.iterrows():
                data = pd.to_datetime(row['Data'], format='%d/%m/%Y')
                if self.year == data.year and self.month == data.month:
                    df_centralizator.loc[df_centralizator['Data'] == data.strftime('%d/%m/%Y'), 'Intrari SGR'] += row['SGR']

        if not df_casa_de_marcat.empty:
            for idx, row in df_casa_de_marcat.iterrows():
                data = pd.to_datetime(row['Data'], format='%d/%m/%Y')
                if self.year == data.year and self.month == data.month:
                    df_centralizator.loc[df_centralizator['Data'] == data.strftime('%d/%m/%Y'), 'Vanzari SGR Casa Marcat'] += row['SGR']

        if not df_facturi.empty:
            for idx, row in df_facturi.iterrows():
                data = pd.to_datetime(row['Data'], format='%d/%m/%Y')
                if self.year == data.year and self.month == data.month:
                    df_centralizator.loc[df_centralizator['Data'] == data.strftime('%d/%m/%Y'), 'Vanzari SGR Factura'] += row['SGR']

        if not df_transfer.empty:
            for idx, row in df_transfer.iterrows():
                data = pd.to_datetime(row['Data'], format='%d/%m/%Y')
                if self.year == data.year and self.month == data.month:
                    df_centralizator.loc[df_centralizator['Data'] == data.strftime('%d/%m/%Y'), 'Transfer SGR'] += row['SGR']

        # Recalculăm soldul final pentru fiecare zi
        for i in range(len(df_centralizator)):
            if i == 0:
                df_centralizator.at[i, 'Sold Final'] = (
                    df_centralizator.at[i, 'Sold Initial']
                    + df_centralizator.at[i, 'Intrari SGR']
                    - df_centralizator.at[i, 'Vanzari SGR Casa Marcat']
                    - df_centralizator.at[i, 'Vanzari SGR Factura']
                    + df_centralizator.at[i, 'Transfer SGR']
                )
            else:
                df_centralizator.at[i, 'Sold Final'] = (
                    df_centralizator.at[i - 1, 'Sold Final']
                    + df_centralizator.at[i, 'Intrari SGR']
                    - df_centralizator.at[i, 'Vanzari SGR Casa Marcat']
                    - df_centralizator.at[i, 'Vanzari SGR Factura']
                    + df_centralizator.at[i, 'Transfer SGR']
                )
        # stergem randul TOTAL dac exista
        df_centralizator = df_centralizator[df_centralizator['Data'] != 'TOTAL']

        # Adăugăm un rând de totaluri
        total_row = pd.DataFrame({
            'Data': ['TOTAL'],
            'Intrari SGR': [df_centralizator['Intrari SGR'].sum()],
            'Vanzari SGR Casa Marcat': [df_centralizator['Vanzari SGR Casa Marcat'].sum()],
            'Vanzari SGR Factura': [df_centralizator['Vanzari SGR Factura'].sum()],
            'Transfer SGR': [df_centralizator['Transfer SGR'].sum()],
            'Sold Final': [df_centralizator['Sold Final'].iloc[-1]],
            'Sold Initial': [''],
            'Numele si Anul': ['']
        })


        year = df_centralizator['Data'].iloc[0].split('/')[-1]
        month = df_centralizator['Data'].iloc[0].split('/')[1]
        how_many_days_are_in_that_month = calendar.monthrange(int(year), int(month))[1]
        row = how_many_days_are_in_that_month
        df_centralizator.loc[row] = total_row.iloc[0]

        # Scriem în fișierul Excel
        writer = pd.ExcelWriter(centralizator_path, engine='xlsxwriter')
        df_centralizator.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
                "bg_color": "#D7E4BC",  # Verde deschis
            }
        )
        total_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#FFDDDD",  # Roz
                "border": 1,
            }
        )

        # Setăm lățimea coloanelor și înălțimea rândurilor
        column_widths = {
            "Data": 10,
            "Intrari SGR": 13,
            "Vanzari SGR Casa Marcat": 23,
            "Vanzari SGR Factura": 19,
            "Transfer SGR": 12,
            "Sold Final": 10,
            "Sold Initial": 12,
            "Numele si Anul": 20
        }

        for col_num, (col_name, width) in enumerate(column_widths.items()):
            worksheet.set_column(col_num, col_num, width)
            worksheet.write(0, col_num, col_name, header_format)

        # Formatăm rândul total
        total_row_index = len(df_centralizator) - 1
        for col_num in range(6):
            worksheet.write(total_row_index + 1, col_num, df_centralizator.iloc[total_row_index, col_num], total_format)

        writer.close()

        QMessageBox.information(self, "Succes", "Centralizatorul a fost actualizat.")
        self.close()




class CentralizatorPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.sold_initial_value = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Centralizator" + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        self.layout = QGridLayout()
        self.setLayout(self.layout)

        # Buton pentru setare sold initial
        self.btn_sold_initial = QPushButton("Setare sold initial", self)
        self.btn_sold_initial.clicked.connect(self.open_sold_initial)
        self.layout.addWidget(self.btn_sold_initial, 0, 0)

        # Buton pentru refresh centralizator
        self.btn_refresh_centralizator = QPushButton("Refresh Centralizator", self)
        self.btn_refresh_centralizator.clicked.connect(self.refresh_centralizator)
        self.layout.addWidget(self.btn_refresh_centralizator, 0, 1)
        
        # Back button
        self.btn_back = QPushButton("Back", self)
        self.btn_back.clicked.connect(self.close)
        self.layout.addWidget(self.btn_back, 0, 2)

    def open_sold_initial(self):
        self.sold_initial_window = QWidget()
        self.sold_initial_window.setWindowTitle("Setare Sold Initial")
        self.sold_initial_window.setGeometry(150, 150, 400, 200)

        layout = QGridLayout()
        self.sold_initial_window.setLayout(layout)

        layout.addWidget(QLabel("Sold Initial"), 0, 0)
        self.sold_initial_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.sold_initial_input, 0, 1)

        layout.addWidget(QLabel("Luna (1-12)"), 1, 0)
        self.month_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.month_input, 1, 1)

        layout.addWidget(QLabel("Anul"), 2, 0)
        self.year_input = QLineEdit(self.sold_initial_window)
        layout.addWidget(self.year_input, 2, 1)

        btn_save = QPushButton("Save", self.sold_initial_window)
        btn_save.clicked.connect(self.save_sold_initial)
        layout.addWidget(btn_save, 3, 0)

        btn_cancel = QPushButton("Cancel", self.sold_initial_window)
        btn_cancel.clicked.connect(self.sold_initial_window.close)
        layout.addWidget(btn_cancel, 3, 1)

        self.sold_initial_window.show()

    
    def format_header(self, file_path):
        # Încarcă fișierul Excel existent
        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            # Dacă fișierul nu există, creează unul nou
            QMessageBox.warning(self, "Eroare", "Fișierul nu există pentru a putea fi formatat.")
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
         
        # Definirea formatarilor dorite pentru header
        header_format = workbook.add_format(
            {
                "bold": True,
                "valign": "top",
                "fg_color": "#D7E4BC",
                "border": 1,
            }
        )
        
        column_widths = {
            "Data": 10,
            "Intrari": 10,
            "Vanzari Numerar": 15,
            "Card": 10,
            "Facturi Op": 10,
            "Reduceri": 10,
            "Total Vanzari": 13,
            "Sold": 10,
            "Sold Initial": 12,
            "Numele si Anul": 20
        }

        for col_num, (col_name, width) in enumerate(column_widths.items()):
            worksheet.set_column(col_num, col_num, width)
            worksheet.write(0, col_num, col_name, header_format)
            
        # Pt finalul listei
        border_format = workbook.add_format({
            'border': 1,
            'bg_color': '#FFDDDD',
            "text_wrap": True,
            "valign": "top"
        })

        total_row_index = len(df) - 1
        for col_num in range(7):
            worksheet.write(total_row_index + 1, col_num, df.iloc[total_row_index, col_num], border_format)

        writer.close()
   
   
   
   
   
    
    def completeaza_centralizator(self):
        centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator.xlsx")

        if not os.path.exists(centralizator_path):
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator nu există.")
            return

        # Citim fișierul centralizator
        try:
            df_centralizator = pd.read_excel(centralizator_path)
        except FileNotFoundError:
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator nu a putut fi citit.")
            return

        # Convertim coloanele la float pentru a evita problemele de compatibilitate
        float_columns = ['Intrari', 'Vanzari Numerar', 'Card', 'Facturi Op', 'Reduceri', 'Total Vanzari', 'Sold']
        for col in float_columns:
            df_centralizator[col] = df_centralizator[col].astype(float)

        # Resetăm coloanele care vor fi recalculată
        df_centralizator[float_columns] = 0.0

        # Citim fișierele de date
        intrari_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - intrari.xlsx")
        casa_de_marcat_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - vanzari casa de marcat.xlsx")
        facturi_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - vanzare cu factura.xlsx")
        transfer_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - transfer.xlsx")

        try:
            df_intrari = pd.read_excel(intrari_path)
        except FileNotFoundError:
            df_intrari = pd.DataFrame()

        try:
            df_casa_de_marcat = pd.read_excel(casa_de_marcat_path)
        except FileNotFoundError:
            df_casa_de_marcat = pd.DataFrame()

        try:
            df_facturi = pd.read_excel(facturi_path)
        except FileNotFoundError:
            df_facturi = pd.DataFrame()

        try:
            df_transfer = pd.read_excel(transfer_path)
        except FileNotFoundError:
            df_transfer = pd.DataFrame()

        # Completăm coloanele din centralizator
        for i, row in df_centralizator.iterrows():
            date_str = row['Data']

            if not df_intrari.empty and date_str in df_intrari['Data'].values:
                df_centralizator.at[i, 'Intrari'] += df_intrari.loc[df_intrari['Data'] == date_str, 'Val cu TVA 19%'].sum()

            if not df_transfer.empty and date_str in df_transfer['Data'].values:
                df_centralizator.at[i, 'Intrari'] += df_transfer.loc[df_transfer['Data'] == date_str, 'Val cu TVA'].sum()

            if not df_casa_de_marcat.empty and date_str in df_casa_de_marcat['Data'].values:
                df_centralizator.at[i, 'Vanzari Numerar'] = df_casa_de_marcat.loc[df_casa_de_marcat['Data'] == date_str, 'Numerar'].sum()
                df_centralizator.at[i, 'Card'] = df_casa_de_marcat.loc[df_casa_de_marcat['Data'] == date_str, 'Card'].sum()
                df_centralizator.at[i, 'Reduceri'] += df_casa_de_marcat.loc[df_casa_de_marcat['Data'] == date_str, 'Reducere/Discount'].sum()

            if not df_facturi.empty and date_str in df_facturi['Data'].values:
                df_centralizator.at[i, 'Facturi Op'] = df_facturi.loc[df_facturi['Data'] == date_str, 'Val cu TVA'].sum()
                df_centralizator.at[i, 'Reduceri'] += df_facturi.loc[df_facturi['Data'] == date_str, 'Reducere/Discount'].sum()

            df_centralizator.at[i, 'Total Vanzari'] = df_centralizator.at[i, 'Vanzari Numerar'] + df_centralizator.at[i, 'Card'] + df_centralizator.at[i, 'Facturi Op'] + np.fabs(df_centralizator.at[i, 'Reduceri'])

            if i == 0:
                df_centralizator.at[i, 'Sold'] = df_centralizator['Sold Initial'].iloc[0] + df_centralizator.at[i, 'Intrari'] - df_centralizator.at[i, 'Total Vanzari']
            else:
                df_centralizator.at[i, 'Sold'] = df_centralizator.at[i-1, 'Sold'] + df_centralizator.at[i, 'Intrari'] - df_centralizator.at[i, 'Total Vanzari']

        # Ștergem rândul TOTAL dacă există
        df_centralizator = df_centralizator[df_centralizator['Data'] != 'TOTAL']

        # Adăugăm rândul TOTAL
        total_row = pd.DataFrame([{
            "Data": "TOTAL",
            "Intrari": df_centralizator['Intrari'].sum(),
            "Vanzari Numerar": df_centralizator['Vanzari Numerar'].sum(),
            "Card": df_centralizator['Card'].sum(),
            "Facturi Op": df_centralizator['Facturi Op'].sum(),
            "Reduceri": df_centralizator['Reduceri'].sum(),
            "Total Vanzari": df_centralizator['Total Vanzari'].sum(),
            "Sold": ""  # Câmpul "Sold" nu se adună
        }])

        year = df_centralizator['Data'].iloc[0].split('/')[-1]
        month = df_centralizator['Data'].iloc[0].split('/')[1]
        how_many_days_are_in_that_month = calendar.monthrange(int(year), int(month))[1]
        row = how_many_days_are_in_that_month
        df_centralizator.loc[row] = total_row.iloc[0]

        # Salvăm fișierul actualizat
        df_centralizator.to_excel(centralizator_path, index=False)
        self.format_header(centralizator_path)
        QMessageBox.information(self, "Succes", f"Fișierul centralizator a fost populat!")
        self.close()




    def save_sold_initial(self):
        try:
            self.sold_initial_value = float(self.sold_initial_input.text())
            self.month = int(self.month_input.text())
            self.year = int(self.year_input.text())
        except ValueError:
            QMessageBox.warning(self, "Eroare", "Vă rugăm să introduceți valori numerice valide pentru sold, lună și an.")
            return

        if self.sold_initial_value < 0:
            QMessageBox.warning(self, "Eroare", "Soldul inițial nu poate fi negativ.")
            return

        centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator.xlsx")

        # Generăm toate zilele lunii respective
        start_date = datetime(self.year, self.month, 1)
        end_date = (start_date + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        dates = pd.date_range(start_date, end_date)

        data = {
            "Data": dates.strftime('%d/%m/%Y'),
            "Intrari": [0] * len(dates),
            "Vanzari Numerar": [0] * len(dates),
            "Card": [0] * len(dates),
            "Facturi Op": [0] * len(dates),
            "Reduceri": [0] * len(dates),
            "Total Vanzari": [0] * len(dates),
            "Sold": [0] * len(dates),
            "Sold Initial": [self.sold_initial_value] * len(dates),
            "Numele si Anul": f"{os.path.basename(self.folder_path)}"
        }

        df = pd.DataFrame(data)
        writer = pd.ExcelWriter(centralizator_path, engine='xlsxwriter')

        # Salvăm în fișierul Excel
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format(
            {
                "bold": True,
                "valign": "top",
                "border": 1,
            }
        )

        # Setăm lățimea coloanelor și înălțimea rândurilor
        column_widths = {
            "Data": 10,
            "Intrari": 10,
            "Vanzari Numerar": 15,
            "Card": 10,
            "Facturi Op": 10,
            "Reduceri": 10,
            "Total Vanzari": 13,
            "Sold": 10,
            "Sold Initial": 12,
            "Numele si Anul": 20
        }
        for col_num, (col_name, width) in enumerate(column_widths.items()):
            worksheet.set_column(col_num, col_num, width)
            worksheet.write(0, col_num, col_name, header_format)

        # Ajustăm înălțimea rândului pentru header
        worksheet.set_row(0, 30)  # Ajustează înălțimea după cum este necesar

        writer.close()

        QMessageBox.information(self, "Succes", f"Fișierul centralizator a fost creat/actualizat")
        self.completeaza_centralizator()
        self.sold_initial_window.close()
        








    def refresh_centralizator(self):
      centralizator_path = os.path.join(self.folder_path, f"{os.path.basename(self.folder_path)} - centralizator.xlsx")

      if not os.path.exists(centralizator_path):
          QMessageBox.warning(self, "Eroare", "Fișierul centralizator nu există.")
          return
      if pd.read_excel(centralizator_path).empty:
            QMessageBox.warning(self, "Eroare", "Fișierul centralizator este gol.")
            return
    #   try:
          # Citim fișierul Excel specificând rândul de cap pentru tabel
      self.completeaza_centralizator() 
    #   except Exception as e:
    #         QMessageBox.warning(self, "Eroare", f"Fișierul centralizator nu a putut fi actualizat: {e}")






class IntrariPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Intrari" + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        layout = QGridLayout()

        layout.addWidget(QLabel("Data (dd/mm/yyyy)"), 0, 0)
        self.entry_data = QLineEdit(self)
        layout.addWidget(self.entry_data, 0, 1)

        layout.addWidget(QLabel("Nr.Doc"), 1, 0)
        self.entry_nr_doc = QLineEdit(self)
        layout.addWidget(self.entry_nr_doc, 1, 1)

        layout.addWidget(QLabel("Val cu TVA 19%"), 2, 0)
        self.entry_val_tva = QLineEdit(self)
        layout.addWidget(self.entry_val_tva, 2, 1)

        layout.addWidget(QLabel("Pret. Cost"), 3, 0)
        self.entry_pred_cost = QLineEdit(self)
        layout.addWidget(self.entry_pred_cost, 3, 1)

        layout.addWidget(QLabel("SGR"), 4, 0)
        self.entry_sgr = QLineEdit(self)
        layout.addWidget(self.entry_sgr, 4, 1)

        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.add_intrare)
        layout.addWidget(self.btn_ok, 5, 1)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.clicked.connect(self.close)
        layout.addWidget(self.btn_cancel, 5, 0)

        self.setLayout(layout)

    def add_intrare(self):
        try:
            # Verificare format data
            datetime.strptime(self.entry_data.text(), "%d/%m/%Y")
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Data nu este în formatul corect (dd/mm/yyyy)!")
            return

        try:
            # Conversie valoare TVA și costuri la float
            val_tva = float(self.entry_val_tva.text())
            pred_cost = float(self.entry_pred_cost.text())
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Valorile TVA și cost trebuie să fie numere!")
            return

        try:
            # Conversie Nr.Doc la numpy.int64
            nr_doc = np.int64(int(self.entry_nr_doc.text()))
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Numărul documentului trebuie să fie un număr valid!")
            return

        tva = round(val_tva / 1.19 * 0.19, 2)
        adaos = round(val_tva - pred_cost - tva, 2)

        new_data = {
            "Data": self.entry_data.text(),
            "Nr.Doc": nr_doc,
            "Val cu TVA 19%": val_tva,
            "Pret. Cost": pred_cost,
            "TVA 19%": tva,
            "Adaos": adaos,
            "SGR": self.entry_sgr.text()
        }

        file_name = os.path.basename(self.folder_path) + " - intrari.xlsx"
        file_path = os.path.join(self.folder_path, file_name)

        df = pd.read_excel(file_path)

        # Verificăm existența numărului documentului
        if "Nr.Doc" in df.columns and nr_doc in df["Nr.Doc"].values:
            QMessageBox.critical(self, "Eroare", "Numărul documentului există deja!")
            return

        # Conversie și filtrare date valide
        if not df.empty: 
            valid_dates = []
            for date in df["Data"]:
                try:
                    valid_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d/%m/%Y")
                    valid_dates.append(valid_date)
                except Exception:
                    valid_dates.append(pd.NaT)

            df["Data"] = valid_dates
            new_data_df = pd.DataFrame([new_data])
            new_data_df["Data"] = pd.to_datetime(new_data_df["Data"], format="%d/%m/%Y").dt.strftime("%d/%m/%Y")

            df = pd.concat([df, new_data_df], ignore_index=True)

            # Sortăm după dată
            df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
            df = df.sort_values(by="Data")
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)
            writer.close()
        else:
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)
            writer.close()

        QMessageBox.information(self, "Succes", "Intrarea a fost adăugată cu succes!")
        self.close()





 
class VanzareFacturaPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Vanzare cu Factura" + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        layout = QGridLayout()

        layout.addWidget(QLabel("Data (dd/mm/yyyy)"), 0, 0)
        self.entry_data = QLineEdit(self)
        layout.addWidget(self.entry_data, 0, 1)

        layout.addWidget(QLabel("Nr"), 1, 0)
        self.entry_nr = QLineEdit(self)
        layout.addWidget(self.entry_nr, 1, 1)

        layout.addWidget(QLabel("Val cu TVA"), 2, 0)
        self.entry_val_tva = QLineEdit(self)
        layout.addWidget(self.entry_val_tva, 2, 1)

        layout.addWidget(QLabel("Valoare Gestiuni sau Aviz"), 3, 0)
        self.entry_val_gestiuni = QLineEdit(self)
        layout.addWidget(self.entry_val_gestiuni, 3, 1)

        layout.addWidget(QLabel("SGR"), 4, 0)
        self.entry_sgr = QLineEdit(self)
        layout.addWidget(self.entry_sgr, 4, 1)

        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.add_vanzare_factura)
        layout.addWidget(self.btn_ok, 5, 1)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.clicked.connect(self.close)
        layout.addWidget(self.btn_cancel, 5, 0)

        self.setLayout(layout)

    def add_vanzare_factura(self):
       try:
           # Verificare format dată
           datetime.strptime(self.entry_data.text(), "%d/%m/%Y")
       except ValueError:
           QMessageBox.critical(self, "Eroare", "Data nu este în formatul corect (dd/mm/yyyy)!")
           return

       try:
           # Conversie valoare TVA și valoare gestiuni la float
           val_tva = float(self.entry_val_tva.text())
           val_gestiuni = float(self.entry_val_gestiuni.text())
       except ValueError:
           QMessageBox.critical(self, "Eroare", "Valorile TVA și valoare gestiuni trebuie să fie numere!")
           return

       try:
           # Conversie Nr la numpy.int64
           nr = np.int64(int(self.entry_nr.text()))
       except ValueError:
           QMessageBox.critical(self, "Eroare", "Numărul trebuie să fie un număr valid!")
           return

       dif_de_scazut = round(val_tva - val_gestiuni, 2)
       tva = round(dif_de_scazut / 1.19 * 0.19, 2)
       adaos = round(dif_de_scazut - tva, 2)

       new_data = {
           "Data": self.entry_data.text(),
           "Nr": nr,
           "Val cu TVA": val_tva,
           "Valoare Gestiuni sau Aviz": val_gestiuni,
           "Reducere/Discount": dif_de_scazut,
           "TVA": tva,
           "Adaos": adaos,
           "SGR": self.entry_sgr.text()
       }

       file_name = os.path.basename(self.folder_path) + " - vanzare cu factura.xlsx"
       file_path = os.path.join(self.folder_path, file_name)

       df = pd.read_excel(file_path)

       # Verificăm existența numărului
       if "Nr" in df.columns and nr in df["Nr"].values:
           QMessageBox.critical(self, "Eroare", "Numărul există deja!")
           return

       # Conversie și filtrare date valide
       if not df.empty:
           valid_dates = []
           for date in df["Data"]:
               try:
                   valid_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d/%m/%Y")
                   valid_dates.append(valid_date)
               except Exception:
                   valid_dates.append(pd.NaT)

           df["Data"] = valid_dates
           new_data_df = pd.DataFrame([new_data])
           new_data_df["Data"] = pd.to_datetime(new_data_df["Data"], format="%d/%m/%Y").dt.strftime("%d/%m/%Y")

           df = pd.concat([df, new_data_df], ignore_index=True)

           # Sortăm după dată
           df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
           df = df.sort_values(by="Data")
           df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

           # Aplicăm formatarea headerului și scriem în fișierul Excel
           writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
           df.to_excel(writer, index=False, sheet_name='Sheet1')
           workbook = writer.book
           worksheet = writer.sheets['Sheet1']

           header_format = workbook.add_format({
               "bold": True,
               "text_wrap": True,
               "valign": "top",
               "border": 1,
           })

           for col_num, value in enumerate(df.columns.values):
               worksheet.write(0, col_num, value, header_format)
               worksheet.set_column(col_num, col_num, 15)
           writer.close()
       else:
           df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

           # Aplicăm formatarea headerului și scriem în fișierul Excel
           writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
           df.to_excel(writer, index=False, sheet_name='Sheet1')
           workbook = writer.book
           worksheet = writer.sheets['Sheet1']

           header_format = workbook.add_format({
               "bold": True,
               "text_wrap": True,
               "valign": "top",
               "border": 1,
           })

           for col_num, value in enumerate(df.columns.values):
               worksheet.write(0, col_num, value, header_format)
               worksheet.set_column(col_num, col_num, 15)
           writer.close()

       QMessageBox.information(self, "Succes", "Intrarea a fost adăugată cu succes!")
       self.close()


class VanzareCasaMarcatPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Vanzare Casa Marcat" + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        layout = QGridLayout()

        layout.addWidget(QLabel("Data (dd/mm/yyyy)"), 0, 0)
        self.entry_data = QLineEdit(self)
        layout.addWidget(self.entry_data, 0, 1)

        layout.addWidget(QLabel("Nr"), 1, 0)
        self.entry_nr = QLineEdit(self)
        layout.addWidget(self.entry_nr, 1, 1)

        layout.addWidget(QLabel("Val cu TVA"), 2, 0)
        self.entry_val_tva = QLineEdit(self)
        layout.addWidget(self.entry_val_tva, 2, 1)

        layout.addWidget(QLabel("Numerar"), 3, 0)
        self.entry_numerar = QLineEdit(self)
        layout.addWidget(self.entry_numerar, 3, 1)

        layout.addWidget(QLabel("Card"), 4, 0)
        self.entry_card = QLineEdit(self)
        layout.addWidget(self.entry_card, 4, 1)

        layout.addWidget(QLabel("Reducere/Discount"), 5, 0)
        self.entry_discount = QLineEdit(self)
        layout.addWidget(self.entry_discount, 5, 1)

        layout.addWidget(QLabel("SGR"), 6, 0)
        self.entry_sgr = QLineEdit(self)
        layout.addWidget(self.entry_sgr, 6, 1)

        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.add_vanzare_casa_marcat)
        layout.addWidget(self.btn_ok, 7, 1)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.clicked.connect(self.close)
        layout.addWidget(self.btn_cancel, 7, 0)

        self.setLayout(layout)

    def add_vanzare_casa_marcat(self):
        try:
            # Verificare format dată
            datetime.strptime(self.entry_data.text(), "%d/%m/%Y")
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Data nu este în formatul corect (dd/mm/yyyy)!")
            return

        try:
            # Conversie valorilor la float
            val_tva = float(self.entry_val_tva.text())
            numerar = float(self.entry_numerar.text())
            card = float(self.entry_card.text())
            discount = float(self.entry_discount.text())
            if discount > 0:
                QMessageBox.critical(self, "Eroare", "Reducerea nu poate fi pozitiva!")
                return
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Valorile trebuie să fie numere!")
            return

        try:
            # Conversie Nr la numpy.int64
            nr = np.int64(int(self.entry_nr.text()))
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Numărul trebuie să fie un număr valid!")
            return

        new_data = {
            "Data": self.entry_data.text(),
            "Nr": nr,
            "Val cu TVA": val_tva,
            "Numerar": numerar,
            "Card": card,
            "Reducere/Discount": discount,
            "SGR": self.entry_sgr.text()
        }

        file_name = os.path.basename(self.folder_path) + " - vanzari casa de marcat.xlsx"
        file_path = os.path.join(self.folder_path, file_name)

        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            df = pd.DataFrame()

        # Verificăm existența numărului
        if "Nr" in df.columns and nr in df["Nr"].values:
            QMessageBox.critical(self, "Eroare", "Numărul există deja!")
            return

        # Conversie și filtrare date valide
        if not df.empty:
            valid_dates = []
            for date in df["Data"]:
                try:
                    valid_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d/%m/%Y")
                    valid_dates.append(valid_date)
                except Exception:
                    valid_dates.append(pd.NaT)

            df["Data"] = valid_dates
            new_data_df = pd.DataFrame([new_data])
            new_data_df["Data"] = pd.to_datetime(new_data_df["Data"], format="%d/%m/%Y").dt.strftime("%d/%m/%Y")

            df = pd.concat([df, new_data_df], ignore_index=True)

            # Sortăm după dată
            df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
            df = df.sort_values(by="Data")
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            writer.close()
        else:
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)
            writer.close()

        QMessageBox.information(self, "Succes", "Intrarea a fost adăugată cu succes!")
        self.close()





class TransferIntreGestPage(QWidget):
    def __init__(self, folder_path, excel_data):
        super().__init__()
        self.folder_path = folder_path
        self.excel_data = excel_data
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Transfer intre Gest" + os.path.basename(self.folder_path))
        self.setGeometry(100, 100, 800, 600)

        layout = QGridLayout()

        layout.addWidget(QLabel("Data (dd/mm/yyyy)"), 0, 0)
        self.entry_data = QLineEdit(self)
        layout.addWidget(self.entry_data, 0, 1)

        layout.addWidget(QLabel("Nr"), 1, 0)
        self.entry_nr = QLineEdit(self)
        layout.addWidget(self.entry_nr, 1, 1)

        layout.addWidget(QLabel("Val cu TVA"), 2, 0)
        self.entry_val_tva = QLineEdit(self)
        layout.addWidget(self.entry_val_tva, 2, 1)

        layout.addWidget(QLabel("Pret de Cost"), 3, 0)
        self.entry_pret_cost = QLineEdit(self)
        layout.addWidget(self.entry_pret_cost, 3, 1)

        layout.addWidget(QLabel("SGR"), 4, 0)
        self.entry_sgr = QLineEdit(self)
        layout.addWidget(self.entry_sgr, 4, 1)

        self.btn_ok = QPushButton("OK", self)
        self.btn_ok.clicked.connect(self.add_transfer_intre_gest)
        layout.addWidget(self.btn_ok, 5, 1)

        self.btn_cancel = QPushButton("Cancel", self)
        self.btn_cancel.clicked.connect(self.close)
        layout.addWidget(self.btn_cancel, 5, 0)

        self.setLayout(layout)

    def add_transfer_intre_gest(self):
        try:
            # Verificare format dată
            datetime.strptime(self.entry_data.text(), "%d/%m/%Y")
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Data nu este în formatul corect (dd/mm/yyyy)!")
            return

        try:
            # Conversie valorilor la float
            val_tva = float(self.entry_val_tva.text())
            pret_cost = float(self.entry_pret_cost.text())
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Valorile trebuie să fie numere!")
            return

        try:
            # Conversie Nr la numpy.int64
            nr = np.int64(int(self.entry_nr.text()))
        except ValueError:
            QMessageBox.critical(self, "Eroare", "Numărul trebuie să fie un număr valid!")
            return

        tva = round(val_tva / 1.19 * 0.19, 2)
        adaos = round(val_tva - pret_cost - tva, 2)

        new_data = {
            "Data": self.entry_data.text(),
            "Nr": nr,
            "Val cu TVA": val_tva,
            "Pret de Cost": pret_cost,
            "TVA": tva,
            "Adaos": adaos,
            "SGR": self.entry_sgr.text()
        }

        file_name = os.path.basename(self.folder_path) + " - transfer.xlsx"
        file_path = os.path.join(self.folder_path, file_name)

        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            df = pd.DataFrame()

        # Verificăm existența numărului
        if "Nr" in df.columns and nr in df["Nr"].values:
            QMessageBox.critical(self, "Eroare", "Numărul există deja!")
            return

        # Conversie și filtrare date valide
        if not df.empty:
            valid_dates = []
            for date in df["Data"]:
                try:
                    valid_date = datetime.strptime(date, "%d/%m/%Y").strftime("%d/%m/%Y")
                    valid_dates.append(valid_date)
                except Exception:
                    valid_dates.append(pd.NaT)

            df["Data"] = valid_dates
            new_data_df = pd.DataFrame([new_data])
            new_data_df["Data"] = pd.to_datetime(new_data_df["Data"], format="%d/%m/%Y").dt.strftime("%d/%m/%Y")

            df = pd.concat([df, new_data_df], ignore_index=True)

            # Sortăm după dată
            df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y")
            df = df.sort_values(by="Data")
            df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            writer.close()
        else:
            df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)

            # Aplicăm formatarea headerului și scriem în fișierul Excel
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                "bold": True,
                "text_wrap": True,
                "valign": "top",
                "border": 1,
            })

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)

            writer.close()

        QMessageBox.information(self, "Succes", "Intrarea a fost adăugată cu succes!")
        self.close()



if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())