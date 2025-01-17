import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QLabel, QSpinBox, QCheckBox, QPushButton,
                            QFileDialog, QTreeWidget, QTreeWidgetItem, QProgressBar,
                            QMessageBox, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QPixmap
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import urllib.parse
from pathlib import Path
import json
import schedule

PASTA_ICONES = Path(__file__).parent / "icons"

class SendMessagesThread(QThread):
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    finished_signal = pyqtSignal(int)

    def __init__(self, df, interval, driver_setup):
        super().__init__()
        self.df = df
        self.interval = interval
        self.driver_setup = driver_setup

    def run(self):
        driver = self.driver_setup()
        if not driver:
            self.finished_signal.emit(0)
            return

        messages_sent = 0
        total_messages = len(self.df)

        try:
            for index, row in self.df.iterrows():
                phone = str(row['telefone'])
                message = str(row['mensagem'])

                self.status_updated.emit(f"Enviando mensagem para {phone}...")

                if self.send_whatsapp_message(driver, phone, message):
                    messages_sent += 1
                    progress = int((messages_sent / total_messages) * 100)
                    self.progress_updated.emit(progress)
                    self.status_updated.emit(f"✅ Mensagem enviada com sucesso para {phone}")
                else:
                    self.status_updated.emit(f"❌ Erro ao enviar mensagem para {phone}")

                time.sleep(self.interval)

        finally:
            driver.quit()
            self.finished_signal.emit(messages_sent)

    def send_whatsapp_message(self, driver, phone, message):
        phone = ''.join(filter(str.isdigit, phone))
        if len(phone) == 11:
            phone = '55' + phone

        message_encoded = urllib.parse.quote(message)
        url = f'https://web.whatsapp.com/send?phone={phone}&text={message_encoded}'

        try:
            driver.get(url)
            message_field = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@aria-placeholder="Digite uma mensagem"]'))
            )
            message_field.send_keys(Keys.ENTER)
            time.sleep(3)
            return True
        except Exception as e:
            return False

    def send_whatsapp_media(self, driver, phone, media_path):
        phone = ''.join(filter(str.isdigit, phone))
        if len(phone) == 11:
            phone = '55' + phone

        driver.get(f'https://web.whatsapp.com/send?phone={phone}')
        time.sleep(5)  

        attach_button = driver.find_element(By.XPATH, '//div[@title="Anexar"]')
        attach_button.click()
        time.sleep(1)

        file_input = driver.find_element(By.XPATH, '//input[@type="file"]')
        file_input.send_keys(media_path)
        time.sleep(2)

        send_button = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_button.click()

class WhatsAppSender(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.initUI()
        self.apply_styles()

    def initUI(self):
        self.setWindowTitle("WhatsApp Message Sender")
        self.setMinimumSize(1000, 800)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        title_image = QLabel()
        title_image.setPixmap(QPixmap(str(PASTA_ICONES / "modal2.png")))
        title_image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_image, alignment=Qt.AlignmentFlag.AlignCenter)

        settings_frame = QFrame()
        settings_frame.setObjectName("settingsFrame")
        settings_layout = QVBoxLayout(settings_frame)

        interval_layout = QHBoxLayout()
        interval_label = QLabel("Intervalo entre mensagens (segundos):")
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(10, 120)
        self.interval_spin.setValue(30)
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spin)
        interval_layout.addStretch()
        settings_layout.addLayout(interval_layout)

        self.preview_check = QCheckBox("Mostrar preview de mensagens")
        self.preview_check.setChecked(True)
        settings_layout.addWidget(self.preview_check)

        layout.addWidget(settings_frame)

        buttons_layout = QHBoxLayout()
        self.download_btn = QPushButton("📥 Download Template")
        self.load_btn = QPushButton("📂 Carregar Arquivo")
        buttons_layout.addWidget(self.download_btn)
        buttons_layout.addWidget(self.load_btn)
        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(['Telefone', 'Mensagem'])
        self.tree.setColumnWidth(0, 200)
        self.tree.setColumnWidth(1, 500)
        layout.addWidget(self.tree)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.status_label = QLabel()
        self.status_label.setObjectName("statusLabel")
        layout.addWidget(self.status_label)

        self.send_button = QPushButton("📤 Enviar Mensagens")
        self.send_button.setObjectName("sendButton")
        layout.addWidget(self.send_button)

        credits = QLabel("💜 PIX: hugorogerio522@gmail.com")
        credits.setObjectName("credits")
        layout.addWidget(credits, alignment=Qt.AlignmentFlag.AlignCenter)

        self.download_btn.clicked.connect(self.download_template)
        self.load_btn.clicked.connect(self.load_file)
        self.send_button.clicked.connect(self.start_sending)

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f6fa;
            }

            #settingsFrame {
                background-color: white;
                border-radius: 10px;
                padding: 20px;
                box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            }

            QPushButton {
                background-color: #25D366;  /* Cor verde do WhatsApp */
                color: white;
                border: none;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #21c354;
            }

            QPushButton:pressed {
                background-color: #1da147;
            }

            #sendButton {
                background-color: #25D366;  /* Cor verde do WhatsApp */
                font-size: 16px;
                padding: 15px;
            }

            QTreeWidget {
                border: 1px solid #dcdde1;
                border-radius: 5px;
                background-color: white;
                padding: 10px;
            }

            QTreeWidget::item {
                padding: 5px;
            }

            QProgressBar {
                border: none;
                border-radius: 5px;
                background-color: #dcdde1;
                height: 10px;
                text-align: center;
            }

            QProgressBar::chunk {
                background-color: #25D366;  /* Cor verde do WhatsApp */
                border-radius: 5px;
            }

            QSpinBox {
                padding: 5px;
                border: 1px solid #dcdde1;
                border-radius: 5px;
                background-color: white;
            }

            QLabel {
                color: #2d3436;
            }

            #statusLabel {
                color: #576574;
                font-size: 14px;
            }

            #credits {
                color: #25D366;  /* Cor verde do WhatsApp */
                font-size: 14px;
                margin-top: 10px;
            }

            QCheckBox {
                spacing: 8px;
            }

            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #dcdde1;
            }

            QCheckBox::indicator:checked {
                background-color: #25D366;  /* Cor verde do WhatsApp */
                border-color: #25D366;  /* Cor verde do WhatsApp */
            }
        """)

        self.setWindowIcon(QIcon(str(PASTA_ICONES / "modal1.png")))

    def download_template(self):
        df = pd.DataFrame({
            'telefone': ['+5511999999999', '11999999999'],
            'mensagem': ['Exemplo de mensagem 1', 'Exemplo de mensagem 2']
        })

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Template",
            "template_whatsapp.xlsx",
            "Excel files (*.xlsx)"
        )

        if file_path:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Sucesso", "Template baixado com sucesso!")

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Abrir Arquivo",
            "",
            "Excel files (*.xlsx)"
        )

        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                if 'telefone' not in self.df.columns or 'mensagem' not in self.df.columns:
                    raise ValueError("O arquivo deve conter as colunas 'telefone' e 'mensagem'")

                self.update_preview()
                QMessageBox.information(self, "Sucesso", "Arquivo carregado com sucesso!")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao processar o arquivo: {str(e)}")

    def update_preview(self):
        self.tree.clear()
        for _, row in self.df.iterrows():
            item = QTreeWidgetItem([str(row['telefone']), str(row['mensagem'])])
            self.tree.addTopLevelItem(item)

    def setup_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument("--remote-debugging-port=9222")
        chrome_options.add_argument('--disable-features=VizDisplay')
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--proxy-server='direct://'")
        chrome_options.add_argument("--proxy-bypass-list=*")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--ignore-certificate-errors')

        try:
            service = Service(ChromeDriverManager().install())
            return webdriver.Chrome(service=service, options=chrome_options)
        except Exception as e:
            QMessageBox.critical(self, "Erro",
                               f"Erro ao inicializar o Chrome: {str(e)}\n"
                               "Verifique se o Google Chrome está instalado corretamente.")
            return None

    def start_sending(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "Erro", "Por favor, carregue um arquivo com dados primeiro!")
            return

        self.send_button.setEnabled(False)
        self.progress.setValue(0)

        self.thread = SendMessagesThread(self.df, self.interval_spin.value(), self.setup_driver)
        self.thread.progress_updated.connect(self.progress.setValue)
        self.thread.status_updated.connect(self.status_label.setText)
        self.thread.finished_signal.connect(self.sending_finished)
        self.thread.start()

    def sending_finished(self, messages_sent):
        self.send_button.setEnabled(True)
        if messages_sent > 0:
            QMessageBox.information(self, "Sucesso",
                                  f"✨ Processo concluído! {messages_sent} mensagens enviadas.")
        else:
            QMessageBox.critical(self, "Erro", "❌ Nenhuma mensagem foi enviada com sucesso.")

    def save_settings(self):
        settings = {
            'interval': self.interval_spin.value(),
            'preview': self.preview_check.isChecked()
        }
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Configurações", "", "JSON files (*.json)")
        if file_path:
            with open(file_path, 'w') as f:
                json.dump(settings, f)
            QMessageBox.information(self, "Sucesso", "Configurações salvas com sucesso!")

    def load_settings(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Carregar Configurações", "", "JSON files (*.json)")
        if file_path:
            with open(file_path, 'r') as f:
                settings = json.load(f)
                self.interval_spin.setValue(settings.get('interval', 30))
                self.preview_check.setChecked(settings.get('preview', True))
            QMessageBox.information(self, "Sucesso", "Configurações carregadas com sucesso!")

    def log_message(self, phone, message, status):
        with open("log.txt", "a") as log_file:
            log_file.write(f"{datetime.now()}: {status} - {phone}: {message}\n")

    def schedule_message(self, phone, message, time):
        schedule.every().day.at(time).do(self.send_whatsapp_message, phone, message)

    def save_template(self):
        template = self.template_input.text()  
        with open("templates.json", "a") as f:
            json.dump(template, f)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = WhatsAppSender()
    window.show()
    sys.exit(app.exec())
