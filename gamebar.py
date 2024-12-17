import os
import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout,
    QHBoxLayout, QWidget, QFileDialog, QLineEdit, QLabel, QSystemTrayIcon, QMenu, QListWidgetItem
)
from PyQt6.QtGui import QIcon, QPixmap, QAction
from PyQt6.QtCore import Qt
from win32com.client import Dispatch
from PIL import Image

class GameShortcutManager(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setStyleSheet("""
            QMainWindow {
                background-color: #2d2d2d;
            }
            QLabel, QPushButton, QLineEdit {
                color: #ffffff;
                font-size: 14px;
            }
            QPushButton {
                background-color: #3b3b3b;
                border: none;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #5c5c5c;
            }
            QListWidget {
                background-color: #3b3b3b;
                color: #ffffff;
            }
        """)

        self.setWindowTitle("Gerenciador de Atalhos de Jogos")
        self.setGeometry(100, 100, 600, 400)
        self.shortcuts = []

        # Layout principal
        self.main_layout = QVBoxLayout()

        # Lista de atalhos
        self.shortcut_list = QListWidget()
        self.shortcut_list.itemClicked.connect(self.display_icon)
        self.main_layout.addWidget(self.shortcut_list)

        # Área para exibir o ícone
        self.icon_label = QLabel("Ícone do Atalho")
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_label.mousePressEvent = self.open_shortcut  # Define ação de clique
        self.main_layout.addWidget(self.icon_label)

        # Campo de entrada para renomear
        self.rename_field = QLineEdit()
        self.rename_field.setPlaceholderText("Insira o novo nome do atalho")
        self.main_layout.addWidget(self.rename_field)

        # Botões
        button_layout = QHBoxLayout()
        import_button = QPushButton("Importar Atalhos")
        import_button.clicked.connect(self.import_shortcuts)
        rename_button = QPushButton("Renomear Selecionado")
        rename_button.clicked.connect(self.rename_shortcut)
        button_layout.addWidget(import_button)
        button_layout.addWidget(rename_button)
        self.main_layout.addLayout(button_layout)

        # Configuração de tray icon
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.png"))
        self.tray_icon.setToolTip("Gerenciador de Atalhos")
        self.tray_icon.activated.connect(self.restore_from_tray)

        tray_menu = QMenu()
        restore_action = QAction("Restaurar", self)
        restore_action.triggered.connect(self.showNormal)
        quit_action = QAction("Sair", self)
        quit_action.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(restore_action)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)

        # Janela principal
        container = QWidget()
        container.setLayout(self.main_layout)
        self.setCentralWidget(container)

    def import_shortcuts(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Selecione os atalhos de jogos", "", "Atalhos (*.lnk)"
        )
        for path in paths:
            if path not in self.shortcuts:
                self.shortcuts.append(path)
                self.add_shortcut_to_list(path)

    def add_shortcut_to_list(self, shortcut_path):
        # Cria um item de lista com o nome do atalho
        item = QListWidgetItem(os.path.basename(shortcut_path))
        
        # Obtém o ícone do atalho
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        icon_path, _ = shortcut.IconLocation.split(",")[0], 0
        if os.path.exists(icon_path):
            icon = Image.open(icon_path)
            icon.thumbnail((64, 64))
            icon.save("temp_icon.png")
            icon_pixmap = QPixmap("temp_icon.png")
            item.setIcon(QIcon(icon_pixmap))
        else:
            item.setIcon(QIcon("default_icon.png"))  # Coloca um ícone padrão se não houver ícone

        self.shortcut_list.addItem(item)

    def rename_shortcut(self):
        selected_item = self.shortcut_list.currentItem()
        if selected_item:
            new_name = self.rename_field.text()
            old_path = self.shortcuts[self.shortcut_list.row(selected_item)]
            new_path = os.path.join(os.path.dirname(old_path), new_name + ".lnk")
            os.rename(old_path, new_path)
            self.shortcuts[self.shortcut_list.row(selected_item)] = new_path
            selected_item.setText(new_name + ".lnk")
            self.rename_field.clear()

    def display_icon(self, item):
        shortcut_path = self.shortcuts[self.shortcut_list.row(item)]
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        icon_path, _ = shortcut.IconLocation.split(",")[0], 0
        if os.path.exists(icon_path):
            icon = Image.open(icon_path)
            icon.thumbnail((64, 64))
            icon.save("temp_icon.png")
            self.icon_label.setPixmap(QPixmap("temp_icon.png"))
        else:
            self.icon_label.setText("Ícone não encontrado")
        # Armazena o atalho associado à imagem
        self.icon_label.shortcut_path = shortcut_path

    def open_shortcut(self, event):
        # Abre o atalho ao clicar na imagem
        if hasattr(self.icon_label, "shortcut_path") and self.icon_label.shortcut_path:
            os.startfile(self.icon_label.shortcut_path)

    def minimize_to_tray(self):
        self.hide()
        self.tray_icon.show()

    def restore_from_tray(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.showNormal()
            self.tray_icon.hide()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    manager = GameShortcutManager()
    manager.show()
    sys.exit(app.exec())
