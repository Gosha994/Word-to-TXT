import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QPushButton, QLineEdit,
                             QLabel, QVBoxLayout, QHBoxLayout, QSizePolicy,
                             QFileDialog, QMessageBox, QProgressBar, QTextEdit)
from PyQt6.QtCore import Qt, QPoint
from PyQt6.QtGui import QIcon, QPixmap
from docx import Document


class WordToTxtConverter(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.dragging = False
        self.drag_position = QPoint()
        self.current_file_path = ""

    def initUI(self):
        # Снос стандартного Title bar окна
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setWindowTitle("Конвертер Word в TXT")
        self.setFixedSize(450, 550)  # Немного увеличил размер для лучшего отображения

        # Иконка приложения
        try:
            self.setWindowIcon(QIcon("converter_icon.png"))
        except:
            pass

        # Тема для всего окна
        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
            }
        """)

        # Основной layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Создаем свой Title bar
        title_bar = self.create_title_bar()
        main_layout.addWidget(title_bar)

        # Контентная область
        content_widget = QWidget()
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)

        # Название приложения
        title_label = QLabel("Конвертер Word → TXT")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("""
            QLabel {
                color: #ffffff;
                font-size: 20px;
                font-weight: bold;
                padding: 10px;
                background-color: #2b2b2b;
                border-radius: 8px;
            }
        """)
        content_layout.addWidget(title_label)

        # Отображение пути к файлу
        self.file_path_display = QLineEdit()
        self.file_path_display.setPlaceholderText("Выберите Word файл...")
        self.file_path_display.setReadOnly(True)
        self.file_path_display.setStyleSheet("""
            QLineEdit {
                background-color: #2b2b2b;
                color: #ffffff;
                border: 2px solid #404040;
                border-radius: 5px;
                font-size: 14px;
                padding: 10px;
                margin: 5px 0px;
            }
        """)
        content_layout.addWidget(self.file_path_display)

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #404040;
                border-radius: 5px;
                text-align: center;
                color: white;
                background-color: #2b2b2b;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
        self.progress_bar.setVisible(False)
        content_layout.addWidget(self.progress_bar)

        # Область для предпросмотра текста
        preview_label = QLabel("Выбранный файл:")
        preview_label.setStyleSheet("color: #cccccc; font-size: 14px;")
        content_layout.addWidget(preview_label)

        # ЗАМЕНА: QLineEdit на QTextEdit для многострочного отображения
        self.preview_display = QTextEdit()
        self.preview_display.setReadOnly(True)
        self.preview_display.setStyleSheet("""
            QTextEdit {
                background-color: #2b2b2b;
                color: #cccccc;
                border: 2px solid #404040;
                border-radius: 5px;
                font-size: 12px;
                padding: 10px;
                margin: 5px 0px;
            }
        """)
        self.preview_display.setFixedHeight(120)  # Фиксированная высота для предпросмотра
        self.preview_display.setText("Здесь будет отображаться предпросмотр текста...")
        content_layout.addWidget(self.preview_display)

        # Layout для кнопок
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(10)

        # Кнопки
        self.open_btn = QPushButton("📁 Открыть Word файл")
        self.convert_btn = QPushButton("🔄 Конвертировать в TXT")
        self.save_btn = QPushButton("💾 Сохранить TXT файл")

        # Стили кнопок
        button_style = """
            QPushButton {
                background-color: #323232;
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                padding: 12px;
                margin: 2px;
            }
            QPushButton:hover {
                background-color: #404040;
                border: 1px solid #505050;
            }
            QPushButton:pressed {
                background-color: #282828;
            }
            QPushButton:disabled {
                background-color: #252525;
                color: #666666;
            }
        """

        for btn in [self.open_btn, self.convert_btn, self.save_btn]:
            btn.setStyleSheet(button_style)
            btn.setFixedHeight(50)

        # Подключаем функции к кнопкам
        self.open_btn.clicked.connect(self.open_word_file)
        self.convert_btn.clicked.connect(self.convert_to_txt)
        self.save_btn.clicked.connect(self.save_txt_file)

        # Изначально делаем кнопки неактивными
        self.convert_btn.setEnabled(False)
        self.save_btn.setEnabled(False)

        buttons_layout.addWidget(self.open_btn)
        buttons_layout.addWidget(self.convert_btn)
        buttons_layout.addWidget(self.save_btn)

        content_layout.addLayout(buttons_layout)
        content_widget.setLayout(content_layout)
        main_layout.addWidget(content_widget)

        self.setLayout(main_layout)

    def create_title_bar(self):
        """Создает кастомную панель заголовка"""
        title_bar = QWidget()
        title_bar.setFixedHeight(40)
        title_bar.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                border-bottom: 1px solid #404040;
            }
        """)

        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(10, 0, 5, 0)
        title_layout.setSpacing(10)

        # Иконка приложения
        icon_label = QLabel()
        icon_label.setFixedSize(20, 20)
        icon_label.setStyleSheet("""
            QLabel {
                background-color: transparent;
                border: none;
            }
        """)
        try:
            pixmap = QPixmap("converter_icon.png").scaled(20, 20, Qt.AspectRatioMode.KeepAspectRatio,
                                                          Qt.TransformationMode.SmoothTransformation)
            icon_label.setPixmap(pixmap)
        except:
            icon_label.setText("📄")

        # Название приложения
        title_label = QLabel("Конвертер Word в TXT")
        title_label.setStyleSheet("""
            QLabel {
                color: #ffffff;
                font-size: 14px;
                background-color: transparent;
            }
        """)

        # Растягивающееся пространство
        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        # Кнопка свертывания
        min_btn = QPushButton("−")
        min_btn.setFixedSize(25, 25)
        min_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 3px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #404040;
            }
            QPushButton:pressed {
                background-color: #505050;
            }
        """)
        min_btn.clicked.connect(self.showMinimized)

        # Кнопка закрытия
        close_btn = QPushButton("×")
        close_btn.setFixedSize(25, 25)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 3px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e81123;
                border: 1px solid #e81123;
            }
            QPushButton:pressed {
                background-color: #f1707a;
            }
        """)
        close_btn.clicked.connect(self.close)

        title_layout.addWidget(icon_label)
        title_layout.addWidget(title_label)
        title_layout.addWidget(spacer)
        title_layout.addWidget(min_btn)
        title_layout.addWidget(close_btn)

        title_bar.setLayout(title_layout)

        # Сохраняем ссылки для обработки перемещения
        title_bar.mousePressEvent = self.title_mouse_press_event
        title_bar.mouseMoveEvent = self.title_mouse_move_event
        title_bar.mouseReleaseEvent = self.title_mouse_release_event

        return title_bar

    def title_mouse_press_event(self, event):
        """Обработка нажатия мыши на заголовок"""
        if event.button() == Qt.MouseButton.LeftButton:
            self.dragging = True
            self.drag_position = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def title_mouse_move_event(self, event):
        """Обработка перемещения мыши при dragging"""
        if self.dragging and event.buttons() == Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_position)
            event.accept()

    def title_mouse_release_event(self, event):
        """Обработка отпускания кнопки мыши"""
        if event.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            event.accept()

    def open_word_file(self):
        """Открытие Word файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите Word файл",
            "",
            "Word Files (*.docx *.doc);;All Files (*)"
        )

        if file_path:
            self.current_file_path = file_path
            self.file_path_display.setText(file_path)
            self.convert_btn.setEnabled(True)
            self.save_btn.setEnabled(False)

            # Показываем имя файла в preview
            file_name = os.path.basename(file_path)
            self.preview_display.setText(
                f"Выбран файл: {file_name}\n\nНажмите 'Конвертировать в TXT' для извлечения текста.")

    def convert_to_txt(self):
        """Конвертация Word в текст"""
        if not self.current_file_path:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите Word файл!")
            return

        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(30)

            # Чтение Word файла
            doc = Document(self.current_file_path)

            self.progress_bar.setValue(60)

            # Извлечение текста из всех параграфов
            full_text = []
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():  # Игнорируем пустые строки
                    full_text.append(paragraph.text)

            self.converted_text = "\n".join(full_text)

            self.progress_bar.setValue(90)

            # Обрезаем текст для предпросмотра (первые 300 символов)
            if self.converted_text.strip():
                preview_text = self.converted_text[:300] + "..." if len(
                    self.converted_text) > 300 else self.converted_text
                # Добавляем информацию о количестве символов
                char_count = len(self.converted_text)
                word_count = len(self.converted_text.split())
                preview_info = f"Текст извлечен успешно!\nСимволов: {char_count}, Слов: {word_count}\n\nПредпросмотр:\n{preview_text}"
                self.preview_display.setText(preview_info)
            else:
                self.preview_display.setText("В документе не найден текст или документ пуст.")
                self.converted_text = ""

            self.progress_bar.setValue(100)

            # Активируем кнопку сохранения только если есть текст
            self.save_btn.setEnabled(bool(self.converted_text.strip()))

            QMessageBox.information(self, "Успех",
                                    "Файл успешно сконвертирован!\nТеперь вы можете сохранить его как TXT.")

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось конвертировать файл:\n{str(e)}")
            self.preview_display.setText(f"Ошибка при конвертации: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)

    def save_txt_file(self):
        """Сохранение текста в TXT файл"""
        if not hasattr(self, 'converted_text') or not self.converted_text.strip():
            QMessageBox.warning(self, "Ошибка", "Нет сконвертированного текста для сохранения!")
            return

        # Предлагаем сохранить с тем же именем но с расширением .txt
        default_name = os.path.splitext(os.path.basename(self.current_file_path))[0] + ".txt"

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить как TXT",
            default_name,
            "Text Files (*.txt);;All Files (*)"
        )

        if file_path:
            try:
                # Добавляем расширение .txt если его нет
                if not file_path.lower().endswith('.txt'):
                    file_path += '.txt'

                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.converted_text)

                QMessageBox.information(self, "Успех", f"Файл успешно сохранен!\n{file_path}")

            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Устанавливаем стиль для всего приложения
    app.setStyle('Fusion')

    ex = WordToTxtConverter()
    ex.show()
    sys.exit(app.exec())