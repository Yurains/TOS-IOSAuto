import sys
import time
import io
import os
import json
import base64

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, 
    QVBoxLayout, QHBoxLayout, QWidget, QLabel, 
    QMessageBox, QListWidget, QMenu, QDialog, QDialogButtonBox
)
from PyQt5.QtCore import (
    Qt, QPoint, QRect, QSize, QThread, pyqtSignal, 
    QEvent, pyqtSlot, QBuffer, QIODevice, QByteArray
)
from PyQt5.QtGui import (
    QScreen, QPixmap, QPainter, QColor, QPen
)

import ddddocr
import numpy as np
from PIL import Image
import pyautogui
from datetime import datetime

JSON_FILE = "saved_captures.json"

class CaptureInfo:
    def __init__(self, x, y, width, height, image, ocr_text):
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.image = image  # QPixmap
        self.ocr_text = ocr_text
        self.click_count = 1  # 可自行調整預設值

    def __str__(self):
        return f"文字: {self.ocr_text} (點擊次數: {self.click_count})"

    def to_dict(self):
        """
        將此擷取信息轉成可保存於 JSON 的字典格式。
        使用 QBuffer 介面將 QPixmap 轉為 base64，避免
        QPixmap.save() 直接對 BytesIO 造成 TypeError。
        """
        byte_array = QByteArray()
        buffer = QBuffer(byte_array)
        buffer.open(QIODevice.WriteOnly)
        self.image.save(buffer, "PNG")
        buffer.close()

        encoded_image = base64.b64encode(byte_array.data()).decode("utf-8")
        
        return {
            "x": self.x,
            "y": self.y,
            "width": self.width,
            "height": self.height,
            "ocr_text": self.ocr_text,
            "click_count": self.click_count,
            "encoded_image": encoded_image
        }

    @staticmethod
    def from_dict(data):
        """
        從字典還原為 CaptureInfo 物件。
        將 base64 字串轉回 QPixmap。
        """
        decoded_image = base64.b64decode(data["encoded_image"])
        qpixmap = QPixmap()
        qpixmap.loadFromData(decoded_image, "PNG")

        cap = CaptureInfo(
            data["x"],
            data["y"],
            data["width"],
            data["height"],
            qpixmap,
            data["ocr_text"]
        )
        cap.click_count = data.get("click_count", 1)
        return cap

class ScreenCaptureWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowState(Qt.WindowFullScreen)
        
        screen = QApplication.primaryScreen()
        self.original_screenshot = screen.grabWindow(0)
        
        self.begin = QPoint()
        self.end = QPoint()
        self.is_drawing = False

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(self.rect(), self.original_screenshot)
        
        # 半透明遮罩
        mask = QColor(0, 0, 0, 100)
        painter.fillRect(self.rect(), mask)
        
        if self.is_drawing:
            pen = QPen(Qt.red, 2, Qt.SolidLine)
            painter.setPen(pen)
            
            rect = QRect(self.begin, self.end)
            painter.drawPixmap(rect, self.original_screenshot, rect)
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.begin = event.pos()
            self.end = self.begin
            self.is_drawing = True
            self.update()

    def mouseMoveEvent(self, event):
        if self.is_drawing:
            self.end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.is_drawing:
            self.is_drawing = False
            if self.begin and self.end:
                x1, y1 = min(self.begin.x(), self.end.x()), min(self.begin.y(), self.end.y())
                x2, y2 = max(self.begin.x(), self.end.x()), max(self.begin.y(), self.end.y())
                
                if x2 - x1 > 0 and y2 - y1 > 0:
                    screenshot = self.original_screenshot.copy(x1, y1, x2 - x1, y2 - y1)
                    self.capture_info = {
                        'x': x1,
                        'y': y1,
                        'width': x2 - x1,
                        'height': y2 - y1,
                        'image': screenshot
                    }
                    self.close()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.close()

class WorkerThread(QThread):
    """使程式持續執行點擊的執行緒。"""
    update_status = pyqtSignal(str)

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.keep_running = True  # 控制執行緒的運行

    def run(self):
        # 不斷循環執行點擊，直到 keep_running 為 False
        while self.keep_running:
            self.main_window.execute_all_clicks(infinite_mode=True)
            time.sleep(1)  # 適度延遲以避免過度頻繁執行

    def stop(self):
        self.keep_running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('多區域OCR自動點擊器')
        self.setGeometry(100, 100, 800, 600)
        
        self.captures = []
        self.ocr = ddddocr.DdddOcr()

        # 用於控制是否處於“持續執行”狀態
        self.worker_thread = None
        
        self.init_ui()
        # 啟動時嘗試從 JSON 載入先前的擷取資訊
        self.load_captures_from_json()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        
        # 左側面板（擷取列表）
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        
        self.capture_button = QPushButton('新增擷取區域')
        self.capture_button.clicked.connect(self.start_capture)
        self.capture_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
        """)

        self.capture_list = QListWidget()
        self.capture_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.capture_list.customContextMenuRequested.connect(self.show_context_menu)
        self.capture_list.itemClicked.connect(self.show_capture_preview)
        
        left_layout.addWidget(self.capture_button)
        left_layout.addWidget(self.capture_list)

        # “開始持續執行”按鈕
        self.start_infinite_button = QPushButton('開始持續執行')
        self.start_infinite_button.clicked.connect(self.start_infinite_execution)
        self.start_infinite_button.setStyleSheet("""
            QPushButton {
                background-color: #ff9800;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        left_layout.addWidget(self.start_infinite_button)

        # ✅ 新增「一鍵刪除全部」按鈕
        self.clear_all_button = QPushButton('一鍵刪除全部')
        self.clear_all_button.clicked.connect(self.clear_all_captures)
        self.clear_all_button.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        left_layout.addWidget(self.clear_all_button)

        # 右側面板（預覽和控制）
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(QSize(400, 300))
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("border: 2px dashed #cccccc;")
        
        self.ocr_result = QLabel('OCR結果')
        self.ocr_result.setWordWrap(True)
        self.ocr_result.setStyleSheet("""
            QLabel {
                padding: 10px;
                background-color: #f5f5f5;
                border-radius: 5px;
            }
        """)
        
        self.execute_button = QPushButton('執行所有點擊')
        self.execute_button.clicked.connect(self.execute_all_clicks)
        self.execute_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        
        # 新增「執行單點」按鈕，只針對選取的擷取項目進行單點操作
        self.execute_single_button = QPushButton('位置點擊測試')
        self.execute_single_button.clicked.connect(self.execute_single_click)
        self.execute_single_button.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        
        self.status_label = QLabel('準備就緒')
        
        right_layout.addWidget(self.preview_label)
        right_layout.addWidget(self.ocr_result)
        right_layout.addWidget(self.execute_button)
        right_layout.addWidget(self.execute_single_button)
        right_layout.addWidget(self.status_label)
        
        # 設置左右面板比例
        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(right_panel, 2)

    def show_context_menu(self, position):
        if not self.capture_list.itemAt(position):
            return
            
        menu = QMenu()
        delete_action = menu.addAction("刪除")
        set_clicks_action = menu.addAction("設置點擊次數")
        
        action = menu.exec_(self.capture_list.mapToGlobal(position))
        
        if action == delete_action:
            self.delete_capture()
        elif action == set_clicks_action:
            self.set_click_count()

    def delete_capture(self):
        current_row = self.capture_list.currentRow()
        if current_row >= 0:
            self.capture_list.takeItem(current_row)
            self.captures.pop(current_row)
            # 每次修改後保存一次
            self.save_captures_to_json()

    def set_click_count(self):
        current_row = self.capture_list.currentRow()
        if current_row >= 0:
            capture = self.captures[current_row]
            counts = [1, 2, 3, 5, 10]
            current_index = counts.index(capture.click_count) if capture.click_count in counts else 0
            capture.click_count = counts[(current_index + 1) % len(counts)]
            self.capture_list.currentItem().setText(str(capture))
            # 每次修改後保存一次
            self.save_captures_to_json()

    def start_capture(self):
        """ 進行螢幕擷取，並自動 OCR。 """
        self.hide()
        QApplication.processEvents()
        
        screen_capture = ScreenCaptureWidget()
        screen_capture.show()
        
        # 等待擷取視窗關閉
        while screen_capture.isVisible():
            QApplication.processEvents()
        
        # 擷取完成後進行 OCR
        if hasattr(screen_capture, 'capture_info'):
            info = screen_capture.capture_info
            pixmap = info['image']
            temp_path = f"temp_capture_{len(self.captures)}.png"
            pixmap.save(temp_path)
            
            try:
                with open(temp_path, 'rb') as f:
                    image_bytes = f.read()
                
                ocr_text = self.ocr.classification(image_bytes)
                
                capture = CaptureInfo(
                    info['x'], info['y'],
                    info['width'], info['height'],
                    pixmap, ocr_text
                )
                
                self.captures.append(capture)
                self.capture_list.addItem(str(capture))
                # 新增後存檔
                self.save_captures_to_json()
                
            except Exception as e:
                QMessageBox.warning(self, "錯誤", f"OCR識別失敗: {str(e)}")
            
            # 刪除暫存檔
            if os.path.exists(temp_path):
                os.remove(temp_path)
        
        self.show()

    def show_capture_preview(self, item):
        index = self.capture_list.row(item)
        if 0 <= index < len(self.captures):
            capture = self.captures[index]
            scaled_pixmap = capture.image.scaled(
                400, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            self.preview_label.setPixmap(scaled_pixmap)
            self.ocr_result.setText(f"OCR結果: {capture.ocr_text}")

    def execute_all_clicks(self, infinite_mode=False):
        """
        執行所有擷取區域的點擊。 
        如果 infinite_mode=True，表示在持續執行的模式下呼叫此函式。
        """
        try:
            if not infinite_mode:
                self.status_label.setText("開始執行點擊...")
            QApplication.processEvents()
            
            # 保存原始滑鼠位置
            original_x, original_y = pyautogui.position()
            
            for capture in self.captures:
                # 先重新擷取該區域並進行OCR，只有當OCR結果符合最初擷取到的內容時才點擊
                new_screenshot = pyautogui.screenshot(
                    region=(capture.x, capture.y, capture.width, capture.height)
                )
                img_bytes = io.BytesIO()
                new_screenshot.save(img_bytes, format="PNG")
                img_bytes.seek(0)
                current_ocr_text = self.ocr.classification(img_bytes.getvalue())

                if current_ocr_text.strip() == capture.ocr_text.strip():
                    click_x = capture.x + capture.width // 2
                    click_y = capture.y + capture.height // 2
                    
                    for i in range(capture.click_count):
                        if not infinite_mode:
                            self.status_label.setText(
                                f"點擊 '{capture.ocr_text}' ({i+1}/{capture.click_count})"
                            )
                        QApplication.processEvents()
                        
                        pyautogui.moveTo(click_x, click_y, duration=0.2)
                        pyautogui.click()
                        time.sleep(0.2)
                else:
                    if not infinite_mode:
                        self.status_label.setText(
                            f"跳過: OCR未符合 '{capture.ocr_text}', 目前為 '{current_ocr_text}'"
                        )
                    QApplication.processEvents()
                    time.sleep(0.5)
            
            # 恢復滑鼠位置
            pyautogui.moveTo(original_x, original_y, duration=0.2)

            if not infinite_mode:
                self.status_label.setText("所有點擊已完成")
            
        except Exception as e:
            if not infinite_mode:
                self.status_label.setText(f"執行錯誤: {str(e)}")
                QMessageBox.warning(self, "錯誤", str(e))

    def execute_single_click(self):
        """
        只對當前選取的擷取區域執行單次點擊操作，
        先重新擷取該區域進行OCR比對，若符合則點擊一次。
        """
        current_row = self.capture_list.currentRow()
        if current_row < 0 or current_row >= len(self.captures):
            QMessageBox.warning(self, "提示", "請先選擇一個擷取區域")
            return

        capture = self.captures[current_row]
        try:
            self.status_label.setText("開始執行單點...")
            QApplication.processEvents()
            
            # 保存原始滑鼠位置
            original_x, original_y = pyautogui.position()

            new_screenshot = pyautogui.screenshot(
                region=(capture.x, capture.y, capture.width, capture.height)
            )
            img_bytes = io.BytesIO()
            new_screenshot.save(img_bytes, format="PNG")
            img_bytes.seek(0)
            current_ocr_text = self.ocr.classification(img_bytes.getvalue())

            if current_ocr_text.strip() == capture.ocr_text.strip():
                click_x = capture.x + capture.width // 2
                click_y = capture.y + capture.height // 2

                self.status_label.setText(f"單點執行: 點擊 '{capture.ocr_text}'")
                QApplication.processEvents()
                
                pyautogui.moveTo(click_x, click_y, duration=0.2)
                pyautogui.click()
                time.sleep(0.2)
            else:
                self.status_label.setText(
                    f"跳過: OCR未符合 '{capture.ocr_text}', 目前為 '{current_ocr_text}'"
                )
                QApplication.processEvents()
                time.sleep(0.5)
            
            # 恢復滑鼠位置
            pyautogui.moveTo(original_x, original_y, duration=0.2)
            self.status_label.setText("單點執行已完成")
            
        except Exception as e:
            self.status_label.setText(f"單點執行錯誤: {str(e)}")
            QMessageBox.warning(self, "錯誤", str(e))

    def start_infinite_execution(self):
        """
        顯示一個視窗說明即將開始無限執行，
        並在視窗中提示按下 Enter 鍵後結束。
        """
        # 如果已經在持續執行，則不重複啟動
        if self.worker_thread and self.worker_thread.isRunning():
            QMessageBox.information(self, "提示", "持續執行已經在進行中。")
            return

        # 彈出對話框，提示即將開始
        dialog = QDialog(self)
        dialog.setWindowTitle("開始持續執行")
        layout = QVBoxLayout(dialog)

        label = QLabel("此程式將不斷地執行點擊。\n請按下 Enter 鍵(或點確定)結束！", dialog)
        label.setWordWrap(True)
        layout.addWidget(label)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec_()

        # 按下確定後，啟動執行緒
        self.worker_thread = WorkerThread(self)
        self.worker_thread.start()

        # 在主視窗安裝事件過濾器，用於偵測鍵盤 Enter
        self.installEventFilter(self)

        # 更新按鈕狀態
        self.status_label.setText("持續執行已開始，按 Enter 結束。")

    def eventFilter(self, obj, event):
        """ 檢測按下 Enter 鍵以結束 worker_thread。 """
        if event.type() == QEvent.KeyPress:
            if event.key() in (Qt.Key_Return, Qt.Key_Enter):
                self.stop_infinite_execution()
                return True
        return super().eventFilter(obj, event)

    @pyqtSlot()
    def stop_infinite_execution(self):
        """ 結束持續執行 """
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.stop()  # 設置 keep_running = False
            self.worker_thread.quit()
            self.worker_thread.wait()
            self.worker_thread = None

        # 移除事件過濾器
        self.removeEventFilter(self)
        self.status_label.setText("已停止持續執行。")

    def save_captures_to_json(self):
        """
        將目前的擷取列表存到 JSON。
        會使用 to_dict()，其中包含將 QPixmap 轉成 base64。
        """
        data = [capture.to_dict() for capture in self.captures]
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load_captures_from_json(self):
        """
        從 JSON 載入擷取列表。
        其中會使用 from_dict()，將 base64 還原成 QPixmap。
        """
        if os.path.exists(JSON_FILE):
            try:
                with open(JSON_FILE, 'r', encoding='utf-8') as f:
                    data_list = json.load(f)
                self.captures.clear()
                self.capture_list.clear()

                for item in data_list:
                    capture = CaptureInfo.from_dict(item)
                    self.captures.append(capture)
                    self.capture_list.addItem(str(capture))

            except Exception as e:
                QMessageBox.warning(self, "讀取錯誤", f"讀取 {JSON_FILE} 失敗: {e}")

    # ✅ 新增：一鍵刪除全部
    def clear_all_captures(self):
        """
        一鍵刪除所有的內容：
        1) 停止持續執行（若在執行）
        2) 清空 captures 與清單
        3) 清空右側預覽與 OCR 顯示
        4) 刪除或重建 JSON 檔
        """
        reply = QMessageBox.question(
            self,
            "確認刪除",
            "此操作將刪除所有擷取與紀錄，且無法復原。\n確定要刪除嗎？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # 1) 停止持續執行
        self.stop_infinite_execution()

        # 2) 清空資料
        self.captures.clear()
        self.capture_list.clear()

        # 3) 清空右側顯示
        self.preview_label.clear()
        self.preview_label.setText("")  # 清除占位
        self.ocr_result.setText("OCR結果")
        self.status_label.setText("已清空所有擷取內容。")

        # 4) 刪除或重建 JSON
        try:
            if os.path.exists(JSON_FILE):
                os.remove(JSON_FILE)
        except Exception as e:
            # 若刪除失敗，至少寫入為空陣列
            with open(JSON_FILE, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
            QMessageBox.warning(self, "檔案處理", f"刪除 {JSON_FILE} 失敗，已改為重置為空內容：{e}")
            return

        # 重新寫入空陣列（可有可無，保證檔案存在且為空）
        with open(JSON_FILE, 'w', encoding='utf-8') as f:
            json.dump([], f, ensure_ascii=False, indent=2)

        QMessageBox.information(self, "完成", "已刪除全部內容並重置。")

    def closeEvent(self, event):
        """ 在關閉視窗時確保執行緒停止，並保存資料 """
        self.stop_infinite_execution()
        self.save_captures_to_json()
        super().closeEvent(event)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
