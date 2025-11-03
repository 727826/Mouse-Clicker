import sys
import time
import threading
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QMessageBox
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import Qt, QTimer, QTime, Signal, QObject, QThread
# from pynput import mouse, keyboard


# ================= 后台工作线程 =================
class Worker(QObject):
    """
    一个在后台线程执行连点任务的 Worker。
    """
    finished = Signal()
    update_countdown = Signal(str)

    def __init__(self, settings):
        super().__init__()
        self.settings = settings
        self.is_running = True
        self.mouse_controller = mouse.Controller()
        self.keyboard_controller = keyboard.Controller()

    def run(self):
        """主执行逻辑"""
        try:
            start_delay = self.settings['start_delay']
            if start_delay > 0:
                for i in range(start_delay, 0, -1):
                    if not self.is_running: return
                    self.update_countdown.emit(f"{i}秒后开始...")
                    time.sleep(1)

            repeat_count = self.settings['repeat_count']
            interval_ms = self.settings['interval_ms']

            # 根据设置重建按键对象
            action_type = self.settings['action_type']
            action_value = self.settings['action_value']

            action = None
            if action_type == 'mouse':
                action = getattr(mouse.Button, action_value)
            elif action_type == 'keyboard_key':
                action = getattr(keyboard.Key, action_value)
            elif action_type == 'keyboard_code':
                action = keyboard.KeyCode.from_char(action_value)

            if action is None:
                raise ValueError("未能识别捕获的按键类型")

            for i in range(repeat_count):
                if not self.is_running: break

                countdown_msg = f"剩余 {repeat_count - i} 次"
                self.update_countdown.emit(countdown_msg)

                if action_type == 'mouse':
                    self.mouse_controller.click(action)
                else:
                    self.keyboard_controller.tap(action)

                time.sleep(interval_ms / 1000.0)

        except Exception as e:
            print(f"Worker 线程出错: {e}")
            self.update_countdown.emit(f"错误: {str(e)}")
        finally:
            if self.is_running:  # 只有在正常完成时才显示 "已停止"
                self.update_countdown.emit("已停止")
            else:  # 如果是被手动停止的
                self.update_countdown.emit("任务已取消")
            self.finished.emit()

    def stop(self):
        """请求停止线程"""
        self.is_running = False


# ================= 事件监听器 =================
class EventListener(QObject):
    """
    全局监听鼠标和键盘事件，用于捕获用户要重复的按键。
    """
    key_captured = Signal(str, str, str)  # 发送: 类型, 值, 显示名称

    def __init__(self):
        super().__init__()
        self.mouse_listener = None
        self.keyboard_listener = None
        self.is_listening = False
        self.capture_lock = threading.Lock()
        self._stop_thread = None  # 用于停止操作的线程

    def _safe_stop_and_join(self):
        """在后台线程中安全地停止和加入监听器"""
        if self.mouse_listener:
            self.mouse_listener.stop()
            self.mouse_listener.join()
            self.mouse_listener = None
        if self.keyboard_listener:
            self.keyboard_listener.stop()
            self.keyboard_listener.join()
            self.keyboard_listener = None
        self.is_listening = False
        print("监听器已安全停止。")

    def on_click(self, x, y, button, pressed):
        with self.capture_lock:
            if pressed and self.is_listening:
                self.key_captured.emit('mouse', button.name, f"鼠标: {button.name}")
                self.is_listening = False

    def on_press(self, key):
        with self.capture_lock:
            if self.is_listening:
                try:
                    if hasattr(key, 'char') and key.char is not None:
                        self.key_captured.emit('keyboard_code', key.char, f"按键: {key.char}")
                    else:
                        self.key_captured.emit('keyboard_key', key.name, f"特殊键: {key.name}")
                except Exception as e:
                    print(f"处理按键时出错: {e}")
                self.is_listening = False

    def start_listening(self):
        if self.is_listening:
            return

        self.is_listening = True
        self.mouse_listener = mouse.Listener(on_click=self.on_click)
        self.keyboard_listener = keyboard.Listener(on_press=self.on_press)
        self.mouse_listener.start()
        self.keyboard_listener.start()

    def stop_listening(self):
        """
        异步地停止监听器，避免阻塞主线程。
        """
        if self.mouse_listener or self.keyboard_listener:
            if self._stop_thread is None or not self._stop_thread.is_alive():
                self._stop_thread = threading.Thread(target=self._safe_stop_and_join, daemon=True)
                self._stop_thread.start()


# ================= 主窗口类 =================
class AutoClicker:

    def __init__(self):
        self.ui = QUiLoader().load('ui/mc.ui')

        self.active = False
        self.captured_action_type = 'mouse'
        self.captured_action_value = 'left'
        self.captured_action_name = "鼠标: left"

        self.thread = QThread()  # 将 QThread 提升为持久成员
        self.worker = None

        self.listener = EventListener()
        self.listener.key_captured.connect(self.on_key_captured)

        # 定时启动功能
        self.countdown_timer = QTimer()
        self.countdown_timer.setInterval(1000)
        self.countdown_timer.timeout.connect(self.check_schedule_time)

        self.ui.start.clicked.connect(self.toggle_run)
        self.ui.captureButton.clicked.connect(self.start_capture)

        self.setup_ui_components()
        self.update_capture_label()

    def setup_ui_components(self):
        """初始化UI控件的设置和状态"""
        self.ui.statusbar.showMessage("就绪")

        self.ui.timeEdit.setTime(QTime.currentTime())
        self.ui.timeEdit.setDisplayFormat("HH:mm:ss")
        self.ui.timeEdit.timeChanged.connect(self.on_time_edit_changed)

        self.ui.repeatCountBox.setRange(1, 99999)
        self.ui.repeatCountBox.setValue(10)

        self.ui.intervalBox.setRange(1, 60000)
        self.ui.intervalBox.setValue(50)

        self.ui.delayBox.setRange(0, 60)
        self.ui.delayBox.setValue(3)

    def on_time_edit_changed(self):
        """当用户修改时间时，如果正在倒计时，则停止"""
        if self.countdown_timer.isActive():
            self.stop_worker()
            QMessageBox.information(self.ui, "提示", "您修改了启动时间，定时任务已取消。")

    def start_capture(self):
        """开始捕获按键"""
        if self.active:
            QMessageBox.warning(self.ui, "警告", "请先停止当前任务再捕获新按键。")
            return
        self.ui.captureLabel.setText("请按下要重复的按键...")
        self.listener.start_listening()

    def on_key_captured(self, action_type, action_value, name):
        """当监听到按键时，延迟处理以确保线程安全"""
        QTimer.singleShot(0, lambda: self._process_capture(action_type, action_value, name))

    def _process_capture(self, action_type, action_value, name):
        """实际处理捕获结果的函数"""
        self.captured_action_type = action_type
        self.captured_action_value = action_value
        self.captured_action_name = name
        self.update_capture_label()
        self.listener.stop_listening()

    def update_capture_label(self):
        """更新显示已捕获按键的标签"""
        self.ui.captureLabel.setText(f"已捕获: {self.captured_action_name}")

    def toggle_run(self):
        """启动或停止连点器"""
        if not self.active:
            scheduled_time = self.ui.timeEdit.time()
            current_time = QTime.currentTime()

            if self.ui.scheduleCheckBox.isChecked():
                if scheduled_time <= current_time:
                    QMessageBox.warning(self.ui, "时间错误", "设定的启动时间已过，请选择一个未来的时间。")
                    return
                self.active = True
                self.ui.start.setText("取消定时 (Cancel)")
                self.set_controls_enabled(False)
                self.ui.statusbar.showMessage("等待定时启动...")
                self.countdown_timer.start()
            else:
                self.start_worker(is_scheduled=False)
        else:
            self.stop_worker()

    def check_schedule_time(self):
        """每秒检查一次是否到达预定时间"""
        scheduled_time = self.ui.timeEdit.time()
        current_time = QTime.currentTime()
        seconds_to_start = current_time.secsTo(scheduled_time)

        if seconds_to_start <= 0:
            self.countdown_timer.stop()
            self.start_worker(is_scheduled=True)
        else:
            self.ui.statusbar.showMessage(f"{seconds_to_start} 秒后自动开始...")

    def start_worker(self, is_scheduled=False):
        """启动后台连点线程"""
        self.active = True
        self.ui.start.setText("停止 (Stop)")
        self.set_controls_enabled(False)
        self.ui.statusbar.showMessage("运行中...")

        settings = {
            'repeat_count': self.ui.repeatCountBox.value(),
            'interval_ms': self.ui.intervalBox.value(),
            'start_delay': 0 if is_scheduled else self.ui.delayBox.value(),
            'action_type': self.captured_action_type,
            'action_value': self.captured_action_value
        }

        if self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()

        self.worker = Worker(settings)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.on_worker_finished)
        self.worker.update_countdown.connect(self.ui.statusbar.showMessage)

        self.thread.start()

    def stop_worker(self):
        """请求停止后台所有活动"""
        if self.countdown_timer.isActive():
            self.countdown_timer.stop()
            self.on_worker_finished()  # 直接重置UI

        if self.worker:
            self.worker.stop()

        if self.thread and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()

    def on_worker_finished(self):
        """线程完全结束后，重置UI状态"""
        print("Worker 线程已结束，正在重置UI。")
        self.active = False
        self.ui.start.setText("开始 (Start)")
        # 只有在倒计时器没在运行时才更新为已停止，否则会被倒计时覆盖
        if not self.countdown_timer.isActive():
            self.ui.statusbar.showMessage("已停止")
        self.set_controls_enabled(True)
        self.worker = None

    def set_controls_enabled(self, enabled):
        """启用或禁用界面上的设置控件"""
        self.ui.repeatCountBox.setEnabled(enabled)
        self.ui.intervalBox.setEnabled(enabled)
        self.ui.delayBox.setEnabled(enabled)
        self.ui.timeEdit.setEnabled(enabled)
        self.ui.scheduleCheckBox.setEnabled(enabled)
        self.ui.captureButton.setEnabled(enabled)

    def closeEvent(self, event):
        """确保退出时停止所有活动"""
        self.stop_worker()
        self.listener.stop_listening()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutoClicker()
    # 将 closeEvent 绑定到主窗口
    window.ui.closeEvent = window.closeEvent
    window.ui.show()
    sys.exit(app.exec())