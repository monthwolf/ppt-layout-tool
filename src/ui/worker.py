from PyQt6.QtCore import QThread, pyqtSignal

class Worker(QThread):
    """通用工作线程，用于执行耗时操作"""
    finished = pyqtSignal(object)
    error = pyqtSignal(Exception)

    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.function(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(e)
