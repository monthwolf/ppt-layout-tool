from PyQt6.QtCore import QThread, pyqtSignal

class Worker(QThread):
    """通用工作线程，用于执行耗时操作"""
    finished = pyqtSignal(object)
    error = pyqtSignal(Exception)
    progress = pyqtSignal(int, int, str)  # 当前进度, 总进度, 描述文本

    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs
        
        # 如果有进度回调函数，提取出来
        self.progress_callback = None
        if 'progress_callback' in kwargs:
            self.progress_callback = kwargs.pop('progress_callback')
            
            # 创建一个进度回调包装器，确保信号被正确发出
            # 注意：我们不保留原始回调，而是直接将进度更新发送到信号
            self.kwargs['progress_callback'] = self._progress_wrapper

    def _progress_wrapper(self, current, total, message=""):
        """进度回调函数的包装器，将回调转换为信号"""
        # 直接发出信号，不调用原始回调
        self.progress.emit(current, total, message)

    def run(self):
        try:
            result = self.function(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(e)
