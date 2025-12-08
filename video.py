import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import cv2
import datetime
import threading
import subprocess
import os
import sys
import logging
import signal
import atexit
from tkinterdnd2 import TkinterDnD, DND_FILES

# Windows特定的导入
if sys.platform == 'win32':
    import ctypes
    from ctypes import wintypes

# 设置日志配置
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)

def send_console_ctrl_event(process_id, ctrl_event=2):
    """发送控制台控制事件到指定进程"""
    if sys.platform == 'win32':
        try:
            kernel32 = ctypes.windll.kernel32
            # CTRL_C_EVENT = 0, CTRL_BREAK_EVENT = 1, CTRL_CLOSE_EVENT = 2
            result = kernel32.GenerateConsoleCtrlEvent(ctrl_event, process_id)
            if result:
                print(f"[DEBUG] 成功发送CTRL_CLOSE_EVENT到进程 {process_id}")
                return True
            else:
                print(f"[DEBUG] 发送CTRL_CLOSE_EVENT到进程 {process_id} 失败")
                return False
        except Exception as e:
            print(f"[DEBUG] 发送控制台事件失败: {e}")
            return False
    else:
        # Linux/Mac系统使用SIGINT
        try:
            os.kill(process_id, signal.SIGINT)
            print(f"[DEBUG] 成功发送SIGINT到进程 {process_id}")
            return True
        except Exception as e:
            print(f"[DEBUG] 发送SIGINT失败: {e}")
            return False

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def find_ffmpeg_path():
    """查找FFmpeg可执行文件路径"""
    # 优先使用环境变量
    env_path = os.getenv('FFMPEG_PATH')
    if env_path and os.path.exists(env_path):
        return env_path

    # 检查是否在打包后的环境中
    if getattr(sys, 'frozen', False):
        # 如果是打包后的可执行文件
        if sys.platform == 'win32':
            # Windows打包环境
            bundle_dir = sys._MEIPASS
            ffmpeg_path = os.path.join(bundle_dir, 'ffmpeg', 'bin', 'ffmpeg.exe')
            if os.path.exists(ffmpeg_path):
                logger.info(f"Found bundled FFmpeg: {ffmpeg_path}")
                return ffmpeg_path

            # 也检查ffprobe
            ffprobe_path = os.path.join(bundle_dir, 'ffmpeg', 'bin', 'ffprobe.exe')
            if os.path.exists(ffprobe_path):
                logger.info(f"Found bundled FFprobe: {ffprobe_path}")
                # 返回ffmpeg路径，但记录ffprobe路径
                return ffmpeg_path
        else:
            # Linux/Mac打包环境
            bundle_dir = sys._MEIPASS
            ffmpeg_path = os.path.join(bundle_dir, 'ffmpeg', 'bin', 'ffmpeg')
            if os.path.exists(ffmpeg_path):
                logger.info(f"Found bundled FFmpeg: {ffmpeg_path}")
                return ffmpeg_path

    # 检查当前目录下的ffmpeg
    if sys.platform == 'win32':
        local_paths = [
            'ffmpeg/bin/ffmpeg.exe',
            'ffmpeg.exe',
            'ffmpeg/ffmpeg.exe'
        ]
    else:
        local_paths = [
            'ffmpeg/bin/ffmpeg',
            'ffmpeg',
            'ffmpeg/ffmpeg'
        ]

    for path in local_paths:
        if os.path.exists(path):
            abs_path = os.path.abspath(path)
            logger.info(f"Found local FFmpeg: {abs_path}")
            return abs_path

    # 使用系统PATH中的ffmpeg
    if sys.platform == 'win32':
        try:
            result = subprocess.run(['where', 'ffmpeg'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip().split('\n')[0]
                logger.info(f"Found system FFmpeg: {path}")
                return path
        except:
            pass
    else:
        try:
            result = subprocess.run(['which', 'ffmpeg'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip()
                logger.info(f"Found system FFmpeg: {path}")
                return path
        except:
            pass

    # 默认值
    return 'ffmpeg'

def find_ffprobe_path():
    """查找FFprobe可执行文件路径"""
    # 优先使用环境变量
    env_path = os.getenv('FFPROBE_PATH')
    if env_path and os.path.exists(env_path):
        return env_path

    # 检查是否在打包后的环境中
    if getattr(sys, 'frozen', False):
        # 如果是打包后的可执行文件
        if sys.platform == 'win32':
            # Windows打包环境
            bundle_dir = sys._MEIPASS
            ffprobe_path = os.path.join(bundle_dir, 'ffmpeg', 'bin', 'ffprobe.exe')
            if os.path.exists(ffprobe_path):
                logger.info(f"Found bundled FFprobe: {ffprobe_path}")
                return ffprobe_path
        else:
            # Linux/Mac打包环境
            bundle_dir = sys._MEIPASS
            ffprobe_path = os.path.join(bundle_dir, 'ffmpeg', 'bin', 'ffprobe')
            if os.path.exists(ffprobe_path):
                logger.info(f"Found bundled FFprobe: {ffprobe_path}")
                return ffprobe_path

    # 检查当前目录下的ffprobe
    if sys.platform == 'win32':
        local_paths = [
            'ffmpeg/bin/ffprobe.exe',
            'ffprobe.exe',
            'ffmpeg/ffprobe.exe'
        ]
    else:
        local_paths = [
            'ffmpeg/bin/ffprobe',
            'ffprobe',
            'ffmpeg/ffprobe'
        ]

    for path in local_paths:
        if os.path.exists(path):
            abs_path = os.path.abspath(path)
            logger.info(f"Found local FFprobe: {abs_path}")
            return abs_path

    # 使用系统PATH中的ffprobe
    if sys.platform == 'win32':
        try:
            result = subprocess.run(['where', 'ffprobe'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip().split('\n')[0]
                logger.info(f"Found system FFprobe: {path}")
                return path
        except:
            pass
    else:
        try:
            result = subprocess.run(['which', 'ffprobe'], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip()
                logger.info(f"Found system FFprobe: {path}")
                return path
        except:
            pass

    # 默认值
    return 'ffprobe'

# 获取FFmpeg路径
FFMPEG_PATH = find_ffmpeg_path()
logger.info(f"Using ffmpeg from: {FFMPEG_PATH}")

# 检查 FFmpeg 是否可用
def check_ffmpeg():
    try:
        result = subprocess.run([FFMPEG_PATH, '-version'], capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            logger.info("FFmpeg 可用")
            return True
        else:
            logger.error(f"FFmpeg 不可用，返回码: {result.returncode}")
            return False
    except FileNotFoundError:
        logger.error(f"找不到 FFmpeg: {FFMPEG_PATH}")
        return False
    except subprocess.TimeoutExpired:
        logger.error(f"FFmpeg 响应超时: {FFMPEG_PATH}")
        return False
    except Exception as e:
        logger.error(f"检查 FFmpeg 时出错: {e}")
        return False

# 在程序启动时检查 FFmpeg
if not check_ffmpeg():
    error_msg = f"找不到 FFmpeg！\n\n请确保：\n1. 已安装 FFmpeg 并添加到系统PATH\n2. 或者将 FFmpeg 放在程序目录下\n\n当前查找路径: {FFMPEG_PATH}"
    try:
        messagebox.showerror("错误", error_msg)
    except:
        print(f"错误: {error_msg}")
    sys.exit(1)

class VideoTrimmerPro:
    def __init__(self):
        # 强制刷新标准输出
        sys.stdout.flush()

        self.root = TkinterDnD.Tk()
        self.root.title("专业视频剪辑工具")
        self.root.geometry("1280x720")
        self.root.configure(bg="#333333")

        # 添加窗口最大化设置
        self.root.state('zoomed')  # Windows系统
        # self.root.attributes('-zoomed', True)  # Linux系统

        # 初始化变量
        self.video_path = ""
        self.cap = None
        self.total_frames = 0
        self.fps = 30
        self.duration = 0.0
        self.is_processing = False
        self.preview_scale = 1.0
        self.last_position = 0
        self.preview_timer = None
        self.is_high_res = False  # 是否高分辨率视频
        self.preview_enabled = True  # 控制预览功能的开关
        self.progress_var = tk.DoubleVar()  # 进度条变量
        self.preview_lock = threading.Lock()  # 添加互斥锁
        self.current_preview_task = None  # 当前预览任务的标识
        self.is_preview_loading = False  # 添加预览加载状态标志
        self._update_timer = None  # 初始化_update_timer属性

        # 新增：视频列表相关变量
        self.video_list = []  # 存储视频文件路径列表
        self.is_merging = False  # 视频合并状态标志

        # 软字幕相关变量
        self.soft_subtitle_cap = None
        self.soft_subtitle_timer = None
        self.soft_current_subtitle = ""
        self.soft_subtitles = []
        self.soft_is_previewing = False
        self.soft_is_generating = False

        # 声音处理相关变量
        self.video_audio_file_path = ""  # 视频文件路径
        self.is_video_audio_processing = False  # 声音处理处理状态
        self.video_audio_preview_cap = None  # 视频预览
        self.denoise_settings = {
            'noise_reduction': 10.0,  # 降噪强度
            'preserve_voice': True,   # 保留人声
            'output_format': '自动'    # 输出格式（自动选择）
        }

        # 进程管理
        self.active_processes = []  # 跟踪所有活跃的FFmpeg进程

        # 创建界面组件
        self.create_widgets()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 注册退出处理函数
        atexit.register(self.cleanup_on_exit)

        # 注册信号处理（仅在非Windows系统上）
        if sys.platform != 'win32':
            signal.signal(signal.SIGINT, self.signal_handler)
            signal.signal(signal.SIGTERM, self.signal_handler)

        # 启动时清理可能残留的FFmpeg进程
        self.cleanup_orphaned_processes()

        self.root.mainloop()

    def create_widgets(self):
        """创建界面组件"""
        # 创建标签页控件
        self.tab_control = ttk.Notebook(self.root)
        self.tab_control.pack(fill=tk.BOTH, expand=True)

        # 剪切标签页
        self.trim_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.trim_tab, text="视频剪切")

        # 合并标签页
        self.merge_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.merge_tab, text="视频合并")

        # 硬字幕标签页
        self.subtitle_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.subtitle_tab, text="硬字幕")

        # 创建主框架
        main_frame = tk.Frame(self.subtitle_tab, bg="#333333")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 配置主框架的网格权重
        main_frame.grid_rowconfigure(1, weight=1)  # 预览区域可扩展
        main_frame.grid_rowconfigure(3, weight=0)  # 控制区域固定大小

        # 文件选择区域 - 放在顶部一行
        file_frame = tk.Frame(main_frame, bg="#333333")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 视频和字幕文件选择放在同一行
        ttk.Label(file_frame, text="视频文件：").pack(side=tk.LEFT, padx=(0, 5))
        self.video_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.video_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="选择视频", command=self.select_video).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(file_frame, text="字幕文件：").pack(side=tk.LEFT, padx=(0, 5))
        self.subtitle_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.subtitle_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="选择字幕", command=self.select_subtitle).pack(side=tk.LEFT)

        # 预览区域 - 最大化显示
        preview_frame = tk.Frame(main_frame, bg="#1e1e1e")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建预览画布
        self.subtitle_preview_canvas = tk.Canvas(preview_frame, bg="#1e1e1e", bd=0, highlightthickness=0)
        self.subtitle_preview_canvas.pack(fill=tk.BOTH, expand=True)

        # 字幕显示标签
        self.subtitle_label = tk.Label(
            preview_frame,
            text="",
            bg="#1e1e1e",
            fg="white",
            font=("微软雅黑", 24),
            wraplength=1200
        )
        self.subtitle_label.pack(side=tk.BOTTOM, pady=20)

        # 控制区域 - 底部
        control_frame = tk.Frame(main_frame, bg="#333333")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # 样式设置区域 - 左侧
        style_frame = tk.Frame(control_frame, bg="#333333")
        style_frame.pack(side=tk.LEFT, fill=tk.X)

        # 原视频比特率显示
        ttk.Label(style_frame, text="原视频比特率：").pack(side=tk.LEFT)
        self.original_bitrate_var = tk.StringVar(value="未选择视频")
        self.original_bitrate_label = ttk.Label(style_frame, textvariable=self.original_bitrate_var, foreground="white")
        self.original_bitrate_label.pack(side=tk.LEFT, padx=(0, 15))

        # 字体大小
        ttk.Label(style_frame, text="字号：").pack(side=tk.LEFT)
        self.font_size_var = tk.StringVar(value="12")
        font_sizes = ["12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"]
        font_size_combo = ttk.Combobox(
            style_frame,
            textvariable=self.font_size_var,
            values=font_sizes,
            state="readonly",
            width=4
        )
        font_size_combo.pack(side=tk.LEFT, padx=(0, 15))
        font_size_combo.bind('<<ComboboxSelected>>', lambda e: self.update_subtitle_style())

        # 字幕颜色
        ttk.Label(style_frame, text="颜色：").pack(side=tk.LEFT)
        self.font_color_var = tk.StringVar(value="白色")
        colors = [
            ("白色", "white"),
            ("黑色", "black"),
            ("红色", "red"),
            ("绿色", "green"),
            ("蓝色", "blue"),
            ("黄色", "yellow"),
            ("青色", "cyan"),
            ("洋红色", "magenta"),
            ("橙色", "orange"),
            ("粉色", "pink"),
            ("紫色", "purple"),
            ("灰色", "gray")
        ]
        color_combo = ttk.Combobox(
            style_frame,
            textvariable=self.font_color_var,
            values=[color[0] for color in colors],
            state="readonly",
            width=6
        )
        color_combo.pack(side=tk.LEFT, padx=(0, 15))
        color_combo.bind('<<ComboboxSelected>>', lambda e: self.update_subtitle_style())

        # 保存颜色映射
        self.color_mapping = dict(colors)
        self.reverse_color_mapping = {v: k for k, v in dict(colors).items()}

        # 字幕位置
        ttk.Label(style_frame, text="位置：").pack(side=tk.LEFT)
        self.position_var = tk.StringVar(value="底部")
        positions = [("顶部", "top"), ("中间", "middle"), ("底部", "bottom")]
        position_combo = ttk.Combobox(
            style_frame,
            textvariable=self.position_var,
            values=[pos[0] for pos in positions],
            state="readonly",
            width=6
        )
        position_combo.pack(side=tk.LEFT, padx=(0, 15))
        position_combo.bind('<<ComboboxSelected>>', lambda e: self.update_subtitle_style())

        # 保存位置映射
        self.position_mapping = dict(positions)
        self.reverse_position_mapping = {v: k for k, v in dict(positions).items()}

        # GPU加速选择
        ttk.Label(style_frame, text="GPU加速：").pack(side=tk.LEFT)
        self.gpu_var = tk.StringVar(value="无GPU")
        gpu_options = [
            {"label": "无GPU", "value": ""},
            {"label": "NVIDIA显卡", "value": "h264_nvenc"},
            {"label": "NVIDIA高性能", "value": "h264_nvenc_fast"},
            {"label": "Intel集成显卡", "value": "hevc_qsv"},
            {"label": "AMD 显卡", "value": "av1_amf"},
            {"label": "VAAPI", "value": "h264_vaapi"}
        ]
        gpu_combo = ttk.Combobox(
            style_frame,
            textvariable=self.gpu_var,
            values=[opt["label"] for opt in gpu_options],
            state="readonly",
            width=10
        )
        gpu_combo.pack(side=tk.LEFT, padx=(0, 15))

        # 保存GPU选项映射
        self.gpu_mapping = {opt["label"]: opt["value"] for opt in gpu_options}

        # 合成视频比特率输入
        ttk.Label(style_frame, text="合成比特率：").pack(side=tk.LEFT)
        self.output_bitrate_var = tk.StringVar()
        self.output_bitrate_entry = ttk.Entry(style_frame, textvariable=self.output_bitrate_var, width=8)
        self.output_bitrate_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(style_frame, text="k").pack(side=tk.LEFT)

        # 进度条区域
        progress_slider_frame = tk.Frame(main_frame, bg="#333333")
        progress_slider_frame.pack(fill=tk.X, pady=(0, 10))

        # 播放进度条
        self.subtitle_progress_slider = ttk.Scale(
            progress_slider_frame,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            command=self.on_progress_change
        )
        self.subtitle_progress_slider.pack(fill=tk.X, padx=5)

        # 按钮区域 - 右侧
        button_frame = tk.Frame(control_frame, bg="#333333")
        button_frame.pack(side=tk.RIGHT)

        # 预览按钮
        self.preview_btn = ttk.Button(button_frame, text="预览字幕", command=self.preview_subtitle, width=10)
        self.preview_btn.pack(side=tk.LEFT, padx=5)

        # 生成按钮
        self.generate_btn = ttk.Button(button_frame, text="生成硬字幕视频", command=self.generate_subtitle_video, width=12)
        self.generate_btn.pack(side=tk.LEFT, padx=5)

        # 保存当前进度按钮
        self.save_progress_btn = ttk.Button(button_frame, text="保存当前进度", command=self.save_current_progress, width=12, state="disabled")
        self.save_progress_btn.pack(side=tk.LEFT, padx=5)

        # 进度条 - 最底部
        progress_frame = tk.Frame(main_frame, bg="#333333")
        progress_frame.pack(fill=tk.X)

        self.subtitle_progress_var = tk.DoubleVar()
        self.subtitle_progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.subtitle_progress_var,
            maximum=100
        )
        self.subtitle_progress_bar.pack(fill=tk.X)

        # 初始化字幕相关变量
        self.subtitle_cap = None
        self.subtitle_timer = None
        self.current_subtitle = ""
        self.subtitles = []
        self.is_previewing = False
        self.is_generating = False

        # 软字幕标签页
        self.soft_subtitle_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.soft_subtitle_tab, text="软字幕")
        self.create_soft_subtitle_tab()

        # 声音处理标签页
        self.audio_denoise_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.audio_denoise_tab, text="声音处理")
        self.create_audio_denoise_tab()

        # 在合并标签页中创建视频列表区域
        merge_frame = tk.Frame(self.merge_tab, bg="#333333")
        merge_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 视频列表区域
        list_frame = tk.Frame(merge_frame, bg="#1e1e1e")
        list_frame.pack(fill=tk.BOTH, expand=True)

        # 创建树形视图显示视频列表
        columns = ("文件名", "格式", "大小", "时长", "码率", "帧率")
        self.video_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")

        # 设置列标题和左对齐
        for col in columns:
            self.video_tree.heading(col, text=col, anchor='w', command=lambda c=col: self.sort_tree_column(c))
            self.video_tree.column(col, width=100, anchor='w')

        # 添加排序状态变量
        self.sort_column = None
        self.sort_reverse = False

        self.video_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.video_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.video_tree.configure(yscrollcommand=scrollbar.set)

        # 绑定拖放事件
        self.video_tree.drop_target_register(DND_FILES)
        self.video_tree.dnd_bind('<<Drop>>', self.handle_drop)

        # 按钮区域
        btn_frame = tk.Frame(merge_frame, bg="#333333")
        btn_frame.pack(fill=tk.X, pady=10)

        # 添加视频按钮
        add_btn = ttk.Button(btn_frame, text="添加视频", command=self.add_video)
        add_btn.pack(side=tk.LEFT, padx=5)

        # 删除视频按钮
        remove_btn = ttk.Button(btn_frame, text="删除", command=self.remove_video)
        remove_btn.pack(side=tk.LEFT, padx=5)

        # 上移按钮
        up_btn = ttk.Button(btn_frame, text="上移", command=lambda: self.move_video(-1))
        up_btn.pack(side=tk.LEFT, padx=5)

        # 下移按钮
        down_btn = ttk.Button(btn_frame, text="下移", command=lambda: self.move_video(1))
        down_btn.pack(side=tk.LEFT, padx=5)

        # 合并按钮
        self.merge_btn = ttk.Button(btn_frame, text="合并选中视频", command=self.merge_videos)
        self.merge_btn.pack(side=tk.RIGHT, padx=5)

        # 清空按钮
        clear_btn = ttk.Button(btn_frame, text="清空", command=self.clear_video_list)
        clear_btn.pack(side=tk.RIGHT, padx=5)

        # 合并标签页进度条
        self.merge_progress_bar = ttk.Progressbar(merge_frame, variable=self.progress_var, maximum=100)
        self.merge_progress_bar.pack(fill=tk.X, padx=20, pady=5)

        # 在剪切标签页中创建预览和控制区域
        preview_frame = tk.Frame(self.trim_tab, bg="#333333")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 让preview_frame能够自适应大小
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        # 文件拖放区域
        self.drop_canvas = tk.Canvas(preview_frame, bg="#1e1e1e", bd=0, highlightthickness=0)
        self.drop_canvas.pack(fill=tk.BOTH, expand=True)

        # 绑定重绘事件和拖放事件
        self.drop_canvas.bind("<Configure>", self.redraw_drop_area)
        self.drop_canvas.bind("<Button-1>", self.open_file_dialog)
        # 注册拖放目标
        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind('<<Drop>>', self.handle_drop)

        # 双滑块控件
        self.slider_frame = tk.Frame(preview_frame, bg="#333333")
        self.slider_frame.pack(fill=tk.X, pady=15)

        # 滑块轨道 - 使用相对坐标
        self.track_canvas = tk.Canvas(self.slider_frame, bg="#444444", height=30, bd=0, highlightthickness=0)
        self.track_canvas.pack(fill=tk.X)

        # 绑定重绘事件
        self.track_canvas.bind('<Configure>', self.redraw_track)

        # 时间显示
        self.time_frame = tk.Frame(preview_frame, bg="#333333")
        self.time_frame.pack(fill=tk.X)

        self.start_label = tk.Label(self.time_frame, text="开始时间：00:00:00.000", bg="#333333", fg="white")
        self.start_label.pack(side=tk.LEFT, padx=20)

        self.end_label = tk.Label(self.time_frame, text="结束时间：00:00:00.000", bg="#333333", fg="white")
        self.end_label.pack(side=tk.RIGHT, padx=20)

        # 控制按钮
        self.control_btn = ttk.Button(preview_frame, text="开始剪辑", command=self.toggle_process)
        self.control_btn.pack(pady=5)

        # 进度条
        self.progress_bar = ttk.Progressbar(preview_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=20, pady=5)

    def handle_drop(self, event):
        """处理文件拖放"""
        files = event.data
        if not isinstance(files, str):
            return

        print(f"[DEBUG] 拖放原始数据: {repr(files)}")

        # 处理Windows下的文件路径格式
        # 拖放数据可能是空格分隔或换行符分隔
        files = files.strip()

        # 智能解析文件路径
        file_list = self._parse_dropped_files(files)

        print(f"[DEBUG] 处理后的文件列表: {file_list}")

        # 过滤出视频文件
        video_files = []
        for f in file_list:
            # 清理路径，移除可能的引号
            clean_path = f.strip().strip('"').strip("'")
            if clean_path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.flv', '.ts', '.wmv')):
                video_files.append(clean_path)

        print(f"[DEBUG] 过滤后的视频文件: {video_files}")

    def _parse_dropped_files(self, files_str):
        """智能解析拖放的文件路径"""
        import re

        file_list = []

        # 先尝试按换行符分割（多文件拖放）
        if '\n' in files_str:
            file_list = [f.strip() for f in files_str.split('\n') if f.strip()]
        else:
            # 检查是否包含花括号包围的路径
            if '{' in files_str and '}' in files_str:
                # 使用正则表达式提取所有被花括号包围的路径
                pattern = r'\{([^}]+)\}'
                matches = re.findall(pattern, files_str)
                if matches:
                    file_list = [match.strip() for match in matches if match.strip()]
                else:
                    # 如果没有匹配到，可能是单个花括号包围的路径
                    if files_str.startswith('{') and files_str.endswith('}'):
                        file_list = [files_str.strip('{}')]
            else:
                # 检查是否以各种符号包围
                if files_str.startswith('"') and files_str.endswith('"'):
                    # 双引号包围的路径，直接使用
                    file_list = [files_str.strip('"')]
                elif files_str.startswith("'") and files_str.endswith("'"):
                    # 单引号包围的路径，直接使用
                    file_list = [files_str.strip("'")]
                else:
                    # 没有包围符号，需要判断是否包含空格
                    # 如果包含空格，可能是单个文件路径
                    if ' ' in files_str:
                        # 检查是否是一个有效的文件路径
                        if os.path.exists(files_str):
                            file_list = [files_str]
                        else:
                            # 尝试按空格分割（可能是多个文件）
                            file_list = [f.strip() for f in files_str.split() if f.strip()]
                    else:
                        # 没有空格，直接使用
                        file_list = [files_str]

        return file_list

    def handle_drop(self, event):
        """处理文件拖放"""
        files = event.data
        if not isinstance(files, str):
            return

        print(f"[DEBUG] 拖放原始数据: {repr(files)}")

        # 处理Windows下的文件路径格式
        # 拖放数据可能是空格分隔或换行符分隔
        files = files.strip()

        # 智能解析文件路径
        file_list = self._parse_dropped_files(files)

        print(f"[DEBUG] 处理后的文件列表: {file_list}")

        # 过滤出视频文件
        video_files = []
        for f in file_list:
            # 清理路径，移除可能的引号
            clean_path = f.strip().strip('"').strip("'")
            if clean_path.lower().endswith(('.mp4', '.avi', '.mov', '.mkv', '.flv', '.ts', '.wmv')):
                video_files.append(clean_path)

        print(f"[DEBUG] 过滤后的视频文件: {video_files}")

        if not video_files:
            return

        # 获取当前标签页
        current_tab = self.tab_control.select()
        print(f"[DEBUG] 当前标签页: {current_tab}")
        print(f"[DEBUG] 所有标签页: {self.tab_control.tabs()}")
        print(f"[DEBUG] 第一个标签页: {self.tab_control.tabs()[0]}")

        if current_tab == self.tab_control.tabs()[0]:  # 剪切标签页
            print(f"[DEBUG] 在剪切标签页，加载第一个视频文件")
            # 只处理第一个视频文件
            self.load_video(video_files[0])
        else:  # 合并标签页
            print(f"[DEBUG] 在合并标签页，开始处理 {len(video_files)} 个文件到合并列表")
            for file_path in video_files:
                try:
                    print(f"[DEBUG] 处理文件: {file_path}")

                    # 验证文件是否存在
                    if not os.path.exists(file_path):
                        print(f"[DEBUG] 文件不存在: {file_path}")
                        continue

                    # 获取绝对路径
                    abs_path = os.path.abspath(file_path)
                    print(f"[DEBUG] 绝对路径: {abs_path}")

                    if abs_path not in self.video_list:
                        print(f"[DEBUG] 文件不在列表中，开始添加: {abs_path}")
                        try:
                            # 获取基本文件信息
                            filename = os.path.basename(abs_path)
                            format_type = os.path.splitext(filename)[1][1:].upper()

                            # 获取文件大小，处理特殊字符路径
                            try:
                                file_size = os.path.getsize(abs_path)
                                size_str = self.format_file_size(file_size)
                                print(f"[DEBUG] 文件大小获取成功: {size_str}")
                            except Exception as e:
                                print(f"[DEBUG] 获取文件大小失败: {e}")
                                file_size = 0
                                size_str = "未知"

                            # 使用快速方法获取视频信息，处理特殊字符路径
                            try:
                                # 对于包含特殊字符的路径，使用原始字符串
                                cap = cv2.VideoCapture(abs_path)
                                if cap.isOpened():
                                    # 获取视频属性
                                    fps = cap.get(cv2.CAP_PROP_FPS)
                                    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                                    duration = frame_count / fps if fps > 0 else 0
                                    duration_str = self.format_duration(duration)
                                    bitrate = file_size * 8 / (duration * 1000) if duration > 0 else 0
                                    bitrate_str = f"{bitrate:.2f} Mbps"
                                    fps_str = f"{fps:.2f} fps"
                                    self.video_list.append(abs_path)

                                    # 将信息添加到树形视图
                                    print(f"[DEBUG] 添加到树形视图: {filename}")
                                    self.video_tree.insert("", "end", values=(filename, format_type, size_str, duration_str, bitrate_str, fps_str))
                                    cap.release()
                                    print(f"[DEBUG] 文件添加成功: {filename}")
                                else:
                                    print(f"[DEBUG] 无法打开视频文件: {abs_path}")
                                    # 即使无法打开视频，也添加到列表中，使用默认值
                                    duration_str = "未知"
                                    bitrate_str = "未知"
                                    fps_str = "未知"
                                    self.video_list.append(abs_path)
                                    self.video_tree.insert("", "end", values=(filename, format_type, size_str, duration_str, bitrate_str, fps_str))
                                    print(f"[DEBUG] 文件添加成功（使用默认值）: {filename}")
                            except Exception as e:
                                print(f"[DEBUG] 打开视频文件失败: {e}")
                                # 即使失败，也添加到列表中，使用默认值
                                duration_str = "未知"
                                bitrate_str = "未知"
                                fps_str = "未知"
                                self.video_list.append(abs_path)
                                self.video_tree.insert("", "end", values=(filename, format_type, size_str, duration_str, bitrate_str, fps_str))
                                print(f"[DEBUG] 文件添加成功（使用默认值）: {filename}")
                        except Exception as e:
                            print(f"[DEBUG] 处理文件信息失败: {e}")
                            # 使用最基本的默认值
                            filename = os.path.basename(abs_path)
                            format_type = os.path.splitext(filename)[1][1:].upper()
                            self.video_list.append(abs_path)
                            self.video_tree.insert("", "end", values=(filename, format_type, "未知", "未知", "未知", "未知"))
                            print(f"[DEBUG] 文件添加成功（使用基本默认值）: {filename}")
                    else:
                        print(f"[DEBUG] 文件已在列表中，跳过: {abs_path}")
                except Exception as e:
                    print(f"[DEBUG] 处理文件出错: {file_path}, 错误: {str(e)}")
                    messagebox.showerror("错误", f"无法读取视频文件：{os.path.basename(file_path)}\n错误信息：{str(e)}")

            print(f"[DEBUG] 处理完成，当前列表中有 {len(self.video_list)} 个文件")

    def format_file_size(self, size):
        """格式化文件大小"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.2f} {unit}"
            size /= 1024.0
        return f"{size:.2f} TB"

    def format_duration(self, seconds):
        """格式化视频时长"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = seconds % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:06.3f}"

    def sort_tree_column(self, col):
        """对树形视图的列进行排序"""
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False

        items = [(self.video_tree.set(item, col), item) for item in self.video_tree.get_children('')]
        items.sort(reverse=self.sort_reverse)

        for index, (_, item) in enumerate(items):
            self.video_tree.move(item, '', index)

        # 更新video_list的顺序以匹配UI排序
        self.update_video_list_order()

    def update_video_list_order(self):
        """根据树形视图的显示顺序更新video_list"""
        new_video_list = []
        for item in self.video_tree.get_children(''):
            filename = self.video_tree.item(item)['values'][0]
            # 根据文件名找到对应的完整路径
            for video_path in self.video_list:
                if os.path.basename(video_path) == filename:
                    new_video_list.append(video_path)
                    break

        # 更新video_list
        self.video_list = new_video_list
        print(f"[DEBUG] 排序后视频列表顺序: {[os.path.basename(v) for v in self.video_list]}")

    def add_video(self):
        """添加视频文件"""
        files = filedialog.askopenfilenames(
            title="选择视频文件",
            filetypes=[
                ("视频文件", "*.mp4;*.avi;*.mov;*.mkv;*.flv;*.ts;*.wmv"),
                ("所有文件", "*.*")
            ]
        )
        if files:
            for file_path in files:
                # 模拟拖放事件
                class DummyEvent:
                    def __init__(self, data):
                        self.data = data
                self.handle_drop(DummyEvent(file_path))

    def remove_video(self):
        """删除选中的视频"""
        selected = self.video_tree.selection()
        if selected:
            item = selected[0]
            filename = self.video_tree.item(item)['values'][0]
            # 从列表和树形视图中删除
            for file_path in self.video_list[:]:  # 使用切片创建副本进行迭代
                if os.path.basename(file_path) == filename:
                    self.video_list.remove(file_path)
                    break
            self.video_tree.delete(item)

    def move_video(self, direction):
        """移动视频位置 - 支持多选（Ctrl+点击、Shift+点击）"""
        selected = self.video_tree.selection()
        if not selected:
            return

        # 获取所有选中项的索引，按顺序排序
        selected_indices = []
        for item in selected:
            index = self.video_tree.index(item)
            selected_indices.append((index, item))

        # 按索引排序，确保移动时保持相对顺序
        selected_indices.sort(key=lambda x: x[0])

        if direction == -1:  # 上移
            # 检查是否可以上移（第一个选中项不能是第一个）
            if selected_indices[0][0] == 0:
                return

            # 从前往后移动，避免索引冲突
            for index, item in selected_indices:
                self.video_tree.move(item, '', index - 1)

            # 更新视频列表
            self._update_video_list_from_tree()

        elif direction == 1:  # 下移
            # 检查是否可以下移（最后一个选中项不能是最后一个）
            if selected_indices[-1][0] == len(self.video_list) - 1:
                return

            # 从后往前移动，避免索引冲突
            for index, item in reversed(selected_indices):
                self.video_tree.move(item, '', index + 1)

            # 更新视频列表
            self._update_video_list_from_tree()

    def _update_video_list_from_tree(self):
        """根据树形视图的当前顺序更新video_list"""
        new_video_list = []
        for item in self.video_tree.get_children(''):
            filename = self.video_tree.item(item)['values'][0]
            # 根据文件名找到对应的完整路径
            for video_path in self.video_list:
                if os.path.basename(video_path) == filename:
                    new_video_list.append(video_path)
                    break

        # 更新video_list
        self.video_list = new_video_list
        print(f"[DEBUG] 更新后的视频列表顺序: {[os.path.basename(v) for v in self.video_list]}")

    def clear_video_list(self):
        """清空视频列表"""
        self.video_list.clear()
        for item in self.video_tree.get_children():
            self.video_tree.delete(item)

    def detect_output_format(self):
        """检测输出格式：如果所有输入文件都是同一格式，则使用该格式；否则使用MP4"""
        if not self.video_list:
            return ".mp4"

        # 获取第一个文件的格式
        first_format = os.path.splitext(self.video_list[0])[1].lower()

        # 检查所有文件是否都是同一格式
        for video in self.video_list:
            current_format = os.path.splitext(video)[1].lower()
            if current_format != first_format:
                # 如果格式不一致，默认使用MP4
                return ".mp4"

        # 所有文件格式一致，使用该格式
        return first_format

    def merge_videos(self):
        """合并视频文件"""
        if len(self.video_list) < 2:
            messagebox.showwarning("警告", "请至少添加两个视频文件")
            return

        if self.is_merging:
            messagebox.showinfo("提示", "正在合并中，请等待...")
            return

        print("[DEBUG] 准备合并视频")
        print("[DEBUG] 待合并视频列表:")
        for video in self.video_list:
            print(f"[DEBUG] - {video}")

        # 自动检测输出格式
        output_format = self.detect_output_format()
        print(f"[DEBUG] 检测到的输出格式: {output_format}")

        # 根据检测到的格式构建文件类型过滤器，将检测到的格式放在第一位
        format_names = {
            ".mp4": "MP4文件",
            ".avi": "AVI文件",
            ".mov": "MOV文件",
            ".mkv": "MKV文件",
            ".flv": "FLV文件",
            ".ts": "TS文件",
            ".wmv": "WMV文件"
        }

        # 构建文件类型列表，检测到的格式放在第一位
        detected_format_name = format_names.get(output_format, "MP4文件")
        detected_format_filter = f"*.{output_format[1:]}"  # 去掉点号

        # 构建所有格式列表
        all_formats = [
            ("MP4文件", "*.mp4"),
            ("AVI文件", "*.avi"),
            ("MOV文件", "*.mov"),
            ("MKV文件", "*.mkv"),
            ("FLV文件", "*.flv"),
            ("TS文件", "*.ts"),
            ("WMV文件", "*.wmv"),
            ("所有文件", "*.*")
        ]

        # 将检测到的格式放在第一位，并移除重复项
        filetypes = [(f"{detected_format_name}", detected_format_filter)]
        for format_name, format_filter in all_formats:
            if format_filter != detected_format_filter:  # 避免重复
                filetypes.append((format_name, format_filter))

        print(f"[DEBUG] 文件类型过滤器: {filetypes}")

        # 选择保存路径
        save_path = filedialog.asksaveasfilename(
            title=f"选择保存位置 (推荐: {detected_format_name})",
            defaultextension=output_format,
            filetypes=filetypes
        )

        if not save_path:
            print("[DEBUG] 用户取消了保存路径选择")
            return

        print(f"[DEBUG] 选择的保存路径: {save_path}")

        # 设置合并状态和UI
        self.is_merging = True
        self.progress_var.set(0)  # 重置进度条

        # 禁用相关按钮
        self.disable_merge_buttons()

        # 创建合并线程
        merge_thread = threading.Thread(target=self._merge_videos_thread, args=(save_path,))
        merge_thread.start()

    def disable_merge_buttons(self):
        """禁用合并相关按钮"""
        if hasattr(self, 'merge_btn'):
            self.merge_btn.config(text="合并中...", state="disabled")

    def enable_merge_buttons(self):
        """启用合并相关按钮"""
        if hasattr(self, 'merge_btn'):
            self.merge_btn.config(text="合并选中视频", state="normal")

    def _merge_videos_thread(self, output_path):
        """视频合并线程"""
        process = None
        try:
            print("开始合并视频")
            print(f"输出路径: {output_path}")

            # 创建临时文件列表
            temp_list = os.path.join(os.path.dirname(output_path), "temp_list.txt")
            print(f"[DEBUG] 临时文件列表路径: {temp_list}")
            with open(temp_list, "w", encoding="utf-8") as f:
                for video in self.video_list:
                    # 使用绝对路径，确保路径格式正确
                    abs_video_path = os.path.abspath(video)
                    # 在Windows上使用正斜杠，并转义特殊字符
                    if sys.platform == 'win32':
                        abs_video_path = abs_video_path.replace('\\', '/')
                    f.write(f"file '{abs_video_path}'\n")
                    print(f"[DEBUG] 添加文件到列表: {abs_video_path}")

            # 打印文件列表内容用于调试
            print(f"[DEBUG] 文件列表内容:")
            with open(temp_list, "r", encoding="utf-8") as f:
                for i, line in enumerate(f, 1):
                    print(f"[DEBUG] {i}: {line.strip()}")

            # 使用流复制模式，直接复制所有流，不重新编码
            cmd = [
                FFMPEG_PATH,
                "-f", "concat",
                "-safe", "0",
                "-i", temp_list,
                "-c", "copy",  # 复制所有流而不重新编码，保证原画质和速度
                "-avoid_negative_ts", "1",
                output_path
            ]
            print(f"[DEBUG] 使用流复制模式合并，保证原画质和速度")

            print("执行命令:", " ".join(cmd))

            # 使用subprocess.Popen执行命令，支持进度条更新
            if sys.platform == 'win32':
                creation_flags = subprocess.CREATE_NO_WINDOW  # 合并视频不需要控制台窗口
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in cmd)

                process = subprocess.Popen(
                    cmd_str,
                    shell=True,
                    stdin=subprocess.PIPE,  # 添加stdin管道
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,  # 合并stderr到stdout
                    universal_newlines=True,
                    encoding='utf-8',
                    creationflags=creation_flags
                )
            else:
                creation_flags = 0

                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,  # 添加stdin管道
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,  # 合并stderr到stdout
                    universal_newlines=True,
                    encoding='utf-8',
                    creationflags=creation_flags
                )

            # 将进程添加到跟踪列表
            self.active_processes.append(process)
            print(f"[DEBUG] 启动FFmpeg进程 PID: {process.pid}")

            # 优化的进度更新逻辑 - 针对流复制模式
            import time

            # 计算总文件大小
            total_size = 0
            for video in self.video_list:
                if os.path.exists(video):
                    total_size += os.path.getsize(video)

            print(f"[DEBUG] 总文件大小: {total_size / (1024*1024):.1f} MB")

            # 进度更新变量
            last_size = 0
            start_time = time.time()

            # 初始进度
            self.root.after(0, lambda: self.progress_var.set(5))

            # 处理FFmpeg输出
            timeout_count = 0
            max_timeout = 30  # 30秒超时

            while True:
                # 检查进程是否结束
                if process.poll() is not None:
                    print(f"[DEBUG] FFmpeg进程已结束，返回码: {process.returncode}")
                    break

                # 读取输出
                try:
                    # 使用非阻塞方式读取
                    import select
                    if sys.platform != 'win32':
                        # Linux/Mac使用select
                        ready, _, _ = select.select([process.stdout], [], [], 0.1)
                        if ready:
                            line = process.stdout.readline()
                            if line:
                                print(f"[FFmpeg] {line.strip()}")
                                timeout_count = 0
                    else:
                        # Windows使用简单的readline
                        line = process.stdout.readline()
                        if line:
                            print(f"[FFmpeg] {line.strip()}")
                            timeout_count = 0
                        else:
                            timeout_count += 1
                except Exception as e:
                    print(f"[DEBUG] 读取输出错误: {e}")
                    timeout_count += 1

                # 基于输出文件大小更新进度
                if os.path.exists(output_path):
                    current_size = os.path.getsize(output_path)
                    if current_size > last_size and total_size > 0:
                        # 计算进度百分比，留10%给最终处理
                        progress = min(90, (current_size / total_size) * 90)
                        self.root.after(0, lambda p=progress: self.progress_var.set(p))

                        # 计算速度
                        elapsed_time = time.time() - start_time
                        if elapsed_time > 0:
                            speed_mbps = (current_size / (1024*1024)) / elapsed_time
                            print(f"[DEBUG] 进度: {progress:.1f}% ({current_size / (1024*1024):.1f}MB / {total_size / (1024*1024):.1f}MB) 速度: {speed_mbps:.1f}MB/s")

                        last_size = current_size
                        timeout_count = 0

                # 检查超时
                if timeout_count > max_timeout:
                    print(f"[DEBUG] FFmpeg进程超时，强制终止")
                    process.terminate()
                    break

                # 短暂休眠
                time.sleep(0.1)

            # 安全关闭输出流
            try:
                if process.stdout:
                    process.stdout.close()
                if process.stderr:
                    process.stderr.close()
            except AttributeError:
                pass

            # 等待进程完全结束并获取返回值
            try:
                returncode = process.wait(timeout=300)  # 5分钟超时
                print(f"[DEBUG] FFmpeg进程正常结束，返回码: {returncode}")
            except subprocess.TimeoutExpired:
                print("[DEBUG] FFmpeg进程超时，强制终止")
                process.terminate()
                try:
                    returncode = process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    returncode = -1

            # 清理临时文件
            try:
                os.remove(temp_list)
            except:
                pass

            # 检查结果
            if returncode == 0 and os.path.exists(output_path):
                # 合并成功，设置进度条为100%
                self.root.after(0, lambda: self.progress_var.set(100))
                self.root.after(0, self.handle_merge_completion, returncode, output_path)
            else:
                error_msg = f"合并失败：FFmpeg返回错误代码 {returncode}"
                print(error_msg)
                self.root.after(0, messagebox.showerror, "错误", error_msg)

        except Exception as e:
            print(f"合并失败：{str(e)}")
            self.root.after(0, messagebox.showerror, "错误", str(e))
        finally:
            # 清理进程资源
            if process is not None:
                try:
                    # 从活跃进程列表中移除
                    if process in self.active_processes:
                        self.active_processes.remove(process)
                        print(f"[DEBUG] 从活跃进程列表中移除合并进程 PID: {process.pid}")

                    # 确保进程已完全结束
                    if process.poll() is None:
                        print(f"[DEBUG] 合并进程 {process.pid} 仍在运行，等待其结束...")
                        try:
                            process.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            print(f"[DEBUG] 合并进程 {process.pid} 超时，强制终止")
                            process.terminate()
                            try:
                                process.wait(timeout=2)
                            except subprocess.TimeoutExpired:
                                process.kill()

                    # 关闭进程的stdin管道
                    if hasattr(process, 'stdin') and process.stdin and not process.stdin.closed:
                        try:
                            process.stdin.close()
                        except:
                            pass

                    print(f"[DEBUG] 合并进程 {process.pid} 清理完成")
                except Exception as e:
                    print(f"[DEBUG] 清理合并进程时出错: {e}")

            self.is_merging = False
            # 延迟重置进度条，让用户看到100%完成状态
            self.root.after(2000, lambda: self.progress_var.set(0))  # 2秒后重置
            self.root.after(0, self.enable_merge_buttons)

    def handle_merge_completion(self, returncode, output_path):
        """处理合并完成回调"""
        try:
            logger.info(f"合并完成，返回码: {returncode}")
            logger.info(f"输出路径: {output_path}")

            if returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                # 立即打开输出文件所在目录
                output_dir = os.path.dirname(output_path)
                print(f"[DEBUG] 打开目录: {output_dir}")
                try:
                    # 确保路径格式正确
                    if sys.platform == 'win32':
                        subprocess.run(['explorer', output_dir], creationflags=subprocess.CREATE_NO_WINDOW)
                    else:
                        subprocess.run(['xdg-open', output_dir])
                except Exception as e:
                    logger.error(f"打开目录失败: {str(e)}")
                    print(f"[DEBUG] 打开目录失败: {str(e)}")
                messagebox.showinfo("完成", f"视频合并成功！\n保存路径：{output_path}")
            else:
                error_msg = "合并失败：\n"
                if returncode != 0:
                    error_msg += f"FFmpeg返回错误代码：{returncode}\n"
                if not os.path.exists(output_path):
                    error_msg += "输出文件未生成\n"
                elif os.path.getsize(output_path) == 0:
                    error_msg += "输出文件大小为0\n"
                    try:
                        os.remove(output_path)  # 删除空文件
                    except Exception as e:
                        logger.error(f"删除空文件失败: {str(e)}")

                logger.error(error_msg)
                logger.error("请检查以下信息：")
                logger.error(f"1. 返回码: {returncode}")
                logger.error(f"2. 输出文件是否存在: {os.path.exists(output_path)}")
                if os.path.exists(output_path):
                    logger.error(f"3. 输出文件大小: {os.path.getsize(output_path)}")

                messagebox.showerror("错误", error_msg + "请检查控制台输出获取详细信息")
        except Exception as e:
            error_msg = f"处理合并结果时发生错误：{str(e)}"
            logger.error(error_msg)
            logger.error(traceback.format_exc())
            messagebox.showerror("错误", error_msg)

    def open_file_dialog(self, event=None):
        """打开文件对话框"""
        file_path = filedialog.askopenfilename(
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.mkv *.flv *.ts *.wmv")]
        )
        if file_path:
            self.load_video(file_path)

    def load_video(self, path):
        """加载视频文件"""
        try:
            # 处理文件路径编码问题
            abs_path = os.path.abspath(path)
            try:
                # 尝试使用utf-8编码
                abs_path = abs_path.encode('utf-8', errors='ignore').decode('utf-8')
            except UnicodeEncodeError:
                # 如果utf-8失败，尝试使用系统默认编码
                abs_path = abs_path.encode(sys.getfilesystemencoding(), errors='ignore').decode(sys.getfilesystemencoding())

            self.cap = cv2.VideoCapture(abs_path)
            if not self.cap.isOpened():
                raise Exception("无法打开视频文件")

            # 获取视频信息
            self.fps = self.cap.get(cv2.CAP_PROP_FPS)
            self.total_frames = int(self.cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.duration = self.total_frames / self.fps
            self.video_path = abs_path

            # 分辨率检测
            width = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            self.is_high_res = width > 3840  # 4K以上视为高分辨率
            self.preview_scale = 0.25 if self.is_high_res else 1.0

            # 初始化界面
            self.drop_canvas.delete("all")
            self.show_frame(0)
            self.update_time_labels(0, self.duration)
        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def show_frame(self, position):
        """智能视频帧显示（解决8K卡顿）"""
        # 如果预览功能被禁用或正在处理中，则不更新预览
        if not self.preview_enabled or self.is_processing:
            return

        # 生成新的任务标识
        new_task_id = id(threading.current_thread())
        self.current_preview_task = new_task_id
        self.is_preview_loading = True  # 设置预览加载状态

        # 取消之前的预览任务
        if self.preview_timer:
            self.root.after_cancel(self.preview_timer)
            self.preview_timer = None

        # 使用ffmpeg快速截取当前帧
        if self.cap and self.video_path:
            try:
                # 如果无法获取锁，说明有其他预览任务在进行，直接返回
                if not self.preview_lock.acquire(blocking=False):
                    return

                try:
                    # 再次检查是否是最新的预览任务
                    if self.current_preview_task != new_task_id:
                        return

                    # 如果已经开始处理，则不继续预览
                    if self.is_processing:
                        return

                    # 计算时间戳
                    timestamp = position
                    output_path = os.path.join(os.path.dirname(self.video_path), 'temp_frame.jpg')

                    # 构建优化后的ffmpeg命令，增强硬件加速支持
                    ffmpeg_cmd = [
                        FFMPEG_PATH,
                        '-y',  # 覆盖已存在的文件
                        '-hwaccel', 'auto',  # 自动选择硬件加速
                        '-hwaccel_device', '0',  # 使用第一个可用的硬件设备
                        '-hwaccel_output_format', 'nv12',  # 优化硬件加速输出格式
                        '-ss', f'{timestamp:.3f}',  # 设置时间戳
                        '-i', self.video_path,
                        '-vframes', '1',  # 只截取一帧
                        '-q:v', '5',  # 稍微降低质量以提高速度
                        '-preset', 'ultrafast',  # 使用最快的编码预设
                        '-tune', 'fastdecode',  # 优化解码速度
                        '-f', 'image2',  # 强制使用image2格式
                        output_path
                    ]

                    # 执行ffmpeg命令 - 修复中文路径编码问题
                    if sys.platform == 'win32':
                        # 将命令列表转换为字符串，确保路径编码正确
                        cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in ffmpeg_cmd)
                        subprocess.run(cmd_str, shell=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
                    else:
                        subprocess.run(ffmpeg_cmd, capture_output=True)

                    # 再次检查是否是最新的预览任务
                    if self.current_preview_task != new_task_id:
                        return

                    # 读取并显示图片
                    if os.path.exists(output_path):
                        img = Image.open(output_path)

                        # 自适应显示
                        canvas_width = self.drop_canvas.winfo_width()
                        canvas_height = self.drop_canvas.winfo_height()

                        img_ratio = img.width / img.height
                        canvas_ratio = canvas_width / canvas_height

                        if img_ratio > canvas_ratio:
                            new_width = canvas_width
                            new_height = int(canvas_width / img_ratio)
                        else:
                            new_height = canvas_height
                            new_width = int(canvas_height * img_ratio)

                        # 使用NEAREST方法进行快速缩放
                        img = img.resize((new_width, new_height), Image.Resampling.NEAREST)
                        self.preview_img = ImageTk.PhotoImage(img)

                        # 清除画布并显示新图片
                        self.drop_canvas.delete("all")
                        self.drop_canvas.create_image(
                            canvas_width // 2, canvas_height // 2,
                            image=self.preview_img, anchor=tk.CENTER
                        )

                        # 立即删除临时文件
                        try:
                            os.remove(output_path)
                        except:
                            pass

                        self.is_preview_loading = False  # 预览加载完成
                finally:
                    # 释放锁
                    self.preview_lock.release()

            except Exception as e:
                print(f"预览更新失败: {str(e)}")
                # 确保锁被释放
                try:
                    self.preview_lock.release()
                except RuntimeError:
                    # 如果锁已经被释放则忽略错误
                    pass

    def _show_frame_impl(self, position):
        """实际的帧显示逻辑"""
        if self.cap:
            current_frame = int(position * self.fps)
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, current_frame)
            ret, frame = self.cap.read()
            if ret:
                # 高性能缩放处理
                scale = self.preview_scale  # 统一使用更低的预览质量
                frame = cv2.resize(frame, (0, 0),
                                   fx=scale,
                                   fy=scale,
                                   interpolation=cv2.INTER_NEAREST)  # 使用最快速的插值方法

                # 转换为RGB格式
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame)

                # 自适应显示
                canvas_width = self.drop_canvas.winfo_width()
                canvas_height = self.drop_canvas.winfo_height()

                img_ratio = img.width / img.height
                canvas_ratio = canvas_width / canvas_height

                if img_ratio > canvas_ratio:
                    new_width = canvas_width
                    new_height = int(canvas_width / img_ratio)
                else:
                    new_height = canvas_height
                    new_width = int(canvas_height * img_ratio)

                img = img.resize((new_width, new_height), Image.Resampling.NEAREST)  # 使用最快的重采样方法
                self.preview_img = ImageTk.PhotoImage(img)
                self.drop_canvas.create_image(
                    canvas_width // 2, canvas_height // 2,
                    image=self.preview_img, anchor=tk.CENTER
                )

    def move_start(self, event):
        """移动开始滑块"""
        width = self.track_canvas.winfo_width()
        margin = 20
        track_width = width - (2 * margin)

        # 计算新位置 - 支持精确剪切，保持至少1秒距离
        end_pos = self.get_slider_position("end")
        if end_pos is not None and self.duration > 0:
            # 计算1秒对应的像素距离
            track_width = width - (2 * margin)
            one_second_pixels = (1.0 / self.duration) * track_width
            # 确保开始按钮不会拖动到结束按钮之后，保持至少1秒距离
            x = max(margin, min(event.x, end_pos - one_second_pixels))
        else:
            # 如果无法获取结束按钮位置，使用默认限制
            x = max(margin, min(event.x, width - margin))
        x = min(x, width - margin)

        # 更新开始按钮位置 - 倒水滴形状
        height = self.track_canvas.winfo_height()
        start_icon_top = x - 8, height/2 - 15
        start_icon_mid = x, height/2 - 5
        start_icon_right = x + 8, height/2 - 15
        start_line_top = x, height/2 - 5
        start_line_bottom = x, height/2 + 15

        self.track_canvas.coords("start",
                                 start_icon_top[0], start_icon_top[1],
                                 start_icon_mid[0], start_icon_mid[1],
                                 start_icon_right[0], start_icon_right[1]
                                 )
        self.track_canvas.coords(self.start_line,
                                 start_line_top[0], start_line_top[1],
                                 start_line_bottom[0], start_line_bottom[1]
                                 )

        # 计算时间
        position_ratio = (x - margin) / track_width
        start_time = position_ratio * self.duration
        self.update_time_labels(start=start_time)

        # 立即取消当前预览任务
        if self.current_preview_task:
            self.current_preview_task = None
        if self.preview_timer:
            self.root.after_cancel(self.preview_timer)
            self.preview_timer = None

        # 延迟显示帧，避免频繁更新
        if hasattr(self, '_update_timer') and self._update_timer is not None:
            self.root.after_cancel(self._update_timer)

        # 在拖动结束后恢复预览并更新帧
        def on_drag_end(e=None):
            self.preview_enabled = True
            self._update_timer = self.root.after(100, lambda: self.show_frame(start_time))
            self.track_canvas.unbind('<ButtonRelease-1>')

        # 绑定鼠标释放事件
        self.track_canvas.bind('<ButtonRelease-1>', on_drag_end)

    def move_end(self, event):
        """移动结束滑块"""
        width = self.track_canvas.winfo_width()
        margin = 20
        track_width = width - (2 * margin)

        # 计算新位置 - 支持精确剪切，保持至少1秒距离
        start_pos = self.get_slider_position("start")
        if start_pos is not None and self.duration > 0:
            # 计算1秒对应的像素距离
            track_width = width - (2 * margin)
            one_second_pixels = (1.0 / self.duration) * track_width
            # 确保结束按钮不会拖动到开始按钮之前，保持至少1秒距离
            x = min(width - margin, max(event.x, start_pos + one_second_pixels))
        else:
            # 如果无法获取开始按钮位置，使用默认限制
            x = min(width - margin, max(event.x, margin))
        x = max(x, margin)

        # 更新结束按钮位置 - 倒水滴形状
        height = self.track_canvas.winfo_height()
        end_icon_top = x - 8, height/2 - 15
        end_icon_mid = x, height/2 - 5
        end_icon_right = x + 8, height/2 - 15
        end_line_top = x, height/2 - 5
        end_line_bottom = x, height/2 + 15

        self.track_canvas.coords("end",
                                 end_icon_top[0], end_icon_top[1],
                                 end_icon_mid[0], end_icon_mid[1],
                                 end_icon_right[0], end_icon_right[1]
                                 )
        self.track_canvas.coords(self.end_line,
                                 end_line_top[0], end_line_top[1],
                                 end_line_bottom[0], end_line_bottom[1]
                                 )

        # 计算时间
        position_ratio = (x - margin) / track_width
        end_time = position_ratio * self.duration
        self.update_time_labels(end=end_time)

        # 立即取消当前预览任务
        if self.current_preview_task:
            self.current_preview_task = None
        if self.preview_timer:
            self.root.after_cancel(self.preview_timer)
            self.preview_timer = None

        # 延迟显示帧，避免频繁更新
        if hasattr(self, '_update_timer') and self._update_timer is not None:
            self.root.after_cancel(self._update_timer)

        # 在拖动结束后恢复预览并更新帧
        def on_drag_end(e=None):
            self.preview_enabled = True
            self._update_timer = self.root.after(100, lambda: self.show_frame(end_time))
            self.track_canvas.unbind('<ButtonRelease-1>')

        # 绑定鼠标释放事件
        self.track_canvas.bind('<ButtonRelease-1>', on_drag_end)

    def position_to_time(self, x):
        """将滑块位置转换为时间"""
        width = self.track_canvas.winfo_width()
        margin = 20
        track_width = width - (2 * margin)
        position_ratio = (x - margin) / track_width
        return position_ratio * self.duration

    def get_slider_position(self, slider):
        """获取滑块当前位置"""
        coords = self.track_canvas.coords(slider)
        if not coords or len(coords) < 6:
            return None
        # 对于倒水滴形状，取中间点的x坐标（底部尖点的x坐标）
        return coords[2]

    def update_time_labels(self, start=None, end=None):
        """更新时间显示"""
        if start is not None:
            self.start_label.config(text=f"开始时间：{self.format_time(start)}")
        if end is not None:
            self.end_label.config(text=f"结束时间：{self.format_time(end)}")

    def format_time(self, seconds):
        """优化后的时间格式化"""
        total_seconds = int(seconds)
        milliseconds = int((seconds - total_seconds) * 1000)
        hours = total_seconds // 3600
        remainder = total_seconds % 3600
        minutes = remainder // 60
        seconds = remainder % 60
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}.{milliseconds:03d}"

    def toggle_process(self):
        """切换处理状态"""
        if not self.is_processing:
            self.start_process()
        else:
            self.stop_process()

    def start_process(self):
        """开始处理"""
        if not self.video_path:
            messagebox.showerror("错误", "请先选择视频文件")
            return

        start_pos = self.get_slider_position("start")
        end_pos = self.get_slider_position("end")

        if start_pos is None or end_pos is None:
            messagebox.showerror("错误", "无法获取滑块位置")
            return

        start_time = self.position_to_time(start_pos)
        end_time = self.position_to_time(end_pos)

        # 允许按钮重叠，但确保时间差至少为0.1秒（支持精确剪切）
        if end_time - start_time < 0.1:
            messagebox.showerror("错误", "剪切区间太短，至少需要0.1秒")
            return

        try:
            # 如果预览正在加载，立即取消预览任务
            if self.is_preview_loading:
                self.current_preview_task = None
                if self.preview_timer:
                    self.root.after_cancel(self.preview_timer)
                    self.preview_timer = None

            # 禁用预览功能
            self.preview_enabled = False
            # 重置进度条
            self.progress_var.set(0)
            output_path = self.generate_output_path()

            # 使用流复制模式进行剪切，避免重新编码
            ffmpeg_cmd = [
                FFMPEG_PATH,
                '-y',
                '-ss', f'{start_time:.3f}',
                '-i', self.video_path,
                '-t', f'{end_time - start_time:.3f}',
                '-c', 'copy',  # 复制所有流而不重新编码
                '-avoid_negative_ts', '1',
                output_path
            ]

            # 在运行前打印命令（重要调试信息）
            print("[DEBUG] 执行命令：", ' '.join(ffmpeg_cmd))
            # 启动处理线程
            self.process_thread = threading.Thread(
                target=self.run_ffmpeg,
                args=(ffmpeg_cmd, output_path)
            )
            self.process_thread.start()
            self.is_processing = True
            self.control_btn.config(text="停止剪辑")

        except Exception as e:
            self.is_processing = False
            self.preview_enabled = True  # 确保预览功能被重新启用
            messagebox.showerror("错误", f"启动失败: {str(e)}")

    def run_ffmpeg(self, cmd, output_path):
        """执行FFmpeg命令"""
        process = None
        try:
            # Windows系统特殊处理 - 修复中文路径编码问题
            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in cmd)
                creation_flags = subprocess.CREATE_NO_WINDOW  # 剪切功能不需要控制台窗口
                encoding = 'utf-8'

                # 使用shell=True来正确处理中文路径
                process = subprocess.Popen(
                    cmd_str,
                    shell=True,
                    stdin=subprocess.PIPE,  # 添加stdin管道
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    universal_newlines=True,
                    encoding=encoding,
                    creationflags=creation_flags
                )
            else:
                creation_flags = 0
                encoding = 'utf-8'

                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,  # 添加stdin管道
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    universal_newlines=True,
                    encoding=encoding,
                    creationflags=creation_flags
                )

            # 将进程添加到跟踪列表
            self.active_processes.append(process)
            print(f"[DEBUG] 启动FFmpeg剪切进程 PID: {process.pid}")

            # 实时输出处理日志并更新进度条
            duration = None
            time = 0
            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    print(line.strip())

                    # 解析FFmpeg输出以更新进度
                    if 'Duration' in line:
                        try:
                            duration_str = line.split('Duration: ')[1].split(',')[0].strip()
                            h, m, s = map(float, duration_str.split(':'))
                            duration = h * 3600 + m * 60 + s
                        except:
                            pass
                    elif 'time=' in line:
                        try:
                            time_str = line.split('time=')[1].split(' ')[0].strip()
                            h, m, s = map(float, time_str.split(':'))
                            time = h * 3600 + m * 60 + s
                            if duration:
                                progress = (time / duration) * 100
                                self.root.after(0, lambda p=progress: self.progress_var.set(min(p, 100)))
                        except:
                            pass

                    if process.poll() is not None:
                        break

            # 安全关闭输出流
            try:
                if process.stdout:
                    process.stdout.close()
            except AttributeError:
                pass

            # 等待进程完全结束并获取返回值
            process.wait()
            returncode = process.returncode

            # 处理完成回调
            if self.is_merging:
                self.root.after(0, self.handle_merge_completion, returncode, output_path)
            else:
                self.root.after(0, self.handle_completion, returncode, output_path)

        except Exception as e:
            self.root.after(0, messagebox.showerror, "错误", f"执行失败: {str(e)}")
        finally:
            # 清理进程资源
            if process is not None:
                try:
                    # 从活跃进程列表中移除
                    if process in self.active_processes:
                        self.active_processes.remove(process)
                        print(f"[DEBUG] 从活跃进程列表中移除剪切进程 PID: {process.pid}")

                    # 确保进程已完全结束
                    if process.poll() is None:
                        print(f"[DEBUG] 进程 {process.pid} 仍在运行，等待其结束...")
                        try:
                            process.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            print(f"[DEBUG] 进程 {process.pid} 超时，强制终止")
                            process.terminate()
                            try:
                                process.wait(timeout=2)
                            except subprocess.TimeoutExpired:
                                process.kill()

                    # 关闭进程的stdin管道
                    if hasattr(process, 'stdin') and process.stdin and not process.stdin.closed:
                        try:
                            process.stdin.close()
                        except:
                            pass

                    print(f"[DEBUG] 剪切进程 {process.pid} 清理完成")
                except Exception as e:
                    print(f"[DEBUG] 清理剪切进程时出错: {e}")

            self.is_processing = False
            self.preview_enabled = True  # 重新启用预览功能
            self.root.after(0, lambda: self.control_btn.config(text="开始剪辑"))
            self.root.after(0, lambda: self.progress_var.set(0))  # 重置进度条

    def handle_completion(self, returncode, output_path):
        """处理完成回调"""
        # 检查输出文件是否存在且大小大于0
        if returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            # 立即打开输出文件所在目录
            output_dir = os.path.dirname(output_path)
            subprocess.run(['explorer', output_dir], creationflags=subprocess.CREATE_NO_WINDOW)
            messagebox.showinfo("完成", f"视频处理成功！\n保存路径：{output_path}")
        else:
            error_msg = "处理失败：\n"
            if returncode != 0:
                error_msg += f"FFmpeg返回错误代码：{returncode}\n"
            if not os.path.exists(output_path):
                error_msg += "输出文件未生成\n"
            elif os.path.getsize(output_path) == 0:
                error_msg += "输出文件大小为0\n"
                try:
                    os.remove(output_path)  # 删除空文件
                except Exception as e:
                    logger.error(f"删除空文件失败: {str(e)}")

            logger.error(error_msg)
            logger.error("请检查以下信息：")
            logger.error(f"1. 返回码: {returncode}")
            logger.error(f"2. 输出文件是否存在: {os.path.exists(output_path)}")
            if os.path.exists(output_path):
                logger.error(f"3. 输出文件大小: {os.path.getsize(output_path)}")

            messagebox.showerror("错误", error_msg + "请检查控制台输出获取详细信息")

    def stop_process(self):
        """停止处理"""
        if hasattr(self, 'process_thread') and self.process_thread.is_alive():
            self.is_processing = False
            self.process_thread.join(timeout=5)
        self.preview_enabled = True  # 重新启用预览功能
        self.progress_var.set(0)  # 重置进度条
        self.control_btn.config(text="开始剪辑")

    def generate_output_path(self):
        """生成合法输出路径"""
        base, ext = os.path.splitext(self.video_path)
        counter = 1
        while True:
            new_path = f"{base}_trimmed_{counter}{ext}"
            if not os.path.exists(new_path):
                return new_path
            counter += 1

    def select_video(self):
        """选择视频文件"""
        file_path = filedialog.askopenfilename(
            title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.mkv *.flv *.ts *.wmv")]
        )
        if file_path:
            self.video_path_var.set(file_path)
            self.load_subtitle_video(file_path)

    def select_subtitle(self):
        """选择字幕文件"""
        file_path = filedialog.askopenfilename(
            title="选择字幕文件",
            filetypes=[("字幕文件", "*.srt *.ass")]
        )
        if file_path:
            self.subtitle_path_var.set(file_path)
            self.load_subtitles(file_path)

    def select_soft_video(self):
        """选择软字幕视频文件"""
        file_path = filedialog.askopenfilename(
            title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.mkv *.flv *.ts *.wmv")]
        )
        if file_path:
            self.soft_video_path_var.set(file_path)
            self.load_soft_subtitle_video(file_path)

    def select_soft_subtitle(self):
        """选择软字幕文件"""
        file_path = filedialog.askopenfilename(
            title="选择字幕文件",
            filetypes=[("字幕文件", "*.srt *.ass")]
        )
        if file_path:
            self.soft_subtitle_path_var.set(file_path)
            self.load_soft_subtitles(file_path)

    def select_video_for_audio_denoise(self):
        """选择用于音频降噪的视频文件"""
        file_path = filedialog.askopenfilename(
            title="选择视频文件",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.mkv *.flv *.ts *.wmv")]
        )
        if file_path:
            self.video_audio_file_var.set(file_path)

    def on_video_audio_file_changed(self, *args):
        """当视频文件路径改变时的回调"""
        video_path = self.video_audio_file_var.get()
        if video_path and os.path.exists(video_path):
            try:
                # 加载视频信息
                if self.video_audio_preview_cap:
                    self.video_audio_preview_cap.release()
                self.video_audio_preview_cap = cv2.VideoCapture(video_path)
                if self.video_audio_preview_cap.isOpened():
                    fps = self.video_audio_preview_cap.get(cv2.CAP_PROP_FPS)
                    total_frames = int(self.video_audio_preview_cap.get(cv2.CAP_PROP_FRAME_COUNT))
                    duration = total_frames / fps if fps > 0 else 0
                    self.video_audio_info_label.config(
                        text=f"视频文件: {os.path.basename(video_path)}\n"
                             f"时长: {duration:.2f}秒\n"
                             f"帧率: {fps:.2f} fps"
                    )
            except Exception as e:
                self.video_audio_info_label.config(text=f"加载视频失败: {str(e)}")

    def on_noise_reduction_changed(self, *args):
        """降噪强度改变时的回调"""
        value = self.noise_reduction_var.get()
        self.noise_value_label.config(text=f"{value:.1f}")

    def on_volume_boost_changed(self, *args):
        """音量放大改变时的回调"""
        value = self.volume_boost_var.get()
        self.volume_value_label.config(text=f"{value:.1f}dB")

    def preview_video_audio(self):
        """预览视频"""
        video_path = self.video_audio_file_var.get()
        if not video_path or not os.path.exists(video_path):
            messagebox.showerror("错误", "请先选择视频文件")
            return

        try:
            if not self.video_audio_preview_cap or not self.video_audio_preview_cap.isOpened():
                self.video_audio_preview_cap = cv2.VideoCapture(video_path)

            ret, frame = self.video_audio_preview_cap.read()
            if ret:
                # 显示第一帧
                self.show_video_audio_frame(frame)
            else:
                messagebox.showinfo("提示", "无法读取视频帧")
        except Exception as e:
            messagebox.showerror("错误", f"预览失败: {str(e)}")

    def show_video_audio_frame(self, frame):
        """显示视频音频预览帧"""
        try:
            # 获取画布尺寸
            canvas_width = self.audio_preview_canvas.winfo_width()
            canvas_height = self.audio_preview_canvas.winfo_height()

            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width = 800
                canvas_height = 600

            # 调整帧大小以适应画布
            frame_height, frame_width = frame.shape[:2]
            scale = min(canvas_width / frame_width, canvas_height / frame_height)
            new_width = int(frame_width * scale)
            new_height = int(frame_height * scale)

            if new_width > 0 and new_height > 0:
                resized_frame = cv2.resize(frame, (new_width, new_height))
                rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)

                from PIL import Image, ImageTk
                pil_image = Image.fromarray(rgb_frame)
                photo = ImageTk.PhotoImage(image=pil_image)

                self.audio_preview_canvas.delete("all")
                self.audio_preview_canvas.create_image(
                    canvas_width // 2, canvas_height // 2,
                    image=photo, anchor=tk.CENTER
                )
                self.audio_preview_canvas.image = photo
        except Exception as e:
            print(f"显示视频帧失败: {str(e)}")

    def start_video_audio_denoise(self):
        """开始声音处理"""
        video_path = self.video_audio_file_var.get()
        if not video_path or not os.path.exists(video_path):
            messagebox.showerror("错误", "请先选择视频文件")
            return

        if self.is_video_audio_processing:
            messagebox.showinfo("提示", "正在处理中，请等待...")
            return

        # 选择保存路径
        video_dir = os.path.dirname(video_path)
        video_name = os.path.splitext(os.path.basename(video_path))[0]
        default_filename = f"{video_name}-denoised.mp4"
        default_path = os.path.join(video_dir, default_filename)

        save_path = filedialog.asksaveasfilename(
            title="选择保存位置",
            initialfile=default_path,
            defaultextension=".mp4",
            filetypes=[("MP4文件", "*.mp4"), ("MKV文件", "*.mkv")]
        )

        if not save_path:
            return

        try:
            # 获取参数
            noise_reduction = self.noise_reduction_var.get()
            volume_boost = self.volume_boost_var.get()
            preserve_voice = self.preserve_voice_var.get()

            # 构建FFmpeg命令
            video_path_clean = os.path.abspath(video_path)

            # 音频滤镜
            audio_filters = []

            # 降噪滤镜
            if noise_reduction > 0:
                if preserve_voice:
                    # 使用highpass和lowpass保留人声频率范围
                    audio_filters.append(f"highpass=f=80,lowpass=f=8000")
                audio_filters.append(f"anlmdn=s={noise_reduction}")

            # 音量放大
            if volume_boost > 0:
                # 将dB转换为线性增益
                gain = 10 ** (volume_boost / 20)
                audio_filters.append(f"volume={gain}")

            # 构建FFmpeg命令
            ffmpeg_cmd = [
                FFMPEG_PATH,
                '-y',
                '-i', video_path_clean,
                '-c:v', 'copy',  # 复制视频流
            ]

            if audio_filters:
                ffmpeg_cmd.extend(['-af', ','.join(audio_filters)])
            else:
                ffmpeg_cmd.extend(['-c:a', 'copy'])  # 如果没有滤镜，复制音频流

            ffmpeg_cmd.append(save_path)

            print("FFmpeg命令:", " ".join(ffmpeg_cmd))

            # 开始处理
            self.is_video_audio_processing = True
            self.audio_progress_var.set(0)
            self.video_audio_denoise_btn.config(state='disabled')
            self.video_audio_stop_btn.config(state='normal')
            self.video_audio_status_label.config(text="正在处理...")

            # 启动处理线程
            process_thread = threading.Thread(
                target=self.run_video_audio_denoise,
                args=(ffmpeg_cmd, save_path)
            )
            process_thread.start()

        except Exception as e:
            print(f"错误：处理失败 - {str(e)}")
            self.is_video_audio_processing = False
            messagebox.showerror("错误", f"处理失败: {str(e)}")
            self.video_audio_denoise_btn.config(state='normal')
            self.video_audio_stop_btn.config(state='disabled')

    def stop_video_audio_denoise(self):
        """停止声音处理"""
        self.is_video_audio_processing = False
        self.video_audio_status_label.config(text="已停止")
        self.video_audio_denoise_btn.config(state='normal')
        self.video_audio_stop_btn.config(state='disabled')

    def run_video_audio_denoise(self, cmd, output_path):
        """执行声音处理"""
        process = None
        final_returncode = -1
        try:
            print("\n=== 开始执行声音处理FFmpeg命令 ===")
            print("1. 检查FFmpeg路径...")
            if not os.path.exists(FFMPEG_PATH):
                raise Exception(f"FFmpeg不存在: {FFMPEG_PATH}")
            print(f"FFmpeg路径: {FFMPEG_PATH}")

            print("2. 检查输入文件...")
            video_path = self.video_audio_file_var.get()
            if not os.path.exists(video_path):
                raise Exception(f"输入视频不存在: {video_path}")
            print(f"输入视频: {video_path}")

            # 获取视频时长用于进度计算
            video_duration = 0
            try:
                info_cmd = [FFMPEG_PATH, '-i', video_path]
                if sys.platform == 'win32':
                    cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in info_cmd)
                    info_result = subprocess.run(
                        cmd_str,
                        shell=True,
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    info_result = subprocess.run(
                        info_cmd,
                        capture_output=True,
                        text=True,
                        encoding='utf-8'
                    )

                for line in info_result.stderr.split('\n'):
                    if 'Duration:' in line:
                        duration = line.split('Duration: ')[1].split(',')[0].strip()
                        h, m, s = map(float, duration.split(':'))
                        video_duration = h * 3600 + m * 60 + s
                        break
            except Exception as e:
                print(f"获取视频时长失败: {e}，将无法显示准确进度")

            print(f"视频时长: {video_duration:.2f}秒")

            print("3. 执行FFmpeg命令...")
            print("命令:", " ".join(cmd))

            # 使用subprocess.Popen执行命令
            if sys.platform == 'win32':
                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    preexec_fn=os.setsid
                )

            # 将进程添加到跟踪列表
            self.active_processes.append(process)
            print(f"[DEBUG] 启动声音处理进程 PID: {process.pid}")

            # 实时读取输出并更新进度
            import time
            last_progress = 0
            last_file_size = 0
            file_size_stable_count = 0
            last_output_time = time.time()

            while True:
                # 检查进程是否已退出
                if process.poll() is not None:
                    print("[DEBUG] 声音处理FFmpeg进程已退出")
                    break

                # 检查是否应该停止处理
                if not self.is_video_audio_processing:
                    print("[DEBUG] 收到停止信号，终止声音处理进程...")
                    process.terminate()
                    break

                try:
                    # 读取一行输出（非阻塞）
                    import select
                    if sys.platform != 'win32':
                        ready, _, _ = select.select([process.stderr], [], [], 0.1)
                        if ready:
                            line = process.stderr.readline()
                        else:
                            line = None
                    else:
                        line = process.stderr.readline()

                    if line:
                        line = line.strip()
                        print(line)
                        last_output_time = time.time()

                        # 解析进度信息
                        if 'time=' in line:
                            try:
                                time_str = line.split('time=')[1].split(' ')[0].strip()
                                h, m, s = map(float, time_str.split(':'))
                                current_time = h * 3600 + m * 60 + s

                                # 计算进度百分比
                                if video_duration > 0:
                                    progress = (current_time / video_duration) * 100
                                    progress = min(progress, 98)

                                    if progress > last_progress:
                                        last_progress = progress
                                        self.root.after(0, lambda p=progress: self.audio_progress_var.set(p))
                                        self.root.after(0, lambda p=progress: self.video_audio_status_label.config(text=f"处理中... {p:.1f}%"))
                                        print(f"进度: {progress:.1f}% ({current_time:.1f}s / {video_duration:.1f}s)")
                            except Exception as e:
                                print(f"进度解析错误: {e}")
                except Exception as e:
                    if "readline" not in str(e).lower():
                        print(f"读取输出错误: {e}")

                # 监控输出文件大小变化
                if os.path.exists(output_path):
                    try:
                        current_file_size = os.path.getsize(output_path)
                        if current_file_size > last_file_size:
                            last_file_size = current_file_size
                            file_size_stable_count = 0
                        elif current_file_size == last_file_size and current_file_size > 0:
                            file_size_stable_count += 1
                        else:
                            file_size_stable_count = 0

                        if file_size_stable_count > 50:
                            print(f"[DEBUG] 文件大小已稳定 {file_size_stable_count * 0.1:.1f}秒")
                    except:
                        pass

                # 检查长时间无输出
                elapsed_since_output = time.time() - last_output_time
                if elapsed_since_output > 30:
                    if process.poll() is None:
                        print(f"[DEBUG] 超过30秒无输出，但进程仍在运行")
                        if os.path.exists(output_path):
                            file_size = os.path.getsize(output_path)
                            if file_size > 0 and last_progress < 95:
                                self.root.after(0, lambda: self.audio_progress_var.set(95))
                                self.root.after(0, lambda: self.video_audio_status_label.config(text="处理中... 95%"))
                                print(f"[DEBUG] 设置进度为95%")

                time.sleep(0.1)

            # 获取最终返回码
            returncode = process.poll()
            if returncode is None:
                try:
                    returncode = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    returncode = process.returncode if process.returncode is not None else -1

            # 读取剩余输出
            try:
                if process.stdin and not process.stdin.closed:
                    process.stdin.close()
                stdout, stderr = process.communicate(timeout=2)
            except:
                stdout, stderr = "", ""

            final_returncode = returncode

            print("4. 检查执行结果...")
            print(f"返回码: {final_returncode}")

            # 处理结果
            if final_returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print("声音处理完成！")
                self.root.after(0, lambda: self.audio_progress_var.set(100))
                self.root.after(0, lambda: self.video_audio_status_label.config(text="处理完成"))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"处理完成:\n{output_path}"))
            else:
                error_msg = f"处理失败（返回码: {final_returncode}）"
                if stderr:
                    error_msg += f"\n{stderr[-500:]}"
                print(error_msg)
                self.root.after(0, lambda msg=error_msg: messagebox.showerror("错误", msg))
                self.root.after(0, lambda: self.audio_progress_var.set(0))
                self.root.after(0, lambda: self.video_audio_status_label.config(text="处理失败"))

        except Exception as e:
            error_msg = f"执行失败: {str(e)}"
            print(error_msg)
            logger.error(error_msg)
            self.root.after(0, lambda msg=error_msg: messagebox.showerror("错误", msg))
            self.root.after(0, lambda: self.audio_progress_var.set(0))
            self.root.after(0, lambda: self.video_audio_status_label.config(text="处理失败"))
        finally:
            print("5. 清理状态和进程资源...")

            # 清理进程资源
            if process is not None:
                try:
                    process_pid = process.pid if hasattr(process, 'pid') else 'unknown'
                    print(f"[DEBUG] 开始清理声音处理进程 PID: {process_pid}")

                    # 关闭所有管道
                    for pipe_name in ['stdin', 'stdout', 'stderr']:
                        pipe = getattr(process, pipe_name, None)
                        if pipe and not pipe.closed:
                            try:
                                pipe.close()
                                print(f"[DEBUG] 已关闭 {pipe_name} 管道")
                            except Exception as e:
                                print(f"[DEBUG] 关闭 {pipe_name} 管道时出错: {e}")

                    # 确保进程已完全结束
                    if process.poll() is None:
                        print(f"[DEBUG] 声音处理进程 {process_pid} 仍在运行，尝试终止...")
                        try:
                            process.terminate()
                            try:
                                process.wait(timeout=2)
                                print(f"[DEBUG] 进程 {process_pid} 已通过terminate结束")
                            except subprocess.TimeoutExpired:
                                print(f"[DEBUG] 进程 {process_pid} terminate后仍运行，强制kill")
                                try:
                                    process.kill()
                                    process.wait(timeout=1)
                                    print(f"[DEBUG] 进程 {process_pid} 已通过kill结束")
                                except:
                                    print(f"[DEBUG] 警告：进程 {process_pid} kill后仍可能运行")
                        except Exception as e:
                            print(f"[DEBUG] 终止进程时出错: {e}")

                    # 从活跃进程列表中移除
                    if process in self.active_processes:
                        self.active_processes.remove(process)
                        print(f"[DEBUG] 从活跃进程列表中移除声音处理进程 PID: {process_pid}")

                    print(f"[DEBUG] 声音处理进程 {process_pid} 清理完成")
                except Exception as e:
                    print(f"[DEBUG] 清理声音处理进程时出错: {e}")

            # 更新状态标志
            self.is_video_audio_processing = False

            # 重新启用按钮
            self.root.after(0, lambda: self.video_audio_denoise_btn.config(state='normal'))
            self.root.after(0, lambda: self.video_audio_stop_btn.config(state='disabled'))
            self.root.after(0, lambda: self.video_audio_preview_btn.config(state='normal'))

            # 根据最终返回码处理进度条
            final_rc = final_returncode if 'final_returncode' in locals() else (process.returncode if process and process.returncode is not None else -1)
            if final_rc != 0:
                self.root.after(0, lambda: self.audio_progress_var.set(0))
                print(f"[DEBUG] 处理失败（返回码: {final_rc}），重置进度条")

            print("=== 声音处理FFmpeg命令执行完成 ===\n")

    def load_subtitle_video(self, video_path):
        """加载视频文件用于字幕预览"""
        try:
            if self.subtitle_cap:
                self.subtitle_cap.release()
            self.subtitle_cap = cv2.VideoCapture(video_path)
            if not self.subtitle_cap.isOpened():
                raise Exception("无法打开视频文件")

            # 获取视频信息
            self.subtitle_fps = self.subtitle_cap.get(cv2.CAP_PROP_FPS)
            self.subtitle_total_frames = int(self.subtitle_cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.subtitle_duration = self.subtitle_total_frames / self.subtitle_fps

            # 使用ffprobe获取视频比特率
            self.get_video_bitrate(video_path)

            # 显示第一帧
            ret, frame = self.subtitle_cap.read()
            if ret:
                self.show_subtitle_frame(frame)
        except Exception as e:
            messagebox.showerror("错误", f"加载视频失败: {str(e)}")

    def load_soft_subtitle_video(self, video_path):
        """加载软字幕视频文件用于预览"""
        try:
            if self.soft_subtitle_cap:
                self.soft_subtitle_cap.release()
            self.soft_subtitle_cap = cv2.VideoCapture(video_path)
            if not self.soft_subtitle_cap.isOpened():
                raise Exception("无法打开视频文件")

            # 获取视频信息（保存用于进度条）
            self.soft_subtitle_fps = self.soft_subtitle_cap.get(cv2.CAP_PROP_FPS)
            self.soft_subtitle_total_frames = int(self.soft_subtitle_cap.get(cv2.CAP_PROP_FRAME_COUNT))
            self.soft_subtitle_duration = self.soft_subtitle_total_frames / self.soft_subtitle_fps if self.soft_subtitle_fps > 0 else 0

            # 显示第一帧
            ret, frame = self.soft_subtitle_cap.read()
            if ret:
                # 显示视频帧到预览画布
                self.show_soft_subtitle_frame(frame)
        except Exception as e:
            messagebox.showerror("错误", f"加载视频失败: {str(e)}")

    def create_styled_ass_file(self, input_subtitle_path, output_ass_path, font_size, font_color_chinese, font_position):
        """创建带样式的ASS字幕文件"""
        # 读取原始字幕文件
        with open(input_subtitle_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 检测字幕格式
        subtitle_ext = os.path.splitext(input_subtitle_path)[1].lower()
        is_ass = subtitle_ext in ['.ass', '.ssa']

        # 解析字幕条目
        subtitle_entries = []

        if is_ass:
            # 解析ASS格式
            # 查找Events章节
            events_start = content.find('[Events]')
            if events_start != -1:
                events_content = content[events_start:]
                # 查找Format行
                format_line = None
                for line in events_content.split('\n'):
                    if line.strip().startswith('Format:'):
                        format_line = line.strip()
                        break

                # 解析Dialogue行
                for line in events_content.split('\n'):
                    if line.strip().startswith('Dialogue:'):
                        parts = line.split(',', 9)  # ASS格式最多10个字段
                        if len(parts) >= 10:
                            start_time = parts[1].strip()
                            end_time = parts[2].strip()
                            text = parts[9].strip()
                            subtitle_entries.append({
                                'start': self._parse_ass_time(start_time),
                                'end': self._parse_ass_time(end_time),
                                'text': text
                            })
        else:
            # 解析SRT格式
            blocks = content.strip().split('\n\n')
            for block in blocks:
                lines = block.strip().split('\n')
                if len(lines) >= 3:
                    time_line = lines[1]
                    if ' --> ' in time_line:
                        start_time, end_time = time_line.split(' --> ')
                        text = ' '.join(lines[2:])
                        # 转换SRT时间格式到秒
                        start_seconds = self._parse_srt_time(start_time)
                        end_seconds = self._parse_srt_time(end_time)
                        subtitle_entries.append({
                            'start': start_seconds,
                            'end': end_seconds,
                            'text': text
                        })

        # 获取样式参数
        font_color_english = self.color_mapping.get(font_color_chinese, 'white')
        position_english = self.position_mapping.get(font_position, 'bottom')

        # 转换颜色格式（ASS使用BGR格式）
        color_map = {
            'white': '&H00FFFFFF&',
            'black': '&H00000000&',
            'red': '&H000000FF&',
            'green': '&H0000FF00&',
            'blue': '&H00FF0000&',
            'yellow': '&H0000FFFF&',
            'cyan': '&H00FFFF00&',
            'magenta': '&H00FF00FF&',
            'orange': '&H0000A5FF&',
            'pink': '&H00C0C0FF&',
            'purple': '&H00800080&',
            'gray': '&H00808080&'
        }
        primary_color = color_map.get(font_color_english, '&H00FFFFFF&')

        # 位置对齐方式（ASS Alignment值：1=左下, 2=中下, 3=右下, 4=左中, 5=中中, 6=右中, 7=左上, 8=中上, 9=右上）
        position_map = {
            'top': '8',      # 中上
            'middle': '5',   # 中中
            'bottom': '2'    # 中下（默认）
        }
        alignment = position_map.get(position_english, '2')

        # 创建ASS文件内容
        ass_content = f"""[Script Info]
Title: Styled Subtitles
ScriptType: v4.00+
ScaledBorderAndShadow: yes

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Microsoft YaHei,{font_size},{primary_color},&H000000FF,&H00000000,&H80000000,0,0,0,0,100,100,0,0,1,2,0,{alignment},10,10,10,1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""

        # 添加字幕条目
        for entry in subtitle_entries:
            start_time_str = self._seconds_to_ass_time(entry['start'])
            end_time_str = self._seconds_to_ass_time(entry['end'])
            text = entry['text'].replace('\n', '\\N')  # ASS格式的换行符
            ass_content += f"Dialogue: 0,{start_time_str},{end_time_str},Default,,0,0,0,,{text}\n"

        # 写入文件
        with open(output_ass_path, 'w', encoding='utf-8') as f:
            f.write(ass_content)

    def _parse_srt_time(self, time_str):
        """解析SRT时间格式 (HH:MM:SS,mmm) 为秒"""
        time_str = time_str.strip()
        # 处理逗号或点作为毫秒分隔符
        if ',' in time_str:
            time_part, ms_part = time_str.split(',')
            milliseconds = int(ms_part)
        elif '.' in time_str:
            time_part, ms_part = time_str.split('.')
            milliseconds = int(ms_part)
        else:
            time_part = time_str
            milliseconds = 0

        parts = time_part.split(':')
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2])

        return hours * 3600 + minutes * 60 + seconds + milliseconds / 1000.0

    def _seconds_to_ass_time(self, seconds):
        """将秒数转换为ASS时间格式 (H:MM:SS.cc)"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        centiseconds = int((seconds % 1) * 100)
        return f"{hours}:{minutes:02d}:{secs:02d}.{centiseconds:02d}"

    def load_soft_subtitles(self, subtitle_path):
        """加载软字幕文件"""
        try:
            self.soft_subtitles = []

            # 获取视频时长
            video_path = self.soft_video_path_var.get()
            if not video_path:
                raise Exception("请先选择视频文件")

            # 使用ffmpeg获取视频信息 - 修复中文路径编码问题
            info_cmd = [FFMPEG_PATH, '-i', video_path]
            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in info_cmd)
                info_result = subprocess.run(
                    cmd_str,
                    shell=True,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                info_result = subprocess.run(
                    info_cmd,
                    capture_output=True,
                    text=True,
                    encoding='utf-8'
                )

            # 解析视频时长
            video_duration = None
            for line in info_result.stderr.split('\n'):
                if 'Duration:' in line:
                    duration = line.split('Duration: ')[1].split(',')[0].strip()
                    h, m, s = map(float, duration.split(':'))
                    video_duration = h * 3600 + m * 60 + s
                    break

            if video_duration is None:
                raise Exception("无法获取视频时长")

            print(f"视频时长: {video_duration:.2f}秒")

            # 读取字幕文件
            with open(subtitle_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 检测字幕格式
            subtitle_ext = os.path.splitext(subtitle_path)[1].lower()
            is_ass_format = subtitle_ext in ['.ass', '.ssa'] or '[Script Info]' in content or '[Events]' in content

            if is_ass_format:
                # 解析ASS/SSA格式字幕
                print("检测到ASS/SSA格式字幕")
                self.soft_subtitles = self._parse_ass_subtitles(content, video_duration)

                # 重新统计（从解析方法的输出中提取，或者重新统计以确保准确）
                for subtitle in self.soft_subtitles:
                    if subtitle['end'] > video_duration:
                        subtitle['end'] = video_duration
            else:
                # 解析SRT格式字幕
                print("检测到SRT格式字幕")
                blocks = content.strip().split('\n\n')

                for block in blocks:
                    lines = block.strip().split('\n')
                    if len(lines) >= 3:
                        time_line = lines[1]
                        text = ' '.join(lines[2:])

                        # 解析时间
                        if ' --> ' not in time_line:
                            continue  # 跳过无效的时间行

                        try:
                            start_time, end_time = time_line.split(' --> ')
                            start_h, start_m, start_s = map(float, start_time.replace(',', '.').split(':'))
                            end_h, end_m, end_s = map(float, end_time.replace(',', '.').split(':'))

                            start_seconds = start_h * 3600 + start_m * 60 + start_s
                            end_seconds = end_h * 3600 + end_m * 60 + end_s

                            # 检查字幕时间是否超出视频时长
                            if start_seconds >= video_duration:
                                continue

                            if end_seconds > video_duration:
                                end_seconds = video_duration

                            self.soft_subtitles.append({
                                'start': start_seconds,
                                'end': end_seconds,
                                'text': text
                            })
                        except Exception as e:
                            print(f"解析字幕行失败: {e}")
                            continue

            print(f"成功加载 {len(self.soft_subtitles)} 条字幕")
        except Exception as e:
            messagebox.showerror("错误", f"加载字幕失败: {str(e)}")

    def show_soft_subtitle_frame(self, frame):
        """显示软字幕视频帧到预览画布"""
        try:
            # 获取画布尺寸
            canvas_width = self.soft_subtitle_preview_canvas.winfo_width()
            canvas_height = self.soft_subtitle_preview_canvas.winfo_height()

            if canvas_width <= 1 or canvas_height <= 1:
                # 如果画布还没有初始化，使用默认尺寸
                canvas_width = 800
                canvas_height = 600

            # 调整帧大小以适应画布
            frame_height, frame_width = frame.shape[:2]
            scale = min(canvas_width / frame_width, canvas_height / frame_height)
            new_width = int(frame_width * scale)
            new_height = int(frame_height * scale)

            if new_width > 0 and new_height > 0:
                resized_frame = cv2.resize(frame, (new_width, new_height))

                # 转换为RGB格式
                rgb_frame = cv2.cvtColor(resized_frame, cv2.COLOR_BGR2RGB)

                # 转换为PIL图像
                from PIL import Image, ImageTk
                pil_image = Image.fromarray(rgb_frame)
                photo = ImageTk.PhotoImage(image=pil_image)

                # 清除画布并显示图像
                self.soft_subtitle_preview_canvas.delete("all")
                self.soft_subtitle_preview_canvas.create_image(
                    canvas_width // 2, canvas_height // 2,
                    image=photo, anchor=tk.CENTER
                )
                self.soft_subtitle_preview_canvas.image = photo  # 保持引用
        except Exception as e:
            print(f"显示视频帧失败: {str(e)}")

    def get_video_bitrate(self, video_path):
        """使用ffprobe获取视频比特率"""
        try:
            # 使用ffprobe获取视频比特率 - 修复中文路径编码问题
            ffprobe_cmd = [
                find_ffprobe_path(),
                '-v', 'error',
                '-select_streams', 'v:0',
                '-show_entries', 'stream=bit_rate',
                '-of', 'default=noprint_wrappers=1:nokey=1',
                video_path
            ]

            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in ffprobe_cmd)
                result = subprocess.run(
                    cmd_str,
                    shell=True,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                result = subprocess.run(
                    ffprobe_cmd,
                    capture_output=True,
                    text=True,
                    encoding='utf-8'
                )

            if result.returncode == 0 and result.stdout.strip():
                # 获取比特率（单位：bps），转换为kbps
                bitrate_bps = int(result.stdout.strip())
                bitrate_kbps = bitrate_bps / 1000

                # 更新显示
                self.original_bitrate_var.set(f"{bitrate_kbps:.0f}k")

                # 如果用户没有输入合成比特率，则设置为原视频比特率
                if not self.output_bitrate_var.get():
                    self.output_bitrate_var.set(f"{bitrate_kbps:.0f}")

                print(f"原视频比特率: {bitrate_kbps:.0f}k")
            else:
                # 如果无法获取比特率，显示未知
                self.original_bitrate_var.set("未知")
                print("无法获取视频比特率")

        except Exception as e:
            print(f"获取视频比特率失败: {str(e)}")
            self.original_bitrate_var.set("获取失败")

    def _parse_ass_subtitles(self, content, video_duration):
        """解析ASS/SSA格式字幕"""
        import re
        subtitles = []
        skipped_count = 0
        truncated_count = 0

        # 查找Events章节
        events_start = content.find('[Events]')
        if events_start == -1:
            raise Exception("ASS格式错误：未找到[Events]章节")

        # 获取Events章节后的内容
        events_content = content[events_start:]

        # 查找Format行，确定字段顺序
        format_line = None
        for line in events_content.split('\n'):
            if line.strip().startswith('Format:'):
                format_line = line.strip()
                break

        # 默认字段顺序（标准ASS格式）
        field_order = ['Layer', 'Start', 'End', 'Style', 'Name', 'MarginL', 'MarginR', 'MarginV', 'Effect', 'Text']
        start_idx = 1
        end_idx = 2
        text_idx = 9

        # 如果有Format行，解析字段顺序
        if format_line:
            fields = [f.strip() for f in format_line.split(':')[1].split(',')]
            try:
                start_idx = fields.index('Start')
                end_idx = fields.index('End')
                text_idx = fields.index('Text')
            except ValueError:
                # 如果字段名不匹配，使用默认顺序
                pass

        # 解析Dialogue行
        for line in events_content.split('\n'):
            line = line.strip()
            if not line.startswith('Dialogue:'):
                continue

            # 移除"Dialogue:"前缀
            dialogue_content = line[9:].strip()

            # 分割字段，但要注意文本字段（最后一个字段）可能包含逗号
            # 标准格式：Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
            # 需要分割前9个字段（索引0-8），剩余部分作为文本字段（索引9）

            # 先按逗号分割所有部分
            all_parts = dialogue_content.split(',')

            # 如果分割后的部分少于10个，说明格式不正确
            if len(all_parts) < 10:
                print(f"警告：字幕行字段数量不足（期望10个，实际{len(all_parts)}个），跳过: {line[:50]}...")
                continue

            # 前9个字段是固定字段
            parts = all_parts[:9]
            # 第10个字段开始（索引9及之后）合并为文本字段（可能包含逗号）
            text = ','.join(all_parts[9:]).strip()

            # 提取开始和结束时间
            try:
                start_time_str = parts[start_idx].strip()
                end_time_str = parts[end_idx].strip()
            except IndexError as e:
                print(f"警告：字段索引错误，跳过: {line[:50]}... 错误: {e}")
                continue

            try:
                # 解析ASS时间格式：H:MM:SS.cc 或 H:MM:SS:cc (SSA格式)
                # 转换为秒数
                start_seconds = self._parse_ass_time(start_time_str)
                end_seconds = self._parse_ass_time(end_time_str)

                # 检查字幕时间是否超出视频时长
                if start_seconds >= video_duration:
                    print(f"跳过超出视频时长的字幕: {text[:50]}...")
                    skipped_count += 1
                    continue

                if end_seconds > video_duration:
                    print(f"截断超出视频时长的字幕: {text[:50]}...")
                    end_seconds = video_duration
                    truncated_count += 1

                # 清理ASS标签（可选，保留原始文本也可以）
                # text = self._clean_ass_tags(text)

                subtitles.append({
                    'start': start_seconds,
                    'end': end_seconds,
                    'text': text
                })
            except (ValueError, IndexError) as e:
                print(f"解析字幕行失败，跳过: {line[:50]}... 错误: {e}")
                continue

        print(f"ASS字幕解析完成: 加载{len(subtitles)}条，跳过{skipped_count}条，截断{truncated_count}条")
        return subtitles

    def _parse_ass_time(self, time_str):
        """解析ASS时间格式为秒数
        ASS格式: H:MM:SS.cc 或 H:MM:SS:cc (SSA格式)
        """
        time_str = time_str.strip()

        # 处理SSA格式（使用冒号分隔百分秒）
        if time_str.count(':') == 3:
            # SSA格式: H:MM:SS:cc
            parts = time_str.split(':')
            hours = int(parts[0])
            minutes = int(parts[1])
            seconds = int(parts[2])
            centiseconds = int(parts[3])
            total_seconds = hours * 3600 + minutes * 60 + seconds + centiseconds / 100.0
        else:
            # ASS格式: H:MM:SS.cc
            # 先按点分割，然后再按冒号分割
            if '.' in time_str:
                time_part, centiseconds_part = time_str.rsplit('.', 1)
                centiseconds = int(centiseconds_part)
            else:
                time_part = time_str
                centiseconds = 0

            # 解析时:分:秒
            parts = time_part.split(':')
            if len(parts) == 3:
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = int(parts[2])
            elif len(parts) == 2:
                hours = 0
                minutes = int(parts[0])
                seconds = int(parts[1])
            else:
                raise ValueError(f"无法解析时间格式: {time_str}")

            total_seconds = hours * 3600 + minutes * 60 + seconds + centiseconds / 100.0

        return total_seconds

    def load_subtitles(self, subtitle_path):
        """加载字幕文件"""
        try:
            self.subtitles = []

            # 获取视频时长
            video_path = self.video_path_var.get()
            if not video_path:
                raise Exception("请先选择视频文件")

            # 使用ffmpeg获取视频信息 - 修复中文路径编码问题
            info_cmd = [FFMPEG_PATH, '-i', video_path]
            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in info_cmd)
                info_result = subprocess.run(
                    cmd_str,
                    shell=True,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                info_result = subprocess.run(
                    info_cmd,
                    capture_output=True,
                    text=True,
                    encoding='utf-8'
                )

            # 解析视频时长
            video_duration = None
            for line in info_result.stderr.split('\n'):
                if 'Duration:' in line:
                    duration = line.split('Duration: ')[1].split(',')[0].strip()
                    h, m, s = map(float, duration.split(':'))
                    video_duration = h * 3600 + m * 60 + s
                    break

            if video_duration is None:
                raise Exception("无法获取视频时长")

            print(f"视频时长: {video_duration:.2f}秒")

            # 读取字幕文件
            with open(subtitle_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 检测字幕格式
            subtitle_ext = os.path.splitext(subtitle_path)[1].lower()
            is_ass_format = subtitle_ext in ['.ass', '.ssa'] or '[Script Info]' in content or '[Events]' in content

            skipped_count = 0
            truncated_count = 0

            if is_ass_format:
                # 解析ASS/SSA格式字幕
                print("检测到ASS/SSA格式字幕")
                self.subtitles = self._parse_ass_subtitles(content, video_duration)

                # 重新统计（从解析方法的输出中提取，或者重新统计以确保准确）
                for subtitle in self.subtitles:
                    if subtitle['end'] > video_duration:
                        subtitle['end'] = video_duration
            else:
                # 解析SRT格式字幕
                print("检测到SRT格式字幕")
                blocks = content.strip().split('\n\n')

                for block in blocks:
                    lines = block.strip().split('\n')
                    if len(lines) >= 3:
                        time_line = lines[1]
                        text = ' '.join(lines[2:])

                        # 解析时间
                        if ' --> ' not in time_line:
                            continue  # 跳过无效的时间行

                        try:
                            start_time, end_time = time_line.split(' --> ')
                            start_h, start_m, start_s = map(float, start_time.replace(',', '.').split(':'))
                            end_h, end_m, end_s = map(float, end_time.replace(',', '.').split(':'))

                            start_seconds = start_h * 3600 + start_m * 60 + start_s
                            end_seconds = end_h * 3600 + end_m * 60 + end_s

                            # 检查字幕时间是否超出视频时长
                            if start_seconds >= video_duration:
                                print(f"跳过超出视频时长的字幕: {text[:50]}...")
                                skipped_count += 1
                                continue

                            if end_seconds > video_duration:
                                print(f"截断超出视频时长的字幕: {text[:50]}...")
                                end_seconds = video_duration
                                truncated_count += 1

                            self.subtitles.append({
                                'start': start_seconds,
                                'end': end_seconds,
                                'text': text
                            })
                        except (ValueError, IndexError) as e:
                            print(f"解析SRT字幕行失败，跳过: {time_line} 错误: {e}")
                            continue

            # 按时间排序
            self.subtitles.sort(key=lambda x: x['start'])

            # 打印字幕加载信息
            print(f"加载字幕数量: {len(self.subtitles)}")
            print(f"跳过的字幕数量: {skipped_count}")
            print(f"截断的字幕数量: {truncated_count}")
            if self.subtitles:
                print(f"字幕时间范围: {self.subtitles[0]['start']:.2f}秒 - {self.subtitles[-1]['end']:.2f}秒")

        except Exception as e:
            messagebox.showerror("错误", f"加载字幕失败: {str(e)}")

    def show_subtitle_frame(self, frame):
        """显示视频帧和字幕"""
        if frame is None:
            return

        # 获取画布尺寸
        canvas_width = self.subtitle_preview_canvas.winfo_width()
        canvas_height = self.subtitle_preview_canvas.winfo_height()

        # 转换为RGB格式
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # 创建 PIL Image
        img = Image.fromarray(frame)

        # 计算缩放比例，保持宽高比
        img_ratio = img.width / img.height
        canvas_ratio = canvas_width / canvas_height

        if img_ratio > canvas_ratio:
            # 如果图片更宽，按宽度缩放
            new_width = canvas_width
            new_height = int(canvas_width / img_ratio)
        else:
            # 如果图片更高，按高度缩放
            new_height = canvas_height
            new_width = int(canvas_height * img_ratio)

        # 缩放图片
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        self.subtitle_preview_photo = ImageTk.PhotoImage(img)

        # 更新画布，确保图像居中显示
        self.subtitle_preview_canvas.delete("all")
        x = (canvas_width - new_width) // 2
        y = (canvas_height - new_height) // 2

        self.subtitle_preview_canvas.create_image(
            x, y,
            image=self.subtitle_preview_photo,
            anchor=tk.NW
        )

    def preview_subtitle(self):
        """预览字幕效果"""
        if not self.subtitle_cap or not self.subtitles:
            messagebox.showerror("错误", "请先选择视频和字幕文件")
            return

        if self.is_previewing:
            self.stop_preview()
        else:
            self.start_preview()

    def start_preview(self):
        """开始预览"""
        self.is_previewing = True
        self.preview_btn.config(text="停止预览")
        self.subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
        self.update_subtitle_preview()

    def stop_preview(self):
        """停止预览"""
        self.is_previewing = False
        self.preview_btn.config(text="预览字幕")
        if self.subtitle_timer:
            self.root.after_cancel(self.subtitle_timer)

    def update_subtitle_preview(self):
        """更新字幕预览"""
        if not self.is_previewing:
            return

        try:
            # 获取当前时间和进度
            current_frame = self.subtitle_cap.get(cv2.CAP_PROP_POS_FRAMES)
            fps = self.subtitle_fps if hasattr(self, 'subtitle_fps') else self.subtitle_cap.get(cv2.CAP_PROP_FPS)
            total_frames = self.subtitle_total_frames if hasattr(self, 'subtitle_total_frames') else int(self.subtitle_cap.get(cv2.CAP_PROP_FRAME_COUNT))
            current_time = current_frame / fps if fps > 0 else 0

            # 更新进度条
            if hasattr(self, 'subtitle_duration') and self.subtitle_duration > 0:
                progress = (current_time / self.subtitle_duration) * 100
                self.subtitle_progress_slider.set(progress)

            # 读取当前帧
            ret, frame = self.subtitle_cap.read()
            if ret:
                # 查找当前时间对应的字幕
                current_subtitle = ""
                for subtitle in self.subtitles:
                    if subtitle['start'] <= current_time <= subtitle['end']:
                        current_subtitle = subtitle['text']
                        break

                # 应用当前样式设置
                font_size = self.font_size_var.get()
                color_name = self.font_color_var.get()
                color = self.color_mapping.get(color_name, "white")

                # 更新字幕显示（应用样式）
                self.subtitle_label.config(
                    text=current_subtitle,
                    font=("微软雅黑", int(font_size)),
                    fg=color
                )

                # 显示帧
                self.show_subtitle_frame(frame)

            # 继续更新
            self.subtitle_timer = self.root.after(50, self.update_subtitle_preview)

            # 检查是否播放完毕
            if current_frame >= total_frames - 1:
                self.subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)

        except Exception as e:
            print(f"预览更新失败: {str(e)}")
            self.stop_preview()

    def detect_hardware_encoders(self):
        """检测系统可用的硬件加速编码器"""
        print("检测系统可用的硬件加速编码器...")
        available_encoders = []

        try:
            # 检测可用的硬件加速方法
            hwaccel_cmd = [FFMPEG_PATH, '-hwaccels']
            hwaccel_result = subprocess.run(
                hwaccel_cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )

            hwaccel_output = hwaccel_result.stdout
            print("可用的硬件加速方法:")
            print(hwaccel_output)

            # 检测可用的H.264编码器
            encoders_cmd = [FFMPEG_PATH, '-encoders']
            encoders_result = subprocess.run(
                encoders_cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
            )

            encoders_output = encoders_result.stdout

            # 检查NVIDIA GPU编码器
            if 'h264_nvenc' in encoders_output:
                available_encoders.append(('h264_nvenc', 'NVIDIA GPU加速'))

            # 检查Intel QSV编码器
            if 'h264_qsv' in encoders_output:
                available_encoders.append(('h264_qsv', 'Intel GPU加速'))

            # 检查AMD AMF编码器
            if 'h264_amf' in encoders_output:
                available_encoders.append(('h264_amf', 'AMD GPU加速'))

            # 检查VAAPI编码器
            if 'h264_vaapi' in encoders_output:
                available_encoders.append(('h264_vaapi', 'VAAPI加速'))

            # 如果没有硬件编码器可用，使用软件编码
            if not available_encoders:
                available_encoders.append(('libx264', '软件编码(CPU)'))

            print(f"检测到的编码器: {available_encoders}")
            return available_encoders

        except Exception as e:
            print(f"检测硬件编码器失败: {str(e)}")
            # 出错时默认使用软件编码
            return [('libx264', '软件编码(CPU)')]

    def generate_subtitle_video(self):
        """生成带字幕的视频"""
        print("\n=== 开始生成字幕视频 ===")
        if not self.subtitle_cap or not self.subtitles:
            print("错误：未选择视频或字幕文件")
            messagebox.showerror("错误", "请先选择视频和字幕文件")
            return

        if self.is_generating:
            print("错误：正在生成中，请等待...")
            messagebox.showinfo("提示", "正在生成中，请等待...")
            return

        print("1. 选择保存路径...")
        # 生成默认输出文件名：视频名称-C.mp4
        video_path = self.video_path_var.get()
        if video_path:
            video_dir = os.path.dirname(video_path)
            video_name = os.path.splitext(os.path.basename(video_path))[0]
            default_filename = f"{video_name}-C.mp4"
            default_path = os.path.join(video_dir, default_filename)
        else:
            default_path = ""

        # 选择保存路径
        save_path = filedialog.asksaveasfilename(
            title="选择保存位置",
            initialfile=default_path,
            defaultextension=".mp4",
            filetypes=[("MP4文件", "*.mp4"), ("TS文件", "*.ts")]
        )

        if not save_path:  # Changed from if not f'"{save_path}"'
            print("用户取消了保存路径选择")
            return

        print(f"2. 保存路径: {save_path}")  # Changed from print(f"2. 保存路径: {f'"{save_path}"'}")

        # 获取字幕样式设置
        print("3. 获取字幕样式设置...")
        font_size = self.font_size_var.get()
        font_color = self.font_color_var.get()
        position = self.position_var.get()
        gpu_option = self.gpu_var.get()
        print(f"   - 字体大小: {font_size}")
        print(f"   - 字体颜色: {font_color}")
        print(f"   - 字幕位置: {position}")
        print(f"   - GPU加速: {gpu_option}")

        try:
            # 获取视频信息
            print("4. 获取视频信息...")
            video_path = self.video_path_var.get()
            subtitle_path = self.subtitle_path_var.get()

            # 验证路径
            print("4.1. 验证文件路径...")
            path_errors = self.validate_paths(video_path, subtitle_path)
            if path_errors:
                error_msg = "路径验证失败：\n" + "\n".join(path_errors)
                print(f"路径验证错误: {error_msg}")
                messagebox.showerror("错误", error_msg)
                return

            # 使用ffmpeg获取视频信息 - 修复中文路径编码问题
            info_cmd = [FFMPEG_PATH, '-i', video_path]
            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in info_cmd)
                info_result = subprocess.run(
                    cmd_str,
                    shell=True,
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                info_result = subprocess.run(
                    info_cmd,
                    capture_output=True,
                    text=True,
                    encoding='utf-8'
                )

            # 解析视频信息
            video_info = {}
            for line in info_result.stderr.split('\n'):
                if 'Duration:' in line:
                    # 解析时长
                    duration = line.split('Duration: ')[1].split(',')[0].strip()
                    h, m, s = map(float, duration.split(':'))
                    video_info['duration'] = h * 3600 + m * 60 + s
                elif 'Stream' in line and 'Video:' in line:
                    # 解析视频流信息
                    if 'fps' in line:
                        fps = float(line.split('fps')[0].split(',')[-1].strip())
                        video_info['fps'] = fps
                    if 'kb/s' in line:
                        bitrate = float(line.split('kb/s')[0].split(',')[-1].strip())
                        video_info['bitrate'] = bitrate * 1000  # 转换为b/s
                elif 'Stream' in line and 'Audio:' in line:
                    # 解析音频流信息
                    if 'Hz' in line:
                        sample_rate = int(line.split('Hz')[0].split(',')[-1].strip())
                        video_info['sample_rate'] = sample_rate
                    if 'kb/s' in line:
                        audio_bitrate = float(line.split('kb/s')[0].split(',')[-1].strip())
                        video_info['audio_bitrate'] = audio_bitrate * 1000  # 转换为b/s

            print("视频信息:", video_info)

            # 根据用户选择设置编码器参数
            encoder = self.gpu_mapping.get(gpu_option, "")

            print("5. 构建FFmpeg命令...")
            # 构建基础FFmpeg命令
            # 确保使用绝对路径，并统一路径格式
            subtitle_path_clean = os.path.abspath(subtitle_path)
            video_path_clean = os.path.abspath(video_path)

            # 构建字幕滤镜，使用force_style参数设置样式
            try:
                # 确保路径使用正斜杠，并转义冒号和反斜杠
                subtitle_path_normalized = subtitle_path.replace('\\', '/').replace(':', '\\:').replace("'", "\\'")

                # 获取字幕样式设置
                font_size = self.font_size_var.get()
                font_color_chinese = self.font_color_var.get()  # 获取中文颜色名称
                font_position = self.position_var.get()

                # 将中文颜色名称转换为英文颜色名称
                font_color_english = self.color_mapping.get(font_color_chinese, 'white')

                # 转换颜色格式（从英文颜色名转换为BGR十六进制）
                # ASS字幕格式使用BGR颜色值，格式为&HAABBGGRR&
                color_map = {
                    'white': '&H00FFFFFF&',    # 白色 (BGR: FFFFFF)
                    'black': '&H00000000&',    # 黑色 (BGR: 000000)
                    'red': '&H000000FF&',      # 红色 (BGR: 0000FF)
                    'green': '&H0000FF00&',    # 绿色 (BGR: 00FF00)
                    'blue': '&H00FF0000&',     # 蓝色 (BGR: FF0000)
                    'yellow': '&H0000FFFF&',   # 黄色 (BGR: 00FFFF)
                    'cyan': '&H00FFFF00&',     # 青色 (BGR: FFFF00)
                    'magenta': '&H00FF00FF&',  # 洋红色 (BGR: FF00FF)
                    'orange': '&H0000A5FF&',   # 橙色 (BGR: 00A5FF)
                    'pink': '&H00C0C0FF&',     # 粉色 (BGR: C0C0FF)
                    'purple': '&H00800080&',   # 紫色 (BGR: 800080)
                    'gray': '&H00808080&'      # 灰色 (BGR: 808080)
                }
                primary_color = color_map.get(font_color_english, '&H00FFFFFF&')
                print(f"颜色映射: {font_color_chinese} -> {font_color_english} -> {primary_color}")

                # 设置字体位置
                # 将中文位置名称转换为英文位置名称
                font_position_english = self.position_mapping.get(font_position, 'bottom')

                position_map = {
                    'top': 'Alignment=2',      # 顶部居中
                    'middle': 'Alignment=5',   # 中间居中
                    'bottom': 'Alignment=2'    # 底部居中（默认）
                }
                alignment = position_map.get(font_position_english, 'Alignment=2')
                print(f"位置映射: {font_position} -> {font_position_english} -> {alignment}")

                # 构建force_style字符串 - 包含位置信息
                force_style = f"FontName=Arial,FontSize={font_size},PrimaryColour={primary_color},{alignment}"

                # FFmpeg字幕滤镜路径转义规则：
                # 1. 路径中的冒号需要转义为\:（FFmpeg滤镜语法要求）
                # 2. 路径中的反斜杠需要转义为\\或统一使用正斜杠
                # 3. 路径中的单引号需要转义为\'（如果使用单引号包围路径）
                # 4. 统一使用正斜杠作为路径分隔符
                subtitle_path_for_filter = subtitle_path.replace('\\', '/')  # 统一使用正斜杠
                # 转义冒号（FFmpeg滤镜语法要求）
                subtitle_path_escaped = subtitle_path_for_filter.replace(':', '\\:')
                # 转义单引号（如果路径中包含单引号）
                subtitle_path_escaped = subtitle_path_escaped.replace("'", "\\'")

                # 使用单引号包围路径和样式（FFmpeg滤镜标准语法）
                subtitle_filter = f"subtitles='{subtitle_path_escaped}':force_style='{force_style}'"

                print(f"使用带样式的字幕滤镜: {subtitle_filter}")
                print(f"样式设置: 字体大小={font_size}, 颜色={font_color_chinese}->{font_color_english}({primary_color}), 位置={font_position}->{font_position_english}({alignment})")
            except Exception as e:
                print(f"构建字幕滤镜失败: {str(e)}")
                # 备用方法：使用最基本的滤镜
                subtitle_filter = 'format=yuv420p'
                print(f"使用无字幕滤镜: {subtitle_filter}")

            # 构建FFmpeg命令列表（不使用额外引号，subprocess会自动处理）
            ffmpeg_cmd = [
                FFMPEG_PATH,
                '-y',
                '-i', video_path_clean,  # 直接使用路径，不加引号
                '-vf', subtitle_filter   # 直接使用滤镜字符串
            ]

            # 获取用户指定的比特率
            output_bitrate = self.output_bitrate_var.get()
            if not output_bitrate:
                # 如果用户没有输入，使用原视频比特率
                try:
                    # 从原视频比特率显示中提取数值
                    original_bitrate_text = self.original_bitrate_var.get()
                    if original_bitrate_text and original_bitrate_text != "未选择视频" and original_bitrate_text != "未知" and original_bitrate_text != "获取失败":
                        output_bitrate = original_bitrate_text.replace('k', '')
                    else:
                        output_bitrate = "3000"  # 默认值
                except:
                    output_bitrate = "3000"  # 默认值

            print(f"使用比特率: {output_bitrate}k")

            # 添加编码器参数
            if encoder:
                if encoder == "h264_nvenc_fast":
                    # NVIDIA高性能模式：使用用户指定的比特率
                    ffmpeg_cmd.extend([
                        '-b:v', f'{output_bitrate}k',  # 使用用户指定的比特率
                        '-r', '30',       # 固定帧率
                        '-vcodec', 'h264_nvenc'  # 使用vcodec参数，与您的命令一致
                    ])
                elif encoder == "h264_nvenc":
                    # NVIDIA标准模式：使用复杂的参数
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-preset', 'p4',  # 使用p4预设平衡性能和质量
                        '-rc', 'vbr',     # 使用可变比特率
                        '-cq', '19',      # 设置质量参数
                        '-b:v', f'{output_bitrate}k',  # 使用用户指定的比特率
                        '-maxrate', f"{int(output_bitrate) * 1.5:.0f}k",  # 设置最大比特率
                        '-bufsize', f"{int(output_bitrate) * 2:.0f}k",    # 设置缓冲区大小
                        '-r', f"{video_info.get('fps', 30):.0f}",
                        '-spatial-aq', '1',  # 启用空间AQ
                        '-temporal-aq', '1', # 启用时间AQ
                        '-rc-lookahead', '32'  # 设置前瞻帧数
                    ])
                elif encoder == "hevc_qsv":
                    # Intel QSV编码器：使用简化的参数（QSV不支持preset等参数）
                    # 只使用基本的比特率控制参数
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-b:v', f'{output_bitrate}k',  # 使用用户指定的比特率
                        '-r', f"{video_info.get('fps', 30):.0f}",
                    ])
                elif encoder in ["h264_qsv", "h265_qsv"]:
                    # Intel QSV的其他编码器：使用简化参数
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-b:v', f'{output_bitrate}k',
                        '-r', f"{video_info.get('fps', 30):.0f}",
                    ])
                elif encoder in ["h264_amf", "av1_amf"]:
                    # AMD AMF编码器
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-quality', 'balanced',  # AMF质量选项
                        '-rc', 'vbr_peak',  # AMF比特率控制
                        '-b:v', f'{output_bitrate}k',
                        '-maxrate', f"{int(output_bitrate) * 1.5:.0f}k",
                        '-bufsize', f"{int(output_bitrate) * 2:.0f}k",
                        '-r', f"{video_info.get('fps', 30):.0f}",
                    ])
                elif encoder == "h264_vaapi":
                    # VAAPI编码器（Linux）
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-b:v', f'{output_bitrate}k',
                        '-maxrate', f"{int(output_bitrate) * 1.5:.0f}k",
                        '-bufsize', f"{int(output_bitrate) * 2:.0f}k",
                        '-r', f"{video_info.get('fps', 30):.0f}",
                    ])
                else:
                    # 其他编码器：使用通用参数
                    ffmpeg_cmd.extend([
                        '-c:v', encoder,
                        '-b:v', f'{output_bitrate}k',
                        '-maxrate', f"{int(output_bitrate) * 1.5:.0f}k",
                        '-bufsize', f"{int(output_bitrate) * 2:.0f}k",
                        '-r', f"{video_info.get('fps', 30):.0f}",
                    ])

                # 对于硬件编码，重新编码音频以确保兼容性
                ffmpeg_cmd.extend([
                    '-c:a', 'aac',
                    '-b:a', f"{video_info.get('audio_bitrate', 192*1000):.0f}",
                    '-ar', f"{video_info.get('sample_rate', 48000)}",
                    '-ac', '2'
                ])
            else:
                ffmpeg_cmd.extend([
                    '-c:v', 'libx264',
                    '-preset', 'medium',
                    '-b:v', f'{output_bitrate}k',  # 使用用户指定的比特率
                    '-r', f"{video_info.get('fps', 30):.0f}",
                    '-c:a', 'aac',
                    '-b:a', f"{video_info.get('audio_bitrate', 192*1000):.0f}",
                    '-ar', f"{video_info.get('sample_rate', 48000)}",
                    '-ac', '2'
                ])

            # 添加其他参数
            ffmpeg_cmd.extend([
                '-avoid_negative_ts', '1',
                '-threads', '4',
                '-max_muxing_queue_size', '1024',
                # 添加参数确保中途退出时文件可播放
                '-movflags', '+faststart',  # 将元数据移到文件开头
                '-f', 'mp4',               # 强制使用MP4格式
                '-reset_timestamps', '1',  # 重置时间戳
                '-fflags', '+genpts',      # 生成时间戳
                save_path  # 直接使用路径，不加引号
            ])

            print("FFmpeg命令:")
            print(" ".join(ffmpeg_cmd))

            # 打印可复制的命令（用于手动测试）
            print("\n=== 可复制的FFmpeg命令（用于手动测试）===")
            copyable_cmd = " ".join(ffmpeg_cmd)
            print(copyable_cmd)
            print("=== 复制上面的命令到控制台执行测试 ===\n")

            # 开始生成
            print("6. 开始生成字幕视频...")
            self.is_generating = True
            self.subtitle_progress_var.set(0)

            # 禁用相关按钮
            self.disable_subtitle_buttons()

            # 启动处理线程
            print("7. 启动处理线程...")
            generate_thread = threading.Thread(
                target=self.run_subtitle_ffmpeg,
                args=(ffmpeg_cmd, save_path, video_info.get('duration', 7730.76))
            )
            generate_thread.start()

        except Exception as e:
            print(f"错误：生成失败 - {str(e)}")
            self.is_generating = False
            messagebox.showerror("错误", f"生成失败: {str(e)}")

    def preview_soft_subtitle(self):
        """预览软字幕效果"""
        if not self.soft_subtitle_cap or not self.soft_subtitles:
            messagebox.showerror("错误", "请先选择视频和字幕文件")
            return

        if self.soft_is_previewing:
            self.stop_soft_preview()
        else:
            self.start_soft_preview()

    def start_soft_preview(self):
        """开始软字幕预览"""
        self.soft_is_previewing = True
        self.soft_preview_btn.config(text="停止预览")
        self.soft_subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
        self.update_soft_subtitle_preview()

    def stop_soft_preview(self):
        """停止软字幕预览"""
        self.soft_is_previewing = False
        self.soft_preview_btn.config(text="预览字幕")
        if self.soft_subtitle_timer:
            self.root.after_cancel(self.soft_subtitle_timer)

    def update_soft_subtitle_style(self):
        """更新软字幕样式"""
        # 获取当前样式设置
        font_size = self.soft_font_size_var.get()
        color_name = self.soft_font_color_var.get()
        position_name = self.soft_position_var.get()

        # 转换为实际值
        color = self.color_mapping.get(color_name, "white")

        # 更新字幕标签样式
        self.soft_subtitle_label.config(
            font=("微软雅黑", int(font_size)),
            fg=color
        )

        # 更新字幕位置
        if position_name == "顶部":
            self.soft_subtitle_label.pack_forget()
            self.soft_subtitle_label.pack(side=tk.TOP, pady=20)
        elif position_name == "中间":
            self.soft_subtitle_label.pack_forget()
            # 使用place方法精确定位到中间
            self.soft_subtitle_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        else:  # 底部
            self.soft_subtitle_label.pack_forget()
            self.soft_subtitle_label.pack(side=tk.BOTTOM, pady=20)

        # 如果正在预览，立即更新当前帧的字幕显示
        if self.soft_is_previewing and hasattr(self, 'soft_subtitle_cap') and self.soft_subtitle_cap:
            current_frame = self.soft_subtitle_cap.get(cv2.CAP_PROP_POS_FRAMES)
            fps = self.soft_subtitle_cap.get(cv2.CAP_PROP_FPS)
            current_time = current_frame / fps if fps > 0 else 0

            # 查找当前时间对应的字幕
            current_subtitle = ""
            for subtitle in self.soft_subtitles:
                if subtitle['start'] <= current_time <= subtitle['end']:
                    current_subtitle = subtitle['text']
                    break

            # 更新字幕显示
            self.soft_subtitle_label.config(text=current_subtitle)

    def on_soft_progress_change(self, value):
        """处理软字幕进度条拖动"""
        if not self.soft_subtitle_cap:
            return

        # 计算目标时间
        progress = float(value)
        if hasattr(self, 'soft_subtitle_duration') and self.soft_subtitle_duration > 0:
            target_time = (progress / 100.0) * self.soft_subtitle_duration

            try:
                # 更新视频帧位置
                if hasattr(self, 'soft_subtitle_fps') and self.soft_subtitle_fps > 0:
                    self.soft_subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, int(target_time * self.soft_subtitle_fps))
                    ret, frame = self.soft_subtitle_cap.read()
                    if ret:
                        # 查找当前时间对应的字幕
                        current_subtitle = ""
                        for subtitle in self.soft_subtitles:
                            if subtitle['start'] <= target_time <= subtitle['end']:
                                current_subtitle = subtitle['text']
                                break

                        # 应用当前样式设置
                        font_size = self.soft_font_size_var.get()
                        color_name = self.soft_font_color_var.get()
                        color = self.color_mapping.get(color_name, "white")

                        # 更新字幕显示（应用样式）
                        self.soft_subtitle_label.config(
                            text=current_subtitle,
                            font=("微软雅黑", int(font_size)),
                            fg=color
                        )

                        # 显示帧
                        self.show_soft_subtitle_frame(frame)

            except Exception as e:
                print(f"预览更新失败: {str(e)}")

    def update_soft_subtitle_preview(self):
        """更新软字幕预览"""
        if not self.soft_is_previewing:
            return

        try:
            # 获取当前时间和进度
            current_frame = self.soft_subtitle_cap.get(cv2.CAP_PROP_POS_FRAMES)
            fps = self.soft_subtitle_fps if hasattr(self, 'soft_subtitle_fps') else self.soft_subtitle_cap.get(cv2.CAP_PROP_FPS)
            total_frames = self.soft_subtitle_total_frames if hasattr(self, 'soft_subtitle_total_frames') else int(self.soft_subtitle_cap.get(cv2.CAP_PROP_FRAME_COUNT))
            current_time = current_frame / fps if fps > 0 else 0

            # 更新进度条
            if hasattr(self, 'soft_subtitle_duration') and self.soft_subtitle_duration > 0:
                progress = (current_time / self.soft_subtitle_duration) * 100
                self.soft_subtitle_progress_slider.set(progress)

            # 读取当前帧
            ret, frame = self.soft_subtitle_cap.read()
            if ret:
                # 查找当前时间对应的字幕
                current_subtitle = ""
                for subtitle in self.soft_subtitles:
                    if subtitle['start'] <= current_time <= subtitle['end']:
                        current_subtitle = subtitle['text']
                        break

                # 应用当前样式设置
                font_size = self.soft_font_size_var.get()
                color_name = self.soft_font_color_var.get()
                color = self.color_mapping.get(color_name, "white")

                # 更新字幕显示（应用样式）
                self.soft_subtitle_label.config(
                    text=current_subtitle,
                    font=("微软雅黑", int(font_size)),
                    fg=color
                )

                # 显示帧
                self.show_soft_subtitle_frame(frame)

            # 继续更新
            self.soft_subtitle_timer = self.root.after(50, self.update_soft_subtitle_preview)

            # 检查是否播放完毕
            if current_frame >= total_frames - 1:
                self.soft_subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, 0)

        except Exception as e:
            print(f"预览更新失败: {str(e)}")
            self.stop_soft_preview()

    def generate_soft_subtitle_video(self):
        """生成软字幕视频（将字幕作为独立轨道）"""
        print("\n=== 开始生成软字幕视频 ===")
        if not self.soft_subtitle_cap or not self.soft_subtitles:
            print("错误：未选择视频或字幕文件")
            messagebox.showerror("错误", "请先选择视频和字幕文件")
            return

        if self.soft_is_generating:
            print("错误：正在生成中，请等待...")
            messagebox.showinfo("提示", "正在生成中，请等待...")
            return

        print("1. 选择保存路径...")
        # 生成默认输出文件名：视频名称-soft.mp4
        video_path = self.soft_video_path_var.get()
        if video_path:
            video_dir = os.path.dirname(video_path)
            video_name = os.path.splitext(os.path.basename(video_path))[0]
            default_filename = f"{video_name}-soft.mp4"
            default_path = os.path.join(video_dir, default_filename)
        else:
            default_path = ""

        # 选择保存路径 - 默认使用MKV格式（支持所有字幕格式）
        save_path = filedialog.asksaveasfilename(
            title="选择保存位置",
            initialfile=default_path.replace('.mp4', '.mkv'),
            defaultextension=".mkv",
            filetypes=[("MKV文件", "*.mkv"), ("MP4文件", "*.mp4")]
        )

        if not save_path:
            print("用户取消了保存路径选择")
            return

        print(f"2. 保存路径: {save_path}")

        try:
            # 获取视频和字幕路径
            video_path = self.soft_video_path_var.get()
            subtitle_path = self.soft_subtitle_path_var.get()

            # 验证路径
            if not os.path.exists(video_path):
                messagebox.showerror("错误", f"视频文件不存在: {video_path}")
                return
            if not os.path.exists(subtitle_path):
                messagebox.showerror("错误", f"字幕文件不存在: {subtitle_path}")
                return

            # 检测输出格式
            output_ext = os.path.splitext(save_path)[1].lower()
            is_mp4 = output_ext == '.mp4'

            # 检测字幕格式
            subtitle_ext = os.path.splitext(subtitle_path)[1].lower()
            is_srt = subtitle_ext in ['.srt']
            is_ass = subtitle_ext in ['.ass', '.ssa']

            # 获取用户选择的样式设置
            font_size = self.soft_font_size_var.get()
            font_color_chinese = self.soft_font_color_var.get()
            font_position = self.soft_position_var.get()

            print(f"3. 应用字幕样式 - 字号: {font_size}, 颜色: {font_color_chinese}, 位置: {font_position}")

            # 对于软字幕，需要将样式嵌入到ASS文件中
            # 无论输入是SRT还是ASS，都转换为带样式的ASS文件
            import tempfile
            temp_ass = tempfile.NamedTemporaryFile(suffix='.ass', delete=False)
            temp_ass_path = temp_ass.name
            temp_ass.close()

            # 创建带样式的ASS文件
            try:
                self.create_styled_ass_file(subtitle_path, temp_ass_path, font_size, font_color_chinese, font_position)
                print(f"[DEBUG] 已创建带样式的ASS文件: {temp_ass_path} (大小: {os.path.getsize(temp_ass_path)} 字节)")
                subtitle_path_clean = os.path.abspath(temp_ass_path)
                subtitle_codec = 'ass'
                use_temp_file = True
            except Exception as e:
                error_msg = f"创建带样式ASS文件失败: {str(e)}"
                print(error_msg)
                messagebox.showerror("错误", error_msg)
                try:
                    os.unlink(temp_ass_path)
                except:
                    pass
                return

            video_path_clean = os.path.abspath(video_path)

            print("4. 构建FFmpeg命令...")
            # 构建FFmpeg命令 - 软字幕不需要重新编码视频，只需要添加字幕轨道
            # 参考硬字幕的实现，使用subprocess.Popen直接传递参数列表
            ffmpeg_cmd = [
                FFMPEG_PATH,
                '-y',
                '-i', video_path_clean,
                '-i', subtitle_path_clean,
                '-c:v', 'copy',  # 复制视频流，不重新编码
                '-c:a', 'copy',  # 复制音频流，不重新编码
                '-map', '0:v',  # 映射第一个输入的视频流
                '-map', '0:a',  # 映射第一个输入的音频流
                '-map', '1:s',  # 映射第二个输入的字幕流
            ]

            # 根据输出格式设置字幕编解码器和容器格式
            # 由于我们已经统一转换为带样式的ASS格式，所以总是使用ASS
            if is_mp4:
                # MP4容器：ASS格式需要转换为mov_text（MP4标准字幕格式）
                # 但mov_text不支持样式，所以先尝试使用ass，如果失败再转换
                # 实际上MP4对ASS支持有限，建议使用MKV格式以保留完整样式
                ffmpeg_cmd.extend(['-c:s', 'mov_text'])
                ffmpeg_cmd.extend(['-f', 'mp4'])
            else:
                # MKV容器：使用ass编解码器以确保样式被正确编码
                ffmpeg_cmd.extend(['-c:s', 'ass'])  # 使用ass编解码器以保留样式
                ffmpeg_cmd.extend(['-f', 'matroska'])

            # 设置字幕语言元数据
            ffmpeg_cmd.extend(['-metadata:s:s:0', 'language=chi'])

            ffmpeg_cmd.append(save_path)

            print("FFmpeg命令:")
            print(" ".join(ffmpeg_cmd))

            # 开始生成
            print("5. 开始生成软字幕视频...")
            self.soft_is_generating = True
            self.soft_subtitle_progress_var.set(0)

            # 禁用相关按钮
            self.soft_generate_btn.config(state='disabled')
            self.soft_preview_btn.config(state='disabled')

            # 启动处理线程
            print("6. 启动处理线程...")
            generate_thread = threading.Thread(
                target=self.run_soft_subtitle_ffmpeg,
                args=(ffmpeg_cmd, save_path, use_temp_file, subtitle_path_clean if use_temp_file else None)
            )
            generate_thread.start()

        except Exception as e:
            print(f"错误：生成失败 - {str(e)}")
            self.soft_is_generating = False
            messagebox.showerror("错误", f"生成失败: {str(e)}")
            self.soft_generate_btn.config(state='normal')
            self.soft_preview_btn.config(state='normal')
            if 'use_temp_file' in locals() and use_temp_file and os.path.exists(subtitle_path_clean):
                try:
                    os.unlink(subtitle_path_clean)
                except:
                    pass

    def run_soft_subtitle_ffmpeg(self, cmd, output_path, use_temp_file=False, temp_file_path=None):
        """执行FFmpeg命令生成软字幕视频"""
        process = None
        final_returncode = -1
        try:
            print("\n=== 开始执行FFmpeg命令 ===")

            # 获取视频时长用于进度计算
            video_path = self.soft_video_path_var.get()
            video_duration = 0.0
            try:
                info_cmd = [FFMPEG_PATH, '-i', video_path]
                if sys.platform == 'win32':
                    cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in info_cmd)
                    info_result = subprocess.run(
                        cmd_str,
                        shell=True,
                        capture_output=True,
                        text=True,
                        encoding='utf-8',
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    info_result = subprocess.run(
                        info_cmd,
                        capture_output=True,
                        text=True,
                        encoding='utf-8'
                    )

                for line in info_result.stderr.split('\n'):
                    if 'Duration:' in line:
                        duration = line.split('Duration: ')[1].split(',')[0].strip()
                        h, m, s = map(float, duration.split(':'))
                        video_duration = h * 3600 + m * 60 + s
                        break
            except Exception as e:
                print(f"获取视频时长失败: {e}，将无法显示准确进度")

            print(f"视频时长: {video_duration:.2f}秒")

            # 执行FFmpeg命令
            if sys.platform == 'win32':
                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    preexec_fn=os.setsid
                )

            # 将进程添加到跟踪列表
            self.active_processes.append(process)
            print(f"[DEBUG] 启动软字幕进程 PID: {process.pid}")

            # 实时读取输出并更新进度
            import time
            last_progress = 0
            last_file_size = 0
            file_size_stable_count = 0
            last_output_time = time.time()

            while True:
                # 检查进程是否已退出
                if process.poll() is not None:
                    print("[DEBUG] 软字幕FFmpeg进程已退出")
                    break

                # 检查是否应该停止处理
                if not self.soft_is_generating:
                    print("[DEBUG] 收到停止信号，终止软字幕进程...")
                    process.terminate()
                    break

                try:
                    # 读取一行输出（非阻塞）
                    import select
                    if sys.platform != 'win32':
                        ready, _, _ = select.select([process.stderr], [], [], 0.1)
                        if ready:
                            line = process.stderr.readline()
                        else:
                            line = None
                    else:
                        line = process.stderr.readline()

                    if line:
                        line = line.strip()
                        print(line)
                        last_output_time = time.time()

                        # 解析进度信息
                        if 'time=' in line:
                            try:
                                time_str = line.split('time=')[1].split(' ')[0].strip()
                                h, m, s = map(float, time_str.split(':'))
                                current_time = h * 3600 + m * 60 + s

                                # 计算进度百分比
                                if video_duration > 0:
                                    progress = (current_time / video_duration) * 100
                                    progress = min(progress, 98)

                                    if progress > last_progress:
                                        last_progress = progress
                                        self.root.after(0, lambda p=progress: self.soft_subtitle_progress_var.set(p))
                                        print(f"进度: {progress:.1f}% ({current_time:.1f}s / {video_duration:.1f}s)")
                            except Exception as e:
                                print(f"进度解析错误: {e}")
                except Exception as e:
                    if "readline" not in str(e).lower():
                        print(f"读取输出错误: {e}")

                # 监控输出文件大小变化
                if os.path.exists(output_path):
                    try:
                        current_file_size = os.path.getsize(output_path)
                        if current_file_size > last_file_size:
                            last_file_size = current_file_size
                            file_size_stable_count = 0
                        elif current_file_size == last_file_size and current_file_size > 0:
                            file_size_stable_count += 1
                        else:
                            file_size_stable_count = 0

                        if file_size_stable_count > 50:
                            print(f"[DEBUG] 文件大小已稳定 {file_size_stable_count * 0.1:.1f}秒")
                    except:
                        pass

                # 检查长时间无输出
                elapsed_since_output = time.time() - last_output_time
                if elapsed_since_output > 30:
                    if process.poll() is None:
                        print(f"[DEBUG] 超过30秒无输出，但进程仍在运行")
                        if os.path.exists(output_path):
                            file_size = os.path.getsize(output_path)
                            if file_size > 0 and last_progress < 95:
                                self.root.after(0, lambda: self.soft_subtitle_progress_var.set(95))
                                print(f"[DEBUG] 设置进度为95%")

                time.sleep(0.1)

            # 获取最终返回码
            returncode = process.poll()
            if returncode is None:
                try:
                    returncode = process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    returncode = process.returncode if process.returncode is not None else -1

            # 读取剩余输出
            try:
                if process.stdin and not process.stdin.closed:
                    process.stdin.close()
                stdout, stderr = process.communicate(timeout=5)
            except subprocess.TimeoutExpired:
                print("[DEBUG] 读取输出超时")
                stdout, stderr = "", ""
            except Exception as e:
                print(f"[DEBUG] 读取输出时出错: {e}")
                stdout, stderr = "", ""

            final_returncode = returncode

            # 打印完整的错误输出用于调试
            if stderr:
                print("=== FFmpeg错误输出 ===")
                print(stderr)
                print("=== 错误输出结束 ===")

            # 检查输出文件
            file_exists = os.path.exists(output_path)
            file_size = 0
            if file_exists:
                try:
                    file_size = os.path.getsize(output_path)
                    print(f"[DEBUG] 输出文件大小: {file_size} 字节 ({file_size / (1024*1024):.2f} MB)")
                except Exception as e:
                    print(f"[DEBUG] 获取文件大小失败: {e}")

            print(f"[DEBUG] 最终返回码: {final_returncode}")
            print(f"[DEBUG] 文件存在: {file_exists}, 文件大小: {file_size}")

            # 处理结果
            if final_returncode == 0 and file_exists and file_size > 0:
                print("软字幕视频生成成功！")
                self.root.after(0, lambda: self.soft_subtitle_progress_var.set(100))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"软字幕视频已生成:\n{output_path}\n文件大小: {file_size / (1024*1024):.2f} MB"))
            else:
                error_msg = f"生成失败"
                if final_returncode != 0:
                    error_msg += f"（返回码: {final_returncode}）"
                if file_exists and file_size == 0:
                    error_msg += "\n输出文件大小为0，可能是编码参数错误"
                    # 删除空文件
                    try:
                        os.remove(output_path)
                        print(f"[DEBUG] 已删除空文件: {output_path}")
                    except:
                        pass
                elif not file_exists:
                    error_msg += "\n输出文件未生成"

                if stderr:
                    # 提取关键错误信息
                    error_lines = stderr.split('\n')
                    key_errors = []
                    for line in error_lines:
                        if any(keyword in line.lower() for keyword in ['error', 'failed', 'invalid', 'cannot', 'unable']):
                            key_errors.append(line.strip())
                    if key_errors:
                        error_msg += "\n\n关键错误信息:\n" + "\n".join(key_errors[-5:])  # 只显示最后5条关键错误
                    else:
                        error_msg += f"\n\n完整错误输出:\n{stderr[-1000:]}"  # 显示最后1000字符

                print(error_msg)
                self.root.after(0, lambda msg=error_msg: messagebox.showerror("错误", msg))
                self.root.after(0, lambda: self.soft_subtitle_progress_var.set(0))

        except Exception as e:
            error_msg = f"执行失败: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
            self.root.after(0, lambda: self.soft_subtitle_progress_var.set(0))
        finally:
            # 清理进程资源
            if process is not None:
                try:
                    if process.poll() is None:
                        process.terminate()
                        try:
                            process.wait(timeout=2)
                        except subprocess.TimeoutExpired:
                            process.kill()
                            process.wait(timeout=1)

                    # 关闭管道
                    for pipe_name in ['stdin', 'stdout', 'stderr']:
                        pipe = getattr(process, pipe_name, None)
                        if pipe and not pipe.closed:
                            try:
                                pipe.close()
                            except:
                                pass

                    # 从活跃进程列表中移除
                    if process in self.active_processes:
                        self.active_processes.remove(process)
                except Exception as e:
                    print(f"清理软字幕进程时出错: {e}")

            # 更新状态标志
            self.soft_is_generating = False

            # 重新启用按钮
            self.root.after(0, self.enable_soft_subtitle_buttons)

            # 根据最终返回码处理进度条
            final_rc = final_returncode if 'final_returncode' in locals() else (process.returncode if process and process.returncode is not None else -1)
            if final_rc != 0:
                self.root.after(0, lambda: self.soft_subtitle_progress_var.set(0))
                print(f"[DEBUG] 处理失败（返回码: {final_rc}），重置进度条")

            # 清理临时文件
            if use_temp_file and temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.unlink(temp_file_path)
                    print(f"[DEBUG] 已删除临时字幕文件: {temp_file_path}")
                except Exception as e:
                    print(f"[DEBUG] 删除临时文件失败: {e}")

            print("=== 软字幕FFmpeg命令执行完成 ===\n")

    def run_subtitle_ffmpeg(self, cmd, output_path, video_duration=7730.76):
        """执行FFmpeg命令生成字幕视频"""
        process = None
        final_returncode = -1  # 初始化返回码，默认失败
        try:
            print("\n=== 开始执行FFmpeg命令 ===")
            print("1. 检查FFmpeg路径...")
            if not os.path.exists(FFMPEG_PATH):
                raise Exception(f"FFmpeg不存在: {FFMPEG_PATH}")
            print(f"FFmpeg路径: {FFMPEG_PATH}")

            print("2. 检查输入文件...")
            if not os.path.exists(self.video_path_var.get()):
                raise Exception(f"输入视频不存在: {self.video_path_var.get()}")
            print(f"输入视频: {self.video_path_var.get()}")

            print("3. 检查字幕文件...")
            if not os.path.exists(self.subtitle_path_var.get()):
                raise Exception(f"字幕文件不存在: {self.subtitle_path_var.get()}")
            print(f"字幕文件: {self.subtitle_path_var.get()}")

            print("4. 执行FFmpeg命令...")
            print("命令:", " ".join(cmd))

            # 添加详细的路径调试信息
            print("4.1. 路径调试信息:")
            print(f"   - FFmpeg路径: {FFMPEG_PATH}")
            print(f"   - 输入视频: {self.video_path_var.get()}")
            print(f"   - 字幕文件: {self.subtitle_path_var.get()}")
            print(f"   - 输出文件: {output_path}")
            print(f"   - 输入视频存在: {os.path.exists(self.video_path_var.get())}")
            print(f"   - 字幕文件存在: {os.path.exists(self.subtitle_path_var.get())}")

            # 检查FFmpeg是否可执行
            try:
                ffmpeg_test = subprocess.run([FFMPEG_PATH, '-version'],
                                             capture_output=True,
                                             timeout=10,
                                             creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
                if ffmpeg_test.returncode != 0:
                    raise Exception(f"FFmpeg测试失败，返回码: {ffmpeg_test.returncode}")
                print("   - FFmpeg可执行: 是")
            except Exception as e:
                raise Exception(f"FFmpeg不可执行: {str(e)}")

            # 使用subprocess.Popen执行命令，直接传递参数列表（不使用shell=True）
            # 这样可以避免Windows shell解析问题，并且Python的subprocess会自动处理路径编码
            if sys.platform == 'win32':
                process = subprocess.Popen(
                    cmd,  # 直接传递参数列表
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',  # 处理编码错误，用替换字符代替
                    creationflags=subprocess.CREATE_NO_WINDOW,  # 不显示控制台窗口
                    preexec_fn=None  # Windows上不需要
                )
            else:
                # Linux/Mac系统
                process = subprocess.Popen(
                    cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding='utf-8',
                    errors='replace',  # 处理编码错误，用替换字符代替
                    preexec_fn=os.setsid  # 创建新的进程组
                )

            # 保存进程引用到实例变量（用于其他地方访问）
            self.process = process

            # 将进程添加到跟踪列表
            self.active_processes.append(process)
            print(f"[DEBUG] 启动FFmpeg字幕进程 PID: {process.pid}")

            # 设置进程为守护进程，确保主程序退出时子进程也会退出
            if hasattr(process, 'daemon'):
                process.daemon = True

            # 实时读取输出并更新进度
            print("5. 开始实时读取FFmpeg输出...")
            import time

            # 使用传入的视频时长
            print(f"视频时长: {video_duration:.2f}秒")

            # 实时读取stderr输出
            last_progress = 0
            last_file_size = 0
            file_size_stable_count = 0  # 文件大小稳定计数器
            last_output_time = time.time()  # 最后输出时间

            while True:
                # 检查进程是否已退出
                if process.poll() is not None:
                    print("[DEBUG] FFmpeg进程已退出")
                    break

                # 检查是否应该停止处理
                if not self.is_generating:
                    print("[DEBUG] 收到停止信号，终止FFmpeg进程...")
                    process.terminate()
                    break

                try:
                    # 读取一行输出（非阻塞）
                    import select
                    if sys.platform != 'win32':
                        # Linux/Mac使用select
                        ready, _, _ = select.select([process.stderr], [], [], 0.1)
                        if ready:
                            line = process.stderr.readline()
                        else:
                            line = None
                    else:
                        # Windows使用简单的readline（可能阻塞，但设置了超时）
                        line = process.stderr.readline()

                    if line:
                        line = line.strip()
                        print(line)
                        last_output_time = time.time()  # 更新最后输出时间

                        # 解析进度信息
                        if 'time=' in line:
                            try:
                                time_str = line.split('time=')[1].split(' ')[0].strip()
                                h, m, s = map(float, time_str.split(':'))
                                current_time = h * 3600 + m * 60 + s

                                # 计算进度百分比
                                progress = (current_time / video_duration) * 100
                                progress = min(progress, 98)  # 限制最大进度为98%，保留2%给最终处理

                                if progress > last_progress:
                                    last_progress = progress
                                    # 更新进度条
                                    self.root.after(0, lambda p=progress: self.subtitle_progress_var.set(p))
                                    print(f"进度: {progress:.1f}% ({current_time:.1f}s / {video_duration:.1f}s)")

                            except Exception as e:
                                print(f"进度解析错误: {e}")
                except Exception as e:
                    # 读取错误，可能是流已关闭
                    if "readline" not in str(e).lower():
                        print(f"读取输出错误: {e}")

                # 监控输出文件大小变化（用于大文件处理）
                if os.path.exists(output_path):
                    try:
                        current_file_size = os.path.getsize(output_path)

                        # 如果文件大小在增长，重置稳定计数器
                        if current_file_size > last_file_size:
                            last_file_size = current_file_size
                            file_size_stable_count = 0
                        elif current_file_size == last_file_size and current_file_size > 0:
                            # 文件大小稳定，增加计数器
                            file_size_stable_count += 1
                        else:
                            file_size_stable_count = 0

                        # 如果文件大小稳定超过5秒，且进程仍在运行，可能是缓冲写入阶段
                        # 这种情况下，进度可能接近完成，但FFmpeg还在写入元数据
                        if file_size_stable_count > 50:  # 5秒（50 * 0.1秒）
                            print(f"[DEBUG] 文件大小已稳定 {file_size_stable_count * 0.1:.1f}秒，可能处于最终写入阶段")
                    except Exception as e:
                        pass  # 忽略文件访问错误

                # 检查是否有长时间无输出（大文件可能在最后阶段）
                elapsed_since_output = time.time() - last_output_time
                if elapsed_since_output > 30:  # 30秒无输出
                    # 检查进程是否还在运行
                    if process.poll() is None:
                        print(f"[DEBUG] 超过30秒无输出，但进程仍在运行，可能是大文件最终处理阶段")
                        # 可以尝试检查文件大小来判断是否完成
                        if os.path.exists(output_path):
                            file_size = os.path.getsize(output_path)
                            if file_size > 0:
                                # 文件存在且有大小，可能是写入阶段，设置一个高进度
                                if last_progress < 95:
                                    self.root.after(0, lambda: self.subtitle_progress_var.set(95))
                                    print(f"[DEBUG] 设置进度为95%（等待最终完成）")

                time.sleep(0.1)  # 短暂休眠

            # 进程已退出，快速获取返回码和输出
            print("[DEBUG] 开始读取最终输出...")
            returncode = None
            stdout = ""
            stderr = ""

            # 进程已经退出，直接获取返回码
            returncode = process.poll()
            if returncode is not None:
                print(f"[DEBUG] 进程已退出，返回码: {returncode}")
            else:
                # 如果还没退出，等待一下（最多5秒）
                try:
                    returncode = process.wait(timeout=5)
                    print(f"[DEBUG] 等待进程结束，返回码: {returncode}")
                except subprocess.TimeoutExpired:
                    print("[DEBUG] 等待进程超时，强制获取返回码")
                    returncode = process.returncode if process.returncode is not None else -1

            # 快速读取剩余输出（不阻塞，设置短超时）
            try:
                if process.stdin and not process.stdin.closed:
                    try:
                        process.stdin.close()
                    except:
                        pass

                # 使用短超时快速读取输出，避免长时间阻塞
                stdout, stderr = process.communicate(timeout=2)
            except subprocess.TimeoutExpired:
                # 超时就忽略，我们已经有了返回码
                print("[DEBUG] 读取输出超时，使用已有返回码继续")
                stdout, stderr = "", ""
            except Exception as e:
                print(f"[DEBUG] 读取输出时出错: {e}")
                stdout, stderr = "", ""

            # 确保返回码有效
            if returncode is None:
                returncode = process.returncode if process.returncode is not None else -1

            print("6. 检查执行结果...")
            print(f"返回码: {returncode}")
            if stdout:
                stdout_preview = stdout[-2000:] if len(stdout) > 2000 else stdout
                print("标准输出:")
                print(stdout_preview)
            if stderr:
                stderr_preview = stderr[-2000:] if len(stderr) > 2000 else stderr
                print("错误输出:")
                print(stderr_preview)

            # 保存returncode供后续使用
            final_returncode = returncode

            # 检查输出文件（快速检查，不进行完整性验证）
            print("6. 检查输出文件...")
            file_exists = False
            file_size = 0
            if os.path.exists(output_path):
                try:
                    file_size = os.path.getsize(output_path)
                    file_exists = True
                    print(f"输出文件存在: {output_path}")
                    print(f"文件大小: {file_size} 字节 ({file_size / (1024*1024):.2f} MB)")
                except Exception as e:
                    print(f"获取文件大小失败: {e}")
            else:
                print("输出文件不存在")

            # 如果返回码为0且文件存在，立即设置进度为100%并调用完成回调
            # 完整性检查对于大文件太慢，FFmpeg返回码0已经说明文件生成成功
            if final_returncode == 0 and file_exists and file_size > 0:
                print("[DEBUG] FFmpeg返回码0且文件存在，立即设置进度为100%并调用完成回调")
                # 立即设置进度条和调用完成回调
                self.root.after(0, lambda: self.subtitle_progress_var.set(100))
                self.root.after(0, self.handle_subtitle_completion, final_returncode, output_path)
            else:
                # 失败情况，也调用完成回调
                print("7. 调用完成回调...")
                self.root.after(0, self.handle_subtitle_completion, final_returncode, output_path)

        except Exception as e:
            error_msg = f"生成失败: {str(e)}"
            print(f"发生异常: {error_msg}")
            logger.error(error_msg)
            self.root.after(0, messagebox.showerror, "错误", error_msg)
        finally:
            print("8. 清理状态和进程资源...")

            # 清理进程资源
            if process is not None:
                try:
                    process_pid = process.pid if hasattr(process, 'pid') else 'unknown'
                    print(f"[DEBUG] 开始清理字幕进程 PID: {process_pid}")

                    # 关闭所有管道
                    for pipe_name in ['stdin', 'stdout', 'stderr']:
                        pipe = getattr(process, pipe_name, None)
                        if pipe and not pipe.closed:
                            try:
                                pipe.close()
                                print(f"[DEBUG] 已关闭 {pipe_name} 管道")
                            except Exception as e:
                                print(f"[DEBUG] 关闭 {pipe_name} 管道时出错: {e}")

                    # 确保进程已完全结束（快速检查，不阻塞）
                    if process.poll() is None:
                        print(f"[DEBUG] 字幕进程 {process_pid} 仍在运行，尝试快速终止...")
                        # 快速终止，不长时间等待
                        try:
                            process.terminate()
                            try:
                                process.wait(timeout=2)  # 只等待2秒
                                print(f"[DEBUG] 进程 {process_pid} 已通过terminate结束")
                            except subprocess.TimeoutExpired:
                                print(f"[DEBUG] 进程 {process_pid} terminate后仍运行，强制kill")
                                try:
                                    process.kill()
                                    process.wait(timeout=1)  # 只等待1秒
                                    print(f"[DEBUG] 进程 {process_pid} 已通过kill结束")
                                except:
                                    print(f"[DEBUG] 警告：进程 {process_pid} kill后仍可能运行，继续清理资源")
                        except Exception as e:
                            print(f"[DEBUG] 终止进程时出错: {e}，继续清理资源")
                    else:
                        print(f"[DEBUG] 进程 {process_pid} 已退出，快速清理资源")

                    # 从活跃进程列表中移除
                    if process in self.active_processes:
                        self.active_processes.remove(process)
                        print(f"[DEBUG] 从活跃进程列表中移除字幕进程 PID: {process_pid}")

                    print(f"[DEBUG] 字幕进程 {process_pid} 清理完成")
                except Exception as e:
                    print(f"[DEBUG] 清理字幕进程时出错: {e}")
                    import traceback
                    traceback.print_exc()

            # 更新状态标志
            self.is_generating = False

            # 重新启用按钮
            self.root.after(0, self.enable_subtitle_buttons)

            # 根据最终返回码处理进度条
            final_rc = final_returncode if 'final_returncode' in locals() else (process.returncode if process and process.returncode is not None else -1)
            if final_rc != 0:
                # 失败时重置进度条
                self.root.after(0, lambda: self.subtitle_progress_var.set(0))
                print(f"[DEBUG] 处理失败（返回码: {final_rc}），重置进度条")
            # 成功时进度条已在前面设置为100%，这里不需要再设置

            print("=== FFmpeg命令执行完成 ===\n")

    def handle_subtitle_completion(self, returncode, output_path):
        """处理字幕视频生成完成"""
        print("\n=== 处理字幕视频生成完成 ===")
        print("1. 检查输出文件...")

        # 无论成功或失败，先设置进度条为100%（如果是成功的话）
        # 这样用户可以看到处理已完成
        if returncode == 0:
            self.subtitle_progress_var.set(100)
            print("[DEBUG] 设置进度条为100%")

        if returncode == 0 and os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            print("生成成功！")
            print(f"输出文件: {output_path}")
            file_size = os.path.getsize(output_path)
            print(f"文件大小: {file_size} 字节 ({file_size / (1024*1024):.2f} MB)")

            # 打开输出文件所在目录
            print("2. 打开输出目录...")
            output_dir = os.path.dirname(output_path)
            try:
                if sys.platform == 'win32':
                    subprocess.run(['explorer', output_dir], creationflags=subprocess.CREATE_NO_WINDOW)
                else:
                    subprocess.run(['xdg-open', output_dir])
                print(f"已打开目录: {output_dir}")
            except Exception as e:
                print(f"打开目录失败: {e}")

            messagebox.showinfo("完成", f"字幕视频生成成功！\n保存路径：{output_path}")

            # 2秒后重置进度条
            self.root.after(2000, lambda: self.subtitle_progress_var.set(0))
        else:
            print("生成失败！")
            error_msg = "生成失败：\n"
            if returncode != 0:
                error_msg += f"FFmpeg返回错误代码：{returncode}\n"
                print(f"FFmpeg错误代码: {returncode}")
            if not os.path.exists(output_path):
                error_msg += "输出文件未生成\n"
                print("输出文件不存在")
            elif os.path.exists(output_path) and os.path.getsize(output_path) == 0:
                error_msg += "输出文件大小为0\n"
                print("输出文件大小为0")
                try:
                    os.remove(output_path)  # 删除空文件
                    print("已删除空文件")
                except Exception as e:
                    print(f"删除空文件失败: {str(e)}")

            # 失败时重置进度条
            self.subtitle_progress_var.set(0)

            logger.error(error_msg)
            messagebox.showerror("错误", error_msg + "请检查控制台输出获取详细信息")
        print("=== 处理完成 ===\n")

    def update_subtitle_style(self):
        """更新字幕样式"""
        # 获取当前样式设置
        font_size = self.font_size_var.get()
        color_name = self.font_color_var.get()
        position_name = self.position_var.get()

        # 转换为实际值
        color = self.color_mapping.get(color_name, "white")

        # 更新字幕标签样式
        self.subtitle_label.config(
            font=("微软雅黑", int(font_size)),
            fg=color
        )

        # 更新字幕位置
        if position_name == "顶部":
            self.subtitle_label.pack_forget()
            self.subtitle_label.pack(side=tk.TOP, pady=20)
        elif position_name == "中间":
            self.subtitle_label.pack_forget()
            # 使用place方法精确定位到中间
            self.subtitle_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        else:  # 底部
            self.subtitle_label.pack_forget()
            self.subtitle_label.pack(side=tk.BOTTOM, pady=20)

        # 如果正在预览，立即更新当前帧的字幕显示
        if self.is_previewing and hasattr(self, 'subtitle_cap'):
            current_frame = self.subtitle_cap.get(cv2.CAP_PROP_POS_FRAMES)
            current_time = current_frame / self.subtitle_fps

            # 查找当前时间对应的字幕
            current_subtitle = ""
            for subtitle in self.subtitles:
                if subtitle['start'] <= current_time <= subtitle['end']:
                    current_subtitle = subtitle['text']
                    break

            # 更新字幕显示
            self.subtitle_label.config(text=current_subtitle)

    def on_progress_change(self, value):
        """处理进度条拖动"""
        if not self.subtitle_cap:
            return

        # 计算目标时间
        progress = float(value)
        if hasattr(self, 'subtitle_duration') and self.subtitle_duration > 0:
            target_time = (progress / 100.0) * self.subtitle_duration

            try:
                # 更新视频帧位置
                if hasattr(self, 'subtitle_fps') and self.subtitle_fps > 0:
                    self.subtitle_cap.set(cv2.CAP_PROP_POS_FRAMES, int(target_time * self.subtitle_fps))
                    ret, frame = self.subtitle_cap.read()
                    if ret:
                        # 查找当前时间对应的字幕
                        current_subtitle = ""
                        for subtitle in self.subtitles:
                            if subtitle['start'] <= target_time <= subtitle['end']:
                                current_subtitle = subtitle['text']
                                break

                        # 应用当前样式设置
                        font_size = self.font_size_var.get()
                        color_name = self.font_color_var.get()
                        color = self.color_mapping.get(color_name, "white")

                        # 更新字幕显示（应用样式）
                        self.subtitle_label.config(
                            text=current_subtitle,
                            font=("微软雅黑", int(font_size)),
                            fg=color
                        )

                        # 显示帧
                        self.show_subtitle_frame(frame)

            except Exception as e:
                print(f"预览更新失败: {str(e)}")

    def convert_color_to_ass(self, color_name):
        """将颜色名称转换为ASS格式的颜色代码"""
        # 颜色映射表
        color_map = {
            "白色": "FFFFFF",
            "黄色": "FFFF00",
            "绿色": "00FF00",
            "青色": "00FFFF",
            "红色": "FF0000",
            "粉色": "FFC0CB",
            "橙色": "FFA500",
            "紫色": "800080"
        }
        color_code = color_map.get(color_name, "FFFFFF")
        # 返回正确的ASS颜色代码格式
        return f"&H00{color_code}"

    def escape_path_for_ffmpeg(self, path):
        """为FFmpeg转义路径，处理中文字符和特殊字符"""
        # 将路径转换为绝对路径
        abs_path = os.path.abspath(path)
        # 在Windows上，使用正斜杠，但不要转义冒号（会导致FFmpeg错误）
        if sys.platform == 'win32':
            # 替换反斜杠为正斜杠
            abs_path = abs_path.replace('\\', '/')
            # 注意：不要转义冒号，FFmpeg在Windows上可以正确处理C:/这样的路径
        return abs_path

    def validate_paths(self, video_path, subtitle_path):
        """验证视频和字幕文件路径"""
        errors = []

        # 检查视频文件
        if not video_path:
            errors.append("未选择视频文件")
        elif not os.path.exists(video_path):
            errors.append(f"视频文件不存在: {video_path}")
        elif not os.path.isfile(video_path):
            errors.append(f"视频路径不是文件: {video_path}")
        else:
            # 检查视频文件是否可读
            try:
                with open(video_path, 'rb') as f:
                    f.read(1024)  # 尝试读取前1KB
            except Exception as e:
                errors.append(f"视频文件无法读取: {str(e)}")

            # 检查视频文件格式
            video_ext = os.path.splitext(video_path)[1].lower()
            if video_ext not in ['.mp4', '.avi', '.mov', '.mkv', '.flv', '.ts', '.wmv', '.m4v']:
                errors.append(f"不支持的视频格式: {video_ext}")

        # 检查字幕文件
        if not subtitle_path:
            errors.append("未选择字幕文件")
        elif not os.path.exists(subtitle_path):
            errors.append(f"字幕文件不存在: {subtitle_path}")
        elif not os.path.isfile(subtitle_path):
            errors.append(f"字幕路径不是文件: {subtitle_path}")
        else:
            # 检查字幕文件是否可读
            try:
                with open(subtitle_path, 'r', encoding='utf-8') as f:
                    f.read(1024)  # 尝试读取前1KB
            except Exception as e:
                errors.append(f"字幕文件无法读取: {str(e)}")

            # 检查字幕文件格式
            subtitle_ext = os.path.splitext(subtitle_path)[1].lower()
            if subtitle_ext not in ['.srt', '.ass', '.ssa']:
                errors.append(f"不支持的字幕格式: {subtitle_ext}")

        return errors

    def get_position_alignment(self, position):
        """获取字幕位置对应的ASS对齐值"""
        position_map = {
            '顶部': '8',
            '中间': '5',
            '底部': '2'
        }
        return position_map.get(position, '2')

    def redraw_track(self, event=None):
        """重绘轨道和滑块，适应新的窗口大小"""
        width = self.track_canvas.winfo_width()
        height = self.track_canvas.winfo_height()

        # 清除现有内容
        self.track_canvas.delete("all")

        # 计算边距
        margin = 20
        track_width = width - (2 * margin)

        # 绘制轨道线
        self.track_canvas.create_line(
            margin, height/2,
                    width - margin, height/2,
            fill="#666666", width=4
        )

        # 获取当前滑块位置的比例
        if hasattr(self, 'start_btn'):
            start_pos = self.get_slider_position("start")
            if start_pos is not None:
                start_ratio = (start_pos - margin) / track_width
            else:
                start_ratio = 0
        else:
            start_ratio = 0

        if hasattr(self, 'end_btn'):
            end_pos = self.get_slider_position("end")
            if end_pos is not None:
                end_ratio = (end_pos - margin) / track_width
            else:
                end_ratio = 1
        else:
            end_ratio = 1

        # 使用比例重新计算滑块位置
        start_x = margin + (track_width * start_ratio)
        end_x = margin + (track_width * end_ratio)

        # 重绘滑块 - 使用倒水滴形状+垂直线设计
        # 开始按钮
        start_icon_top = start_x - 8, height/2 - 15  # 倒水滴顶部左点
        start_icon_mid = start_x, height/2 - 5      # 倒水滴底部尖点
        start_icon_right = start_x + 8, height/2 - 15  # 倒水滴顶部右点
        start_line_top = start_x, height/2 - 5      # 垂直线顶部
        start_line_bottom = start_x, height/2 + 15  # 垂直线底部

        self.start_btn = self.track_canvas.create_polygon(
            start_icon_top[0], start_icon_top[1],
            start_icon_mid[0], start_icon_mid[1],
            start_icon_right[0], start_icon_right[1],
            fill="#00FF00", outline="#00FF00", width=1, tags="start"
        )
        self.start_line = self.track_canvas.create_line(
            start_line_top[0], start_line_top[1],
            start_line_bottom[0], start_line_bottom[1],
            fill="#00FF00", width=2, tags="start"
        )

        # 结束按钮
        end_icon_top = end_x - 8, height/2 - 15  # 倒水滴顶部左点
        end_icon_mid = end_x, height/2 - 5      # 倒水滴底部尖点
        end_icon_right = end_x + 8, height/2 - 15  # 倒水滴顶部右点
        end_line_top = end_x, height/2 - 5      # 垂直线顶部
        end_line_bottom = end_x, height/2 + 15  # 垂直线底部

        self.end_btn = self.track_canvas.create_polygon(
            end_icon_top[0], end_icon_top[1],
            end_icon_mid[0], end_icon_mid[1],
            end_icon_right[0], end_icon_right[1],
            fill="#FF0000", outline="#FF0000", width=1, tags="end"
        )
        self.end_line = self.track_canvas.create_line(
            end_line_top[0], end_line_top[1],
            end_line_bottom[0], end_line_bottom[1],
            fill="#FF0000", width=2, tags="end"
        )

        # 重新绑定事件
        self.track_canvas.tag_bind("start", "<B1-Motion>", self.move_start)
        self.track_canvas.tag_bind("end", "<B1-Motion>", self.move_end)

    def redraw_drop_area(self, event=None):
        """重绘拖放区域，使其在窗口最大化时能够自适应居中显示"""
        width = self.drop_canvas.winfo_width()
        height = self.drop_canvas.winfo_height()

        # 清除现有内容
        self.drop_canvas.delete("all")

        # 居中显示+号和文字
        self.drop_canvas.create_text(width/2, height/2 - 40, text="+", font=("Arial", 72), fill="#666666", tags="plus")
        self.drop_canvas.create_text(width/2, height/2 + 40, text="拖放文件到这里", font=("微软雅黑", 12), fill="#888888", tags="text")

    def on_closing(self):
        """窗口关闭时的处理"""
        print("[DEBUG] 程序正在关闭，清理所有进程...")

        # 设置退出标志
        self.is_generating = False

        # 终止所有活跃的FFmpeg进程
        self.terminate_all_processes()

        # 强制退出程序
        try:
            self.root.quit()  # 退出主循环
            self.root.destroy()  # 销毁窗口
        except:
            pass

        # 如果程序仍然没有退出，强制终止
        import os
        import time
        time.sleep(2)  # 给进程一些时间完成清理
        os._exit(0)  # 强制退出

    def terminate_all_processes(self):
        """终止所有活跃的FFmpeg进程"""
        print(f"[DEBUG] 发现 {len(self.active_processes)} 个活跃进程")

        for process in self.active_processes:
            try:
                if process.poll() is None:  # 进程还在运行
                    print(f"[DEBUG] 尝试优雅终止FFmpeg进程 PID: {process.pid}")

                    # 方法1：向FFmpeg的stdin发送'q'命令（最可靠的方式）
                    try:
                        if process.stdin and not process.stdin.closed:
                            process.stdin.write('q\n')
                            process.stdin.flush()
                            print(f"[DEBUG] 已向进程 {process.pid} 发送'q'命令")

                            # 等待FFmpeg优雅退出
                            try:
                                process.wait(timeout=5)  # 给FFmpeg 5秒时间完成清理
                                print(f"[DEBUG] 进程 {process.pid} 已通过'q'命令优雅退出")
                                continue  # 成功退出，继续处理下一个进程
                            except subprocess.TimeoutExpired:
                                print(f"[DEBUG] 进程 {process.pid} 未在5秒内通过'q'命令退出")
                    except Exception as e:
                        print(f"[DEBUG] 发送'q'命令失败: {e}")

                    # 方法2：如果'q'命令失败，尝试SIGTERM
                    try:
                        process.terminate()
                        print(f"[DEBUG] 已向进程 {process.pid} 发送SIGTERM信号")

                        # 等待SIGTERM生效
                        try:
                            process.wait(timeout=3)  # 给FFmpeg 3秒时间完成清理
                            print(f"[DEBUG] 进程 {process.pid} 已通过SIGTERM信号退出")
                            continue  # 成功退出，继续处理下一个进程
                        except subprocess.TimeoutExpired:
                            print(f"[DEBUG] 进程 {process.pid} 未在3秒内通过SIGTERM退出")
                    except Exception as e:
                        print(f"[DEBUG] 发送SIGTERM失败: {e}")

                    # 方法3：最后尝试强制终止
                    try:
                        process.kill()
                        print(f"[DEBUG] 已向进程 {process.pid} 发送SIGKILL信号")
                        try:
                            process.wait(timeout=1)
                            print(f"[DEBUG] 进程 {process.pid} 已被强制终止")
                        except subprocess.TimeoutExpired:
                            print(f"[DEBUG] 进程 {process.pid} 无法终止")
                    except Exception as e:
                        print(f"[DEBUG] 强制终止失败: {e}")

            except Exception as e:
                print(f"[DEBUG] 终止进程时出错: {e}")

        # 清空进程列表
        self.active_processes.clear()

        # 不再进行系统级强制清理，避免损坏视频文件
        print("[DEBUG] 进程清理完成（已避免强制终止以保护视频文件完整性）")

    def cleanup_orphaned_processes(self):
        """清理可能残留的FFmpeg进程"""
        try:
            if sys.platform == 'win32':
                print("[DEBUG] 检查并清理残留的FFmpeg进程...")
                # 检查是否有ffmpeg.exe进程在运行
                result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq ffmpeg.exe'],
                                        capture_output=True, text=True,
                                        creationflags=subprocess.CREATE_NO_WINDOW)

                if 'ffmpeg.exe' in result.stdout:
                    print("[DEBUG] 发现残留的FFmpeg进程，正在清理...")
                    # 强制终止所有ffmpeg进程
                    subprocess.run(['taskkill', '/F', '/IM', 'ffmpeg.exe'],
                                   stdout=subprocess.DEVNULL,
                                   stderr=subprocess.DEVNULL,
                                   creationflags=subprocess.CREATE_NO_WINDOW)
                    print("[DEBUG] 残留进程清理完成")
                else:
                    print("[DEBUG] 未发现残留的FFmpeg进程")
        except Exception as e:
            print(f"[DEBUG] 清理残留进程时出错: {e}")

    def signal_handler(self, signum, frame):
        """信号处理函数"""
        print(f"[DEBUG] 收到信号 {signum}，开始优雅退出...")
        self.cleanup_on_exit()
        sys.exit(0)

    def cleanup_on_exit(self):
        """程序退出时的清理函数"""
        print("[DEBUG] 程序退出清理开始...")
        try:
            # 设置退出标志
            self.is_generating = False

            # 优雅清理：优先使用'q'命令
            if self.active_processes:
                print("[DEBUG] 优雅终止所有FFmpeg进程...")

                for process in self.active_processes:
                    if process.poll() is None:
                        try:
                            print(f"[DEBUG] 优雅终止进程 {process.pid}")

                            # 方法1：向FFmpeg的stdin发送'q'命令（最可靠的方式）
                            try:
                                if process.stdin and not process.stdin.closed:
                                    process.stdin.write('q\n')
                                    process.stdin.flush()
                                    print(f"[DEBUG] 已向进程 {process.pid} 发送'q'命令")

                                    # 等待FFmpeg优雅退出
                                    try:
                                        process.wait(timeout=10)  # 给FFmpeg 10秒时间完成清理
                                        print(f"[DEBUG] 进程 {process.pid} 已通过'q'命令优雅退出")
                                        continue  # 成功退出，继续处理下一个进程
                                    except subprocess.TimeoutExpired:
                                        print(f"[DEBUG] 进程 {process.pid} 未在10秒内通过'q'命令退出")
                            except Exception as e:
                                print(f"[DEBUG] 发送'q'命令失败: {e}")

                            # 方法2：如果'q'命令失败，尝试SIGTERM
                            try:
                                process.terminate()
                                print(f"[DEBUG] 已向进程 {process.pid} 发送SIGTERM信号")
                            except Exception as e:
                                print(f"[DEBUG] 发送SIGTERM失败: {e}")

                        except Exception as e:
                            print(f"[DEBUG] 终止进程失败: {e}")

                # 等待2秒让进程完成清理
                import time
                time.sleep(2)

            # 不再进行系统级强制清理
            print("[DEBUG] 退出清理完成（已避免强制终止以保护视频文件完整性）")
        except Exception as e:
            print(f"[DEBUG] 退出清理时出错: {e}")

    def check_video_file_integrity(self, file_path):
        """检查视频文件完整性"""
        try:
            # 使用FFmpeg检查文件完整性 - 修复中文路径编码问题
            check_cmd = [FFMPEG_PATH, '-v', 'error', '-i', file_path, '-f', 'null', '-']

            # 根据文件大小动态设置超时时间
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)  # 转换为MB
            timeout_seconds = max(60, int(file_size_mb / 100))  # 每100MB给1分钟，最少60秒

            print(f"文件大小: {file_size_mb:.1f}MB, 设置超时时间: {timeout_seconds}秒")

            if sys.platform == 'win32':
                # 将命令列表转换为字符串，确保路径编码正确
                cmd_str = ' '.join(f'"{arg}"' if ' ' in arg or any(ord(c) > 127 for c in arg) else arg for arg in check_cmd)
                result = subprocess.run(
                    cmd_str,
                    shell=True,
                    capture_output=True,
                    text=True,
                    timeout=timeout_seconds,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
            else:
                result = subprocess.run(
                    check_cmd,
                    capture_output=True,
                    text=True,
                    timeout=timeout_seconds
                )

            # 如果返回码为0且没有错误输出，说明文件完整
            if result.returncode == 0 and not result.stderr.strip():
                return True
            else:
                print(f"文件完整性检查失败: {result.stderr}")
                return False

        except Exception as e:
            print(f"文件完整性检查出错: {e}")
            return False

    def disable_subtitle_buttons(self):
        """禁用字幕相关按钮"""
        if hasattr(self, 'generate_btn'):
            self.generate_btn.config(text="生成中...", state="disabled")
        if hasattr(self, 'preview_btn'):
            self.preview_btn.config(state="disabled")
        if hasattr(self, 'save_progress_btn'):
            self.save_progress_btn.config(state="normal")  # 生成中时启用保存进度按钮

    def enable_subtitle_buttons(self):
        """启用字幕相关按钮"""
        if hasattr(self, 'generate_btn'):
            self.generate_btn.config(text="生成硬字幕视频", state="normal")
        if hasattr(self, 'preview_btn'):
            self.preview_btn.config(state="normal")
        if hasattr(self, 'save_progress_btn'):
            self.save_progress_btn.config(state="disabled")  # 完成后禁用保存进度按钮

    def enable_soft_subtitle_buttons(self):
        """启用软字幕相关按钮"""
        if hasattr(self, 'soft_generate_btn'):
            self.soft_generate_btn.config(state="normal")
        if hasattr(self, 'soft_preview_btn'):
            self.soft_preview_btn.config(state="normal")

    def save_current_progress(self):
        """保存当前进度"""
        if not self.is_generating or not self.active_processes:
            messagebox.showwarning("警告", "当前没有正在生成的视频")
            return

        try:
            # 使用优雅的方式让FFmpeg保存当前进度
            for process in self.active_processes:
                if process.poll() is None:
                    print(f"[DEBUG] 发送保存进度信号...")

                    # 方法1：向FFmpeg的stdin发送'q'命令（最可靠的方式）
                    try:
                        if process.stdin and not process.stdin.closed:
                            process.stdin.write('q\n')
                            process.stdin.flush()
                            print(f"[DEBUG] 已发送'q'命令保存进度")

                            # 等待FFmpeg优雅退出
                            try:
                                process.wait(timeout=10)  # 给FFmpeg 10秒时间完成清理
                                print(f"[DEBUG] 进程已通过'q'命令优雅退出")
                                messagebox.showinfo("提示", "进度保存完成！")
                                return
                            except subprocess.TimeoutExpired:
                                print(f"[DEBUG] 进程未在10秒内通过'q'命令退出")
                    except Exception as e:
                        print(f"[DEBUG] 发送'q'命令失败: {e}")

                    # 方法2：如果'q'命令失败，尝试SIGTERM
                    try:
                        process.terminate()
                        print(f"[DEBUG] 已发送SIGTERM信号保存进度")
                    except Exception as e:
                        print(f"[DEBUG] 发送SIGTERM失败: {e}")

            messagebox.showinfo("提示", "正在保存当前进度，请稍等...")

        except Exception as e:
            messagebox.showerror("错误", f"保存进度失败: {str(e)}")

    def create_soft_subtitle_tab(self):
        """创建软字幕标签页（界面与硬字幕一致，但命令不同）"""
        # 创建主框架
        main_frame = tk.Frame(self.soft_subtitle_tab, bg="#333333")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 配置主框架的网格权重
        main_frame.grid_rowconfigure(1, weight=1)  # 预览区域可扩展
        main_frame.grid_rowconfigure(3, weight=0)  # 控制区域固定大小

        # 文件选择区域 - 放在顶部一行
        file_frame = tk.Frame(main_frame, bg="#333333")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 视频和字幕文件选择放在同一行
        ttk.Label(file_frame, text="视频文件：").pack(side=tk.LEFT, padx=(0, 5))
        self.soft_video_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.soft_video_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="选择视频", command=self.select_soft_video).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(file_frame, text="字幕文件：").pack(side=tk.LEFT, padx=(0, 5))
        self.soft_subtitle_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.soft_subtitle_path_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="选择字幕", command=self.select_soft_subtitle).pack(side=tk.LEFT)

        # 预览区域 - 最大化显示
        preview_frame = tk.Frame(main_frame, bg="#1e1e1e")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建预览画布
        self.soft_subtitle_preview_canvas = tk.Canvas(preview_frame, bg="#1e1e1e", bd=0, highlightthickness=0)
        self.soft_subtitle_preview_canvas.pack(fill=tk.BOTH, expand=True)

        # 字幕显示标签
        self.soft_subtitle_label = tk.Label(
            preview_frame,
            text="",
            bg="#1e1e1e",
            fg="white",
            font=("微软雅黑", 24),
            wraplength=1200
        )
        self.soft_subtitle_label.pack(side=tk.BOTTOM, pady=20)

        # 进度条拖动区域 - 在预览区域和控制区域之间
        progress_slider_frame = tk.Frame(main_frame, bg="#333333")
        progress_slider_frame.pack(fill=tk.X, pady=(0, 10))

        # 播放进度条（可拖动）
        self.soft_subtitle_progress_slider = ttk.Scale(
            progress_slider_frame,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            command=self.on_soft_progress_change
        )
        self.soft_subtitle_progress_slider.pack(fill=tk.X, padx=5)

        # 控制区域 - 底部
        control_frame = tk.Frame(main_frame, bg="#333333")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # 样式设置区域 - 左侧
        style_frame = tk.Frame(control_frame, bg="#333333")
        style_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 字体大小
        ttk.Label(style_frame, text="字号：").pack(side=tk.LEFT)
        self.soft_font_size_var = tk.StringVar(value="12")
        font_sizes = ["12", "14", "16", "18", "20", "22", "24", "26", "28", "30", "32", "34", "36", "38", "40", "42", "44", "46", "48"]
        soft_font_size_combo = ttk.Combobox(
            style_frame,
            textvariable=self.soft_font_size_var,
            values=font_sizes,
            state="readonly",
            width=4
        )
        soft_font_size_combo.pack(side=tk.LEFT, padx=(0, 15))
        soft_font_size_combo.bind('<<ComboboxSelected>>', lambda e: self.update_soft_subtitle_style())

        # 字幕颜色
        ttk.Label(style_frame, text="颜色：").pack(side=tk.LEFT)
        self.soft_font_color_var = tk.StringVar(value="白色")
        colors = [
            ("白色", "white"),
            ("黑色", "black"),
            ("红色", "red"),
            ("绿色", "green"),
            ("蓝色", "blue"),
            ("黄色", "yellow"),
            ("青色", "cyan"),
            ("洋红色", "magenta"),
            ("橙色", "orange"),
            ("粉色", "pink"),
            ("紫色", "purple"),
            ("灰色", "gray")
        ]
        soft_color_combo = ttk.Combobox(
            style_frame,
            textvariable=self.soft_font_color_var,
            values=[color[0] for color in colors],
            state="readonly",
            width=6
        )
        soft_color_combo.pack(side=tk.LEFT, padx=(0, 15))
        soft_color_combo.bind('<<ComboboxSelected>>', lambda e: self.update_soft_subtitle_style())

        # 保存颜色映射（如果还没有）
        if not hasattr(self, 'color_mapping'):
            self.color_mapping = dict(colors)
            self.reverse_color_mapping = {v: k for k, v in dict(colors).items()}

        # 字幕位置
        ttk.Label(style_frame, text="位置：").pack(side=tk.LEFT)
        self.soft_position_var = tk.StringVar(value="底部")
        positions = [("顶部", "top"), ("中间", "middle"), ("底部", "bottom")]
        soft_position_combo = ttk.Combobox(
            style_frame,
            textvariable=self.soft_position_var,
            values=[pos[0] for pos in positions],
            state="readonly",
            width=6
        )
        soft_position_combo.pack(side=tk.LEFT, padx=(0, 15))
        soft_position_combo.bind('<<ComboboxSelected>>', lambda e: self.update_soft_subtitle_style())

        # 保存位置映射（如果还没有）
        if not hasattr(self, 'position_mapping'):
            self.position_mapping = dict(positions)

        # 按钮区域 - 右侧
        button_frame = tk.Frame(control_frame, bg="#333333")
        button_frame.pack(side=tk.RIGHT)

        # 预览按钮
        self.soft_preview_btn = ttk.Button(button_frame, text="预览字幕", command=self.preview_soft_subtitle, width=10)
        self.soft_preview_btn.pack(side=tk.LEFT, padx=5)

        # 生成按钮
        self.soft_generate_btn = ttk.Button(button_frame, text="生成软字幕视频", command=self.generate_soft_subtitle_video, width=12)
        self.soft_generate_btn.pack(side=tk.LEFT, padx=5)

        # 进度条 - 最底部
        progress_frame = tk.Frame(main_frame, bg="#333333")
        progress_frame.pack(fill=tk.X)

        self.soft_subtitle_progress_var = tk.DoubleVar()
        self.soft_subtitle_progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.soft_subtitle_progress_var,
            maximum=100
        )
        self.soft_subtitle_progress_bar.pack(fill=tk.X)

    def create_audio_denoise_tab(self):
        """创建音频降噪标签页界面"""
        # 主框架
        main_frame = tk.Frame(self.audio_denoise_tab, bg="#333333")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 配置主框架的网格权重
        main_frame.grid_rowconfigure(1, weight=1)  # 预览区域可扩展
        main_frame.grid_rowconfigure(3, weight=0)  # 控制区域固定大小

        # 文件选择区域
        file_frame = tk.Frame(main_frame, bg="#333333")
        file_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(file_frame, text="视频文件：").pack(side=tk.LEFT, padx=(0, 5))
        self.video_audio_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.video_audio_file_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(file_frame, text="选择视频", command=self.select_video_for_audio_denoise).pack(side=tk.LEFT, padx=(0, 10))

        # 拖放支持
        self.video_audio_file_var.trace('w', self.on_video_audio_file_changed)

        # 预览区域
        preview_frame = tk.Frame(main_frame, bg="#1e1e1e")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 音频预览画布
        self.audio_preview_canvas = tk.Canvas(preview_frame, bg="#1e1e1e", bd=0, highlightthickness=0)
        self.audio_preview_canvas.pack(fill=tk.BOTH, expand=True)

        # 视频信息显示
        self.video_audio_info_label = tk.Label(preview_frame, text="请选择视频文件",
                                               bg="#1e1e1e", fg="white", font=("Arial", 12))
        self.video_audio_info_label.pack(pady=20)

        # 控制区域
        control_frame = tk.Frame(main_frame, bg="#333333")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # 降噪参数设置
        params_frame = tk.Frame(control_frame, bg="#333333")
        params_frame.pack(fill=tk.X, pady=(0, 10))

        # 降噪强度
        ttk.Label(params_frame, text="降噪强度：").pack(side=tk.LEFT, padx=(0, 5))
        self.noise_reduction_var = tk.DoubleVar(value=self.denoise_settings['noise_reduction'])
        noise_scale = ttk.Scale(params_frame, from_=0, to=50, variable=self.noise_reduction_var,
                                orient=tk.HORIZONTAL, length=200)
        noise_scale.pack(side=tk.LEFT, padx=(0, 10))
        self.noise_value_label = tk.Label(params_frame, text="10.0", bg="#333333", fg="white")
        self.noise_value_label.pack(side=tk.LEFT, padx=(0, 20))

        # 声音放大
        ttk.Label(params_frame, text="声音放大：").pack(side=tk.LEFT, padx=(0, 5))
        self.volume_boost_var = tk.DoubleVar(value=0.0)  # 0表示不放大
        volume_scale = ttk.Scale(params_frame, from_=0, to=100, variable=self.volume_boost_var,
                                 orient=tk.HORIZONTAL, length=200)
        volume_scale.pack(side=tk.LEFT, padx=(0, 10))
        self.volume_value_label = tk.Label(params_frame, text="0.0dB", bg="#333333", fg="white")
        self.volume_value_label.pack(side=tk.LEFT, padx=(0, 20))

        # 保留人声选项
        self.preserve_voice_var = tk.BooleanVar(value=self.denoise_settings['preserve_voice'])
        preserve_check = ttk.Checkbutton(params_frame, text="保留人声", variable=self.preserve_voice_var)
        preserve_check.pack(side=tk.LEFT, padx=(0, 20))

        # 输出格式选择 - 根据输入视频自动选择
        ttk.Label(params_frame, text="输出格式：").pack(side=tk.LEFT, padx=(0, 5))
        self.video_output_format_var = tk.StringVar(value="自动")
        format_combo = ttk.Combobox(params_frame, textvariable=self.video_output_format_var,
                                    values=['自动', 'mp4', 'avi', 'mov', 'mkv'], width=8, state='readonly')
        format_combo.pack(side=tk.LEFT)

        # 绑定降噪强度和音量变化事件
        self.noise_reduction_var.trace('w', self.on_noise_reduction_changed)
        self.volume_boost_var.trace('w', self.on_volume_boost_changed)

        # 按钮区域
        button_frame = tk.Frame(control_frame, bg="#333333")
        button_frame.pack(fill=tk.X)

        # 预览按钮
        self.video_audio_preview_btn = ttk.Button(button_frame, text="预览视频", command=self.preview_video_audio)
        self.video_audio_preview_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 开始降噪按钮
        self.video_audio_denoise_btn = ttk.Button(button_frame, text="开始声音处理", command=self.start_video_audio_denoise)
        self.video_audio_denoise_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 停止按钮
        self.video_audio_stop_btn = ttk.Button(button_frame, text="停止", command=self.stop_video_audio_denoise, state='disabled')
        self.video_audio_stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        # 进度条
        self.audio_progress_var = tk.DoubleVar()
        self.audio_progress_bar = ttk.Progressbar(main_frame, variable=self.audio_progress_var, maximum=100)
        self.audio_progress_bar.pack(fill=tk.X, pady=(10, 0))

        # 状态标签
        self.video_audio_status_label = tk.Label(main_frame, text="", bg="#333333", fg="white")
        self.video_audio_status_label.pack(pady=(5, 0))

    def __del__(self):
        """清理资源"""
        try:
            self.terminate_all_processes()
        except:
            pass


if __name__ == "__main__":
    # 确保异常能够被正确打印
    sys.stdout = sys.__stdout__
    sys.stderr = sys.__stderr__

    try:
        app = VideoTrimmerPro()
    except Exception as e:
        logger.error(f"程序启动失败: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
