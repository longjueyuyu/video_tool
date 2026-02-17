# -*- coding: utf-8 -*-
import yt_dlp
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import sys
import time
import subprocess
from pathlib import Path
import browser_cookie3
import json


def get_program_dir():
    """获取程序所在目录：脚本运行时为脚本目录，打包成 exe 后为 exe 所在目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_ffmpeg_path():
    """
    返回程序目录下 ffmpeg 可执行文件的路径，供 yt-dlp 合并音视频使用。
    优先：程序目录/ffmpeg.exe，其次：程序目录/_internal/ffmpeg.exe。
    未找到则返回 None（yt-dlp 会使用系统 PATH 中的 ffmpeg）。
    """
    base = get_program_dir()
    for name in ("ffmpeg.exe", "ffmpeg"):
        path = os.path.join(base, name)
        if os.path.isfile(path):
            return path
    internal = os.path.join(base, "_internal")
    if os.path.isdir(internal):
        for name in ("ffmpeg.exe", "ffmpeg"):
            path = os.path.join(internal, name)
            if os.path.isfile(path):
                return path
    return None

class YouTubeDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("YouTube视频下载器")
        
        # 将窗口居中显示
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - 800) // 2
        y = (screen_height - 800) // 2
        self.root.geometry(f"800x800+{x}+{y}")
        
        self.root.configure(bg='#f0f0f0')
        
        # 设置样式
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TEntry', padding=5)
        style.configure('TLabel', padding=5)
        style.configure('Header.TLabel', font=('Arial', 10, 'bold'))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # URL输入框和下载按钮行
        url_frame = ttk.Frame(main_frame)
        url_frame.pack(fill=tk.X, pady=(0, 10))
        
        url_label = ttk.Label(url_frame, text="视频链接:")
        url_label.pack(side=tk.LEFT, padx=(0, 5))
        
        # 创建支持右键菜单的输入框
        self.url_entry = ttk.Entry(url_frame)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # 添加右键菜单
        self.context_menu = tk.Menu(self.url_entry, tearoff=0)
        self.context_menu.add_command(label="粘贴", command=self.paste_url)
        self.url_entry.bind("<Button-3>", self.show_context_menu)
        
        self.download_button = ttk.Button(url_frame, text="开始下载", command=self.start_download)
        self.download_button.pack(side=tk.RIGHT)
        
        # 保存路径输入框和按钮行
        path_frame = ttk.Frame(main_frame)
        path_frame.pack(fill=tk.X, pady=(0, 20))
        
        path_label = ttk.Label(path_frame, text="保存位置:")
        path_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.path_entry = ttk.Entry(path_frame)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_button = ttk.Button(path_frame, text="选择目录", command=self.browse_path)
        browse_button.pack(side=tk.LEFT, padx=(0, 10))
        
        cancel_button = ttk.Button(path_frame, text="取消下载", command=self.cancel_download)
        cancel_button.pack(side=tk.RIGHT)
        
        # 创建下载信息表格（初始隐藏）
        self.info_frame = ttk.Frame(main_frame)
        
        # 创建表头框架
        headers_frame = ttk.Frame(self.info_frame)
        headers_frame.pack(fill=tk.X)
        
        # 表头和列宽配置
        headers = ["进度", "速度", "已下载总大小", "剩余时间"]
        col_widths = [10, 15, 25, 10]  # 设置每列的宽度比例
        
        for i, (header, width) in enumerate(zip(headers, col_widths)):
            headers_frame.grid_columnconfigure(i, weight=width)
            label = ttk.Label(headers_frame, text=header, style='Header.TLabel', anchor=tk.CENTER)
            label.grid(row=0, column=i, sticky=tk.EW, padx=5, pady=5)
            
        # 创建进度列表框架（带滚动条）
        list_frame = ttk.Frame(self.info_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建进度列表
        self.progress_list = tk.Text(list_frame, height=20, yscrollcommand=scrollbar.set)
        self.progress_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.progress_list.yview)
        
        # 设置默认保存路径：程序所在目录/YouTube（支持选择目录修改）
        default_path = os.path.join(get_program_dir(), "YouTube")
        os.makedirs(default_path, exist_ok=True)  # 确保目录存在
        self.path_entry.insert(0, default_path)
        
        # 下载状态变量
        self.is_downloading = False
        self.current_download = None
        self.download_thread = None
        self.retry_count = 0
        self.max_retries = 10
        self.total_retry_count = 0

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def paste_url(self):
        try:
            url = self.root.clipboard_get()
            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, url)
        except Exception as e:
            print(f"Error pasting URL: {str(e)}")

    def update_progress_list(self, progress, speed, size, eta, is_status=False):
        try:
            if not self.is_downloading:
                return
                
            if is_status:
                self.progress_list.insert(tk.END, f"\n>>> {progress}\n")
            else:
                progress_text = f"{progress:.1f}%".ljust(10)
                if isinstance(speed, (int, float)):
                    speed_text = f"{speed:.1f}MB/s".ljust(15)
                else:
                    speed_text = str(speed).ljust(15)
                size_text = str(size).ljust(25)
                eta_text = f"{eta}s" if str(eta).isdigit() else str(eta)
                
                progress_line = f"{progress_text}{speed_text}{size_text}{eta_text}\n"
                self.progress_list.insert(tk.END, progress_line)
            
            self.progress_list.see(tk.END)
            self.root.update()
        except Exception as e:
            print(f"Error updating progress list: {str(e)}")

    def get_cookies(self):
        try:
            return browser_cookie3.chrome(domain_name='.youtube.com')
        except:
            try:
                return browser_cookie3.firefox(domain_name='.youtube.com')
            except:
                return None

    def save_cookies(self):
        """保存cookies到文件（放在程序所在目录，便于 exe 与脚本共用）"""
        cookies_file = os.path.join(get_program_dir(), 'youtube_cookies.txt')
        if not os.path.exists(cookies_file):
            # 创建一个示例cookie文件
            example_cookies = [
                {
                    "domain": ".youtube.com",
                    "name": "CONSENT",
                    "path": "/",
                    "value": "YES+cb"
                }
            ]
            with open(cookies_file, 'w') as f:
                for cookie in example_cookies:
                    f.write(f"{cookie['domain']}\tTRUE\t{cookie['path']}\tFALSE\t2147483647\t{cookie['name']}\t{cookie['value']}\n")
        return cookies_file

    def download_video(self, url):
        try:
            save_path = self.path_entry.get()
            os.makedirs(save_path, exist_ok=True)
            
            self.update_progress_list("开始处理下载请求...", None, None, None, True)
            
            # 获取或创建cookies文件
            cookies_file = self.save_cookies()
            
            # 新增下载选项配置（合并音视频需要 ffmpeg，优先使用程序目录内捆绑的 ffmpeg）
            ydl_opts = {
                'format': 'bestvideo[height<=4320]+bestaudio/best',  # 下载分辨率小于等于8K的最好的视频和音频
                'outtmpl': os.path.join(save_path, '%(title)s.%(ext)s'),
                'noplaylist': True,  # 仅下载单个视频，不下载播放列表
                'retries': 5,  # 设置重试次数
                'retry_sleep': 5,  # 设置每次重试之间的间隔时间，单位为秒
                'socket_timeout': 30,  # 设置连接超时时间
                'max_filesize': 10000000000,  # 限制最大文件大小，单位字节，设置为10GB
                'progress_hooks': [self.progress_hook],  # 显示下载进度
                'verbose': True,  # 打开详细输出
                'merge_output_format': 'mp4',
                'postprocessors': [{
                    'key': 'FFmpegVideoConvertor',
                    'preferedformat': 'mp4',
                }],
                'http_headers': {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                    'Accept-Language': 'en-us,en;q=0.5',
                    'Sec-Fetch-Mode': 'navigate',
                },
                'cookies': cookies_file
            }
            # 使用程序目录下的 ffmpeg（打包发给别人时无需对方本机安装 ffmpeg）
            ffmpeg_path = get_ffmpeg_path()
            if ffmpeg_path:
                ydl_opts['ffmpeg_location'] = ffmpeg_path
            
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                try:
                    self.current_download = ydl
                    info_dict = ydl.extract_info(url, download=False)
                    video_title = info_dict.get('title', 'Unknown')
                    video_duration = info_dict.get('duration')
                    if video_duration:
                        mins, secs = divmod(video_duration, 60)
                        hours, mins = divmod(mins, 60)
                        duration_str = f"{hours:02d}:{mins:02d}:{secs:02d}" if hours else f"{mins:02d}:{secs:02d}"
                    else:
                        duration_str = "未知"
                    
                    video_resolution = info_dict.get('height', 'Unknown')
                    
                    # 显示视频信息
                    info_msg = f"视频信息:\n标题: {video_title}\n时长: {duration_str}\n分辨率: {video_resolution}p"
                    self.update_progress_list(info_msg, None, None, None, True)
                    
                    # 开始下载
                    self.update_progress_list("正在下载视频...", None, None, None, True)
                    ydl.download([url])
                except Exception as e:
                    if "Sign in to confirm you're not a bot" in str(e):
                        # 如果遇到验证码，提示用户
                        self.update_progress_list("需要登录YouTube账号以继续下载...", None, None, None, True)
                        if messagebox.askyesno("提示", "需要登录YouTube账号才能下载。是否打开YouTube网站登录？"):
                            import webbrowser
                            webbrowser.open("https://www.youtube.com")
                            time.sleep(5)  # 等待用户登录
                            # 重试下载
                            ydl.download([url])
                    else:
                        raise e
                
            self.retry_count = 0
            self.total_retry_count = 0
            self.update_progress_list("下载完成！正在打开下载目录...", None, None, None, True)
            self.open_download_directory()
                
        except Exception as e:
            if not self.is_downloading:
                return
                
            error_msg = f"下载出错: {str(e)}"
            self.update_progress_list(error_msg, None, None, None, True)
            
            if self.retry_count < self.max_retries:
                self.retry_count += 1
                self.total_retry_count += 1
                retry_msg = f"正在进行第 {self.retry_count} 次自动重试..."
                self.update_progress_list(retry_msg, None, None, None, True)
                time.sleep(2)
                self.retry_download(url, clear_progress=False)
            else:
                if self.total_retry_count >= 20:
                    retry_msg = "已重试多次仍然失败，是否继续尝试？\n点击\"是\"继续重试，点击\"否\"取消下载"
                    if messagebox.askyesno("警告", retry_msg):
                        self.retry_count = 0
                        self.retry_download(url, clear_progress=False)
                    else:
                        self.is_downloading = False
                        self.download_button.configure(text="开始下载")
                else:
                    if messagebox.askyesno("错误", "自动重试失败，是否手动重试？"):
                        self.retry_count = 0
                        self.retry_download(url, clear_progress=False)
                    else:
                        self.is_downloading = False
                        self.download_button.configure(text="开始下载")
        finally:
            if self.is_downloading:
                self.is_downloading = False
                self.download_button.configure(text="开始下载")

    def progress_hook(self, d):
        if not self.is_downloading:
            return
            
        if d['status'] == 'downloading':
            try:
                # 计算进度
                if d.get('total_bytes'):
                    progress = (d['downloaded_bytes'] / d['total_bytes']) * 100
                    total_size = self.format_size(d['total_bytes'])
                elif d.get('total_bytes_estimate'):
                    progress = (d['downloaded_bytes'] / d['total_bytes_estimate']) * 100
                    total_size = self.format_size(d['total_bytes_estimate'])
                else:
                    progress = 0
                    total_size = "未知"
                
                # 计算下载速度 (转换为MB/s)
                speed = d.get('speed', 0)
                if speed:
                    speed_mb = speed / (1024 * 1024)
                else:
                    speed_mb = "计算中..."
                
                # 计算已下载大小
                downloaded = self.format_size(d['downloaded_bytes'])
                size_str = f"{downloaded} / {total_size}"
                
                # 格式化剩余时间（秒）
                eta = d.get('eta', '未知')
                if isinstance(eta, (int, float)):
                    eta_str = str(int(eta))
                else:
                    eta_str = '未知'
                
                # 更新进度列表
                self.update_progress_list(progress, speed_mb, size_str, eta_str)
                
            except Exception as e:
                print(f"Error updating progress: {str(e)}")
                
        elif d['status'] == 'finished':
            self.update_progress_list(100, 0, "下载完成", "0")

    def retry_download(self, url, clear_progress=True):
        self.is_downloading = True
        self.download_button.configure(text="取消下载")
        
        if clear_progress:
            self.progress_list.delete(1.0, tk.END)
        
        self.download_thread = threading.Thread(target=self.download_video, args=(url,))
        self.download_thread.daemon = True
        self.download_thread.start()

    def start_download(self):
        if self.is_downloading:
            self.cancel_download()
        else:
            url = self.url_entry.get().strip()
            if not url:
                messagebox.showerror("错误", "请输入视频链接")
                return
            
            save_path = self.path_entry.get().strip()
            if not save_path:
                messagebox.showerror("错误", "请选择或输入保存位置")
                return
                
            # 显示进度信息框架并清空进度列表
            self.info_frame.pack(fill=tk.BOTH, expand=True)
            self.progress_list.delete(1.0, tk.END)
                
            self.is_downloading = True
            self.download_button.configure(text="取消下载")
            self.retry_count = 0
            self.total_retry_count = 0
            
            self.download_thread = threading.Thread(target=self.download_video, args=(url,))
            self.download_thread.daemon = True
            self.download_thread.start()

    def cancel_download(self):
        if self.is_downloading:
            if messagebox.askyesno("确认", "确定要取消当前下载吗？"):
                self.is_downloading = False
                if self.current_download:
                    self.current_download.params['quiet'] = True
                self.download_button.configure(text="开始下载")

    def browse_path(self):
        path = filedialog.askdirectory(initialdir=self.path_entry.get())
        if path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, path)

    def format_size(self, bytes):
        try:
            for unit in ['B', 'KB', 'MB', 'GB']:
                if bytes < 1024:
                    return f"{bytes:.2f} {unit}"
                bytes /= 1024
            return f"{bytes:.2f} TB"
        except Exception as e:
            print(f"Error formatting size: {str(e)}")
            return "未知"

    def open_download_directory(self):
        try:
            path = self.path_entry.get()
            if os.path.exists(path):
                if os.name == 'nt':  # Windows
                    os.startfile(path)
                elif sys.platform == 'darwin':  # macOS
                    subprocess.run(['open', path], check=False)
                else:  # Linux 等
                    subprocess.run(['xdg-open', path], check=False)
        except Exception as e:
            print(f"Error opening directory: {str(e)}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = YouTubeDownloader(root)
        root.mainloop()
    except Exception as e:
        print(f"Application error: {str(e)}") 