# ruff: noqa: F401,F403,F405,E402,F541,E722
import asyncio
import websockets
import json
import os
import ctypes
import sys
import time
import winreg
import tkinter as tk
from tkinter import ttk
import customtkinter
import subprocess
import random
import threading
import atexit
import psutil
import queue
from dataclasses import dataclass
from queue import PriorityQueue
from typing import List
from flask import Flask, jsonify
from flask_cors import CORS
from threading import Thread
# 创建一个队列
gui_queue = queue.Queue()

# 创建一个线程安全的标志
stop_flag = threading.Event()


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def disable_proxy():
    internet_settings = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        0,
        winreg.KEY_ALL_ACCESS,
    )

    # 设置代理使能为0（禁用）
    winreg.SetValueEx(internet_settings, "ProxyEnable", 0, winreg.REG_DWORD, 0)
    # 清空代理服务器信息
    winreg.SetValueEx(internet_settings, "ProxyServer", 0, winreg.REG_SZ, "")

    winreg.CloseKey(internet_settings)
    print("Proxy disabled.")


def initGame():
    print("Init game.")
    pass


@dataclass
class UserBean:
    user_id: str
    nickname: str
    gifts: int = 0
    likes: int = 0
    messages: List[str] = None
    last_time: str = ""
    
    def __init__(self, user_id, nickname):
        self.user_id = user_id
        self.nickname = nickname
        self.gifts = 0
        self.likes = 0
        self.messages = []
        self.last_time = time.strftime("%H:%M:%S")
    
    def __lt__(self, other):
        # 优先级比较：先按礼���数，再按点赞数
        if self.gifts != other.gifts:
            return self.gifts > other.gifts
        return self.likes > other.likes


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("无人信号v2")
        self.geometry("800x600")
        
        # 左侧统计信息面板
        self.stats_frame = ttk.Frame(self)
        self.stats_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # 添加所有统计标签
        self.current_viewers_label = ttk.Label(self.stats_frame, text="在线人数: 0 人")
        self.current_viewers_label.grid(row=0, column=0, sticky="w")
        
        self.total_viewers_label = ttk.Label(self.stats_frame, text="总入场统计人数: 0 人")
        self.total_viewers_label.grid(row=1, column=0, sticky="w")
        
        self.gender_ratio_label = ttk.Label(self.stats_frame, text="男女比例: 0:0")
        self.gender_ratio_label.grid(row=2, column=0, sticky="w")
        
        self.total_likes_label = ttk.Label(self.stats_frame, text="总点赞数: 0")
        self.total_likes_label.grid(row=3, column=0, sticky="w")
        
        self.new_followers_label = ttk.Label(self.stats_frame, text="收获粉丝人数: 0 人")
        self.new_followers_label.grid(row=4, column=0, sticky="w")
        
        self.total_gifts_label = ttk.Label(self.stats_frame, text="总礼物收入: 0")
        self.total_gifts_label.grid(row=5, column=0, sticky="w")

        # 用户列表
        self.LogFrame = LogFrame(self)
        self.LogFrame.grid(row=1, column=0, sticky="nsew")

    def update_stats(self, stats_data):
        """更新统计数据"""
        if isinstance(stats_data, dict):
            self.current_viewers_label.config(text=f"在线人数: {stats_data.get('OnlineUserCount', 0)} 人")
            self.total_viewers_label.config(text=f"总入场统计人数: {stats_data.get('TotalUserCount', 0)} 人")
            # 其他统计数据更新...


class LogFrame(customtkinter.CTkFrame):
    def __init__(self, master):
        super().__init__(master)
        
        # 初始化Flask服务器
        self.app = Flask(__name__)
        CORS(self.app)
        self.init_flask_routes()
        self.start_flask_server()
        
        # 初始化用户队列和映射
        self.user_queue = PriorityQueue()
        self.user_map = {}  # user_id -> UserBean映射
        self.current_sort_key = "gifts"  # 默认排序键
        
        # 创建主框架
        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # 创建排序按钮框架，放在表格上方
        self.sort_frame = ttk.Frame(self.main_frame)
        self.sort_frame.grid(row=0, column=0, sticky="nw", padx=5, pady=5)
        
        # 添加排序按钮
        self.gift_sort_btn = ttk.Button(
            self.sort_frame, 
            text="按礼物排序", 
            command=lambda: self.sort_users("gifts")
        )
        self.gift_sort_btn.grid(row=0, column=0, padx=5)
        
        self.like_sort_btn = ttk.Button(
            self.sort_frame, 
            text="按点赞排序", 
            command=lambda: self.sort_users("likes")
        )
        self.like_sort_btn.grid(row=0, column=1, padx=5)

        # 创建表格框架
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree_frame.grid(row=1, column=0, sticky="nsew", pady=(5,0))  # 添加上边距

        # 设置列表列
        self.treeview = ttk.Treeview(
            self.tree_frame,
            columns=("礼物数", "点赞数", "观众昵称", "消息", "最新时间", "操作"),
            show="headings"
        )
        
        # 配置列宽和标题
        self.treeview.column("礼物数", width=50, anchor="center", minwidth=50)
        self.treeview.column("点赞数", width=50, anchor="center", minwidth=50)
        self.treeview.column("观众昵称", width=150, anchor="w", minwidth=100)
        self.treeview.column("消息", width=300, anchor="w", minwidth=200)
        self.treeview.column("最新时间", width=100, anchor="center", minwidth=100)
        self.treeview.column("操作", width=50, anchor="center", minwidth=50)

        # 设置表头
        self.treeview.heading("礼物数", text="礼物数")
        self.treeview.heading("点赞数", text="点赞数")
        self.treeview.heading("观众昵称", text="观众昵称")
        self.treeview.heading("消息", text="消息")
        self.treeview.heading("最新时间", text="最新时间")
        self.treeview.heading("操作", text="操作")

        # 设置表格样式
        style = ttk.Style()
        style.configure(
            "Treeview", 
            rowheight=60,  # 设置基础行高
            font=('微软雅黑', 9),  # 设置内容字体
            background="#ffffff",
            fieldbackground="#ffffff"
        )
        style.configure(
            "Treeview.Heading",
            font=('微软雅黑', 9, 'bold'),  # 设置表头字体
            padding=(5, 2)  # 设置表头padding
        )

        # 绑定事件处理函数
        self.treeview.bind('<Motion>', self.handle_motion)  # 鼠标移动事件
        self.treeview.tag_configure('hover', background='#f0f0f0')  # 悬停效果

        # 添加垂直滚动条
        self.scrollbar_y = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.treeview.yview)
        self.scrollbar_y.grid(row=0, column=1, sticky="ns")
        
        # 添加水平滚动条
        self.scrollbar_x = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.treeview.xview)
        self.scrollbar_x.grid(row=1, column=0, sticky="ew")
        
        # 配置滚动条
        self.treeview.configure(
            yscrollcommand=self.scrollbar_y.set,
            xscrollcommand=self.scrollbar_x.set
        )
        
        # 调整表格位置以适应两个滚动条
        self.treeview.grid(row=0, column=0, sticky="nsew")

        # 配置grid权重，使表格可以自动扩展
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # 设置按钮样式
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")

        # 添加选择和复制功能的绑定
        self.treeview.bind('<Button-1>', self.handle_click)
        self.treeview.bind('<Control-c>', self.copy_cell_content)
        
        # 用于跟踪当前选中的单元格
        self.selected_item = None
        self.selected_column = None
        
        # 创建右键菜单
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="复制", command=self.copy_cell_content)
        self.treeview.bind('<Button-3>', self.show_context_menu)  # 右键菜单
        
        # 添加选中单元格的样式
        style = ttk.Style()
        style.configure("Selected.Treeview.Cell", background="#e0e0ff")

    def init_flask_routes(self):
        """初始化Flask路由"""
        
        @self.app.route('/api/next_user', methods=['GET'])
        def get_next_user():
            """获取下一个待处理用户"""
            try:
                if not self.user_queue.empty():
                    user = self.user_queue.get_nowait()
                    return jsonify({
                        'success': True,
                        'user': {
                            'user_id': user.user_id,
                            'nickname': user.nickname
                        }
                    })
                return jsonify({
                    'success': True,
                    'user': None
                })
            except Exception as e:
                return jsonify({
                    'success': False,
                    'error': str(e)
                })
        
        @self.app.route('/api/complete/<user_id>', methods=['POST'])
        def complete_processing(user_id):
            """标记用户处理完成"""
            try:
                if user_id in self.user_map:
                    del self.user_map[user_id]
                    self.refresh_user_list()
                return jsonify({'success': True})
            except Exception as e:
                return jsonify({
                    'success': False,
                    'error': str(e)
                })
    
    def start_flask_server(self):
        """在新线程中启动Flask服务器"""
        def run_flask():
            self.app.run(port=5000, threaded=True)
        
        Thread(target=run_flask, daemon=True).start()

    def handle_motion(self, event):
        """处理鼠标悬停效果"""
        item = self.treeview.identify_row(event.y)
        if item:
            # 清除所有项的悬停效果
            for child in self.treeview.get_children():
                if child != item:
                    self.treeview.item(child, tags='')
            # 为当前项添加悬停效果
            self.treeview.item(item, tags=('hover',))

    def update_user_data(self, user_id, msg_type, msg_data):
        """更新用户数据"""
        # 获取或创建UserBean
        if user_id not in self.user_map:
            user = msg_data.get("User", {})
            user_bean = UserBean(
                user_id=user_id,
                nickname=user.get("Nickname", "")
            )
            self.user_map[user_id] = user_bean
        else:
            user_bean = self.user_map[user_id]
        
        # 更新用户数据
        if msg_type == 2:  # 点赞消息
            count = msg_data.get("Count", 1)
            user_bean.likes += count
        elif msg_type == 5:  # 礼物消���
            count = msg_data.get("GiftCount", 1)
            user_bean.gifts += count
        elif msg_type == 1:  # 弹幕消息
            content = msg_data.get("Content", "")
            user_bean.messages.append(content)  # 只添加原始消息内容
        
        # 使用消息中的时间戳
        user_bean.last_time = msg_data.get("Time", "")  # 从WebSocket消息获取时间
        
        # 限制消息历史记录数量
        if len(user_bean.messages) > 3:
            user_bean.messages = user_bean.messages[-3:]  # 只保留最近3条消息
        
        # 刷新显示
        self.refresh_user_list()

    def refresh_user_list(self):
        """刷新用户列表显示"""
        # 保存当前选中状态
        current_selection = (self.selected_item, self.selected_column)
        
        # 清空现有列表
        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        # 创建临时优先级队列进行排序
        temp_queue = PriorityQueue()
        sorted_users = []
        
        # 将所有用户放入临时队列
        for user_bean in self.user_map.values():
            temp_queue.put(user_bean)
        
        # 按优先级取出用户
        while not temp_queue.empty():
            user_bean = temp_queue.get()
            sorted_users.append(user_bean)
        
        # 更新显示
        for user_bean in sorted_users:
            # 格式化消息显示
            messages = user_bean.messages[-3:]  # 获取最近3条消息
            formatted_messages = "\n".join(msg for msg in messages if msg)  # 使用换行符连接非空消息
            
            item = self.treeview.insert("", "end", values=(
                user_bean.gifts,
                user_bean.likes,
                user_bean.nickname,
                formatted_messages,
                user_bean.last_time,
                "删除"
            ))
            
            # 根据消息行数动态调整行高
            msg_lines = len(messages)
            if msg_lines > 1:
                # 计算所需的行高 (每行20像素 + padding)
                required_height = (msg_lines * 20) + 10
                self.treeview.item(item, tags=(f'height{required_height}',))
                
                # 创建自定义标签样式
                style = ttk.Style()
                style.configure(
                    f'Treeview{required_height}', 
                    rowheight=required_height
                )

        # 如果之前有选中的单元格，尝试恢复选中状态
        if all(current_selection):
            item, column = current_selection
            if item in self.treeview.get_children():
                self.selected_item = item
                self.selected_column = column
                self.treeview.tag_configure(f'selected_{item}_{column}', 
                                          background="#e0e0ff")
                self.treeview.item(item, tags=(f'selected_{item}_{column}',))

    def sort_users(self, sort_by):
        """更改排序方式"""
        self.current_sort_key = sort_by
        # 更新UserBean的比较方法
        UserBean.__lt__ = lambda self, other: (
            getattr(self, sort_by) > getattr(other, sort_by)
        )
        self.refresh_user_list()

    def sort_by_column(self, column):
        """根据列标题点击排序"""
        sort_map = {
            "礼物数": "gifts",
            "点赞数": "likes"
        }
        self.sort_users(sort_map.get(column, "gifts"))

    def display_message(self, message):
        """处理接收到的消息"""
        try:
            data = json.loads(message)
            msg_type = data.get("Type")
            msg_data = json.loads(data.get("Data", "{}"))
            
            if msg_type == 6:  # 统计消息
                self.master.update_stats(msg_data)
            
            elif msg_type in [1, 2, 5]:  # 弹幕、点赞、礼物消息
                user = msg_data.get("User", {})
                user_id = user.get("Id")
                if user_id:
                    self.update_user_data(user_id, msg_type, msg_data)
                    
                # 更新总统计数据
                total_stats = {
                    "OnlineUserCount": msg_data.get("CurrentCount", 0),
                    "TotalLikes": msg_data.get("Total", 0),
                }
                self.master.update_stats(total_stats)
                
        except json.JSONDecodeError:
            print(f"Invalid JSON message: {message}")
        except Exception as e:
            print(f"Error processing message: {e}")

    def handle_click(self, event):
        """处理单击事件，选中单元格"""
        # 获取点击的区域信息
        region = self.treeview.identify("region", event.x, event.y)
        if region == "cell":
            # 获取点击的单元格信息
            column = self.treeview.identify_column(event.x)
            item = self.treeview.identify_row(event.y)
            
            # 清除之前的选中状态
            if self.selected_item and self.selected_column:
                self.treeview.tag_configure(f'selected_{self.selected_item}_{self.selected_column}', 
                                          background="")
            
            # 设置新的选中状态
            self.selected_item = item
            self.selected_column = column
            
            # 高亮显示选中的单元格
            self.treeview.tag_configure(f'selected_{item}_{column}', 
                                      background="#e0e0ff")
            self.treeview.item(item, tags=(f'selected_{item}_{column}',))

    def show_context_menu(self, event):
        """显示右键菜单"""
        if self.selected_item and self.selected_column:
            self.context_menu.post(event.x_root, event.y_root)

    def copy_cell_content(self, event=None):
        """复制选中单元格的内容"""
        if self.selected_item and self.selected_column:
            # 获取选中单元格的值
            values = self.treeview.item(self.selected_item)['values']
            if values:
                # 获取列索引（去掉'#'��缀并转换为整数）
                col_idx = int(self.selected_column.replace('#', '')) - 1
                if 0 <= col_idx < len(values):
                    content = str(values[col_idx])
                    # 复制到剪贴板
                    self.clipboard_clear()
                    self.clipboard_append(content)
                    self.update()  # 刷新剪贴板


# async def send_ping(websocket):
#     while True:
#         try:
#             if websocket.open:
#                 await websocket.ping()
#                 await asyncio.sleep(5)
#             else:
#                 break
#         except websockets.exceptions.ConnectionClosedOK:
#             # 连接已关闭，尝试重新连接
#             websocket = await websockets.connect("ws://127.0.0.1:8888")


async def receive_messages():
    while True:
        try:
            async with websockets.connect(
                "ws://127.0.0.1:8888", ping_timeout=None, ping_interval=None
            ) as websocket:
                while True:
                    message = await websocket.recv()
                    if message is not None:
                        try:
                            json.loads(message)
                            app.LogFrame.display_message(message)
                        except json.JSONDecodeError:
                            print(
                                f"Received a message that could not be parsed as JSON: {message}"
                            )
        except (
            websockets.ConnectionClosed,
            websockets.exceptions.ConnectionClosedError,
            websockets.exceptions.ConnectionClosedOK,
        ) as e:
            print(f"Connection closed, retrying...{e}")
            continue
        except Exception as e:
            print(f"An error occurred on ws server: {e}")
            continue


def main():
    global listener_process
    global app
    # Check if the process is already running
    for proc in psutil.process_iter():
        if proc.name() == "WssBarrageServer.exe":
            print("The process is already running.")
            proc.kill()
            time.sleep(1)

    # Start a new process
    # 将WssBarrageServer.exe放至同目录下dy-barrage-grab文件夹
    # 获取当前脚本的绝对路径
    script_dir = os.path.dirname(os.path.realpath(__file__))
    # 构建WssBarrageServer.exe的绝对路径
    exe_path = os.path.join(script_dir, "dy-barrage-grab", "WssBarrageServer.exe")

    # 使用绝对路径启动进程
    listener_process = subprocess.Popen(exe_path)
    # 创建一个新的事件循环
    loop = asyncio.new_event_loop()
    # 设置这个事件循环为当前线程的事件循环
    asyncio.set_event_loop(loop)
    app = App()

    # 点击关闭按钮时调用的函数
    def on_close():
        print("Closing...")
        # 设置停止标志
        stop_flag.set()
        # 关闭WssBarrageServer
        listener_process.terminate()
        # 关闭代理
        disable_proxy()
        # 退出程序
        os._exit(0)

    # 设置关闭按钮的回调函数
    app.protocol("WM_DELETE_WINDOW", on_close)
    # 在事件循环中创建新的异步任务
    loop.create_task(receive_messages())

    def update():
        # 更新Tkinter的界面
        app.update()
        # 在事件循环中定期调用update方法
        loop.call_soon(update)

        # 检查队列中是否有任务
        while not gui_queue.empty():
            # 获取任务并执行
            task = gui_queue.get()
            task()

    # 在事件循环中定期调用update方法
    loop.call_soon(update)
    try:
        # 运行事件循环
        loop.run_forever()
    except Exception as e:
        print(f"An error occurred in loop: {e}")
    finally:
        # 关闭事件循环
        loop.close()


if __name__ == "__main__":
    if is_admin():
        # 如果已经是管理员权限，那么直接运行��的代码
        print("Running as administrator.")
        pass
    else:
        # 如果不是管理员权限，那么以管理员权限重新启动程序
        print("Running as non-administrator.")
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
    # try:
    main()
# except Exception as e:
# print(f"An error occurred: {e}")
# finally:
#     # 这里结束进程
#     listener_process.terminate()


# 注册一个函数，在程序结束时调用
def on_exit():
    listener_process.terminate()


atexit.register(on_exit)
