import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import threading
import time
from datetime import datetime
import traceback

# ==========================================
# 动态导入模块
# 确保父目录在路径中以便导入
# ==========================================
# 这一行确保了 'zhongzi' 和 'qijian' 文件夹可以被正确导入
# 注意：在实际运行环境中，您可能需要调整 sys.path.append 的路径
# sys.path.append(os.path.dirname(os.path.abspath(__file__))) 

try:
    # 导入所有子模块的 GUI 类
    from path_a.Rin_FSV3004 import RinGUI as Rin_FSV3004_GUI
    from path_a.Rin_4051 import Rin_4051_GUI
    from path_a.LineWidth_FSV3004 import LineWidthGUI

    from path_a.TimeDomain import TimeDomainGUI
    from path_a.SpectrumSNR import SpectrumSNRGUI
    from path_a.SingleFrequency import SingleFrequencyGUI
    from qijian.CT_W import CT_W_GUI
    from qijian.CT_P import CT_P_GUI
    from qijian.CT_L import CT_L_GUI
except ImportError as e:
    # 如果导入失败，会给出提示，但不终止程序
    print(f"模块导入错误: {e}")
    print("请检查目录结构是否包含 'zhongzi' 和 'qijian' 文件夹，且包含正确的脚本。")

# ==========================================
# 配置定义
# ==========================================
# 定义模块映射：名称 -> (类, 默认启动方法名, 所属分组)
MODULE_MAP = {
    "Rin_FSV3004": {"class": Rin_FSV3004_GUI, "start_method": "start_rin", "group": "zhongzi"},
    "Rin_4051": {"class": Rin_4051_GUI, "start_method": "start_test", "group": "zhongzi"},
    "线宽": {"class": LineWidthGUI, "start_method": "start_measurement", "group": "zhongzi"},
    "时域": {"class": TimeDomainGUI, "start_method": "start_test", "group": "zhongzi"},
    "信噪比": {"class": SpectrumSNRGUI, "start_method": "start_test", "group": "zhongzi"},
    "单频": {"class": SingleFrequencyGUI, "start_method": "start", "group": "zhongzi"},
    "CT-波长": {"class": CT_W_GUI, "start_method": "start_group1", "group": "qijian"},
    "CT-功率": {"class": CT_P_GUI, "start_method": "start_group1", "group": "qijian"},
    "CT-线宽": {"class": CT_L_GUI, "start_method": "start_group1", "group": "qijian"},
}

# 按分组整理模块
MODULE_GROUPS = {
    "种子": [name for name, info in MODULE_MAP.items() if info["group"] == "zhongzi"],
    "器件": [name for name, info in MODULE_MAP.items() if info["group"] == "qijian"],
}

CONFIG_FILE = "integration_config.json"

# ==========================================
# 集成平台主类
# ==========================================
class IntegratedPlatform:
    def __init__(self, root):
        self.root = root
        self.root.title("PTS")
        self.root.geometry("2000x1240")
        try:
            self.root.iconbitmap("PreciLasers.ico")
        except:
            pass

        # 状态变量和映射
        self.active_modules = {} # 存储 {name: gui_instance}
        self.check_vars = {}     # 存储 {name: BooleanVar}
        self.name_to_tab_id = {} # 存储 {name: tab_frame_widget}
        
        self.saved_params = self.load_config()

        self.setup_ui()
        
        # 绑定 Notebook 的页签关闭事件
        self.notebook.bind("<<NotebookTabClosed>>", self.on_tab_closed)
        
        # 绑定关闭事件以保存参数
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # 绑定快捷键 (全选: Ctrl+A, 取消全选: Ctrl+D)
        self.root.bind('<Control-a>', lambda event: self.select_all())
        self.root.bind('<Control-d>', lambda event: self.deselect_all())

    def setup_ui(self):
        # 样式设置 (提升用户体验)
        self.style = ttk.Style()
        self.style.theme_use('vista')
        # 调整 notebook tab 样式以实现流畅切换效果
        self.style.configure("TNotebook.Tab", padding=[10, 5], font=("Microsoft YaHei", 10))
        self.style.map("TNotebook.Tab", background=[("selected", "#c0c0c0")])
        # 自定义 Checkbutton 样式（尝试设置背景色/前景色）
        # 注意：不同主题对 background 的支持不同，若无效可改用 tk.Checkbutton
        try:
            self.style.configure("Custom.TCheckbutton", background="#ffffff", foreground="#333333")
            self.style.map("Custom.TCheckbutton",
                           background=[('active', '#ffffff'), ('!active', '#ffffff')],
                           foreground=[('disabled', '#a3a3a3'), ('!disabled', '#333333')])
        except Exception:
            pass
        
        # 主分割窗格 (左右布局)
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)

        # === 左侧：测试项选择区域 (Left Panel) ===
        # width=250 是初始宽度
        self.left_panel = tk.Frame(self.paned_window, bg="#ffffff", width=380)
        
        # 【关键修改】禁止 Frame 根据内部子控件自动调整大小
        # 这样即使内部控件内容很少或很多，Frame 都会保持设定的 width=250
        self.left_panel.pack_propagate(False) 
        self.left_panel.grid_propagate(False)

        # weight=0: 窗口拉伸时不分配额外空间给左侧
        self.paned_window.add(self.left_panel, weight=0)

        # 标题
        tk.Label(self.left_panel, text="测试项目", bg="#ffffff", 
                 font=("Microsoft YaHei", 14, "bold"), fg="#333").pack(pady=10, padx=10, anchor="center")

        # 全选/反选 (快捷键支持)
        btn_frame = tk.Frame(self.left_panel, bg="#ffffff")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(btn_frame, text="全选 (Ctrl+A)", command=self.select_all, width=12).pack(side=tk.LEFT, padx=1)
        ttk.Button(btn_frame, text="清空 (Ctrl+D)", command=self.deselect_all, width=12).pack(side=tk.RIGHT, padx=1)

        # 核心修改：使用 Notebook 实现“种子”和“器件”页签
        self.module_notebook = ttk.Notebook(self.left_panel)
        self.module_notebook.pack(fill=tk.BOTH, expand=True, padx=10)

        for group_name, module_list in MODULE_GROUPS.items():
            # 为每个分组创建一个 Frame 作为页签内容
            group_frame = ttk.Frame(self.module_notebook)
            self.module_notebook.add(group_frame, text=f" {group_name} ") # 增加空格美化
            
            # 使用 Scrollable Frame 包含勾选框
            canvas = tk.Canvas(group_frame, bg="#ffffff", highlightthickness=0)
            scrollbar = ttk.Scrollbar(group_frame, orient="vertical", command=canvas.yview)
            check_frame = tk.Frame(canvas, bg="#ffffff") # 内部 Frame

            canvas.create_window((0, 0), window=check_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            scrollbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            
            check_frame.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion = c.bbox("all")))

            # 动态生成勾选框
            for name in module_list:
                var = tk.BooleanVar()
                self.check_vars[name] = var
                
                # 使用 row_frame 保证对齐
                row_frame = tk.Frame(check_frame, bg="#ffffff")
                row_frame.pack(anchor="w", pady=1, fill=tk.X)
                
                # 测试项勾选框，使用修改后的样式
                cb = ttk.Checkbutton(row_frame, text=name, variable=var,
                                     command=lambda n=name: self.toggle_module(n),
                                     style="Custom.TCheckbutton")
                cb.pack(side=tk.LEFT, anchor="w")

        # 底部控制区
        ctrl_frame = tk.Frame(self.left_panel, bg="#ffffff", bd=1, relief=tk.RAISED)
        ctrl_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # 一键测试按钮 (醒目)
        self.btn_run_all = tk.Button(ctrl_frame, text="▶ 一键测试", 
                                     bg="#13A80B", fg="white", activebackground="#45a049", activeforeground="white",
                                     font=("Microsoft YaHei", 12, "bold"),
                                     command=self.run_selected_tests)
        self.btn_run_all.pack(pady=15, padx=10, fill=tk.X)

        # 测试进度的可视化展示
        tk.Label(ctrl_frame, text="总测试进度:", bg="#ffffff", font=("Microsoft YaHei", 9)).pack(fill=tk.X, padx=10)
        self.progress = ttk.Progressbar(ctrl_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.progress_label = tk.Label(ctrl_frame, text="未执行 (0/0)", bg="#ffffff", font=("Microsoft YaHei", 9))
        self.progress_label.pack(fill=tk.X, padx=10, pady=(0, 5))

        # === 右侧：测试内容显示区域 (Right Panel - Notebook) ===
        self.right_panel = tk.Frame(self.paned_window, bg="white")
        self.paned_window.add(self.right_panel, weight=1)  # weight=1 表示右侧自动伸缩填充
        
        # 右侧采用页签式设计
        self.notebook = ttk.Notebook(self.right_panel)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 欢迎页
        welcome_frame = ttk.Frame(self.notebook, style='TFrame')

        # 欢迎页：以 Markdown 格式展示基础使用说明（只读）
        welcome_md = """
        频准测试系统 (PTS)

        欢迎使用一体化测试系统。下面提供一些基础说明，帮助你在实际连接仪器并运行测试前完成必要准备。

        一、连接配置

            1.已在主机上安装并配置好 VISA 后端（搜索框输入"NI"，出现"NI MAX"，则配置成功）。
            2.仪器开启远程控制功能，有些仪器需设置控制方式，如YOKOGAWA光谱仪需设置为NET(VXI-11)。
            3.配置好仪器IP；主机IP地址设为静态IP，且与仪器处于同一网段。
                主机：IP地址-192.168.7.7，子网掩码-255.255.255.0，网关-192.168.7.1，首选DNS-1.1.1.1。
                仪器：IP地址-对应程序默认地址，其余同上。
                PS：主机若控制两台仪器，第二个IP地址设置为192.168.7.8，其余同上。
            4.将主机与仪器通过网线连接。

        二、使用方式

            1. 网盘 "\\\\192.168.110.5\\\\信息部\\\\PTS\\\\集成软件" 中可找到最新软件，复制到本地即可。
            2. 在左侧“测试项目”里勾选需要的模块（或使用“全选/清空”）。
            3. 勾选后对应模块页签会出现在右侧，打开页签进行参数设置。
            4. 点击模块内的“开始测试”或在左侧使用“▶ 一键测试”启动所有选中项。
            5. 测试运行过程中请查看各模块页签内的运行日志与左侧下方的进度条

        三、输出与保存

            1.测试数据默认保存到模块配置中指定的输出目录（可以在模块参数中修改）。
            2.程序会保存 CSV/DAT 等格式的数据文件，并生成可视化图片供保存。

        四、常见故障与排查

            1.无法连接仪器：检查 IP 是否可达（ping）、VISA 是否安装、仪器远程控制方式是否正确。
            2.二进制读取失败：程序会回退到 ASCII 读取并在日志中提示，若频繁失败请检查仪器固件和命令兼容性。
            3.GUI 无响应：可能是长时间测量或阻塞的查询，可尝试停止后重新连接。

        如需进一步帮助，请联系开发人员（张珂）。
        """

        # 使用只读 Text 控件显示 Markdown 文本（保留原始 Markdown 格式）
        txt_frame = tk.Frame(welcome_frame, padx=10, pady=10)
        txt_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(txt_frame, orient=tk.VERTICAL)
        text_widget = tk.Text(txt_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, bg="#ffffff",
                      font=("Microsoft YaHei", 11), relief=tk.FLAT)
        scrollbar.config(command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        text_widget.insert(tk.END, welcome_md)
        text_widget.configure(state=tk.DISABLED)

        self.notebook.add(welcome_frame, text="🏠 首页")

    # ================= 核心逻辑：添加/移除页签 =================

    def toggle_module(self, name):
        """勾选框回调：添加或移除页签"""
        is_checked = self.check_vars[name].get()
        
        if is_checked:
            if name not in self.active_modules:
                self.add_tab(name)
        else:
            if name in self.active_modules:
                # 找到页签，并关闭 (这将触发 on_tab_closed)
                if name in self.name_to_tab_id:
                    # 获取页签索引
                    tab_widget = self.name_to_tab_id[name]
                    tab_index = self.notebook.index(tab_widget)
                    self.notebook.forget(tab_index)
                    self.remove_tab(name)

    def add_tab(self, name):
        """实例化模块GUI并添加到Notebook"""
        try:
            module_info = MODULE_MAP[name]
            GuiClass = module_info["class"]
            
            # 创建页签容器 (用于嵌入子程序)
            tab_frame = ttk.Frame(self.notebook, padding=5)
            
            # 实例化GUI，传入tab_frame作为parent
            gui_instance = GuiClass(parent=tab_frame)
            
            # 添加页签
            self.notebook.add(tab_frame, text=name, sticky="nsew")
            
            self.active_modules[name] = gui_instance
            self.name_to_tab_id[name] = tab_frame
            self.notebook.select(tab_frame)
            
            # 尝试恢复保存的参数
            self.restore_module_params(name, gui_instance)
            
        except Exception as e:
            msg = f"无法加载模块 {name}，请检查该文件是否已按要求修改：\n{str(e)}\n{traceback.format_exc()}"
            self.log("加载错误", msg)
            self.check_vars[name].set(False) # 加载失败则取消勾选

    def remove_tab(self, name):
        """清理模块实例和状态"""
        if name in self.active_modules:
            # 1. 保存当前参数
            self.save_module_params(name, self.active_modules[name])
            
            gui_instance = self.active_modules[name]
            
            # 2. 尝试调用关闭/清理方法 (如停止线程)
            if hasattr(gui_instance, "stop") and callable(gui_instance.stop):
                 try:
                    gui_instance.stop()
                 except:
                    pass

            # 3. 删除引用
            del self.active_modules[name]
            if name in self.name_to_tab_id:
                del self.name_to_tab_id[name]
        
            # 4. 更新进度
            self.update_overall_progress()

    def on_tab_closed(self, event):
        """Notebook页签关闭操作，自动取消勾选并移除实例"""
        try:
            # 获取被关闭页签的 widget id
            selected_tab_id = self.notebook.select()
            closed_tab_text = self.notebook.tab(selected_tab_id, "text")
        except:
             # 如果是最后一个 tab 被关了，会找不到 select()
             return

        # 遍历找到被关闭的页签名称
        module_name = None
        for name, tab_id in self.name_to_tab_id.items():
            if tab_id == self.notebook.nametowidget(selected_tab_id):
                 module_name = name
                 break
        
        if module_name:
            # 移除模块 (保存参数，停止线程等)
            self.remove_tab(module_name)
            
            # 自动取消左侧勾选
            if module_name in self.check_vars:
                 self.check_vars[module_name].set(False)


    # ================= 核心功能：运行控制 =================

    def run_selected_tests(self):
        """一键运行所有选中的测试"""
        selected = [name for name, var in self.check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("提示", "请先勾选至少一个测试项")
            return

        # 启动进度指示
        self.progress.config(mode='indeterminate')
        self.progress.start(15)
        self.btn_run_all.config(state="disabled", text="测试启动中...")
        
        # 使用线程启动，防止界面卡死
        threading.Thread(target=self._execute_tests, args=(selected,), daemon=True).start()

    def _execute_tests(self, selected_names):
        """后台执行逻辑：按顺序发送启动命令"""
        
        total_tests = len(selected_names)
        completed_count = 0
        
        for name in selected_names:
            self.update_overall_progress(current=completed_count, total=total_tests, text=f"正在启动: {name}")

            if name in self.active_modules:
                instance = self.active_modules[name]
                method_name = MODULE_MAP[name]["start_method"]
                
                # 尝试调用启动方法
                if hasattr(instance, method_name) and callable(getattr(instance, method_name)):
                    try:
                        method = getattr(instance, method_name)
                        # 在UI线程中调用，防止非线程安全的GUI操作报错
                        self.root.after(0, method)
                        # TODO: 实际的测试状态更新需要依赖子模块的日志反馈或状态变量
                    except Exception as e:
                        print(f"[{name}] 启动失败: {e}")
                else:
                    print(f"[{name}] 未找到启动方法 {method_name}")
            
            completed_count += 1
            # 简单的间隔，防止瞬间并发导致VISA资源冲突
            time.sleep(1) 

        # 启动完成后，切换到确定模式，显示总进度 (例如，依赖于所有模块完成)
        self.update_overall_progress(current=total_tests, total=total_tests, text="所有任务已启动")
        
        # 恢复按钮状态
        self.root.after(1000, self._reset_run_button)

    def _reset_run_button(self):
        self.progress.config(mode='determinate') # 切换到确定模式 (等待所有完成)
        self.progress.stop() # 停止不确定模式动画
        self.btn_run_all.config(state="normal", text="▶ 一键测试")
        # messagebox.showinfo("完成", "所有选中测试的启动命令已发送。\n请查看各页签日志确认运行状态。")
        
    def update_overall_progress(self, current=None, total=None, text=None):
        """更新总进度条和标签"""
        selected = [name for name, var in self.check_vars.items() if var.get()]
        active_count = len(selected)
        
        if current is None and total is None:
            # 仅刷新 label
            self.progress_label.config(text=f"已选中 {active_count} 个任务")
        else:
            # 更新进度条
            if total > 0:
                percent = int(current / total * 100)
                self.progress['value'] = percent
                self.progress_label.config(text=f"{text} ({current}/{total}, {percent}%)")
            else:
                self.progress_label.config(text="未执行 (0/0)")


    # ================= 快捷操作 =================
    
    def select_all(self):
        # 只选择当前左侧 module_notebook 激活的分组（例如 '种子' 或 '器件'）
        try:
            sel = self.module_notebook.select()
            if not sel:
                raise Exception("no selection")
            tab_text = self.module_notebook.tab(sel, "text")
            # 页面创建时为 text=f" {group_name} "，去掉空白并匹配
            group_name = tab_text.strip()
        except Exception:
            # 回退：如果无法确定当前页签，则选择所有模块（兼容旧行为）
            group_name = None

        if group_name and group_name in MODULE_GROUPS:
            target_list = MODULE_GROUPS[group_name]
        else:
            target_list = list(MODULE_MAP.keys())

        for name in target_list:
            # 如果 check_vars 中没有该 name（理论上不应发生），先创建变量
            if name not in self.check_vars:
                self.check_vars[name] = tk.BooleanVar(value=False)

            self.check_vars[name].set(True)
            if name not in self.active_modules:
                self.add_tab(name)

        self.update_overall_progress()

    def deselect_all(self):
        # 仅取消当前左侧 module_notebook 选中页签下的测试项
        try:
            sel = self.module_notebook.select()
            if not sel:
                raise Exception("no selection")
            tab_text = self.module_notebook.tab(sel, "text")
            group_name = tab_text.strip()
        except Exception:
            group_name = None

        if group_name and group_name in MODULE_GROUPS:
            target_list = MODULE_GROUPS[group_name]
        else:
            # 回退为清空所有已激活模块（兼容旧行为）
            target_list = list(self.active_modules.keys())

        # 遍历目标列表并关闭对应页签与实例
        for name in list(target_list):
            if name in self.name_to_tab_id:
                try:
                    tab_widget = self.name_to_tab_id[name]
                    tab_index = self.notebook.index(tab_widget)
                    self.notebook.forget(tab_index)
                except Exception:
                    pass

            if name in self.check_vars:
                self.check_vars[name].set(False)

            # remove_tab 会安全地保存参数并删除实例引用
            if name in self.active_modules:
                self.remove_tab(name)

        self.update_overall_progress()


    # ================= 数据持久化 =================
    
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置失败: {e}. 使用空配置。")
                return {}
        return {}

    def save_config(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.saved_params, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"保存配置失败: {e}")

    def save_module_params(self, name, instance):
        """尝试从GUI实例中提取参数并保存"""
        params = {}
        # 策略1: 检查是否有 get_params 方法 (CT_系列脚本有)
        if hasattr(instance, "get_params") and callable(instance.get_params):
            try:
                params = instance.get_params()
            except:
                pass
        # 策略2: 检查是否有 entries 字典 (通用)
        elif hasattr(instance, "entries") and isinstance(instance.entries, dict):
            for k, entry in instance.entries.items():
                try:
                    params[k] = entry.get()
                except:
                    pass
        # 策略3: 检查 params 字典 (Rin, LineWidth系列有)
        elif hasattr(instance, "params") and isinstance(instance.params, dict):
             # 仅保存可序列化的简单值
             for k, v in instance.params.items():
                 if isinstance(v, (str, int, float, bool)):
                     params[k] = v
        
        if params:
            self.saved_params[name] = params

    def restore_module_params(self, name, instance):
        """将保存的参数回填到GUI"""
        if name not in self.saved_params:
            return

        params = self.saved_params[name]
        
        # 优先回填到 entries 字典
        if hasattr(instance, "entries") and isinstance(instance.entries, dict):
            for k, val in params.items():
                if k in instance.entries:
                    entry = instance.entries[k]
                    # 清空并填入
                    try:
                        entry.delete(0, tk.END)
                        entry.insert(0, str(val))
                    except:
                        pass
        
        # 其次同步更新内部 params 字典
        if hasattr(instance, "params") and isinstance(instance.params, dict):
            for k, val in params.items():
                if k in instance.params:
                    # 尝试进行类型转换，防止出错
                    orig_type = type(instance.params[k])
                    try:
                        instance.params[k] = orig_type(val)
                    except:
                        instance.params[k] = val

    def on_close(self):
        """关闭窗口时保存所有活跃模块的参数"""
        print("正在保存配置并关闭平台...")
        for name, instance in list(self.active_modules.items()):
            # 确保在退出前停止子模块线程
            if hasattr(instance, "stop") and callable(instance.stop):
                 try:
                    instance.stop()
                 except:
                    pass
            self.save_module_params(name, instance)
        self.save_config()
        self.root.destroy()
        # 强制退出，确保所有后台线程结束
        os._exit(0) 

if __name__ == "__main__":
    # 多进程支持（如果底层脚本用到）
    import multiprocessing
    multiprocessing.freeze_support() # 仅在打包EXE时需要

    root = tk.Tk()
    app = IntegratedPlatform(root)
    root.mainloop()
    # pyinstaller package.spec