# PerformanceCurveConfigDialog.py - 修復版本

import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.colors as mcolors

class PerformanceCurveConfigDialog:
    def __init__(self, parent, default_ranges, available_curves=None):
        self.parent = parent
        self.default_ranges = default_ranges
        
        # 定義所有可能的曲線類型
        self.all_possible_curves = ['ps', 'pt', 'h', 'eff_s', 'eff_t', 'n', 'i', 'v', 'f', 'temp']
        
        # 確保包含所有10種曲線類型
        if available_curves is None:
            self.available_curves = self.all_possible_curves
        else:
            # 過濾掉不在 all_possible_curves 中的曲線
            self.available_curves = [curve for curve in available_curves if curve in self.all_possible_curves]
            print(f"過濾後的可用曲線: {self.available_curves}")

        self.result = None

        print(f"最終可用的曲線: {self.available_curves}")

        # 擴展到10條曲線的樣式預設值
        self.curve_styles = {
            'ps': {'color': 'blue', 'marker': 'o', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'pt': {'color': 'red', 'marker': 's', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'h': {'color': 'green', 'marker': '^', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'eff_s': {'color': 'purple', 'marker': 'D', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'eff_t': {'color': 'orange', 'marker': 'v', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'n': {'color': 'brown', 'marker': 'p', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'i': {'color': 'pink', 'marker': 'h', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'v': {'color': 'cyan', 'marker': '+', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'f': {'color': 'gray', 'marker': 'x', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8},
            'temp': {'color': 'magenta', 'marker': '*', 'marker_filled': True, 'line_style': '-', 'linewidth': 2.0, 'markersize': 8}
        }
        
        # 擴展曲線顯示名稱映射
        self.curve_display_names = {
            'ps': '靜壓 (Ps)',
            'pt': '全壓 (Pt)',
            'h': '輸入功率 (H)',
            'eff_s': '靜壓效率 (ηs)',
            'eff_t': '全壓效率 (ηt)',
            'n': '轉速 (N)',
            'i': '電流 (I)',
            'v': '電壓 (V)',
            'f': '頻率 (F)',
            'temp': '溫度 (Temp)'
        }
        # 新增：軸格式選項
        self.axis_format_options = {
            '一般數值': 'normal',
            '科學記號': 'scientific',
            '工程記號': 'engineering',
            '對數座標': 'log'
        }
        # 確保所有 available_curves 都有對應的顯示名稱和樣式
        for curve in self.available_curves:
            if curve not in self.curve_display_names:
                # 為未知曲線創建默認顯示名稱
                self.curve_display_names[curve] = f'{curve.upper()}'
                print(f"為未知曲線 {curve} 創建默認顯示名稱")
            
            if curve not in self.curve_styles:
                # 為未知曲線創建默認樣式
                default_colors = ['blue', 'red', 'green', 'purple', 'orange', 'brown', 'pink', 'cyan', 'gray', 'magenta']
                default_markers = ['o', 's', '^', 'D', 'v', 'p', 'h', '+', 'x', '*']
                
                color_index = len(self.curve_styles) % len(default_colors)
                marker_index = len(self.curve_styles) % len(default_markers)
                
                self.curve_styles[curve] = {
                    'color': default_colors[color_index],
                    'marker': default_markers[marker_index],
                    'marker_filled': True,
                    'line_style': '-',
                    'linewidth': 2.0,
                    'markersize': 8
                }
                print(f"為未知曲線 {curve} 創建默認樣式")

        # 軸系映射
        self.axis_mapping = {
            '左軸 (主縱軸)': 'y1',
            '右軸1': 'y2', 
            '右軸2': 'y3',
            '右軸3': 'y4',
            '右軸4': 'y5',
            '右軸5': 'y6',
            '右軸6': 'y7',
            '右軸7': 'y8',
            '右軸8': 'y9'
        }

        # 圖例位置選項
        self.legend_locations = {
            '無圖例': 'none',  # 新增無圖例選項
            '圖上方標題下方': 'upper center',
            '圖內左上角': 'upper left',
            '圖內右上角': 'upper right',
            '圖內左下角': 'lower left',
            '圖內右下角': 'lower right',
            '圖下方': 'lower center',
            '圖中央': 'center'
        }
        
        # 可用顏色列表 - 擴展更多顏色
        self.colors = (list(mcolors.TABLEAU_COLORS.keys()) + 
                      list(mcolors.BASE_COLORS.keys()) + 
                      ['darkred', 'darkgreen', 'darkblue', 'darkorange', 'darkviolet'])
        
        # 軸線設定相關變數
        self.x_axis_linewidth_var = None
        self.y_axis_linewidth_var = None
        self.x_label_color_var = None
        self.y_label_color_var = None
        self.show_grid_var = None
        self.grid_style_var = None
        self.grid_alpha_var = None
        self.x_label_fontsize_var = None
        self.y_label_fontsize_var = None
        self.x_tick_fontsize_var = None
        self.y_tick_fontsize_var = None
    
        # 標題設定相關變數
        self.show_title_var = None
        self.title_text_var = None
        self.title_fontsize_var = None
        self.title_fontweight_var = None
    
        # 圖例設定相關變數
        self.legend_location_var = None
        self.legend_vertical_offset_var = None
        self.show_legend_frame_var = None
        self.legend_alpha_var = None
        self.legend_fontsize_var = None
    
        self.dialog = None
        self.create_dialog()

    def create_dialog(self):
        try:
            print("創建設定對話框窗口...")
        
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("性能曲線設定 - 10曲線版本")
            self.dialog.geometry("1100x650")
        
            self.dialog.resizable(True, True)
            self.dialog.minsize(900, 550)
            self.dialog.maxsize(1400, 900)
        
            self.dialog.transient(self.parent)
            self.dialog.grab_set()
            self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
            # 主框架
            main_frame = tk.Frame(self.dialog)
            main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
            # 頂部按鈕框架
            top_btn_frame = tk.Frame(main_frame, bg="#F0F0F0", height=45)
            top_btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))
            top_btn_frame.pack_propagate(False)
        
            # 左側標題
            title_label = tk.Label(top_btn_frame, text="📊 性能曲線設定 (10曲線)", 
                                  font=("Microsoft JhengHei", 11, "bold"),
                                  bg="#F0F0F0", fg="#2E8B57")
            title_label.pack(side=tk.LEFT, padx=15, pady=10)
        
            # 右側按鈕
            btn_subframe = tk.Frame(top_btn_frame, bg="#F0F0F0")
            btn_subframe.pack(side=tk.RIGHT, padx=10, pady=8)
        
            ok_btn = tk.Button(btn_subframe, text="✅ 確定", command=self.on_ok, 
                              bg="#32CD32", fg="white", 
                              font=("Microsoft JhengHei", 10, "bold"),
                              width=8, height=1)
            ok_btn.pack(side=tk.LEFT, padx=5)
        
            cancel_btn = tk.Button(btn_subframe, text="❌ 取消", command=self.on_cancel,
                                  bg="#FF6B6B", fg="white",
                                  font=("Microsoft JhengHei", 10, "bold"),
                                  width=8, height=1)
            cancel_btn.pack(side=tk.LEFT, padx=5)
        
            # 筆記本框架
            notebook_container = tk.Frame(main_frame, bg="white")
            notebook_container.pack(fill=tk.BOTH, expand=True)
        
            # 創建筆記本控件
            notebook = ttk.Notebook(notebook_container)
            notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
            # 創建各頁籤
            basic_frame = ttk.Frame(notebook)
            notebook.add(basic_frame, text="🔧 曲線與軸系")
        
            style_frame = ttk.Frame(notebook)
            notebook.add(style_frame, text="🎨 曲線樣式")
        
            legend_frame = ttk.Frame(notebook)
            notebook.add(legend_frame, text="📖 圖例設定")
        
            axis_frame = ttk.Frame(notebook)
            notebook.add(axis_frame, text="📐 軸線設定")
        
            self.create_basic_tab(basic_frame)
            self.create_style_tab(style_frame)
            self.create_legend_tab(legend_frame)
            self.create_axis_tab(axis_frame)
        
            # 底部說明文字
            bottom_frame = tk.Frame(main_frame, bg="#F8F8F8", height=30)
            bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
            bottom_frame.pack_propagate(False)
        
            help_label = tk.Label(bottom_frame, 
                                 text="💡 支援10條曲線，可自由分配至9個座標軸（1個左軸 + 8個右軸）",
                                 font=("Microsoft JhengHei", 9), 
                                 bg="#F8F8F8", fg="#666666")
            help_label.pack(pady=5)
        
            # 讓確定按鈕獲得焦點
            self.dialog.after(100, lambda: ok_btn.focus_set())
            self.dialog.bind('<Return>', lambda e: self.on_ok())
            self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
            # 等待窗口關閉
            self.parent.wait_window(self.dialog)
        
        except Exception as e:
            print(f"創建設定對話框時發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"創建設定對話框時發生錯誤：\n{str(e)}")

    def create_basic_tab(self, parent):
        """創建基本設定創建 - 包含軸系分配、格式設置和圖例名稱"""
        print(f"創建基本創建，可用的曲線: {self.available_curves}")
    
        # 創建滾動框架
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ========== 添加X軸範圍設置 ==========
        x_axis_frame = ttk.LabelFrame(scrollable_frame, text="X軸範圍設置 (流量)", padding=8)
        x_axis_frame.pack(fill=tk.X, padx=5, pady=3)
    
        # X轴最小值
        x_min_frame = tk.Frame(x_axis_frame)
        x_min_frame.pack(fill=tk.X, pady=2)
        tk.Label(x_min_frame, text="流量最小值 (Qmin):", width=15).pack(side=tk.LEFT)
        default_x_range = self.default_ranges.get('x', (0, 100))
        self.x_min_var = tk.DoubleVar(value=default_x_range[0])
        x_min_entry = tk.Entry(x_min_frame, textvariable=self.x_min_var, width=10)
        x_min_entry.pack(side=tk.LEFT, padx=5)
    
        # X軸最大值
        x_max_frame = tk.Frame(x_axis_frame)
        x_max_frame.pack(fill=tk.X, pady=2)
        tk.Label(x_max_frame, text="流量最大值 (Qmax):", width=15).pack(side=tk.LEFT)
        self.x_max_var = tk.DoubleVar(value=default_x_range[1])
        x_max_entry = tk.Entry(x_max_frame, textvariable=self.x_max_var, width=10)
        x_max_entry.pack(side=tk.LEFT, padx=5)
    
        # X軸標籤
        x_label_frame = tk.Frame(x_axis_frame)
        x_label_frame.pack(fill=tk.X, pady=2)
        tk.Label(x_label_frame, text="X軸標籤:", width=15).pack(side=tk.LEFT)
        self.x_label_var = tk.StringVar(value="流量 (CMM)")
        x_label_entry = tk.Entry(x_label_frame, textvariable=self.x_label_var, width=20)
        x_label_entry.pack(side=tk.LEFT, padx=5)

        # X軸格式選擇
        x_format_frame = tk.Frame(x_axis_frame)
        x_format_frame.pack(fill=tk.X, pady=2)
        tk.Label(x_format_frame, text="X軸格式:", width=15).pack(side=tk.LEFT)
        self.x_format_var = tk.StringVar(value='一般數值')
        x_format_combo = ttk.Combobox(x_format_frame, textvariable=self.x_format_var,
                                      values=list(self.axis_format_options.keys()),
                                      state="readonly", width=15)
        x_format_combo.pack(side=tk.LEFT, padx=5)

        # 曲線與軸系對應設定
        axis_curve_frame = ttk.LabelFrame(scrollable_frame, text="曲線軸系分配 (支援10條曲線)", padding=8)
        axis_curve_frame.pack(fill=tk.X, padx=5, pady=3)

        # ========== 新增：快速選擇按钮 ==========
        button_frame = tk.Frame(axis_curve_frame)
        button_frame.grid(row=0, column=0, columnspan=9, padx=5, pady=5, sticky="ew")  # 改为9列

        # All Clear 按钮
        all_clear_btn = tk.Button(button_frame, text="🗑️ All Clear", 
                                 command=self.clear_all_curves,
                                 bg="#FF6B6B", fg="white",
                                 font=("Microsoft JhengHei", 9),
                                 width=12, height=1)
        all_clear_btn.pack(side=tk.LEFT, padx=5)

        # Default 按钮
        default_btn = tk.Button(button_frame, text="🔧 Default", 
                               command=self.set_default_curves,
                               bg="#32CD32", fg="white",
                               font=("Microsoft JhengHei", 9),
                               width=12, height=1)
        default_btn.pack(side=tk.LEFT, padx=5)

        # 狀態標籤
        self.selection_status_var = tk.StringVar(value="預設狀態：全部不勾選")
        status_label = tk.Label(button_frame, textvariable=self.selection_status_var,
                              font=("Microsoft JhengHei", 9), fg="blue")
        status_label.pack(side=tk.LEFT, padx=20)

        # ========== 擴展表格標頭，添加圖例名稱、軸格式與範圍設定 ==========
        headers = ['曲線', '顯示', '圖例名稱', '對應軸系', '軸標籤', '格式', '最小值', '最大值', '設為主縱軸']
        for i, header in enumerate(headers):
            width = 12
            if i == 0:  # 曲線名稱
                width = 12
            elif i in [1, 8]:  # 顯示、設為主縱軸
                width = 8
            elif i in [6, 7]:  # 最小值、最大值
                width = 10
            elif i == 2:  # 圖例名稱
                width = 15
            else:  # 對應軸系、軸標籤、格式
                width = 15
            tk.Label(axis_curve_frame, text=header, font=("Microsoft JhengHei", 9, "bold"), 
                    width=width).grid(row=1, column=i, padx=2, pady=5, sticky="ew")  # 改為 row=1

        # 為每一列設定權重，讓其自動伸展
        for i in range(len(headers)):
            axis_curve_frame.grid_columnconfigure(i, weight=1)

        self.curve_axis_vars = {}
        self.main_axis_var = tk.StringVar(value="")  # 紀錄主縱軸選擇
        curves_to_show = []
        for curve_key in self.available_curves:
            if curve_key in self.curve_display_names:
                display_name = self.curve_display_names[curve_key]
                curves_to_show.append((curve_key, display_name))
                #print(f"添加曲線到基本頁籤: {curve_key} -> {display_name}")

        # 從 row=2 開始（因为按鈕在 row=0，標頭在 row=1）
        for row, (key, label) in enumerate(curves_to_show, 2):  # 改為從2開始
            # 曲線名稱
            tk.Label(axis_curve_frame, text=label, width=12, anchor="w").grid(
                row=row, column=0, padx=2, pady=2, sticky="w")
    
            # 顯示選擇 - 預設都不勾選
            show_var = tk.BooleanVar(value=False)  # 改為預設 False
            show_cb = tk.Checkbutton(axis_curve_frame, variable=show_var)
            show_cb.grid(row=row, column=1, padx=2, pady=2)
    
            # ========== 新增：圖例名稱編輯字段 ==========
            legend_name_var = tk.StringVar(value=self.curve_display_names[key])
            legend_name_entry = tk.Entry(axis_curve_frame, textvariable=legend_name_var, width=15)
            legend_name_entry.grid(row=row, column=2, padx=2, pady=2, sticky="ew")
    
            # 軸系選擇 - 動態選項
            axis_options = list(self.axis_mapping.keys())
            axis_var = tk.StringVar(value=self.get_default_axis(key))
            axis_combo = ttk.Combobox(axis_curve_frame, textvariable=axis_var,
                                    values=axis_options, state="readonly", width=15)
            axis_combo.grid(row=row, column=3, padx=2, pady=2, sticky="ew")
    
            # 軸編籤自定義
            label_var = tk.StringVar(value=self.get_default_axis_label(key))
            label_entry = tk.Entry(axis_curve_frame, textvariable=label_var, width=15)
            label_entry.grid(row=row, column=4, padx=2, pady=2, sticky="ew")
    
            # 軸格式選擇
            format_var = tk.StringVar(value='一般數值')
            format_combo = ttk.Combobox(axis_curve_frame, textvariable=format_var,
                                       values=list(self.axis_format_options.keys()),
                                       state="readonly", width=15)
            format_combo.grid(row=row, column=5, padx=2, pady=2, sticky="ew")

            # ========== 添加軸範圍設定 ==========
            # 獲取預設範圍
            axis_key = self.axis_mapping[axis_var.get()]
            default_min, default_max = self.get_default_range(axis_key)
            # 最小值输入框
            min_var = tk.DoubleVar(value=default_min)
            min_entry = tk.Entry(axis_curve_frame, textvariable=min_var, width=10)
            min_entry.grid(row=row, column=6, padx=2, pady=2, sticky="ew")
            # 最大值输入框
            max_var = tk.DoubleVar(value=default_max)
            max_entry = tk.Entry(axis_curve_frame, textvariable=max_var, width=10)
            max_entry.grid(row=row, column=7, padx=2, pady=2, sticky="ew")
    
            # 作為主縱軸選擇 (單選按鈕)
            main_axis_rb = tk.Radiobutton(axis_curve_frame, variable=self.main_axis_var, 
                                        value=key, text="")
            main_axis_rb.grid(row=row, column=8, padx=2, pady=2)
    
            # 儲存變量 - 新增 legend_name
            self.curve_axis_vars[key] = {
                'show': show_var,
                'legend_name': legend_name_var,  # 新增圖例名稱變量
                'axis': axis_var, 
                'label': label_var,
                'format': format_var,
                'min': min_var,
                'max': max_var
            }
    
            # 绑定事件：当选择"左轴 (主纵轴)"时自动设为主纵轴
            def create_axis_callback(curve_key):
                def callback(*args):
                    if self.curve_axis_vars[curve_key]['axis'].get() == '左軸 (主縱軸)':
                        self.main_axis_var.set(curve_key)
                return callback
            axis_var.trace('w', create_axis_callback(key))

        # 如果没有選擇主縱軸，自動選擇第一個（但保持不勾選狀態）
        if curves_to_show and not self.main_axis_var.get():
            first_curve = curves_to_show[0][0]
            self.main_axis_var.set(first_curve)
            self.curve_axis_vars[first_curve]['axis'].set('左軸 (主縱軸)')
            print(f"自動設定主縱軸: {first_curve}")

        # 圖表標題設定
        title_frame = ttk.LabelFrame(scrollable_frame, text="圖表標題", padding=8)
        title_frame.pack(fill=tk.X, padx=5, pady=3)
    
        # 第一行：顯示標題 + 標题文字
        title_row1 = tk.Frame(title_frame)
        title_row1.pack(fill=tk.X, pady=2)
        self.show_title_var = tk.BooleanVar(value=True)
        show_title_cb = tk.Checkbutton(title_row1, text="顯示標題",
                                     variable=self.show_title_var,
                                     font=("Microsoft JhengHei", 9))
        show_title_cb.pack(side=tk.LEFT, padx=5)
        tk.Label(title_row1, text="標題文字:", width=8).pack(side=tk.LEFT, padx=(20, 5))
        self.title_text_var = tk.StringVar(value="性能曲線")
        title_text_entry = tk.Entry(title_row1, textvariable=self.title_text_var, width=20)
        title_text_entry.pack(side=tk.LEFT, padx=5)
    
        # 第二行：字體大小 + 粗細
        title_row2 = tk.Frame(title_frame)
        title_row2.pack(fill=tk.X, pady=2)
        tk.Label(title_row2, text="字體大小:", width=8).pack(side=tk.LEFT, padx=5)
        self.title_fontsize_var = tk.IntVar(value=14)
        title_font_spinbox = tk.Spinbox(title_row2, from_=8, to=24, 
                                      textvariable=self.title_fontsize_var,
                                      width=5)
        title_font_spinbox.pack(side=tk.LEFT, padx=5)
        tk.Label(title_row2, text="字體粗细:", width=8).pack(side=tk.LEFT, padx=(20, 5))
        self.title_fontweight_var = tk.StringVar(value="bold")
        bold_radio = tk.Radiobutton(title_row2, text="粗體", 
                                  variable=self.title_fontweight_var, value="bold",
                                  font=("Microsoft JhengHei", 8))
        bold_radio.pack(side=tk.LEFT, padx=2)
        normal_radio = tk.Radiobutton(title_row2, text="正常", 
                                    variable=self.title_fontweight_var, value="normal",
                                    font=("Microsoft JhengHei", 8))
        normal_radio.pack(side=tk.LEFT, padx=2)

        # 在 create_basic_tab 方法的結尾，添加以下代碼：

        # 確保所有曲線都有有效的軸系選擇
        for key, vars_dict in self.curve_axis_vars.items():
            current_axis = vars_dict['axis'].get()
            if not current_axis or current_axis not in self.axis_mapping:
                # 設置默認軸系
                default_axis = self.get_default_axis(key)
                vars_dict['axis'].set(default_axis)
                print(f"初始化曲線 {key} 的軸系為: {default_axis}")


    def clear_all_curves(self):
        """清除所有曲線的勾選狀態"""
        for key, vars_dict in self.curve_axis_vars.items():
            vars_dict['show'].set(False)
    
        self.selection_status_var.set("狀態：全部清除")
        print("已清除所有曲線選擇")

    def set_default_curves(self):
        """設定預設的曲線選擇（根據常見使用情境）"""
        # 預設選擇常見的性能曲線：靜壓、全壓、功率、效率
        default_curves = ['ps', 'pt', 'h', 'eff_s', 'eff_t']
    
        for key, vars_dict in self.curve_axis_vars.items():
            if key in default_curves:
                vars_dict['show'].set(True)
            else:
                vars_dict['show'].set(False)
    
        # 自動設定主縱軸為靜壓
        if 'ps' in self.curve_axis_vars:
            self.main_axis_var.set('ps')
            self.curve_axis_vars['ps']['axis'].set('左軸 (主縱軸)')
    
        self.selection_status_var.set("狀態：已套用預設選擇")
        print("已套用預設曲線選擇")



    def get_default_axis(self, curve_key):
        """根據曲線類型返回預設軸系"""
        default_axis_mapping = {
            'ps': '左軸 (主縱軸)',
            'pt': '左軸 (主縱軸)', 
            'h': '右軸1',
            'eff_s': '右軸2',
            'eff_t': '右軸2',
            'n': '右軸3',
            'i': '右軸4',
            'v': '右軸5',
            'f': '右軸6',
            'temp': '右軸7'
        }
        return default_axis_mapping.get(curve_key, '右軸8')

    def get_default_axis_label(self, curve_key):
        """根據曲線類型返回預設軸標籤"""
        default_label_mapping = {
            'ps': '靜壓 (mmAq)',
            'pt': '全壓 (mmAq)',
            'h': '輸入功率 (kW)',
            'eff_s': '靜壓效率 (%)',
            'eff_t': '全壓效率 (%)', 
            'n': '轉速 (RPM)',
            'i': '電流 (A)',
            'v': '電壓 (V)',
            'f': '頻率 (Hz)',
            'temp': '溫度 (°C)'
        }
        return default_label_mapping.get(curve_key, '數值')

    def get_default_range(self, axis_key):
        """根據軸键獲取默認範圍"""
        # 首先檢查是否有用户自定義的范围
        if hasattr(self, 'user_ranges') and axis_key in self.user_ranges:
            return self.user_ranges[axis_key]
    
        if axis_key in self.default_ranges:
            min_val, max_val = self.default_ranges[axis_key]
            # 確保範圍是合理的數值
            if min_val is not None and max_val is not None and min_val < max_val:
                return (min_val, max_val)

        # 默認範圍 - 使用更靈活的範圍
        default_ranges = {
            'x': (0, 100),      # 流量範圍
            'y1': (0, 100),     # 壓力範圍
            'y2': (0, 100),     # 功率範圍
            'y3': (0, 100),     # 效率範圍
            'y4': (0, 100),     # 轉速範圍
            'y5': (0, 100),     # 電流範圍
            'y6': (0, 100),     # 電压範圍
            'y7': (0, 100),     # 频率範圍
            'y8': (0, 100),     # 温度範圍
            'y9': (0, 100)      # 備用範圍
        }

        return default_ranges.get(axis_key, (0, 100))


    def create_style_tab(self, parent):
        """創建樣式設定頁籤 - 支援10條曲線"""
        print(f"創建樣式頁籤，可用的曲線: {self.available_curves}")
    
        # 標記樣式
        self.markers = {
            '圓圈': 'o',
            '方塊': 's', 
            '三角形': '^',
            '倒三角形': 'v',
            '菱形': 'D',
            '五角星': 'p',
            '六角形': 'h',
            '加號': '+',
            '叉號': 'x',
            '星號': '*',
            '點': '.'
        }
    
        # 線條樣式
        self.line_styles = {
            '實線': '-',
            '虛線': '--',
            '點線': ':',
            '點劃線': '-.'
        }
    
        # 創建滾動框架
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
    
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
    
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
    
        # 只為有數據的曲線創建設定
        curves_to_show = []
        for curve_key in self.available_curves:
            if curve_key in self.curve_display_names:
                display_name = self.curve_display_names[curve_key]
                curves_to_show.append((curve_key, display_name))
    
        # 使用網格佈局來節省空間
        for i, (key, label) in enumerate(curves_to_show):
            row = i // 2
            col = (i % 2) * 3
            
            frame = ttk.LabelFrame(scrollable_frame, text=label, padding=8)
            frame.grid(row=row, column=col, columnspan=3, padx=5, pady=5, sticky="ew")
            
            self.create_curve_style_widgets(frame, key)
    
        # 如果沒有曲線可顯示，顯示提示訊息
        if not curves_to_show:
            no_data_label = tk.Label(scrollable_frame, 
                                   text="沒有可設定的曲線樣式\n請確認Excel文件包含有效的性能數據",
                                   font=("Microsoft JhengHei", 10), fg="gray")
            no_data_label.pack(pady=20)
    
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def create_curve_style_widgets(self, parent, curve_key):
        """為單個曲線創建樣式設定控件"""
        # 確保曲線鍵存在於樣式預設值中
        if curve_key not in self.curve_styles:
            # 使用循環的顏色以確保唯一性
            default_colors = ['blue', 'red', 'green', 'purple', 'orange', 'brown', 
                            'pink', 'cyan', 'gray', 'magenta', 'olive', 'navy']
            default_markers = ['o', 's', '^', 'D', 'v', 'p', 'h', '+', 'x', '*', '.', '1']
        
            color_index = len(self.curve_styles) % len(default_colors)
            marker_index = len(self.curve_styles) % len(default_markers)
        
            self.curve_styles[curve_key] = {
                'color': default_colors[color_index],
                'marker': default_markers[marker_index],
                'marker_filled': True,
                'line_style': '-',
                'linewidth': 2.0,
                'markersize': 8
            }
    
        style = self.curve_styles[curve_key]
    
        # 使用網格佈局來節省空間
        # 第一行：顏色和標記
        row1_frame = tk.Frame(parent)
        row1_frame.pack(fill=tk.X, pady=2)
    
        tk.Label(row1_frame, text="顏色:", width=6).pack(side=tk.LEFT)
        color_var = tk.StringVar(value=style['color'])
        color_combo = ttk.Combobox(row1_frame, textvariable=color_var, 
                                  values=self.colors, state="readonly", width=12)
        color_combo.pack(side=tk.LEFT, padx=2)
        setattr(self, f'{curve_key}_color', color_var)
    
        tk.Label(row1_frame, text="標記:", width=6).pack(side=tk.LEFT, padx=(10,0))
        marker_var = tk.StringVar(value=self.get_marker_name(style['marker']))
        marker_combo = ttk.Combobox(row1_frame, textvariable=marker_var,
                                   values=list(self.markers.keys()), state="readonly", width=10)
        marker_combo.pack(side=tk.LEFT, padx=2)
        setattr(self, f'{curve_key}_marker', marker_var)
    
        # 第二行：線條樣式和填充
        row2_frame = tk.Frame(parent)
        row2_frame.pack(fill=tk.X, pady=2)
    
        tk.Label(row2_frame, text="線條:", width=6).pack(side=tk.LEFT)
        line_var = tk.StringVar(value=self.get_line_style_name(style['line_style']))
        line_combo = ttk.Combobox(row2_frame, textvariable=line_var,
                                 values=list(self.line_styles.keys()), state="readonly", width=10)
        line_combo.pack(side=tk.LEFT, padx=2)
        setattr(self, f'{curve_key}_line_style', line_var)
    
        tk.Label(row2_frame, text="填充:", width=6).pack(side=tk.LEFT, padx=(10,0))
        fill_var = tk.BooleanVar(value=style['marker_filled'])
        filled_radio = tk.Radiobutton(row2_frame, text="實心", variable=fill_var, value=True,
                                     font=("Microsoft JhengHei", 8))
        filled_radio.pack(side=tk.LEFT)
        hollow_radio = tk.Radiobutton(row2_frame, text="空心", variable=fill_var, value=False,
                                     font=("Microsoft JhengHei", 8))
        hollow_radio.pack(side=tk.LEFT)
        setattr(self, f'{curve_key}_filled', fill_var)
    
        # 第三行：線條粗細和符號大小
        row3_frame = tk.Frame(parent)
        row3_frame.pack(fill=tk.X, pady=2)
    
        tk.Label(row3_frame, text="線粗:", width=6).pack(side=tk.LEFT)
        linewidth_var = tk.DoubleVar(value=style['linewidth'])
        linewidth_scale = tk.Scale(row3_frame, from_=0.5, to=5.0, 
                                  resolution=0.5, orient=tk.HORIZONTAL,
                                  variable=linewidth_var, length=80, showvalue=True)
        linewidth_scale.pack(side=tk.LEFT, padx=2)
        setattr(self, f'{curve_key}_linewidth', linewidth_var)
    
        tk.Label(row3_frame, text="符號:", width=6).pack(side=tk.LEFT, padx=(10,0))
        markersize_var = tk.IntVar(value=style['markersize'])
        markersize_scale = tk.Scale(row3_frame, from_=4, to=20, 
                                   orient=tk.HORIZONTAL,
                                   variable=markersize_var, length=80, showvalue=True)
        markersize_scale.pack(side=tk.LEFT, padx=2)
        setattr(self, f'{curve_key}_markersize', markersize_var)


    def create_legend_tab(self, parent):
        """創建圖例設定頁籤"""
        # 圖例位置選擇
        location_frame = ttk.LabelFrame(parent, text="圖例位置", padding=10)
        location_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(location_frame, text="選擇圖例位置:", 
                font=("Microsoft JhengHei", 10)).pack(anchor="w", pady=5)
        
        # 初始化圖例位置變數
        self.legend_location_var = tk.StringVar(value="圖上方標題下方")
        
        locations = [
            '無圖例',  # 新增無圖例選項
            '圖上方標題下方',
            '圖內左上角', 
            '圖內右上角',
            '圖內左下角',
            '圖內右下角',
            '圖下方',
            '圖中央'
        ]
        
        for i, location in enumerate(locations):
            rb = tk.Radiobutton(location_frame, text=location, 
                              variable=self.legend_location_var, value=location,
                              font=("Microsoft JhengHei", 9))
            rb.pack(anchor="w", padx=20, pady=2)
        
        # 圖例垂直位置微調（僅適用於圖上方標題下方）
        vertical_frame = tk.Frame(location_frame)
        vertical_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(vertical_frame, text="垂直位置微調:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)
        
        self.legend_vertical_offset_var = tk.DoubleVar(value=0.95)
        vertical_scale = tk.Scale(vertical_frame, from_=0.85, to=0.99, 
                                resolution=0.01, orient=tk.HORIZONTAL,
                                variable=self.legend_vertical_offset_var,
                                length=200, showvalue=True)
        vertical_scale.pack(side=tk.LEFT, padx=10)
        
        # 圖例框線設定
        frame_style_frame = ttk.LabelFrame(parent, text="圖例樣式", padding=10)
        frame_style_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 顯示圖例框線
        self.show_legend_frame_var = tk.BooleanVar(value=True)
        frame_cb = tk.Checkbutton(frame_style_frame, text="顯示圖例框線",
                                variable=self.show_legend_frame_var,
                                font=("Microsoft JhengHei", 10))
        frame_cb.pack(anchor="w", pady=5)
        
        # 圖例背景透明度
        transparency_frame = tk.Frame(frame_style_frame)
        transparency_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(transparency_frame, text="背景透明度:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)
        
        self.legend_alpha_var = tk.DoubleVar(value=0.9)
        alpha_scale = tk.Scale(transparency_frame, from_=0.1, to=1.0, 
                              resolution=0.1, orient=tk.HORIZONTAL,
                              variable=self.legend_alpha_var,
                              length=200, showvalue=True)
        alpha_scale.pack(side=tk.LEFT, padx=10)
        
        # 圖例字體大小
        font_frame = tk.Frame(frame_style_frame)
        font_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(font_frame, text="字體大小:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)
        
        self.legend_fontsize_var = tk.IntVar(value=10)
        font_spinbox = tk.Spinbox(font_frame, from_=8, to=16, 
                                 textvariable=self.legend_fontsize_var,
                                 width=5)
        font_spinbox.pack(side=tk.LEFT, padx=10)
    
    def create_axis_tab(self, parent):
        """創建軸線設定頁籤"""
        # 軸線粗細設定
        linewidth_frame = ttk.LabelFrame(parent, text="軸線粗細", padding=10)
        linewidth_frame.pack(fill=tk.X, padx=10, pady=5)

        # X軸線粗細
        x_axis_frame = tk.Frame(linewidth_frame)
        x_axis_frame.pack(fill=tk.X, pady=2)

        tk.Label(x_axis_frame, text="X軸線粗細:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.x_axis_linewidth_var = tk.DoubleVar(value=1.5)
        x_axis_scale = tk.Scale(x_axis_frame, from_=0.5, to=3.0, 
                              resolution=0.1, orient=tk.HORIZONTAL,
                              variable=self.x_axis_linewidth_var,
                              length=200, showvalue=True)
        x_axis_scale.pack(side=tk.LEFT, padx=10)

        # Y軸線粗細
        y_axis_frame = tk.Frame(linewidth_frame)
        y_axis_frame.pack(fill=tk.X, pady=2)

        tk.Label(y_axis_frame, text="Y軸線粗細:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.y_axis_linewidth_var = tk.DoubleVar(value=1.5)
        y_axis_scale = tk.Scale(y_axis_frame, from_=0.5, to=3.0, 
                              resolution=0.1, orient=tk.HORIZONTAL,
                              variable=self.y_axis_linewidth_var,
                              length=200, showvalue=True)
        y_axis_scale.pack(side=tk.LEFT, padx=10)

        # 軸標籤顏色設定
        label_color_frame = ttk.LabelFrame(parent, text="軸標籤顏色", padding=10)
        label_color_frame.pack(fill=tk.X, padx=10, pady=5)

        # X軸標籤顏色
        x_label_color_frame = tk.Frame(label_color_frame)
        x_label_color_frame.pack(fill=tk.X, pady=2)

        tk.Label(x_label_color_frame, text="X軸標籤顏色:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.x_label_color_var = tk.StringVar(value="black")
        x_label_color_combo = ttk.Combobox(x_label_color_frame, 
                                         textvariable=self.x_label_color_var,
                                         values=self.colors, 
                                         state="readonly", width=15)
        x_label_color_combo.pack(side=tk.LEFT, padx=10)

        # Y軸標籤顏色（左側 - 主縱軸）
        y_label_color_frame = tk.Frame(label_color_frame)
        y_label_color_frame.pack(fill=tk.X, pady=2)

        tk.Label(y_label_color_frame, text="主縱軸標籤顏色:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.y_label_color_var = tk.StringVar(value="blue")
        y_label_color_combo = ttk.Combobox(y_label_color_frame, 
                                         textvariable=self.y_label_color_var,
                                         values=self.colors, 
                                         state="readonly", width=15)
        y_label_color_combo.pack(side=tk.LEFT, padx=10)

        # 說明文字
        note_label = tk.Label(label_color_frame, 
                             text="提示：右側Y軸顏色會自動對應曲線顏色，可在「曲線樣式」頁籤中修改",
                             font=("Microsoft JhengHei", 8), 
                             fg="gray", wraplength=400, justify="left")
        note_label.pack(anchor="w", pady=5)

        # 格線設定
        grid_frame = ttk.LabelFrame(parent, text="格線設定", padding=10)
        grid_frame.pack(fill=tk.X, padx=10, pady=5)

        # 顯示格線
        self.show_grid_var = tk.BooleanVar(value=True)
        grid_cb = tk.Checkbutton(grid_frame, text="顯示格線",
                               variable=self.show_grid_var,
                               font=("Microsoft JhengHei", 10))
        grid_cb.pack(anchor="w", pady=5)

        # 格線樣式
        grid_style_frame = tk.Frame(grid_frame)
        grid_style_frame.pack(fill=tk.X, pady=5)

        tk.Label(grid_style_frame, text="格線樣式:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.grid_style_var = tk.StringVar(value="--")
        grid_style_combo = ttk.Combobox(grid_style_frame, textvariable=self.grid_style_var,
                                      values=['-', '--', ':', '-.', 'None'],
                                      state="readonly", width=10)
        grid_style_combo.pack(side=tk.LEFT, padx=10)

        # 格線透明度
        grid_alpha_frame = tk.Frame(grid_frame)
        grid_alpha_frame.pack(fill=tk.X, pady=5)

        tk.Label(grid_alpha_frame, text="格線透明度:", 
                font=("Microsoft JhengHei", 9), width=12).pack(side=tk.LEFT)

        self.grid_alpha_var = tk.DoubleVar(value=0.3)
        grid_alpha_scale = tk.Scale(grid_alpha_frame, from_=0.1, to=1.0, 
                                  resolution=0.1, orient=tk.HORIZONTAL,
                                  variable=self.grid_alpha_var,
                                  length=200, showvalue=True)
        grid_alpha_scale.pack(side=tk.LEFT, padx=10)

        # 軸標籤字體大小
        label_font_frame = ttk.LabelFrame(parent, text="軸標籤字體大小", padding=10)
        label_font_frame.pack(fill=tk.X, padx=10, pady=5)

        # X軸標籤字體大小
        x_label_frame = tk.Frame(label_font_frame)
        x_label_frame.pack(fill=tk.X, pady=2)

        tk.Label(x_label_frame, text="X軸標籤字體大小:", 
                font=("Microsoft JhengHei", 9), width=15).pack(side=tk.LEFT)

        self.x_label_fontsize_var = tk.IntVar(value=12)
        x_label_spinbox = tk.Spinbox(x_label_frame, from_=8, to=20, 
                                   textvariable=self.x_label_fontsize_var,
                                   width=5)
        x_label_spinbox.pack(side=tk.LEFT, padx=10)

        # Y軸標籤字體大小
        y_label_frame = tk.Frame(label_font_frame)
        y_label_frame.pack(fill=tk.X, pady=2)

        tk.Label(y_label_frame, text="Y軸標籤字體大小:", 
                font=("Microsoft JhengHei", 9), width=15).pack(side=tk.LEFT)

        self.y_label_fontsize_var = tk.IntVar(value=12)
        y_label_spinbox = tk.Spinbox(y_label_frame, from_=8, to=20, 
                                   textvariable=self.y_label_fontsize_var,
                                   width=5)
        y_label_spinbox.pack(side=tk.LEFT, padx=10)

        # 刻度標籤字體大小
        tick_font_frame = ttk.LabelFrame(parent, text="刻度標籤字體大小", padding=10)
        tick_font_frame.pack(fill=tk.X, padx=10, pady=5)

        # X軸刻度字體大小
        x_tick_frame = tk.Frame(tick_font_frame)
        x_tick_frame.pack(fill=tk.X, pady=2)

        tk.Label(x_tick_frame, text="X軸刻度字體大小:", 
                font=("Microsoft JhengHei", 9), width=15).pack(side=tk.LEFT)

        self.x_tick_fontsize_var = tk.IntVar(value=10)
        x_tick_spinbox = tk.Spinbox(x_tick_frame, from_=6, to=16, 
                                   textvariable=self.x_tick_fontsize_var,
                                   width=5)
        x_tick_spinbox.pack(side=tk.LEFT, padx=10)

        # Y軸刻度字體大小
        y_tick_frame = tk.Frame(tick_font_frame)
        y_tick_frame.pack(fill=tk.X, pady=2)

        tk.Label(y_tick_frame, text="Y軸刻度字體大小:", 
                font=("Microsoft JhengHei", 9), width=15).pack(side=tk.LEFT)

        self.y_tick_fontsize_var = tk.IntVar(value=10)
        y_tick_spinbox = tk.Spinbox(y_tick_frame, from_=6, to=16, 
                                   textvariable=self.y_tick_fontsize_var,
                                   width=5)
        y_tick_spinbox.pack(side=tk.LEFT, padx=10)


    def get_marker_name(self, marker_value):
        """根據標記值返回對應的名稱"""
        for name, value in self.markers.items():
            if value == marker_value:
                return name
        return '圓圈'

    def get_line_style_name(self, line_value):
        """根據線條值返回對應的名稱"""
        for name, value in self.line_styles.items():
            if value == line_value:
                return name
        return '實線'

    def on_ok(self):
        """確定按鈕處理 - 支援10條曲線和動態軸系"""
        try:
            # 收集曲線選擇和軸系分配
            curves = {}
            axis_assignments = {}
            axis_labels = {}
            ranges = {}
            legend_names = {}  # 新增：儲存圖力例名稱

            # ========== 收集X軸設置 ==========
            x_min = self.x_min_var.get()
            x_max = self.x_max_var.get()
            if x_min < x_max:  # 確保範圍有效
                ranges['x'] = (x_min, x_max)
    
            x_label = self.x_label_var.get().strip()
            if x_label:
                axis_labels['x'] = x_label

            # 獲取主縱軸
            main_axis_curve = self.main_axis_var.get()
            if not main_axis_curve and self.curve_axis_vars:
                # 如果没有選擇主縱軸，選擇第一個顯示的曲線
                for key in self.curve_axis_vars:
                    if self.curve_axis_vars[key]['show'].get():
                        main_axis_curve = key
                        break

            # ========== 收集軸範圍設定 ==========
            for key, vars_dict in self.curve_axis_vars.items():
                curves[key] = vars_dict['show'].get()

                if curves[key]:
                    # 收集圖例名稱
                    legend_name = vars_dict['legend_name'].get().strip()
                    if legend_name:
                        legend_names[key] = legend_name
                    else:
                        # 如果用户没有輸入，使用預設名稱
                        legend_names[key] = self.curve_display_names[key]
                
                    # 修復：檢查軸系選擇是否有效
                    axis_display_name = vars_dict['axis'].get()
                    if not axis_display_name:
                        # 如果軸系为空，使用默認軸系
                        axis_display_name = self.get_default_axis(key)
                        vars_dict['axis'].set(axis_display_name)
                
                    # 修復：確保軸系映射存在
                    if axis_display_name in self.axis_mapping:
                        axis_key = self.axis_mapping[axis_display_name]
                        axis_assignments[key] = axis_key
    
                        # 收集軸標籤
                        axis_label = vars_dict['label'].get().strip()
                        if axis_label:
                            axis_labels[axis_key] = axis_label
                        else:
                            # 如果軸標籤為空，使用默認標籤
                            axis_labels[axis_key] = self.get_default_axis_label(key)
    
                        # 收集軸範圍
                        min_val = vars_dict['min'].get()
                        max_val = vars_dict['max'].get()
                        if min_val < max_val:  # 確保範圍有效
                            ranges[axis_key] = (min_val, max_val)
                        else:
                            # 如果範圍無效，使用默認範圍
                            default_range = self.get_default_range(axis_key)
                            ranges[axis_key] = default_range
                    else:
                        print(f"⚠️ 警告：曲線 {key} 的軸系選擇 '{axis_display_name}' 無效，跳過該曲線")
                        curves[key] = False  # 禁用該曲線

            # ========== 修復：正確收集軸格式設定 ==========
            axis_format_settings = {}
            # X軸格式設定
            x_format_display = self.x_format_var.get()  # 從 UI 讀取 X 軸格式選擇
            axis_format_settings['x'] = self.axis_format_options.get(x_format_display, 'normal')
    
            # Y軸格式設定
            for key, vars_dict in self.curve_axis_vars.items():
                if vars_dict['show'].get() and key in axis_assignments:  # 只處理有效的曲線
                    axis_display_name = vars_dict['axis'].get()
                    if axis_display_name in self.axis_mapping:
                        axis_key = self.axis_mapping[axis_display_name]
                
                        # 穫取格式設定
                        format_display = vars_dict['format'].get()
                        format_value = self.axis_format_options.get(format_display, 'normal')
                        axis_format_settings[axis_key] = format_value
            
                        print(f"📊 曲線{key} -> 軸 {axis_key}: 格式={format_value}")

            print(f"✅ 最终軸格式設定: {axis_format_settings}")

            # 確保主縱軸在左側
            if main_axis_curve in axis_assignments:
                axis_assignments[main_axis_curve] = 'y1'
                # 確保主縱軸有標籤
                if 'y1' not in axis_labels:
                    axis_labels['y1'] = self.get_default_axis_label(main_axis_curve)

            # 如果没有設定範圍，使用默認範圍
            for axis_key in ['x', 'y1', 'y2', 'y3', 'y4', 'y5', 'y6', 'y7', 'y8', 'y9']:
                if axis_key not in ranges and axis_key in self.default_ranges:
                    ranges[axis_key] = self.default_ranges[axis_key]

            # 收集曲線樣式
            curve_styles = {}
            for key in self.curve_axis_vars.keys():
                try:
                    if hasattr(self, f'{key}_color'):
                        color_var = getattr(self, f'{key}_color')
                        marker_var = getattr(self, f'{key}_marker')
                        filled_var = getattr(self, f'{key}_filled')
                        line_style_var = getattr(self, f'{key}_line_style')
                        linewidth_var = getattr(self, f'{key}_linewidth')
                        markersize_var = getattr(self, f'{key}_markersize')
    
                        curve_styles[key] = {
                            'color': color_var.get(),
                            'marker': self.markers[marker_var.get()],
                            'marker_filled': filled_var.get(),
                            'line_style': self.line_styles[line_style_var.get()],
                            'linewidth': linewidth_var.get(),
                            'markersize': markersize_var.get()
                        }
    
                except Exception as e:
                    print(f"收集曲線 {key} 樣式時發生錯誤: {e}")
                    if key in self.curve_styles:
                        curve_styles[key] = self.curve_styles[key].copy()

            # 收集圖例設定
            legend_settings = {
                'location': self.legend_locations[self.legend_location_var.get()],
                'vertical_offset': self.legend_vertical_offset_var.get(),
                'show_frame': self.show_legend_frame_var.get(),
                'alpha': self.legend_alpha_var.get(),
                'fontsize': self.legend_fontsize_var.get()
            }

            # 收集軸線設定
            axis_settings = {
                'x_axis_linewidth': self.x_axis_linewidth_var.get(),
                'y_axis_linewidth': self.y_axis_linewidth_var.get(),
                'x_label_fontsize': self.x_label_fontsize_var.get(),
                'y_label_fontsize': self.y_label_fontsize_var.get(),
                'x_tick_fontsize': self.x_tick_fontsize_var.get(),
                'y_tick_fontsize': self.y_tick_fontsize_var.get(),
                # 標題設定
                'show_title': self.show_title_var.get(),
                'title_text': self.title_text_var.get(),
                'title_fontsize': self.title_fontsize_var.get(),
                'title_fontweight': self.title_fontweight_var.get(),
                # 軸標籤颜色設定
                'x_label_color': self.x_label_color_var.get(),
                'y_label_color': self.y_label_color_var.get(),
                # 主縱軸信息
                'main_axis_curve': main_axis_curve,
                # 格線設定
                'show_grid': self.show_grid_var.get(),
                'grid_style': self.grid_style_var.get(),
                'grid_alpha': self.grid_alpha_var.get(),
                # ========== 新增：軸格式設定 ==========
                'axis_format_settings': axis_format_settings
            }

            self.result = {
                'curves': curves,
                'axis_assignments': axis_assignments,
                'axis_labels': axis_labels,
                'ranges': ranges,
                'curve_styles': curve_styles,
                'legend_settings': legend_settings,
                'axis_settings': axis_settings,
                'legend_names': legend_names  # 新增圖例名稱
            }

            print(f"✅ 設定收集完成: {len(curve_styles)} 條曲線樣式")
            print(f"📊 軸系分配: {axis_assignments}")
            print(f"🎯 主縱軸: {main_axis_curve}")
            print(f"📈 曲線線式狀態: {curves}")
            print(f"🔢 X軸範圍: {ranges.get('x', '未設置')}")
            print(f"📏 軸線粗细 - X: {axis_settings['x_axis_linewidth']}, Y: {axis_settings['y_axis_linewidth']}")
            print(f"🔧 軸格式設定: {axis_format_settings}")
            print(f"🏷️ 圖例名稱: {legend_names}")

            self.dialog.destroy()

        except Exception as e:
            print(f"❌ 設定收集過程中發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"設定收集失敗：\n{str(e)}")


    def on_cancel(self):
        """取消按鈕處理"""
        self.result = None
        self.dialog.destroy()

    def show(self):
        """顯示對話框並返回結果"""
        return self.result