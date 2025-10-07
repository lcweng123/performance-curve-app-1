# PerformanceCurveApp.py - 完整版本（包含交點線和所有10條曲線支持）

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog, messagebox
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from scipy.interpolate import CubicSpline
import os
import tempfile
import datetime
from PerformanceCurveConfigDialog import PerformanceCurveConfigDialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
matplotlib.use('Agg') 

matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False
matplotlib.rcParams['mathtext.fontset'] = 'stix'  # 使用 STIX 字體支援數學符號


class PerformanceCurvePlotter:
    """支援10條曲線的通用性能曲線繪製器"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📈 通用性能曲線繪製器 (10曲線)")
        self.root.geometry("900x650")
        self.root.resizable(True, True)
        self.root.configure(bg="#F0F8FF")
        
        self.root.resizable(True, True)
        self.root.minsize(800, 600)
        self.root.maxsize(1400, 1000)

        # 初始化數據
        self.data = None
        self.fig = None
        self.canvas = None
        self.current_window = None

        # 新增：文字編輯相關變量
        self.current_fig = None
        self.current_canvas = None
        self.dragging_text = None
        self.text_objects = []
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.text_start_x = 0  # 新增
        self.text_start_y = 0  # 新增

        # 創建界面
        self.create_widgets()

    def create_widgets(self):
        """創建主界面"""
        main_frame = tk.Frame(self.root, bg="#F0F8FF")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 標題
        title_label = tk.Label(main_frame, text="📊 通用性能曲線繪製器 (支援10條曲線)", 
                              bg="#F0F8FF", fg="#4169E1",
                              font=("Microsoft JhengHei", 16, "bold"))
        title_label.pack(pady=(0, 20))

        # 說明文字
        desc_label = tk.Label(main_frame, 
                             text="支援10條曲線、9個座標軸（1個左軸 + 8個右軸）\n可自由分配曲線到任意軸系，適用於風機、泵、電機等各種設備",
                             bg="#F0F8FF", font=("Microsoft JhengHei", 10), 
                             wraplength=800, justify="center")
        desc_label.pack(pady=(0, 15))

        # 載入文件按鈕
        load_frame = tk.Frame(main_frame, bg="#F0F8FF")
        load_frame.pack(fill=tk.X, pady=10)
        load_btn = tk.Button(load_frame, text="📂 載入Excel文件",
                            command=self.load_excel_file,
                            bg="#32CD32", fg="white",
                            font=("Microsoft JhengHei", 11),
                            width=20, height=2)
        load_btn.pack(side=tk.LEFT, padx=10)
        self.file_label = tk.Label(load_frame, text="尚未載入檔案",
                                  bg="#F0F8FF", font=("Microsoft JhengHei", 10), fg="gray")
        self.file_label.pack(side=tk.LEFT, padx=20)

        # 除錯按鈕
        test_frame = tk.Frame(main_frame, bg="#F0F8FF")
        test_frame.pack(fill=tk.X, pady=5)
    
        test_btn = tk.Button(test_frame, text="🐛 除錯數據匹配",
                            command=self.debug_data_matching,
                            bg="#FF6B6B", fg="white",
                            font=("Microsoft JhengHei", 9),
                            width=15, height=1)
        test_btn.pack(side=tk.LEFT, padx=10)

        # 格式化數據按鈕
        format_frame = tk.Frame(main_frame, bg="#F0F8FF")
        format_frame.pack(fill=tk.X, pady=5)
        format_btn = tk.Button(format_frame, text="🔄 格式化Excel數據",
                              command=self.format_excel_data,
                              bg="#FF69B4", fg="white",
                              font=("Microsoft JhengHei", 10),
                              width=20, height=1)
        format_btn.pack(side=tk.LEFT, padx=10)
        self.format_status_label = tk.Label(format_frame, text="",
                                          bg="#F0F8FF", font=("Microsoft JhengHei", 9), fg="blue")
        self.format_status_label.pack(side=tk.LEFT, padx=10)

        # 繪製曲線按鈕
        plot_btn = tk.Button(main_frame, text="📈 繪製性能曲線 (10曲線)",
                            command=self.plot_performance_curve,
                            bg="#FF8C00", fg="white",
                            font=("Microsoft JhengHei", 12, "bold"),
                            width=25, height=2)
        plot_btn.pack(pady=20)

        # 功能說明框
        info_frame = ttk.LabelFrame(main_frame, text="💡 功能特色", padding=10)
        info_frame.pack(fill=tk.X, pady=10)
        
        features = [
            "• 支援最多10條性能曲線同時顯示",
            "• 9個座標軸（1個左軸 + 8個右軸）自由分配",
            "• 可設定任意曲線為主縱軸",
            "• 自動數據匹配和格式化",
            "• 高品質圖表輸出和剪貼簿複製"
        ]
        
        for feature in features:
            tk.Label(info_frame, text=feature, font=("Microsoft JhengHei", 9),
                    bg="white", justify="left").pack(anchor="w", pady=2)

        # 設計者資訊和日期
        designer_frame = tk.Frame(main_frame, bg="#F0F8FF")
        designer_frame.pack(fill=tk.X, pady=15)
    
        current_date = datetime.datetime.now().strftime("%Y/%m/%d")
    
        designer_label = tk.Label(designer_frame, text=f"設計者：翁淩家 | 日期：{current_date} | 版本：10曲線專業版",
                                bg="#F0F8FF", fg="#666666",
                                font=("Microsoft JhengHei", 10))
        designer_label.pack()

        # 狀態欄
        status_frame = tk.Frame(self.root, bg="#D3D3D3", height=25, relief="sunken", bd=1)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_label = tk.Label(status_frame, text="就緒 - 支援10條曲線繪製", 
                                   bg="#D3D3D3", fg="black",
                                   font=("Microsoft JhengHei", 9))
        self.status_label.pack(side=tk.LEFT, padx=10, pady=2)

    def format_excel_data(self):
        """格式化Excel數據，簡化欄位名稱和數值格式"""
        if self.data is None:
            messagebox.showwarning("警告", "請先載入Excel文件！")
            return
        
        try:
            formatted_data = self.data.copy()
            
            # 擴展欄位名稱映射（支援10種曲線類型）
            column_simplification = {
                'Qstd(CMM)': '流量(CMM)', '流量': '流量(CMM)', 'Q': '流量(CMM)', 'Flow': '流量(CMM)',
                'Psstd(mmAq)': '靜壓(mmAq)', '靜壓': '靜壓(mmAq)', 'Ps': '靜壓(mmAq)',
                'Ptstd(mmAq)': '全壓(mmAq)', '全壓': '全壓(mmAq)', 'Pt': '全壓(mmAq)',
                'Hstd(kW)': '功率(kW)', '功率': '功率(kW)', 'H': '功率(kW)', 'Input Power': '功率(kW)',
                'ηs(%)': '靜壓效率(%)', '靜壓效率': '靜壓效率(%)', 'ηs': '靜壓效率(%)',
                'ηt(%)': '全壓效率(%)', '全壓效率': '全壓效率(%)', 'ηt': '全壓效率(%)',
                'Nstd(rpm)': '轉速(rpm)', '轉速': '轉速(rpm)', 'N': '轉速(rpm)', 'Speed': '轉速(rpm)',
                '電流': '電流(A)', 'Current': '電流(A)', 'I': '電流(A)',
                '電壓': '電壓(V)', 'Voltage': '電壓(V)', 'V': '電壓(V)',
                '頻率': '頻率(Hz)', 'Frequency': '頻率(Hz)', 'F': '頻率(Hz)',
                '溫度': '溫度(°C)', 'Temperature': '溫度(°C)', 'Temp': '溫度(°C)'
            }
            
            # 重命名欄位
            new_columns = []
            for col in formatted_data.columns:
                if col in column_simplification:
                    new_columns.append(column_simplification[col])
                else:
                    col_lower = str(col).lower().replace('(', '').replace(')', '').replace(' ', '')
                    for key, value in column_simplification.items():
                        key_lower = str(key).lower().replace('(', '').replace(')', '').replace(' ', '')
                        if key_lower in col_lower or col_lower in key_lower:
                            new_columns.append(value)
                            break
                    else:
                        new_columns.append(col)
            
            formatted_data.columns = new_columns
            
            # 數值格式化
            numeric_columns = {
                '流量(CMM)': 1,
                '靜壓(mmAq)': 1,
                '全壓(mmAq)': 1,
                '功率(kW)': 2,
                '靜壓效率(%)': 2,
                '全壓效率(%)': 2,
                '轉速(rpm)': 0,
                '電流(A)': 2,
                '電壓(V)': 2,
                '頻率(Hz)': 1,
                '溫度(°C)': 1
            }
            
            for col, decimals in numeric_columns.items():
                if col in formatted_data.columns:
                    if decimals == 0:
                        formatted_data[col] = formatted_data[col].round(decimals).astype(int)
                    else:
                        formatted_data[col] = formatted_data[col].round(decimals)
            
            # 詢問是否保存
            save_choice = messagebox.askyesno("格式化完成", 
                                            "數據格式化完成！是否要保存為新的Excel檔案？")
            
            if save_choice:
                file_path = filedialog.asksaveasfilename(
                    title="保存格式化數據",
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                    initialfile="格式化_性能數據.xlsx"
                )
                
                if file_path:
                    formatted_data.to_excel(file_path, index=False, engine='openpyxl')
                    messagebox.showinfo("成功", f"格式化數據已保存至:\n{file_path}")
            
            self.data = formatted_data
            self.format_status_label.config(text="✅ 數據已格式化", fg="green")
            self.status_label.config(text="✅ 數據格式化完成")
            
            self.show_data_preview(formatted_data)
            
        except Exception as e:
            messagebox.showerror("格式化錯誤", f"數據格式化失敗:\n{str(e)}")
            self.format_status_label.config(text="❌ 格式化失敗", fg="red")

    def show_data_preview(self, data):
        """顯示數據預覽"""
        preview_text = "格式化後數據預覽:\n"
        preview_text += f"欄位: {', '.join(data.columns)}\n"
        preview_text += f"數據筆數: {len(data)}\n"
        preview_text += "前3筆數據:\n"
        
        for i in range(min(3, len(data))):
            row_data = []
            for col in data.columns:
                value = data.iloc[i][col]
                row_data.append(f"{value}")
            preview_text += f"第{i+1}筆: {', '.join(row_data)}\n"
        
        self.format_status_label.config(text=preview_text)

    def debug_data_matching(self):
        """除錯數據匹配過程"""
        if self.data is None:
            print("沒有數據載入")
            return
        """   
        print("=== 數據匹配除錯信息 ===")
        print(f"數據欄位: {list(self.data.columns)}")
        print(f"數據形狀: {self.data.shape}")
    
        print("前3行數據:")
        print(self.data.head(3))
    
        print("各欄位NaN數量:")
        print(self.data.isna().sum())
    
        print("各欄位數據範圍:")
        """
        for col in self.data.columns:
            if self.data[col].dtype in ['float64', 'int64']:
                min_val = self.data[col].min()
                max_val = self.data[col].max()
                print(f"  {col}: {min_val} ~ {max_val}")
    
        #print("=== 除錯結束 ===")

    def load_excel_file(self):
        """載入Excel文件"""
        file_path = filedialog.askopenfilename(
            title="選擇Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            try:
                self.data = pd.read_excel(file_path, engine='openpyxl')
                if self.canvas:
                    self.canvas.get_tk_widget().destroy()
                    self.canvas = None
                
                self.file_label.config(text=f"已載入: {os.path.basename(file_path)}", fg="green")
                self.status_label.config(text=f"✅ 已載入 {len(self.data)} 行數據")
                self.format_status_label.config(text="點擊「格式化Excel數據」進行優化", fg="blue")
            
                columns_info = f"偵測到欄位: {', '.join(self.data.columns)}"
                self.format_status_label.config(text=columns_info)
            
                self.debug_data_matching()
            
            except Exception as e:
                messagebox.showerror("載入錯誤", f"無法載入Excel檔案:{e}")
                self.file_label.config(text="載入失敗", fg="red")
                self.status_label.config(text="❌ 載入失敗")

    def get_data_columns(self):
        """自動匹配數據列 - 支援10種曲線類型"""
        if self.data is None:
            return None

        #print(f"數據列名稱: {list(self.data.columns)}")

        # 擴展數據列映射（支援10種曲線類型）
        column_mapping = {
            'flow': ['流量', 'Q', 'Flow', '流量(CMM)', 'Q (CMM)', 'flow', 'q', '風量'],
            'static_pressure': ['靜壓', 'Ps', 'Static Pressure', '靜壓(mmAq)', 'Ps (mmAq)', 'static_pressure', 'ps'],
            'total_pressure': ['全壓', 'Pt', 'Total Pressure', '全壓(mmAq)', 'Pt (mmAq)', 'total_pressure', 'pt'],
            'power': ['輸入功率', 'H', 'Input Power', '輸入功率(kW)', 'H (kW)', 'power', 'h', '功率', '功率(kW)'],
            'efficiency_static': ['靜壓效率', 'ηs', 'Static Efficiency', '靜壓效率(%)', 'ηs (%)', 'efficiency_static', 'eta_s'],
            'efficiency_total': ['全壓效率', 'ηt', 'Total Efficiency', '全壓效率(%)', 'ηt (%)', 'efficiency_total', 'eta_t'],
            'speed': ['轉速', 'N', 'Speed', '轉速(rpm)', 'N (rpm)', 'speed', 'n'],
            'current': ['電流', 'Current', 'I', '電流(A)', 'I (A)', 'current', 'amp'],
            'voltage': ['電壓', 'Voltage', 'V', '電壓(V)', 'V (V)', 'voltage'],
            'frequency': ['頻率', 'Frequency', 'F', '頻率(Hz)', 'F (Hz)', 'frequency', 'hz'],
            'temperature': ['溫度', 'Temperature', 'Temp', '溫度(°C)', '溫度(C)', 'temp']
        }

        matched_columns = {}
        for key, candidates in column_mapping.items():
            found = False
            for col in self.data.columns:
                col_clean = str(col).strip().lower().replace(' ', '').replace('(', '').replace(')', '').replace('_', '')
                for candidate in candidates:
                    candidate_clean = str(candidate).strip().lower().replace(' ', '').replace('(', '').replace(')', '').replace('_', '')
                    if candidate_clean in col_clean or col_clean in candidate_clean:
                        matched_columns[key] = col
                        #print(f"匹配成功: {key} -> {col}")
                        found = True
                        break
                if found:
                    break

        # 檢查必要列
        required_keys = ['flow']
        missing_keys = []
        for key in required_keys:
            if not matched_columns.get(key):
                missing_keys.append(key)

        if missing_keys:
            print(f"缺少必要數據列: {missing_keys}")
            messagebox.showwarning("警告", 
                                 f"未找到必需的流量數據列\n\n"
                                 f"偵測到的欄位: {', '.join(self.data.columns)}")
            return None

        #print(f"最終匹配結果: {matched_columns}")
        return matched_columns

    def plot_performance_curve(self):
        """繪製性能曲線 - 支援10條曲線"""
        try:
            #print("開始繪製性能曲線...")

            if self.data is None:
                messagebox.showwarning("警告", "請先載入Excel文件！")
                return

            #print(f"數據載入成功，共 {len(self.data)} 行")

            columns = self.get_data_columns()
            if columns is None:
                return

            #print(f"匹配到的數據列: {columns}")

            # 動態提取數據
            data_arrays = {}

            # 必需欄位 - 流量 (映射為 'x')
            if 'flow' in columns:
                Q = self.data[columns['flow']].values
                data_arrays['x'] = Q  # 這裡改為 'x'
            else:
                messagebox.showerror("錯誤", "未找到流量數據！")
                return

            # 擴展可選欄位（支援10種曲線）- 修正映射關係
            optional_columns = {
                'static_pressure': ('ps', 'Ps'),
                'total_pressure': ('pt', 'Pt'), 
                'power': ('h', 'H'),
                'efficiency_static': ('eff_s', 'Eta_s'),
                'efficiency_total': ('eff_t', 'Eta_t'),
                'speed': ('n', 'N'),
                'current': ('i', 'I'),
                'voltage': ('v', 'V'),
                'frequency': ('f', 'F'),
                'temperature': ('temp', 'Temp')
            }

            available_curves = []
            for data_key, (curve_key, data_array_key) in optional_columns.items():
                if data_key in columns:
                    values = self.data[columns[data_key]].values
                    if len(values) > 0 and not np.all(np.isnan(values)) and not np.all(values == 0):
                        # 只有當長度與 x 相同時才加入
                        if len(values) == len(Q):
                            data_arrays[data_array_key] = values
                            data_arrays[curve_key] = values
                            available_curves.append(curve_key)
                            print(f"載入 {data_key} -> {curve_key} 數據: {len(values)} 筆")
                        else:
                            print(f"警告：{data_key} 數據長度 ({len(values)}) 與流量 ({len(Q)}) 不一致，跳過")

            # 篩選出真正有數據的曲線（長度 > 0 且非全為 NaN）
            valid_curves = []
            for curve_key in available_curves:
                if curve_key in data_arrays and len(data_arrays[curve_key]) > 0:
                    if not np.all(np.isnan(data_arrays[curve_key])):
                        valid_curves.append(curve_key)

            #print(f"有效曲線鍵: {valid_curves}")

            # 計算預設軸範圍
            default_ranges = self.calculate_default_ranges(data_arrays)
            #print("準備顯示設定對話框...")

            # 彈出設定對話框
            dialog = PerformanceCurveConfigDialog(self.root, default_ranges, valid_curves)

            #print(f"設定對話框結果: {dialog.result}")

            if not dialog.result:
                print("用戶取消了設定")
                return

            #print("開始繪製圖表...")

            # 使用選定的配置繪製圖表
            axis_settings = dialog.result['axis_settings']
            axis_format_settings = axis_settings.get('axis_format_settings', {})

            # 新增：獲取圖例名稱
            legend_names = dialog.result.get('legend_names', {})

            fig = self.create_performance_chart_figure(
                data_arrays, 
                dialog.result['curves'], 
                dialog.result['axis_assignments'],
                dialog.result['axis_labels'],
                dialog.result['ranges'],
                dialog.result['curve_styles'],
                dialog.result['legend_settings'],
                dialog.result['axis_settings'],
                legend_names  # 新增：傳遞圖例名稱
            )
            if fig:
                self.display_chart_in_window(fig)
                print("圖表繪製完成")
            else:
                print("圖表創建失敗")

        except Exception as e:
            print(f"繪製過程中發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("錯誤", f"繪製性能曲線時發生錯誤：\n{str(e)}")


    def calculate_default_ranges(self, data_arrays):
        """動態計算預設軸範圍 - 支援10條曲線"""
        ranges = {
            'x': (0, 100),
            'y1': (0, 100), 'y2': (0, 100), 'y3': (0, 100), 'y4': (0, 100),
            'y5': (0, 100), 'y6': (0, 100), 'y7': (0, 100), 'y8': (0, 100), 'y9': (0, 100)
        }

        # 流量範圍 - 使用 'x' 鍵
        if 'x' in data_arrays:
            Q = data_arrays['x']
            ranges['x'] = (max(0, min(Q) * 0.9), max(Q) * 1.1)

        # 壓力範圍 (Y1軸) - 使用曲線鍵
        pressure_values = []
        for key in ['ps', 'pt']:  # 改為使用曲線鍵
            if key in data_arrays:
                pressure_values.extend(data_arrays[key])
        if pressure_values:
            ranges['y1'] = (max(0, min(pressure_values) * 0.9), max(pressure_values) * 1.1)

        # 功率範圍 (Y2軸)
        if 'h' in data_arrays:  # 改為使用曲線鍵
            H = data_arrays['h']
            ranges['y2'] = (max(0, min(H) * 0.9), max(H) * 1.1)

        # 效率範圍 (Y3軸)
        efficiency_values = []
        for key in ['eff_s', 'eff_t']:  # 改為使用曲線鍵
            if key in data_arrays:
                efficiency_values.extend(data_arrays[key])
        if efficiency_values:
            ranges['y3'] = (max(0, min(efficiency_values) * 0.9), min(100, max(efficiency_values) * 1.1))

        # 轉速範圍 (Y4軸)
        if 'n' in data_arrays:  # 改為使用曲線鍵
            N = data_arrays['n']
            ranges['y4'] = (max(0, min(N) * 0.9), max(N) * 1.1)

        # 電氣參數範圍
        electrical_keys = {
            'i': 'y5',  # 電流
            'v': 'y6',  # 電壓
            'f': 'y7',  # 頻率
            'temp': 'y8'  # 溫度
        }
    
        for key, axis_key in electrical_keys.items():
            if key in data_arrays:
                values = data_arrays[key]
                ranges[axis_key] = (max(0, min(values) * 0.9), max(values) * 1.1)

        return ranges


    def create_performance_chart_figure(self, data_arrays, curves, axis_assignments,
                                      axis_labels, ranges, curve_styles,
                                      legend_settings, axis_settings, legend_names=None):
        """創建性能曲線圖表（支援軸刻度格式自訂）- 完全修復版本"""
        try:
            # 初始化圖例名稱
            if legend_names is None:
                legend_names = {}
            
            # 從 axis_settings 中提取格式設定
            axis_format_settings = axis_settings.get('axis_format_settings', {})
            #print(f"📊 接收到的軸格式設定: {axis_format_settings}")  # 診斷信息

            # 檢查數據
            if not data_arrays or 'x' not in data_arrays:
                print("❌ 錯誤：缺少X軸數據")
                return None

            # 排序數據
            sorted_data = {}
            for key, values in data_arrays.items():
                if key == 'x':
                    sorted_indices = np.argsort(values)
                    sorted_data[key] = np.array(values)[sorted_indices]
                else:
                    sorted_data[key] = np.array(values)[sorted_indices] if key in data_arrays else np.array([])

            # 創建圖表和主軸
            fig, host = plt.subplots(figsize=(12, 8))
            fig.subplots_adjust(right=0.75)

            # 創建右側軸系
            axes = {'y1': host}
            for i in range(2, 10):
                axes[f'y{i}'] = host.twinx()

            # 調整右側軸的位置
            for i, axis_key in enumerate(['y2', 'y3', 'y4', 'y5', 'y6', 'y7', 'y8', 'y9']):
                if axis_key in axes:
                    offset = 1.0 + i * 0.1
                    axes[axis_key].spines['right'].set_position(('axes', offset))
                    axes[axis_key].set_frame_on(True)
                    axes[axis_key].patch.set_visible(False)

            # ========== 設定軸範圍 ==========
            # X軸範圍
            if 'x' in ranges:
                x_min, x_max = ranges['x']
                host.set_xlim(x_min, x_max)

            # Y軸範圍
            for axis_key in ['y1', 'y2', 'y3', 'y4', 'y5', 'y6', 'y7', 'y8', 'y9']:
                if axis_key in ranges and axis_key in axes:
                    y_min, y_max = ranges[axis_key]
                    axes[axis_key].set_ylim(y_min, y_max)

            # ========== 修復：軸刻度格式設定 - 完全重寫 ==========
            from matplotlib.ticker import FuncFormatter, ScalarFormatter, LogFormatter, MaxNLocator
            import matplotlib.ticker as ticker

            # 定義更可靠的工程記號格式化函數
            def engineering_formatter(x, pos):
                if x == 0:
                    return '0'
                try:
                    abs_x = abs(x)
                    if abs_x < 1e-100:
                        return '0'
                    exponent = np.floor(np.log10(abs_x))
                    engineering_exponent = int(np.floor(exponent / 3) * 3)
                    mantissa = x / (10 ** engineering_exponent)
                    if engineering_exponent == 0:
                        return f'{mantissa:.1f}'
                    else:
                        return f'{mantissa:.1f}×10$^{{{engineering_exponent}}}$'
                except:
                    return f'{x:.1f}'

            # 定義科學記號格式化函數
            def scientific_formatter(x, pos):
                """科學記號格式化 - 直接控制版本"""
                if x == 0:
                    return '0'
        
                try:
                    abs_x = abs(x)
                    if abs_x < 1e-100:
                        return '0'
            
                    exponent = np.floor(np.log10(abs_x))
                    mantissa = x / (10 ** exponent)
            
                    return f'{mantissa:.1f}×10$^{{{int(exponent)}}}$'
                except:
                    return f'{x:.1e}'

            #print(f"🔧 開始套用軸格式設定...")

            # 先收集所有使用的軸
            used_axes = set()
            for curve_key in curves:
                if curves[curve_key] and curve_key in axis_assignments:
                    axis_key = axis_assignments[curve_key]
                    used_axes.add(axis_key)

            # 對每個使用的軸套用格式
            for axis_key in ['x'] + [f'y{i}' for i in range(1, 10)]:
                if axis_key == 'x':
                    ax = host
                    axis_name = 'x'
                else:
                    if axis_key not in axes or axis_key not in used_axes:
                        continue
                    ax = axes[axis_key]
                    axis_name = 'y'

                fmt_type = axis_format_settings.get(axis_key, 'normal')
                print(f"  → {axis_key}: {fmt_type}")

                # 清除現有格式
                if axis_name == 'x':
                    ax.xaxis.set_major_formatter(ticker.ScalarFormatter())
                    ax.xaxis.set_minor_formatter(ticker.NullFormatter())
                else:
                    ax.yaxis.set_major_formatter(ticker.ScalarFormatter())
                    ax.yaxis.set_minor_formatter(ticker.NullFormatter())

                # 根據格式類型設定
                if fmt_type == 'scientific':
                    # 科學記號
                    formatter = FuncFormatter(scientific_formatter)
                    if axis_name == 'x':
                        ax.xaxis.set_major_formatter(formatter)
                        # 確保使用足夠的刻度
                        ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
                    else:
                        ax.yaxis.set_major_formatter(formatter)
                        ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
                    print(f"    ✅ 科學記號已套用至 {axis_key}")

                elif fmt_type == 'engineering':
                    # 工程記號
                    formatter = FuncFormatter(engineering_formatter)
                    if axis_name == 'x':
                        ax.xaxis.set_major_formatter(formatter)
                        ax.xaxis.set_major_locator(MaxNLocator(nbins=6))
                    else:
                        ax.yaxis.set_major_formatter(formatter)
                        ax.yaxis.set_major_locator(MaxNLocator(nbins=6))
                    print(f"    ✅ 工程記號已套用至 {axis_key}")

                elif fmt_type == 'log':
                    # 對數座標
                    try:
                        if axis_name == 'x':
                            # 檢查數據是否適合對數座標
                            x_data = sorted_data['x']
                            if np.all(x_data > 0):
                                ax.set_xscale('log')
                                # 使用對數格式化器
                                ax.xaxis.set_major_formatter(ticker.LogFormatterSciNotation())
                                print(f"    ✅ 對數座標已套用至 {axis_key}")
                            else:
                                print(f"    ⚠️  {axis_key} 包含非正數數據，跳過對數座標")
                        else:
                            # 找到對應的數據
                            for curve_key, assigned_axis in axis_assignments.items():
                                if assigned_axis == axis_key and curve_key in sorted_data:
                                    y_data = sorted_data[curve_key]
                                    if np.all(y_data > 0):
                                        ax.set_yscale('log')
                                        ax.yaxis.set_major_formatter(ticker.LogFormatterSciNotation())
                                        print(f"    ✅ 對數座標已套用至 {axis_key}")
                                        break
                            else:
                                print(f"    ⚠️  {axis_key} 無正數數據，跳過對數座標")
                    except Exception as e:
                        print(f"    ❌ 對數座標失敗: {e}")

            # 設定軸標籤和其他樣式
            host.set_xlabel(axis_labels.get('x', '流量 (CMM)'),
                           fontsize=axis_settings.get('x_label_fontsize', 12))

            if 'y1' in axis_labels:
                host.set_ylabel(axis_labels['y1'],
                               fontsize=axis_settings.get('y_label_fontsize', 12))

            # ========== 繪製所有曲線 ==========
            first_curve_for_axis = {}
            legend_handles = []
            legend_labels = []
            line_only_markers = {'+', 'x', '|', '_', 'X', '1', '2', '3', '4'}

            for curve_key in curves:
                if not curves[curve_key] or curve_key not in axis_assignments:
                    continue
        
                if curve_key not in data_arrays or len(data_arrays[curve_key]) == 0:
                    continue
        
                axis_key = axis_assignments[curve_key]
                style = curve_styles.get(curve_key, {})
    
                # 記錄每個軸的第一條曲線
                if axis_key not in first_curve_for_axis:
                    first_curve_for_axis[axis_key] = curve_key
    
                # 獲取數據
                x_data_sorted = sorted_data['x']
                y_data_sorted = sorted_data[curve_key]
    
                # 檢查數據有效性
                valid_mask = ~np.isnan(x_data_sorted) & ~np.isnan(y_data_sorted)
                x_valid = x_data_sorted[valid_mask]
                y_valid = y_data_sorted[valid_mask]
    
                if len(x_valid) < 2:
                    continue
        
                # 繪製樣式
                marker_style = style.get('marker', 'o')
                marker_size = style.get('markersize', 6)
                line_style = style.get('line_style', '-')
                line_width = style.get('linewidth', 2.0)
                color = style.get('color', 'blue')
    
                # 繪製數據點
                if marker_style in line_only_markers:
                    axes[axis_key].scatter(
                        x_valid, y_valid,
                        marker=marker_style, s=marker_size**2,
                        color=color, alpha=0.6, zorder=3
                    )
                elif style.get('marker_filled', True):
                    axes[axis_key].scatter(
                        x_valid, y_valid,
                        marker=marker_style, s=marker_size**2,
                        facecolor=color, edgecolor=color,
                        alpha=0.6, zorder=3
                    )
                else:
                    axes[axis_key].scatter(
                        x_valid, y_valid,
                        marker=marker_style, s=marker_size**2,
                        facecolor='none', edgecolor=color,
                        alpha=0.6, zorder=3
                    )
        
                # 繪製平滑曲線
                if len(x_valid) >= 4:
                    try:
                        x_smooth = np.linspace(min(x_valid), max(x_valid), 300)
                        cs = CubicSpline(x_valid, y_valid)
                        y_smooth = cs(x_smooth)
                        line = axes[axis_key].plot(
                            x_smooth, y_smooth,
                            color=color, linestyle=line_style,
                            linewidth=line_width, zorder=2
                        )[0]
                    except Exception:
                        line = axes[axis_key].plot(
                            x_valid, y_valid,
                            color=color, linestyle=line_style,
                            linewidth=line_width, zorder=2
                        )[0]
                else:
                    line = axes[axis_key].plot(
                        x_valid, y_valid,
                        color=color, linestyle=line_style,
                        linewidth=line_width, zorder=2
                    )[0]
        
                # 創建圖例
                from matplotlib.lines import Line2D
            
                # 使用自定義圖例名稱或默認名稱
                if curve_key in legend_names:
                    display_name = legend_names[curve_key]
                else:
                    display_name = self.get_curve_display_name(curve_key)
                
                legend_line = Line2D([0], [0],
                                   color=color, linestyle=line_style,
                                   linewidth=line_width, marker=marker_style,
                                   markersize=marker_size,
                                   markerfacecolor=color if style.get('marker_filled', True) else 'none',
                                   markeredgecolor=color,
                                   label=display_name)  # 使用自定義名稱
                legend_handles.append(legend_line)
                legend_labels.append(display_name)
    
                # 設定右側軸顏色
                if axis_key != 'y1' and first_curve_for_axis[axis_key] == curve_key:
                    axes[axis_key].tick_params(axis='y', colors=color)
                    if axis_key in axis_labels:
                        axes[axis_key].set_ylabel(axis_labels[axis_key],
                                                color=color,
                                                fontsize=axis_settings.get('y_label_fontsize', 12))

            # 隱藏未使用的軸
            for axis_key in ['y2', 'y3', 'y4', 'y5', 'y6', 'y7', 'y8', 'y9']:
                if axis_key in axes and axis_key not in used_axes:
                    axes[axis_key].spines['right'].set_visible(False)
                    axes[axis_key].tick_params(axis='y', which='both', 
                                             left=False, right=False, 
                                             labelleft=False, labelright=False)

            # 格線和標題
            if axis_settings.get('show_grid', True):
                host.grid(True, linestyle=axis_settings.get('grid_style', '--'),
                         alpha=axis_settings.get('grid_alpha', 0.3))

            if axis_settings.get('show_title', True):
                host.set_title(axis_settings.get('title_text', '性能曲線'),
                              fontsize=axis_settings.get('title_fontsize', 14),
                              fontweight=axis_settings.get('title_fontweight', 'bold'))

            # 圖例
            if legend_handles:
                location = legend_settings.get('location', 'upper center')
                if location != 'none':
                    vertical_offset = legend_settings.get('vertical_offset', 0.95)
                    if location == 'upper center':
                        legend = host.legend(legend_handles, legend_labels,
                                           loc='upper center',
                                           bbox_to_anchor=(0.5, vertical_offset),
                                           ncol=min(5, len(legend_handles)),
                                           frameon=legend_settings.get('show_frame', True),
                                           fontsize=legend_settings.get('fontsize', 10))
                    else:
                        legend = host.legend(legend_handles, legend_labels,
                                           loc=location,
                                           frameon=legend_settings.get('show_frame', True),
                                           fontsize=legend_settings.get('fontsize', 10))
                    if legend_settings.get('show_frame', True):
                        legend.get_frame().set_alpha(legend_settings.get('alpha', 0.9))

            plt.tight_layout()
    
            # 最終強制重繪
            fig.canvas.draw()
            print(f"🎉 圖表繪製完成")
            return fig

        except Exception as e:
            print(f"❌ 創建圖表時發生錯誤: {e}")
            import traceback
            traceback.print_exc()
            return None


    def get_curve_display_name(self, curve_key):
        """根据曲線鍵獲取顯示名稱"""
        display_names = {
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
        return display_names.get(curve_key, curve_key.upper())


    def add_text_to_chart(self, fig, canvas):
        """添加文字到圖表 - 使用螢幕座標"""
        try:
            text_dialog = tk.Toplevel(self.current_window)
            text_dialog.title("添加文字")
            text_dialog.geometry("350x280")
            text_dialog.transient(self.current_window)
            text_dialog.grab_set()
            main_frame = tk.Frame(text_dialog, padx=20, pady=20)
            main_frame.pack(fill=tk.BOTH, expand=True)
            tk.Label(main_frame, text="文字內容:", font=("Microsoft JhengHei", 10)).pack(anchor="w", pady=5)
            text_var = tk.StringVar(value="請輸入文字")
            text_entry = tk.Entry(main_frame, textvariable=text_var, width=30, font=("Microsoft JhengHei", 10))
            text_entry.pack(fill=tk.X, pady=5)
            text_entry.select_range(0, tk.END)
            text_entry.focus_set()
        
            font_frame = tk.Frame(main_frame)
            font_frame.pack(fill=tk.X, pady=5)
            tk.Label(font_frame, text="字體大小:", width=10).pack(side=tk.LEFT)
            font_size_var = tk.IntVar(value=12)
            tk.Spinbox(font_frame, from_=8, to=24, textvariable=font_size_var, width=5).pack(side=tk.LEFT, padx=5)
        
            color_frame = tk.Frame(main_frame)
            color_frame.pack(fill=tk.X, pady=5)
            tk.Label(color_frame, text="文字顏色:", width=10).pack(side=tk.LEFT)
            color_var = tk.StringVar(value="black")
            ttk.Combobox(color_frame, textvariable=color_var,
                        values=['black','red','blue','green','purple','orange','brown','gray','darkred','darkblue'],
                        state="readonly", width=10).pack(side=tk.LEFT, padx=5)
        
            bg_frame = tk.Frame(main_frame)
            bg_frame.pack(fill=tk.X, pady=5)
            tk.Label(bg_frame, text="底色:", width=10).pack(side=tk.LEFT)
            bg_color_var = tk.StringVar(value="lightyellow")
            ttk.Combobox(bg_frame, textvariable=bg_color_var,
                        values=['lightyellow','white','lightblue','lightgreen','pink','lavender','none'],
                        state="readonly", width=10).pack(side=tk.LEFT, padx=5)
        
            edge_frame = tk.Frame(main_frame)
            edge_frame.pack(fill=tk.X, pady=5)
            tk.Label(edge_frame, text="邊框:", width=10).pack(side=tk.LEFT)
            edge_color_var = tk.StringVar(value="gray")
            ttk.Combobox(edge_frame, textvariable=edge_color_var,
                        values=['gray','black','blue','red','green','none'],
                        state="readonly", width=10).pack(side=tk.LEFT, padx=5)
        
            def on_ok():
                text_content = text_var.get().strip()
                if not text_content:
                    return
            
                ax = fig.axes[0]
            
                bbox_dict = None
                if bg_color_var.get() != 'none':
                    edgecolor = None if edge_color_var.get() == 'none' else edge_color_var.get()
                    bbox_dict = dict(
                        boxstyle="round,pad=0.3",
                        facecolor=bg_color_var.get(),
                        alpha=0.7,
                        edgecolor=edgecolor
                    )
            
                # 使用 figure 座標系（0-1 範圍），預設在圖表中心
                text_obj = fig.text(
                    0.5, 0.5, text_content,
                    fontsize=font_size_var.get(),
                    color=color_var.get(),
                    ha='center', va='center',
                    bbox=bbox_dict,
                    picker=True,
                    transform=fig.transFigure  # 使用 figure 座標系
                )
            
                text_obj.original_bbox = bbox_dict
                self.text_objects.append(text_obj)
                canvas.draw()
                text_dialog.destroy()
        
            btn_frame = tk.Frame(main_frame)
            btn_frame.pack(pady=10)
            tk.Button(btn_frame, text="確定", command=on_ok, bg="#32CD32", fg="white", width=10).pack(side=tk.LEFT, padx=5)
            tk.Button(btn_frame, text="取消", command=text_dialog.destroy, bg="#FF6B6B", fg="white", width=10).pack(side=tk.LEFT, padx=5)
            text_dialog.bind('<Return>', lambda e: on_ok())
        except Exception as e:
            messagebox.showerror("錯誤", f"添加文字失敗: {str(e)}")

    def on_chart_press(self, event):
        """處理滑鼠按下 - 使用像素座標"""
        if event.button != 1:
            self.dragging_text = None
            return
    
        for text_obj in self.text_objects:
            try:
                contains, _ = text_obj.contains(event)
                if contains:
                    self.dragging_text = text_obj
                
                    # 記錄像素座標
                    self.drag_start_x = event.x
                    self.drag_start_y = event.y
                
                    # 獲取當前文字的 figure 座標
                    self.text_start_pos = text_obj.get_position()
                
                    # 高亮
                    text_obj.set_bbox(dict(
                        boxstyle="round,pad=0.3", 
                        facecolor='lightblue', 
                        alpha=0.7, 
                        edgecolor='blue',
                        linewidth=2
                    ))
                    self.current_canvas.draw()
                    return
            except:
                continue
    
        self.dragging_text = None

    def on_chart_drag(self, event):
        """拖曳文字 - 使用像素座標"""
        if self.dragging_text is None:
            return
    
        if event.x is None or event.y is None:
            return
    
        try:
            # 計算像素位移
            dx_pixels = event.x - self.drag_start_x
            dy_pixels = event.y - self.drag_start_y
        
            # 轉換為 figure 座標（0-1 範圍）
            fig_width = self.current_fig.get_figwidth() * self.current_fig.dpi
            fig_height = self.current_fig.get_figheight() * self.current_fig.dpi
        
            dx_fig = dx_pixels / fig_width
            dy_fig = -dy_pixels / fig_height  # Y 軸反向
        
            # 計算新位置
            new_x = self.text_start_pos[0] + dx_fig
            new_y = self.text_start_pos[1] + dy_fig
        
            # 限制在圖表範圍內
            new_x = max(0, min(1, new_x))
            new_y = max(0, min(1, new_y))
        
            self.dragging_text.set_position((new_x, new_y))
            self.current_canvas.draw_idle()
        
        except Exception as e:
            print(f"拖曳錯誤: {e}")

    def setup_chart_interaction(self, fig, canvas):
        """設置圖表交互功能"""
        self.dragging_text = None
        self.text_objects = []
        self.drag_start_x = 0
        self.drag_start_y = 0
        self.text_start_pos = (0, 0)
    
        canvas.mpl_connect('button_press_event', self.on_chart_press)
        canvas.mpl_connect('motion_notify_event', self.on_chart_drag)
        canvas.mpl_connect('button_release_event', self.on_chart_release)
    
        print("✅ 圖表交互功能已啟用")


    def on_chart_release(self, event):
        """釋放文字"""
        if self.dragging_text is not None:
            print(f"✅ 釋放文字: {self.dragging_text.get_text()}")
        
            # 恢復原始樣式
            original_bbox = getattr(self.dragging_text, 'original_bbox', None)
            self.dragging_text.set_bbox(original_bbox)
            self.dragging_text = None
            self.current_canvas.draw()


    def display_chart_in_window(self, fig):
        """在窗口中顯示圖表"""
        if fig is None:
            messagebox.showerror("错誤", "圖表創建失敗，無法顯示")
            return
        
        chart_window = tk.Toplevel(self.root)
        chart_window.title("性能曲線圖 - 10曲線專業版")
        chart_window.geometry("1400x900")

        # 頂部按鈕框架
        top_frame = tk.Frame(chart_window, bg="#F0F0F0", height=60)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        top_frame.pack_propagate(False)

        # 左侧按钮群组
        left_button_frame = tk.Frame(top_frame, bg="#F0F0F0")
        left_button_frame.pack(side=tk.LEFT, padx=15, pady=10)

        # 保存按鈕
        tk.Button(left_button_frame, text="💾 保存圖表", 
                 command=lambda: self.save_performance_chart(fig),
                 bg="#4169E1", fg="white", 
                 font=("Microsoft JhengHei", 10, "bold"),
                 width=12, height=1,
                 relief="raised", bd=2).pack(side=tk.LEFT, padx=5)

        # 複製到剪貼簿按鈕
        tk.Button(left_button_frame, text="📋 複製到剪貼簿", 
                 command=lambda: self.copy_to_clipboard(fig),
                 bg="#32CD32", fg="white", 
                 font=("Microsoft JhengHei", 10, "bold"),
                 width=15, height=1,
                 relief="raised", bd=2).pack(side=tk.LEFT, padx=5)

        # 添加文字按鈕
        tk.Button(left_button_frame, text="📝 添加文字", 
                 command=lambda: self.add_text_to_chart(fig, canvas),
                 bg="#FF69B4", fg="white", 
                 font=("Microsoft JhengHei", 10, "bold"),
                 width=12, height=1,
                 relief="raised", bd=2).pack(side=tk.LEFT, padx=5)

        # 重新繪製按鈕
        tk.Button(left_button_frame, text="🔄 重新設定", 
                 command=lambda: [chart_window.destroy(), self.plot_performance_curve()],
                 bg="#FF8C00", fg="white", 
                 font=("Microsoft JhengHei", 10, "bold"),
                 width=12, height=1,
                 relief="raised", bd=2).pack(side=tk.LEFT, padx=5)

        # 標題標籤
        title_label = tk.Label(top_frame, text="性能曲線圖 (10曲線專業版)", 
                font=("Microsoft JhengHei", 14, "bold"),
                bg="#F0F0F0", fg="#2E8B57")
        title_label.pack(side=tk.LEFT, padx=20, expand=True)

        # 嵌入matplotlib圖表
        canvas = FigureCanvasTkAgg(fig, master=chart_window)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
        # 添加缩放控制
        toolbar_frame = tk.Frame(chart_window)
        toolbar_frame.pack(side=tk.BOTTOM, fill=tk.X)
    
        from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
        toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
        toolbar.update()
    
        # 存儲圖表和畫布的引用
        self.current_fig = fig
        self.current_canvas = canvas
        self.current_window = chart_window
    
        # 後用圖表交互
        self.setup_chart_interaction(fig, canvas)




    def on_text_pick(self, event):
        """處理文字拾取事件 - 修複版本"""
        if event.artist in self.text_objects:
            self.dragging_text = event.artist
            # 存储点击时的文字位置和鼠标位置
            self.text_start_x, self.text_start_y = self.dragging_text.get_position()
            self.drag_start_x = event.mouseevent.xdata
            self.drag_start_y = event.mouseevent.ydata
        
            # 高亮选中的文字
            self.dragging_text.set_bbox(dict(boxstyle="round,pad=0.3", 
                                           facecolor='lightblue', 
                                           alpha=0.7,
                                           edgecolor='blue'))
            self.current_canvas.draw()
            return True
        return False




    def on_chart_click(self, event, fig):
        """處理圖表點擊事件"""
        if event.inaxes is None:
            return
    
        # 檢查是否點擊了文字對象
        if event.inaxes == fig.axes[0]:
            # 檢查所有文字對象
            for text_obj in self.text_objects:
                contains, _ = text_obj.contains(event)
                if contains:
                    self.dragging_text = text_obj
                    # 存儲點擊時的文字位置和鼠標位置
                    self.text_start_x, self.text_start_y = text_obj.get_position()
                    self.drag_start_x = event.xdata
                    self.drag_start_y = event.ydata
                    # 高亮選中的文字
                    text_obj.set_bbox(dict(boxstyle="round,pad=0.3", 
                                         facecolor='lightblue', 
                                         alpha=0.7,
                                         edgecolor='blue'))
                    self.current_canvas.draw()
                    break


    def save_performance_chart(self, fig):
        """保存性能圖表"""
        try:
            size_dialog = ImageSizeDialog(self.root)
            if not size_dialog.result:
                return
    
            width, height = size_dialog.result
    
            file_path = filedialog.asksaveasfilename(
                title="保存性能圖表",
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf"), 
                          ("JPG files", "*.jpg"), ("SVG files", "*.svg"),
                          ("All files", "*.*")],
                initialfile="性能曲線_10曲線"
            )
    
            if file_path:
                original_size = fig.get_size_inches()
                fig.set_size_inches(width/100, height/100)
        
                if file_path.lower().endswith('.pdf'):
                    fig.savefig(file_path, dpi=300, bbox_inches='tight', format='pdf')
                elif file_path.lower().endswith('.jpg') or file_path.lower().endswith('.jpeg'):
                    fig.savefig(file_path, dpi=300, bbox_inches='tight', format='jpg', quality=95)
                elif file_path.lower().endswith('.svg'):
                    fig.savefig(file_path, dpi=300, bbox_inches='tight', format='svg')
                else:
                    fig.savefig(file_path, dpi=300, bbox_inches='tight', format='png')
            
                # 恢复原始尺寸
                fig.set_size_inches(original_size)
                messagebox.showinfo("成功", f"圖表已保存到:\n{file_path}\n尺寸: {width}x{height} pixels")
        
        except Exception as e:
            messagebox.showerror("錯誤", f"保存失敗:\n{str(e)}")

    def copy_to_clipboard(self, fig):
        """複製圖表到鍵貼簿"""
        try:
            size_dialog = ImageSizeDialog(self.root)
            if not size_dialog.result:
                return
    
            width, height = size_dialog.result
    
            original_size = fig.get_size_inches()
            fig.set_size_inches(width/100, height/100)
    
            import io
            import win32clipboard
            from PIL import Image
    
            buf = io.BytesIO()
            fig.savefig(buf, format='png', dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
            buf.seek(0)
    
            image = Image.open(buf)
            output = io.BytesIO()
            image.save(output, 'BMP')
            data = output.getvalue()[14:]
    
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
    
            fig.set_size_inches(original_size)
            buf.close()
            output.close()
    
            messagebox.showinfo("成功", f"圖表以複製到鍵貼簿！\n\n尺寸: {width}x{height} pixels")
    
        except ImportError:
            self.fallback_copy_to_clipboard(fig)
        except Exception as e:
            messagebox.showerror("錯誤", f"複製到新剪貼簿失敗：{str(e)}")

    def fallback_copy_to_clipboard(self, fig):
        """備用方法：保存到暫存档案"""
        try:
            import tempfile
            import os
            import subprocess
    
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, "performance_chart_10curves.png")
    
            fig.savefig(temp_file, dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
    
            try:
                if os.name == 'nt':
                    os.startfile(temp_file)
                elif os.name == 'posix':
                    subprocess.run(['xdg-open', temp_file])
            except:
                pass
    
            messagebox.showinfo("複製圖表", 
                              f"圖表已保存到暫存档案：\n{temp_file}\n\n"
                              "請手動複制：\n"
                              "1. 開啟該圖片檔案\n"
                              "2. 按 Ctrl+A (全選)\n"
                              "3. 按 Ctrl+C (複製)\n"
                              "4. 在目標應中按 Ctrl+V (贴上)")
    
        except Exception as e:
            messagebox.showerror("錯誤", f"無法複製圖表：{str(e)}")


    def fallback_copy_to_clipboard(self, fig):
        """備用方法：保存到暫存檔案"""
        try:
            import tempfile
            import os
            import subprocess
        
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, "performance_chart_10curves.png")
        
            fig.savefig(temp_file, dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
        
            try:
                if os.name == 'nt':
                    os.startfile(temp_file)
                elif os.name == 'posix':
                    subprocess.run(['xdg-open', temp_file])
            except:
                pass
        
            messagebox.showinfo("複製圖表", 
                              f"圖表已保存到暫存檔案：\n{temp_file}\n\n"
                              "請手動複製：\n"
                              "1. 開啟該圖片檔案\n"
                              "2. 按 Ctrl+A (全選)\n"
                              "3. 按 Ctrl+C (複製)\n"
                              "4. 在目標應用中按 Ctrl+V (貼上)")
        
        except Exception as e:
            messagebox.showerror("錯誤", f"無法複製圖表：{str(e)}")


    def run(self):
        """啟動應用"""
        self.root.mainloop()

# ImageSizeDialog 類保持不變
class ImageSizeDialog:
    def __init__(self, parent):
        self.parent = parent
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("選擇圖片尺寸")
        self.dialog.geometry("400x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.create_dialog()
    
    def create_dialog(self):
        main_frame = tk.Frame(self.dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="選擇圖片儲存尺寸", 
                font=("Microsoft JhengHei", 12, "bold")).pack(pady=(0, 15))
        
        size_frame = ttk.LabelFrame(main_frame, text="預設尺寸 (16:10 比例)", padding=10)
        size_frame.pack(fill=tk.X, pady=5)
        
        self.size_var = tk.StringVar(value="1280x800")
        
        sizes = [
            ("1920x1200 (16:10) - 高解析度", "1920x1200"),
            ("1280x800 (16:10) - 標準", "1280x800"),
            ("640x400 (16:10) - 小尺寸", "640x400")
        ]
        
        for text, value in sizes:
            tk.Radiobutton(size_frame, text=text, variable=self.size_var, 
                          value=value, font=("Microsoft JhengHei", 9)).pack(anchor="w", pady=2)
        
        custom_frame = ttk.LabelFrame(main_frame, text="自訂尺寸", padding=10)
        custom_frame.pack(fill=tk.X, pady=5)
        
        width_frame = tk.Frame(custom_frame)
        width_frame.pack(fill=tk.X, pady=2)
        tk.Label(width_frame, text="寬度:", width=8).pack(side=tk.LEFT)
        self.custom_width_var = tk.StringVar(value="1280")
        tk.Entry(width_frame, textvariable=self.custom_width_var, width=10).pack(side=tk.LEFT, padx=5)
        
        height_frame = tk.Frame(custom_frame)
        height_frame.pack(fill=tk.X, pady=2)
        tk.Label(height_frame, text="高度:", width=8).pack(side=tk.LEFT)
        self.custom_height_var = tk.StringVar(value="800")
        tk.Entry(height_frame, textvariable=self.custom_height_var, width=10).pack(side=tk.LEFT, padx=5)
        
        self.use_custom_var = tk.BooleanVar(value=False)
        custom_cb = tk.Checkbutton(custom_frame, text="使用自訂尺寸",
                                 variable=self.use_custom_var,
                                 font=("Microsoft JhengHei", 9))
        custom_cb.pack(anchor="w", pady=5)
        
        self.lock_aspect_var = tk.BooleanVar(value=True)
        lock_cb = tk.Checkbutton(custom_frame, text="鎖定 16:10 比例",
                               variable=self.lock_aspect_var,
                               font=("Microsoft JhengHei", 9))
        lock_cb.pack(anchor="w", pady=2)
        
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="確定", command=self.on_ok, 
                 bg="#32CD32", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=self.on_cancel,
                 bg="#FF6B6B", fg="white", width=10).pack(side=tk.LEFT, padx=5)
        
        self.parent.wait_window(self.dialog)
    
    def on_ok(self):
        try:
            if self.use_custom_var.get():
                width = int(self.custom_width_var.get())
                height = int(self.custom_height_var.get())
                
                if width <= 0 or height <= 0:
                    messagebox.showerror("錯誤", "請輸入有效的尺寸數值！")
                    return
                    
                self.result = (width, height)
            else:
                size_str = self.size_var.get()
                width, height = map(int, size_str.split('x'))
                self.result = (width, height)
                
            self.dialog.destroy()
            
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數字！")
    
    def on_cancel(self):
        self.result = None
        self.dialog.destroy()

if __name__ == "__main__":
    app = PerformanceCurvePlotter()
    app.run()