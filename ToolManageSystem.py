import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
import datetime
import subprocess
import tkinter.font as tkFont

class ToolManagementApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("工具库管理系统")
        self.geometry("900x700")
    
        self.conn = pyodbc.connect(f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ=D:\\ToolManagementSystem\\Database1.accdb;")
        self.cursor = self.conn.cursor()

        # 设置较大的默认字体
        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(size=12)  # 调整字体大小
        self.option_add("*Font", default_font)
        
        # 创建主布局的框架
        main_frame = tk.Frame(self)
        main_frame.grid(row=0, column=0, sticky='nsew')
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=3)
        main_frame.grid_columnconfigure(1, weight=1)

        # 创建标签页
        self.tab_control = ttk.Notebook(main_frame)
        
        self.borrow_tab = ttk.Frame(self.tab_control)
        self.return_tab = ttk.Frame(self.tab_control)
        self.add_tool_tab = ttk.Frame(self.tab_control)
        self.scrap_tool_tab = ttk.Frame(self.tab_control)

        # 添加标签页
        self.tab_control.add(self.borrow_tab, text='从工具库中借出工具')
        self.tab_control.add(self.return_tab, text='向工具库归还工具')
        self.tab_control.add(self.add_tool_tab, text='登记新工具入库')
        self.tab_control.add(self.scrap_tool_tab, text='登记工具报废')

        # 使用 grid 方法放置 tab_control
        self.tab_control.grid(row=0, column=0, sticky='nsew')

        # 创建侧边栏：不设置宽度，使之自适应大小
        sidebar = tk.Frame(main_frame)
        sidebar.grid(row=0, column=1, sticky='nsew')
        sidebar.grid_rowconfigure(0, weight=1) # 允许垂直方向拉伸
        sidebar.grid_columnconfigure(0, weight=1) # 允许水平方向拉伸

        # 在侧边栏中添加搜索栏
        tk.Label(sidebar, text='搜索工具：').pack(fill='x')
        self.search_tool_entry = tk.Entry(sidebar)
        self.search_tool_entry.pack(fill='x')
        search_tool_button = tk.Button(sidebar, text='搜索', command=self.search_tool)
        search_tool_button.pack(fill='x')

        # 在侧边栏中创建用于显示搜索结果的 Treeview
        self.search_results_tree = ttk.Treeview(sidebar, columns=("工具编号", "工具名称", "位置", "状态"))
        self.search_results_tree.heading("#1", text="工具编号")
        self.search_results_tree.heading("#2", text="工具名称")
        self.search_results_tree.heading("#3", text="位置")
        self.search_results_tree.heading("#4", text="状态")
        self.search_results_tree.pack(fill='both', expand=True)

        # 快捷操作-绑定双击事件
        self.search_results_tree.bind("<Double-1>", self.on_select_tool)

        # 配置Treeview的列宽度
        self.search_results_tree.column("#0", width=0, stretch=tk.NO)
        self.search_results_tree.column("#1", width=60, stretch=tk.NO)  # 调整列宽度
        self.search_results_tree.column("#2", width=120, stretch=tk.NO)
        self.search_results_tree.column("#3", width=60, stretch=tk.NO)
        self.search_results_tree.column("#4", width=60, stretch=tk.NO)

        # 分页控件
        self.page_number = 0 # 初始化页码为0
        self.results_per_page = 15 # 每页显示10条记录
        self.search_results = [] # 存储搜索结果

        paging_frame = tk.Frame(sidebar)
        paging_frame.pack(fill='x')
        self.prev_page_button = tk.Button(paging_frame, text='上一页', command=self.prev_page)
        self.prev_page_button.pack(side='left')
        self.next_page_button = tk.Button(paging_frame, text='下一页', command=self.next_page)
        self.next_page_button.pack(side='right')

        # 各标签页的用户界面
        self.borrow_tool_ui()
        self.return_tool_ui()
        self.add_tool_ui()
        self.scrap_tool_ui()

        # 创建按钮来打开数据与统计窗口
        open_data_window_button = tk.Button(self, text="--------------------打开数据窗口---------------------", command=self.open_data_window)
        
        # 使用 grid 方法而不是 pack 方法来放置按钮
        open_data_window_button.grid(row=1, column=0)  # 适当调整 row 和 column 的值

        # 绑定窗口关闭事件
        self.protocol("WM_DELETE_WINDOW", self.on_window_close)

    def open_data_window(self):
        data_window = DataWindow(self, self.conn)
    
    def add_tool_ui(self):
        tk.Label(self.add_tool_tab, text='工具名称：').grid(column=0, row=0)
        self.add_tool_name = tk.Entry(self.add_tool_tab)
        self.add_tool_name.grid(column=1, row=0)
        
        tk.Label(self.add_tool_tab, text='位置：').grid(column=0, row=1)
        self.add_location = tk.Entry(self.add_tool_tab)
        self.add_location.grid(column=1, row=1)
        
        tk.Label(self.add_tool_tab, text='增加原因：').grid(column=0, row=2)
        self.add_reasons_combobox = ttk.Combobox(self.add_tool_tab, values=["新购入", "修理后回库", "操作错误后重新登记", "其它，填写在备注中"])
        self.add_reasons_combobox.grid(column=1, row=2)
        self.add_reasons_combobox.current(0)
        
        tk.Label(self.add_tool_tab, text='备注：').grid(column=0, row=3)
        self.add_notes = tk.Entry(self.add_tool_tab)
        self.add_notes.grid(column=1, row=3)
        
        add_tool_button = tk.Button(self.add_tool_tab, text='确定增加', command=self.adjust_inventory)
        add_tool_button.grid(column=1, row=4)

        tk.Label(self.add_tool_tab, text='----↓下方为同名工具批量入库操作区域↓----').grid(column=0, row=5)

        tk.Label(self.add_tool_tab, text='工具名称：').grid(column=0, row=6)
        self.tool_name_entry = tk.Entry(self.add_tool_tab)
        self.tool_name_entry.grid(column=1, row=6)

        tk.Label(self.add_tool_tab, text='这次批量操作的工具被放置在同一个位置还是同一层：').grid(column=0, row=7)
        self.tool_change_in_stock_combobox = ttk.Combobox(self.add_tool_tab, values=["同一位置", "同一层"])
        self.tool_change_in_stock_combobox.grid(column=1, row=7)
        self.tool_change_in_stock_combobox.current(0)

        tk.Label(self.add_tool_tab, text='本次操作入库数量：').grid(column=0, row=8)
        self.tool_num_entry = tk.Entry(self.add_tool_tab)
        self.tool_num_entry.grid(column=1, row=8)

        tk.Label(self.add_tool_tab, text='首个工具位置：').grid(column=0, row=9)
        self.tool_location_1st_entry = tk.Entry(self.add_tool_tab)
        self.tool_location_1st_entry.grid(column=1, row=9)
        
        bulk_add_tool_button = tk.Button(self.add_tool_tab, text='批量添加工具', command=self.bulk_insert)
        bulk_add_tool_button.grid(column=1, row=10)

    def scrap_tool_ui(self):
        tk.Label(self.scrap_tool_tab, text='工具编号（输入编号按下回车将会显示工具名称）：').grid(column=0, row=0)
        self.scrap_tool_id = tk.Entry(self.scrap_tool_tab)
        self.scrap_tool_id.grid(column=1, row=0)
        self.scrap_tool_id.bind("<Return>", lambda event: self.fetch_tool_name(self.scrap_tool_id, self.scrap_tool_name_label))

        tk.Label(self.scrap_tool_tab, text='工具名称：').grid(column=0, row=1)
        self.scrap_tool_name_label = tk.Label(self.scrap_tool_tab, text="")
        self.scrap_tool_name_label.grid(column=1, row=1)

        tk.Label(self.scrap_tool_tab, text='备注：').grid(column=0, row=2)
        self.scrap_notes = tk.Entry(self.scrap_tool_tab)
        self.scrap_notes.grid(column=1, row=2)

        scrap_tool_button = tk.Button(self.scrap_tool_tab, text='确定报废', command=self.scrap_tool)
        scrap_tool_button.grid(column=1, row=3)
    
    def borrow_tool_ui(self):
        tk.Label(self.borrow_tab, text='工具编号（输入编号按下回车将会显示工具名称）：').grid(column=0, row=0)
        self.borrow_tool_id = tk.Entry(self.borrow_tab)
        self.borrow_tool_id.grid(column=1, row=0)
        self.borrow_tool_id.bind("<Return>", lambda event: self.fetch_tool_name(self.borrow_tool_id, self.borrow_tool_name_label))
 
        tk.Label(self.borrow_tab, text='工具名称：').grid(column=0, row=1)
        self.borrow_tool_name_label = tk.Label(self.borrow_tab, text="")  # 将 Entry 改为 Label
        self.borrow_tool_name_label.grid(column=1, row=1)
  
        tk.Label(self.borrow_tab, text='借出者：').grid(column=0, row=2)
        self.borrow_borrower = tk.Entry(self.borrow_tab)
        self.borrow_borrower.grid(column=1, row=2)

        # 借出者部门
        tk.Label(self.borrow_tab, text='借出者部门：').grid(column=0, row=3)
        self.borrow_department = tk.Entry(self.borrow_tab)
        self.borrow_department.grid(column=1, row=3)

        # 使用天数输入框
        tk.Label(self.borrow_tab, text='使用天数：').grid(column=0, row=5)
        self.use_days = tk.Entry(self.borrow_tab)
        self.use_days.grid(column=1, row=5)
        self.use_days.bind("<KeyRelease>", self.calculate_due_date)  # 绑定事件

        
        # 归还日期输入框
        tk.Label(self.borrow_tab, text='预期归还日期（yyyy/mm/dd）：').grid(column=0, row=6)
        self.expected_return_date = tk.Entry(self.borrow_tab)
        self.expected_return_date.grid(column=1, row=6)
        
        # 确定借出按钮
        borrow_button = tk.Button(self.borrow_tab, text='确定借出', command=self.borrow_tool)
        borrow_button.grid(column=1, row=7)

        # 创建一个Treeview来显示借出者和所属部门信息
        self.borrower_department_tree = ttk.Treeview(self.borrow_tab, columns=("借出者", "部门"))
        self.borrower_department_tree.heading("#1", text="借出者")
        self.borrower_department_tree.heading("#2", text="部门")
        self.borrower_department_tree.grid(column=0, row=10, columnspan=2, sticky='nsew')
        self.borrower_department_tree.bind("<Double-1>", self.on_select_borrower_department)

        # 配置列宽度
        self.borrower_department_tree.column("#0", width=0, stretch=tk.NO)
        self.borrower_department_tree.column("#1", width=100)
        self.borrower_department_tree.column("#2", width=100)

        # 绑定双击事件
        self.borrower_department_tree.bind("<Double-1>", self.on_select_borrower_department)

        # 创建并配置滚动条
        borrower_dept_scrollbar = ttk.Scrollbar(self.borrow_tab, orient="vertical", command=self.borrower_department_tree.yview)
        self.borrower_department_tree.configure(yscrollcommand=borrower_dept_scrollbar.set)
        borrower_dept_scrollbar.grid(column=2, row=10, sticky='ns')

        # 初始化 Treeview 内容
        self.fill_borrower_department_tree()

    def on_select_borrower_department(self, event):
        selected_item = self.borrower_department_tree.item(self.borrower_department_tree.selection())
        borrower, department = selected_item['values']
        self.borrow_borrower.delete(0, tk.END)
        self.borrow_borrower.insert(0, borrower)
        self.borrow_department.delete(0, tk.END)
        self.borrow_department.insert(0, department)


    def return_tool_ui(self):
        tk.Label(self.return_tab, text='工具编号（输入编号按下回车将会显示工具名称）：').grid(column=0, row=0)
        self.return_tool_id = tk.Entry(self.return_tab)
        self.return_tool_id.grid(column=1, row=0)
        self.return_tool_id.bind("<Return>", lambda event: self.fetch_tool_name(self.return_tool_id, self.return_tool_name_label))

        tk.Label(self.return_tab, text='工具名称：').grid(column=0, row=1)
        self.return_tool_name_label = tk.Label(self.return_tab, text="")
        self.return_tool_name_label.grid(column=1, row=1)
        
        return_button = tk.Button(self.return_tab, text='确定归还', command=self.return_tool)
        return_button.grid(column=1, row=2)

        # 创建 Treeview 来显示未归还工具信息
        self.unreturned_tree = ttk.Treeview(self.return_tab, columns=("工具编号", "工具名称", "借出者", "借出日期"))
        self.unreturned_tree.heading("#1", text="工具编号")
        self.unreturned_tree.heading("#2", text="工具名称")
        self.unreturned_tree.heading("#3", text="借出者")
        self.unreturned_tree.heading("#4", text="借出日期")
        self.unreturned_tree.grid(column=0, row=3, columnspan=2, padx=10, pady=10, sticky="nsew")

        # 快捷操作-在未归还工具表格中双击选中工具
        self.unreturned_tree.bind("<Double-1>", self.on_select_tool)

        # 列表的列配置
        self.unreturned_tree.column("#0", width=0, stretch=tk.NO)
        self.unreturned_tree.column("#1", width=100)
        self.unreturned_tree.column("#2", width=200)
        self.unreturned_tree.column("#3", width=150)
        self.unreturned_tree.column("#4", width=100)

        # 尚未归还的工具Treeview滚动条
        scrollbar = ttk.Scrollbar(self.return_tab, orient="vertical", command=self.unreturned_tree.yview)
        self.unreturned_tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(column=2, row=3, sticky="ns")

        # 在初始化时显示未归还工具信息
        self.show_unreturned_tools()

    '''
    下面完成输入工具编号就从数据库中获取工具名称的函数
    '''

    def fetch_tool_name(self, tool_id_entry, tool_name_label):
        try:
            # 获取工具编号的原始输入
            tool_id_raw = tool_id_entry.get()

            # 检查输入是否为空
            if not tool_id_raw:
                messagebox.showerror("Error", "请输入工具编号")
                return

            # 检查输入是否为数字，并尝试转换为整数
            if tool_id_raw.isdigit():
                tool_id = int(tool_id_raw)
            else:
                messagebox.showerror("Error", "工具编号必须是数字")
                return

            # 从数据库中检索工具名称
            self.cursor.execute("SELECT ToolName FROM Tools WHERE ToolID = ?", (tool_id,))
            tool_name = self.cursor.fetchone()

            if tool_name:
                # 如果找到了工具，更新工具名称标签
                tool_name_label.config(text=tool_name[0])
            else:
                # 如果没有找到工具，显示错误信息
                messagebox.showerror("Error", "没有找到对应的工具")

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)

    '''
    GUI配置结束。下面是对后端操作。
    '''

    def tool_change_in_stock(self, tool_id, in_stock=True):
        # 公用函数，用于告知主表工具变化状态

        try:
            # 查找工具记录是否存在
            self.cursor.execute("SELECT * FROM Tools WHERE ToolID = ?", [tool_id])
            tool_exists = self.cursor.fetchone()

            # 如果存在，更新InStock状态
            if tool_exists:
                self.cursor.execute("UPDATE Tools SET InStock = ? WHERE ToolID = ?", [in_stock, tool_id])
            else:
                messagebox.showerror("Error", "工具记录不存在")

            # 提交更改
            self.conn.commit()
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", str(e))


    def borrow_tool(self):
        try:
            tool_id = self.borrow_tool_id.get()
            borrower = self.borrow_borrower.get()
            department = self.borrow_department.get()

            # 初始化预期归还时间
            expected_return_date = None

            # 检查输入的使用天数和预期归还日期
            use_days_input = self.use_days.get()
            return_date_input = self.expected_return_date.get()

            # 如果用户只填写了使用天数
            if use_days_input and not return_date_input:
                try:
                    use_days = int(use_days_input)
                    expected_return_date = datetime.date.today() + datetime.timedelta(days=use_days)
                except ValueError:
                    messagebox.showerror("Error", "使用天数必须是整数")
                    return

            # 如果用户只填写了预期归还日期
            elif return_date_input and not use_days_input:
                try:
                    expected_return_date = datetime.datetime.strptime(return_date_input, '%Y/%m/%d').date()
                except ValueError:
                    try:
                        # 尝试另一种日期格式
                        expected_return_date = datetime.datetime.strptime(return_date_input, '%Y-%m-%d').date()
                    except ValueError:
                        messagebox.showerror("Error", "预期归还日期格式错误，正确格式为 YYYY/MM/DD 或 YYYY/M/D")
                        return

            # 如果用户两者都填写了，则检查一致性
            elif use_days_input and return_date_input:
                try:
                    use_days = int(use_days_input)
                    calculated_due_date = datetime.date.today() + datetime.timedelta(days=use_days)
                    input_due_date = datetime.datetime.strptime(return_date_input, '%Y/%m/%d').date()
                    if calculated_due_date != input_due_date:
                        messagebox.showerror("Error", "使用天数和预期归还日期不一致，请检查")
                        return
                    expected_return_date = calculated_due_date
                except ValueError as e:
                    messagebox.showerror("Error", f"输入错误: {e}")
                    return
                
            # 插入数据库之前，确保日期格式正确
            if expected_return_date:
                formatted_date = expected_return_date.strftime('%Y/%m/%d')
                # ... 插入数据库的操作，使用 formatted_date 作为日期值

            # 检查工具是否在库存中
            self.cursor.execute("SELECT InStock FROM Tools WHERE ToolID = ?", (tool_id,))
            result = self.cursor.fetchone()

            if result is not None:
                in_stock = result[0]
                if in_stock != 1:
                    messagebox.showerror("Error", "该工具无法借出（已被借出或已报废）")
                    return

            # 直接从标签获取工具名称
            tool_name = self.borrow_tool_name_label.cget("text")

            if tool_name and borrower and department:
                current_date = datetime.date.today()

                # 插入借用记录
                self.cursor.execute("INSERT INTO BorrowRecords (ToolID, ToolName, Borrower, BorrowerDepartment, BorrowDate, ExpectDate) VALUES (?, ?, ?, ?, ?, ?)",
                                     (tool_id, tool_name, borrower, department, current_date, formatted_date))

                # 更新工具状态为被借出
                self.cursor.execute("UPDATE Tools SET InStock = 0 WHERE ToolID = ?", (tool_id,))

                # 提交更改
                self.conn.commit()

                self.refresh_treeviews()
                self.fill_borrower_department_tree()

            else:
                messagebox.showerror("Error", "请确保已经填写了所有信息，如果没有显示工具名称，就先在工具编号输入框点击并回车。")
                return

        except Exception as e:
            messagebox.showerror("Unexpected Error", f"An unexpected error occurred: {e}")
        else:
            # 清空输入
            self.borrow_tool_id.delete(0, tk.END)
            self.borrow_borrower.delete(0, tk.END)
            self.borrow_department.delete(0, tk.END)
            self.use_days.delete(0, tk.END)
            self.expected_return_date.delete(0, tk.END)
            self.borrow_tool_name_label.config(text="")

            messagebox.showinfo("Success", f"编号为： {tool_id} 的工具，已被： {borrower} 借出。")

    def calculate_due_date(self, event=None):
        use_days_str = self.use_days.get()
        if use_days_str.isdigit():
            use_days = int(use_days_str)
            due_date = datetime.date.today() + datetime.timedelta(days=use_days)
            self.expected_return_date.delete(0, tk.END)
            # 使用正确的格式 yyyy/mm/dd
            self.expected_return_date.insert(0, due_date.strftime('%Y/%m/%d'))
        elif use_days_str.strip() == "":  # 允许清空输入
            self.expected_return_date.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "使用天数必须是整数")
    
    def adjust_inventory(self):
        tool_name = self.add_tool_name.get()  # 获取工具名称
        location = self.add_location.get()  # 获取位置
        reason = self.add_reasons_combobox.get()  # 获取增加原因
        notes = self.add_notes.get()  # 获取备注信息

        # 检查输入是否为空
        if not tool_name or not location:
            messagebox.showerror("Error", "请填写所有必要信息")
            return

        # 使用正则表达式检查工具位置格式
        import re
        pattern = r'^\d+-\d+-\d+$'  # 匹配格式为X-X-XX的字符串，其中X都是数字
        if not re.match(pattern, location):
            # 如果格式不匹配，弹出提示框
            messagebox.showerror("格式错误", "工具位置格式错误，请输入正确的格式（X-X-XX）")
            self.add_location.delete(0, tk.END)  # 清空输入框内容
            return

        current_date = datetime.date.today()

        try:
            # 添加库存调整记录
            self.cursor.execute("INSERT INTO InventoryAdjustments (ToolName, AdjustmentDate, Reason, Notes, Location) VALUES (?, ?, ?, ?, ?)",
                                (tool_name, current_date, reason, notes, location))

            # 插入新的工具记录
            self.cursor.execute("INSERT INTO Tools (ToolName, Location, InStock) VALUES (?, ?, ?)",
                                (tool_name, location, 1))  # InStock 设置为 1

            # 提交更改
            self.conn.commit()

            # 查询刚插入的工具的ToolID和Location
            self.cursor.execute("SELECT ToolID, Location FROM Tools WHERE ToolName = ?", (tool_name,))
            tool_info = self.cursor.fetchone()

            if tool_info:
                tool_id, tool_location = tool_info
                messagebox.showinfo("Success", f"Inventory adjusted for tool {tool_name}\n工具编号: {tool_id}, 工具位置: {tool_location}")
            else:
                messagebox.showinfo("Success", f"Inventory adjusted for tool {tool_name}")

            # 清空输入
            self.add_tool_name.delete(0, tk.END)
            self.add_location.delete(0, tk.END)
            self.add_notes.delete(0, tk.END)

            self.refresh_treeviews()

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)

    def bulk_insert(self):
        # 获取用户输入的工具名称、入库数量、首个工具位置、增加在同一位置还是同一层
        tool_name = self.tool_name_entry.get()
        quantity_str = self.tool_num_entry.get()
        same_loc_or_same_floor = self.tool_change_in_stock_combobox.get()
        location_str = self.tool_location_1st_entry.get()

        # 检查工具名称、数量和位置是否有效
        if not tool_name:
            messagebox.showerror("错误", "请输入工具名称")
            return
        
        if not quantity_str.isdigit() or int(quantity_str) <= 0:  # 先检查是否为数字，然后转换为整数
            messagebox.showerror("错误", "请输入有效的入库数量（必须是大于零的整数）")
            return

        try:
            quantity = int(quantity_str)
            if quantity <= 0:
                raise ValueError("入库数量必须大于零")

            # 分解位置字符串
            parts = location_str.split('-')
            if len(parts) != 3:
                raise ValueError("初始位置格式不正确，正确格式应为X-X-XX")
            cabinet_number, shelf_number, start_sequence_number = map(int, parts)
            
            # 生成工具位置列表
            if same_loc_or_same_floor == "同一层":
                tool_locations = [f"{cabinet_number}-{shelf_number}-{start_sequence_number + i:02d}" for i in range(quantity)]
            else:
                tool_locations = [location_str for _ in range(quantity)]

            # 插入工具记录
            for location in tool_locations:
                self.cursor.execute("INSERT INTO Tools (ToolName, Location, InStock) VALUES (?, ?, ?)",
                                    (tool_name, location, 1))  # InStock 设置为 1

            # 提交更改
            self.conn.commit()
            messagebox.showinfo("成功", f"成功插入 {quantity} 条工具记录")

        except ValueError as e:
            messagebox.showerror("输入错误", str(e))
            return
        except Exception as e:
            messagebox.showerror("数据库错误", f"插入工具记录时发生错误: {str(e)}")
            return

        # 清空输入框
        self.tool_name_entry.delete(0, tk.END)
        self.tool_num_entry.delete(0, tk.END)
        self.tool_location_1st_entry.delete(0, tk.END)

        self.refresh_treeviews()



    def scrap_tool(self):
        tool_id_raw = self.scrap_tool_id.get()  # 获取原始输入
        if tool_id_raw.isdigit():
            tool_id = int(tool_id_raw)
        else:
            messagebox.showerror("Error", "工具编号必须是数字")
            return

        notes = self.scrap_notes.get()  # 获取备注信息

        try:
            # 检查工具是否在库存中
            self.cursor.execute("SELECT InStock FROM Tools WHERE ToolID = ?", (tool_id,))
            result = self.cursor.fetchone()
            if result is None:
                messagebox.showerror("Error", "找不到对应的工具")
                return

            in_stock = result[0]
            if in_stock != 1:
                messagebox.showerror("Error", "工具无法报废（已被借出或已报废）")
                return

            # 获取工具名称
            self.cursor.execute("SELECT ToolName FROM Tools WHERE ToolID = ?", (tool_id,))
            tool_name_result = self.cursor.fetchone()
            if not tool_name_result:
                messagebox.showerror("Error", "找不到对应的工具名称")
                return
            tool_name = tool_name_result[0]

            # 添加损失和损坏记录
            self.cursor.execute("INSERT INTO LossOrDamageRecords (ToolName, AdjustmentDate, Reason, Notes) VALUES (?, ?, ?, ?)",
                                (tool_name, datetime.date.today(), 'Scrap', notes))

            # 将工具标记为已报废
            self.cursor.execute("UPDATE Tools SET InStock = 2 WHERE ToolID = ?", (tool_id,))

            # 提交更改
            self.conn.commit()

            self.refresh_treeviews()
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)
        else:
            messagebox.showinfo("成功", f"编号为 {tool_id} 的工具已登记报废")
            # 清空输入
            self.scrap_tool_id.delete(0, tk.END)
            self.scrap_notes.delete(0, tk.END)

    def return_tool(self):
        tool_id_raw = self.return_tool_id.get()
        if tool_id_raw.isdigit():
            tool_id = int(tool_id_raw)
        else:
            messagebox.showerror("Error", "工具编号必须是数字")
            return

        try:
            # 检查工具是否在借出状态
            self.cursor.execute("SELECT InStock FROM Tools WHERE ToolID = ?", (tool_id,))
            result = self.cursor.fetchone()
            if result is None:
                messagebox.showerror("Error", "找不到对应的工具")
                return

            in_stock = result[0]
            if in_stock != 0:
                messagebox.showerror("Error", "该工具不在借出状态，无法归还")
                return

            # 更新 BorrowRecords 表中的对应记录
            current_date = datetime.date.today()
            self.cursor.execute("UPDATE BorrowRecords SET ReturnDate = ? WHERE ToolID = ?",
                                (current_date, tool_id))

            # 将工具标记为在库中
            self.cursor.execute("UPDATE Tools SET InStock = 1 WHERE ToolID = ?", (tool_id,))

            # 提交更改
            self.conn.commit()
            messagebox.showinfo("成功", f"编号为 {tool_id} 的工具已成功归还")

            self.refresh_treeviews()

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)

        # 清空输入
        self.return_tool_id.delete(0, tk.END)
        self.return_tool_name_label.config(text="")

    def refresh_treeviews(self):
        self.show_unreturned_tools()  # 刷新未归还工具表格
        self.search_tool()  # 刷新搜索结果表格
        self.on_select_borrower_department # 刷新借用人部门表格
    
    def search_tool(self):
        search_query = self.search_tool_entry.get()
        self.search_results_tree.delete(*self.search_results_tree.get_children())

        try:
            # 执行数据库查询
            self.cursor.execute("SELECT ToolID, ToolName, Location, InStock FROM Tools WHERE ToolName LIKE ?", ('%' + search_query + '%',))
            query_results = self.cursor.fetchall()
            
            # 清空当前存储的搜索结果
            self.search_results.clear()

            # 处理每条记录，并添加到 self.search_results
            for tool in query_results:
                tool_id, tool_name, location, in_stock = tool
                status = "在库" if in_stock == 1 else "借出" if in_stock == 0 else "报废"
                self.search_results.append((tool_id, tool_name, location, status))

            # 确保有结果后再显示第一页
            if self.search_results:
                self.show_page(0)
            else:
                messagebox.showinfo("Search", "未找到符合条件的工具。")

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)

    '''
    ---↓↓↓ 侧边栏和快捷操作中心 ↓↓↓---
    '''
    
    def show_page(self, page_number):
        self.search_results_tree.delete(*self.search_results_tree.get_children())
        
        start = page_number * self.results_per_page
        end = start + self.results_per_page
        
        for item in self.search_results[start:end]:
            self.search_results_tree.insert("", "end", values=item)
        
        self.page_number = page_number
        self.update_paging_button()
    
    def update_paging_button(self):
        self.prev_page_button['state'] = 'normal' if self.page_number > 0 else 'disabled'
        self.next_page_button['state'] = 'normal' if (self.page_number + 1) * self.results_per_page < len(self.search_results) else 'disabled'
    def prev_page(self):
        if self.page_number > 0:
            self.show_page(self.page_number - 1)

    def next_page(self):
        if (self.page_number + 1) * self.results_per_page < len(self.search_results):
            self.show_page(self.page_number + 1)

    def show_unreturned_tools(self):
        self.unreturned_tree.delete(*self.unreturned_tree.get_children())

        try:
            self.cursor.execute("SELECT ToolID, ToolName, Borrower, BorrowDate, ExpectDate FROM BorrowRecords WHERE ReturnDate IS NULL")
            unreturned_tools = self.cursor.fetchall()

            for tool in unreturned_tools:
                tool_id, tool_name, borrower, borrow_date, expect_date = tool
                formatted_borrow_date = borrow_date.strftime("%Y-%m-%d")
                overdue_status = self.calculate_overdue(expect_date)
                
                # Determine if the row should be tagged as overdue
                tags = ('overdue',) if overdue_status == "Overdue" else ()

                self.unreturned_tree.insert("", "end", values=(tool_id, tool_name, borrower, formatted_borrow_date, expect_date.strftime("%Y-%m-%d") if expect_date else "N/A", overdue_status), tags=tags)

            # Highlight overdue items
            self.unreturned_tree.tag_configure('overdue', background='red')

        except pyodbc.Error as e:
            messagebox.showerror("Database Error", e)
   
    def on_select_tool(self, event):
        event_widget = event.widget

        # 在侧边栏中选择工具
        if event_widget == self.search_results_tree:
            selected_item = self.search_results_tree.item(self.search_results_tree.selection())
            tool_id, tool_name, _, in_stock_status = selected_item['values']

            # 判断工具的在库状态，并跳转到相应的标签页
            if in_stock_status == "在库":
                self.tab_control.select(self.borrow_tab)  # 切换到借出标签页
                self.fill_tool_info(self.borrow_tool_id, self.borrow_tool_name_label, tool_id, tool_name)
            elif in_stock_status == "借出":
                self.tab_control.select(self.return_tab)  # 切换到归还标签页
                self.fill_tool_info(self.return_tool_id, self.return_tool_name_label, tool_id, tool_name)
            elif in_stock_status == "报废":
                # 这里可以弹出提示，或者进行其他操作
                messagebox.showinfo("信息", f"工具 {tool_name} (编号: {tool_id}) 已报废。")
        
        # 在未归还的工具列表中选择工具
        elif event_widget == self.unreturned_tree:
            selected_item = self.unreturned_tree.item(self.unreturned_tree.selection())
            tool_id, tool_name = selected_item['values'][0], selected_item['values'][1]
            self.fill_tool_info(self.return_tool_id, self.return_tool_name_label, tool_id, tool_name)

        # 在借用人部门列表中选择部门
        elif event_widget == self.borrower_department_tree:
            selected_item = event_widget.item(event_widget.selection())
            if 'values' in selected_item and len(selected_item['values']) == 2:
                borrower, department = selected_item['values']
                self.borrow_borrower.delete(0, tk.END)
                self.borrow_borrower.insert(0, borrower)
                self.borrow_department.delete(0, tk.END)
                self.borrow_department.insert(0, department)

    def setup_borrower_department_treeview(self):
        # 设置 Treeview 列
        self.borrower_department_tree["columns"] = ("借出者", "部门")
        self.borrower_department_tree.column("#0", width=0, stretch=tk.NO)
        self.borrower_department_tree.column("借出者", anchor=tk.W, width=120)
        self.borrower_department_tree.column("部门", anchor=tk.W, width=120)

        # 设置 Treeview 表头
        self.borrower_department_tree.heading("借出者", text="借出者", anchor=tk.W)
        self.borrower_department_tree.heading("部门", text="部门", anchor=tk.W)

        # 调用方法填充数据
        self.fill_borrower_department_tree()

        # 添加滚动条
        borrower_dept_scrollbar = ttk.Scrollbar(self.borrow_tab, orient="vertical", command=self.borrower_department_tree.yview)
        self.borrower_department_tree.configure(yscrollcommand=borrower_dept_scrollbar.set)
        borrower_dept_scrollbar.grid(column=2, row=4, sticky='ns')


    def fill_treeview(self, treeview, query, columns):
        treeview.delete(*treeview.get_children())
        try:
            self.cursor.execute(query)
            for row in self.cursor.fetchall():
                # 确保 row 中的数据是平展开的
                if isinstance(row, tuple):
                    treeview.insert("", "end", values=row)
                else:
                    treeview.insert("", "end", values=(row,))
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", str(e))
    
    def fill_borrower_department_tree(self):
        self.borrower_department_tree.delete(*self.borrower_department_tree.get_children())
        try:
            self.cursor.execute("SELECT DISTINCT Borrower, BorrowerDepartment FROM BorrowRecords")
            for borrower, department in self.cursor.fetchall():
                self.borrower_department_tree.insert("", "end", values=(borrower, department))
        except pyodbc.Error as e:
            messagebox.showerror("Database Error", str(e))


    def fill_tool_info(self, id_entry_widget, name_label_widget, tool_id, tool_name):
        id_entry_widget.delete(0, tk.END)
        id_entry_widget.insert(0, tool_id)
        name_label_widget.config(text=tool_name)

    def calculate_overdue(self, expected_return_date):
        if expected_return_date is not None:
            if isinstance(expected_return_date, datetime.datetime):
                expected_return_date = expected_return_date.date()
            return "Overdue" if expected_return_date < datetime.date.today() else ""
        return ""
    
    def on_window_close(self):
        try:
            # 在窗口关闭时使用cmd运行另一个Python程序
            subprocess.run(["cmd", "/c", "python", "D:\\ToolManagementSystem\\ForBackup.py"])
        except Exception as e:
            # 处理错误，例如显示错误消息
            print(f"An error occurred: {e}")

        # 关闭窗口
        self.destroy()

'''
↓↓↓查询与统计中心↓↓↓
'''

class DataWindow(tk.Toplevel):
    def __init__(self, parent, conn):
        super().__init__(parent)
        self.conn = conn
        self.title("数据库数据")
        self.geometry("800x600")
        self.parent = parent
        self.create_widgets()
    def create_widgets(self):
        # 创建一个 Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both")

        # 创建每个表的标签页
        self.create_table_tab("Tools", ["ToolID", "ToolName", "Location", "InStock"])
        self.create_table_tab("BorrowRecords", ["RecordID", "ToolID", "Borrower", "BorrowDate", "ExpectDate", "ReturnDate"])
        self.create_table_tab("LossOrDamageRecords", ["RecordID", "ToolID", "ToolName", "AdjustmentDate", "Reason"])
        self.create_table_tab("InventoryAdjustments", ["RecordID", "ToolID", "ToolName", "AdjustmentDate", "Reason", "Notes", "Location"])

        # 创建统计信息的标签页
        self.statistics_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.statistics_tab, text="统计信息")
        self.create_statistics_ui()

        self.stats_tree.bind("<<TreeviewOpen>>", self.on_item_open)

    def on_load_details(self):
        selected_item = self.stats_tree.selection()
        if selected_item:  # Checking if there is an item selected
            self.load_detailed_tool_data(selected_item[0])

    def create_table_tab(self, table_name, columns):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=table_name)

        # Pass the newly created tab as the parent for the treeview
        self.create_treeview(tab, table_name, columns)

    def create_table_tab_for_each(self, table_name, columns):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=table_name)
        self.create_treeview(tab, table_name, columns)

    def create_treeview(self, parent, table_name, columns):
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        tree.pack(expand=True, fill="both")

        for col in columns:
            tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(tree, _col, False))
            tree.column(col, width=tkFont.Font().measure(col.title()))

        self.load_data(tree, table_name, columns)

    def load_data(self, tree, table_name, columns):
        try:
            cursor = self.conn.cursor()
            query = f"SELECT {', '.join(columns)} FROM {table_name}"
            cursor.execute(query)
            for row in cursor.fetchall():
                # Format row data properly before inserting into tree
                formatted_row = []
                for item in row:
                    if isinstance(item, datetime.datetime):
                        # Format datetime object as a string
                        formatted_row.append(item.strftime('%Y-%m-%d %H:%M:%S'))
                    else:
                        # Convert other items to string and strip unwanted characters
                        formatted_row.append(str(item).strip("(),'"))
                tree.insert("", "end", values=formatted_row)
        except Exception as e:
            messagebox.showerror("Database Error", str(e))
            print(str(e))

    def treeview_sort_column(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.treeview_sort_column(tree, col, not reverse))

    def create_statistics_ui(self):
        # Create a treeview in the statistics tab
        self.stats_tree = ttk.Treeview(self.statistics_tab, columns=("ToolName", "Total", "InStock", "Out", "Overdue"))
        self.stats_tree.heading("ToolName", text="工具名称")
        self.stats_tree.heading("Total", text="总计")
        self.stats_tree.heading("InStock", text="库存")
        self.stats_tree.heading("Out", text="借出")
        self.stats_tree.heading("Overdue", text="逾期")
        
        # This sets up the treeview to have a hierarchical structure
        self.stats_tree.column("#0", width=0, stretch=tk.NO)
        self.stats_tree.column("ToolName", anchor=tk.W, width=120)
        self.stats_tree.column("Total", anchor=tk.W, width=80)
        self.stats_tree.column("InStock", anchor=tk.W, width=80)
        self.stats_tree.column("Out", anchor=tk.W, width=80)
        self.stats_tree.column("Overdue", anchor=tk.W, width=80)
        self.stats_tree.pack(expand=True, fill="both")
        
        # Load data into the statistics treeview
        self.load_statistics_data()
        
        # Note about alerts
        note_label = tk.Label(self.statistics_tab, text="红色警报说明库存不足或逾期归还。")
        note_label.pack(side="bottom")

    def load_statistics_data(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT ToolName, COUNT(*) as TotalCount,
                SUM(IIF(InStock=1, 1, 0)) as TotalInStock
                FROM Tools
                GROUP BY ToolName
            """)

            for row in cursor.fetchall():
                tool_name, total_count, total_in_stock = row
                inventory_percentage = (total_in_stock / total_count) * 100
                tag = 'low_inventory' if inventory_percentage < 20 else ''
                # Insert the summary data into the Treeview
                parent = self.stats_tree.insert("", "end", text=tool_name, values=(tool_name, total_count, total_in_stock), tags=(tag,))
                # Insert a dummy child item so the "+" icon appears
                self.stats_tree.insert(parent, "end", text="Loading...", values=("Loading...",))
            
            self.stats_tree.tag_configure('low_inventory', background='red')

        except Exception as e:
            messagebox.showerror("Database Error", str(e))

        # Bind the event to the treeview
        self.stats_tree.bind("<<TreeviewOpen>>", self.on_item_open)

    def load_detailed_tool_data(self, parent_item):
        # Clear existing children to avoid duplicates if the item is reopened
        for child in self.stats_tree.get_children(parent_item):
            self.stats_tree.delete(child)

        tool_name = self.stats_tree.item(parent_item, "values")[0]  # Assuming the tool name is the first value
        cursor = self.conn.cursor()
        detailed_query = """
            SELECT ToolID, BorrowDate, ExpectDate, Borrower, BorrowerDepartment
            FROM BorrowRecords
            WHERE ToolName = ? AND ReturnDate IS NULL
        """
        cursor.execute(detailed_query, (tool_name,))

        for row in cursor.fetchall():
            tool_id, last_borrow_time, expected_return_time, borrower, borrower_department = row
            overdue = self.calculate_overdue(expected_return_time)
            tag = 'overdue' if overdue == "Overdue" else ''
            # Insert the detailed information as a child of the tool item
            self.stats_tree.insert(parent_item, "end", values=(tool_id, last_borrow_time, expected_return_time, borrower, borrower_department), tags=(tag,))
            
        self.stats_tree.tag_configure('overdue', background='red')
        cursor.close()
    
    def on_item_open(self, event):
        item = self.stats_tree.focus()  # Get the item that was opened
        self.load_detailed_tool_data(item)

    def on_item_expand(self, event):
        item = self.stats_tree.focus()  # Get the item that was opened
        # Check if the item has children; if not, it's assumed to be a parent that needs loading
        if not self.stats_tree.get_children(item):
            self.load_detailed_tool_data(item)

    def calculate_overdue(self, expected_return_date):
        if expected_return_date and expected_return_date < datetime.date.today():
            return "Overdue"
        return ""

# 运行应用
app = ToolManagementApp()
app.mainloop()