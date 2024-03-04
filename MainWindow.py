from UseDB import MyDB
import tkinter as tk
from tkinter import messagebox  # 弹出框
from tkinter import ttk  #下拉菜单
from tkinter import filedialog
import xlwt  #向excel中写入数据
import xlrd  #从excel中读取数据
import xerox  #复制文本到剪切板

# from tkinter.ttk import Separator # 利用Separator创建分割线时引入


class window:
    # 初始化数据库
    db = MyDB()

    # 创建窗口
    mainWindow = tk.Tk()
    mainWindow.title('账号密码管理系统')
    screenwidth = mainWindow.winfo_screenwidth()
    screenheight = mainWindow.winfo_screenheight()
    mainWindow.geometry('800x500+%d+%d' %
                        ((screenwidth - 800) / 2,
                         (screenheight - 500) / 2))  # 设置窗口产生位置
    mainWindow.resizable(False, False)  #横纵均不允许调整 禁止缩放

    # 初始化界面
    text_label_num = tk.StringVar()  # 数据库中记录数量 文字变量储存器
    tmp_num = 0  # 记录数量

    # 创建画布
    canvas = tk.Canvas(mainWindow, height=500, width=200)
    img_file_head = tk.PhotoImage(file='./img/head.png')

    # 创建头像
    image = canvas.create_image(55, 40, anchor='nw', image=img_file_head)

    # 创建分割线
    line = canvas.create_line(200, 0, 200, 500, fill='gray')

    # 控制画布位置
    canvas.place(x=0, y=0)

    # Button:导入密码库
    btn_inpwd = tk.Button(mainWindow, text="导入密码库", width=13, height=1)
    btn_inpwd.place(x=50, y=190)

    # Button:导出密码库
    btn_outpwd = tk.Button(mainWindow, text="导出密码库", width=13, height=1)
    btn_outpwd.place(x=50, y=250)

    # Label:显示记录数量
    label_shownum = tk.Label(mainWindow,
                             textvariable=text_label_num,
                             font=('Arial', 12),
                             width=15,
                             height=2)
    label_shownum.place(x=30, y=320)

    # Label:名称
    label_name = tk.Label(
        mainWindow,
        text="名称",
        font=('Arial', 10),
        width=6,
        height=1,
        borderwidth=2,
        relief="ridge",
    )
    label_name.place(x=250, y=49)

    # Entry:输入项目名称
    entry_name = tk.Entry(mainWindow, width=20)
    entry_name.place(x=304, y=50)

    # Label:账号
    label_account = tk.Label(
        mainWindow,
        text="账号",
        font=('Arial', 10),
        width=6,
        height=1,
        borderwidth=2,
        relief="ridge",
    )
    label_account.place(x=500, y=49)

    # Entry:输入账号
    entry_account = tk.Entry(mainWindow, width=20)
    entry_account.place(x=554, y=50)

    # Label:备注
    label_remark = tk.Label(
        mainWindow,
        text="备注",
        font=('Arial', 10),
        width=6,
        height=1,
        borderwidth=2,
        relief="ridge",
    )
    label_remark.place(x=250, y=109)

    # Entry:输入备注
    entry_remark = tk.Entry(mainWindow, width=20)
    entry_remark.place(x=304, y=110)

    # Label:密码
    label_password = tk.Label(
        mainWindow,
        text="密码",
        font=('Arial', 10),
        width=6,
        height=1,
        borderwidth=2,
        relief="ridge",
    )
    label_password.place(x=500, y=109)

    # Entry:输入密码
    entry_password = tk.Entry(mainWindow, width=20)
    entry_password.place(x=554, y=110)

    # Button:添加/修改项目
    btn_add_info = tk.Button(mainWindow, text="添加/修改", width=8, height=1)
    btn_add_info.place(x=260, y=160)

    # Button:删除项目
    btn_del_info = tk.Button(mainWindow, text="删除", width=8, height=1)
    btn_del_info.place(x=350, y=160)

    # Entry:输入搜索
    entry_search = tk.Entry(mainWindow, width=16)
    entry_search.place(x=500, y=165)

    # Combobox:下拉菜单
    combobox_selected = ttk.Combobox(mainWindow, state="readonly", width=7)
    combobox_selected.place(x=620, y=165)
    combobox_selected['value'] = ("名称", "账号", "密码", "备注")
    combobox_selected.current(0)

    # Button:点击搜索
    img_search = tk.PhotoImage(file='./img/search.png')
    img_search = img_search.subsample(2, 2)
    btn_search_info = tk.Button(mainWindow, image=img_search)
    btn_search_info.place(x=700, y=165)

    # 表格:显示项目内容
    cols = ('id', '名称', '账号', '密码', '备注')
    tree = ttk.Treeview(mainWindow, show='headings', columns=cols, height=14)
    # 表头设置
    for col in cols:
        tree.heading(col, text=col)  #列标题
    # self.tree.column(col, width=125, anchor='w')  #设置col列的宽度,anchor表示内容位置
    tree.column('id', width=45, anchor='w')
    tree.column('名称', width=105, anchor='w')
    tree.column('账号', width=155, anchor='w')
    tree.column('密码', width=150, anchor='w')
    tree.column('备注', width=140, anchor='w')
    tree.place(x=202, y=195)

    def __init__(self):
        # Button:导入密码库
        self.btn_inpwd['command'] = self.btn_inpwd_func
        # Button:导出密码库
        self.btn_outpwd['command'] = self.btn_outpwd_func
        # 刷新记录数量
        self.refresh_records_num()
        # Button:添加/修改项目
        self.btn_add_info['command'] = self.btn_add_func
        # Button:删除项目
        self.btn_del_info['command'] = self.btn_del_func
        # Button:点击搜索
        self.btn_search_info['command'] = self.btn_searchinfo_func
        # 添加数据到表格中
        self.add_info_in_table()
        # 绑定单击离开事件===========
        self.tree.bind('<ButtonRelease-1>', self.tree_item_Click)
        # 界面运行
        self.mainWindow.mainloop()
        return

    # 方法:计算密码库中的记录数量
    def number_of_records(self):
        global db
        sql_count_record = 'select count(*) from information'
        num = self.db.executeQuery(sql_count_record)[0][0]  # [(0,)]
        return num

    # 方法: 刷新记录数量
    def refresh_records_num(self):
        global tmp_num
        global text_label_num
        tmp_num = self.number_of_records()
        self.text_label_num.set(f"共有 {tmp_num} 条记录")
        return

    # 事件: 删除列表中所有元素
    def del_info_in_table(self):
        global tree
        x = self.tree.get_children()
        for item in x:
            self.tree.delete(item)

    # 事件: 刷新列表中项目
    def add_info_in_table(self):
        global tree
        self.del_info_in_table()
        ans = self.db.executeQuery('select * from information')
        tmp_i = 1  # 设置编号
        for i in ans:
            self.tree.insert("",
                             "end",
                             str(tmp_i),
                             values=(i[0], i[1], i[2], i[3], i[4]))
            tmp_i = tmp_i + 1
        return

    # 事件:导入密码
    def btn_inpwd_func(self):
        # 导入前提示
        messagebox.showinfo(
            title="注意",
            message=
            "导入会覆盖之前保存的记录!\n\n导入文件为Excel表格时:\n每行数据从左至右顺序应为(id,名称,账号,密码,备注) \"备注\"一列可为空\n\n导入文件为txt文本时:\n每行数据从左至右顺序应为(id,名称,账号,密码,备注)，数据之间用TAB分开"
        )

        # 导入前备份
        ans = self.db.executeQuery('select * from information')

        # 选择文件开始导入
        try:
            FilePath = FilePath = filedialog.askopenfilename(
                filetypes=[("Excel", "*.xls"), ("Excel",
                                                "*.xlsx"), ("txt文本", "*.txt")])
            t = FilePath.split('.')[1]
            #如果为Excel表格
            if t == 'xls' or t == 'xlsx':
                # 读取选择的文件
                book = xlrd.open_workbook(FilePath)
                # 获取第一个表格
                sheet1 = book.sheet_by_index(0)
                # 不为4列或5列 出错
                if sheet1.ncols != 4 and sheet1.ncols != 5:
                    messagebox.showwarning(title='导入失败', message='表格格式有误，导入错误')
                    return
                # 准备空列表放数据
                indb = []
                for i in range(1, sheet1.nrows):
                    row = sheet1.row_values(i)
                    indb.append((i, row[1], row[2], row[3], row[4]))
                # 清空数据库总内容
                sql_clear_table = 'delete from information'
                self.db.executeUpdate(sql_clear_table)
                # 放入导入的数据
                sql_add_info = 'insert into information values(?,?,?,?,?)'
                self.db.executeUpdate(sql_add_info, indb)
            elif t == 'txt':
                # 准备空列表放数据
                indb = []
                # 打开文件
                with open(FilePath, "r") as f:
                    # 读取所有行
                    data = f.readlines()
                    # 循环取出每行数据 放入indb
                    for i in range(0, len(data)):
                        bz = data[i].split('\t')
                        indb.append((i + 1, bz[1], bz[2], bz[3], bz[4]))
                # 清空数据库总内容
                sql_clear_table = 'delete from information'
                self.db.executeUpdate(sql_clear_table)
                # 放入导入的数据
                sql_add_info = 'insert into information values(?,?,?,?,?)'
                self.db.executeUpdate(sql_add_info, indb)
        except:
            sql_clear_table = 'delete from information'
            self.db.executeUpdate(sql_clear_table)
            sql_add_info = 'insert into information values(?,?,?,?,?)'
            self.db.executeUpdate(sql_add_info, ans)
            messagebox.showwarning(title='导入失败', message="导入失败")
        self.refresh_records_num()
        self.add_info_in_table()
        return

    # 事件:导出密码
    def btn_outpwd_func(self):
        # 判断是否可以导出
        if tmp_num == 0:
            messagebox.showwarning(title='', message='暂无记录，无法导出！')
            return
        # 取出所有记录
        ans = self.db.executeQuery('select * from information')
        # 创建workbook
        book = xlwt.Workbook()
        # 创建sheet
        sheet1 = book.add_sheet("账号信息")
        # sheet1.write(0, 0, 'id')
        # sheet1.write(0, 1, '名称')
        # sheet1.write(0, 2, '账号')
        # sheet1.write(0, 3, '密码')
        # sheet1.write(0, 4, '备注')
        sheet1.col(0).width = 256 * 5
        sheet1.col(1).width = 256 * 15
        sheet1.col(2).width = 256 * 25
        sheet1.col(3).width = 256 * 25
        sheet1.col(4).width = 256 * 50
        sheet1.row(0).height_mismatch = True  # 允许行高自定义
        sheet1.row(0).height = 20 * 15  # 行高以 1/20 个点为单位
        # 写数据 (行索引,列索引,数据)
        for i in range(0, len(ans)):
            sheet1.write(i, 0, i + 1)
            sheet1.write(i, 1, ans[i][1])
            sheet1.write(i, 2, ans[i][2])
            sheet1.write(i, 3, ans[i][3])
            sheet1.write(i, 4, ans[i][4])
            sheet1.row(i).height_mismatch = True
            sheet1.row(i).height = 20 * 15
        # 保存在选择的位置
        FolderPath = filedialog.askdirectory(title="选择文件保存位置")  # 选择文件夹
        if FolderPath == '':
            messagebox.showwarning(title="导出失败", message="导出失败")
        book.save(f"{FolderPath}/账号信息.xls")
        return

    # 事件:添加/修改项目
    def btn_add_func(self):
        global tmp_num
        # 获取内容
        name = self.entry_name.get()
        remark = self.entry_remark.get()
        account = self.entry_account.get()
        password = self.entry_password.get()
        # 查询所有记录
        ans = self.db.executeQuery("select * from information")
        #记录是否有替换
        pd = False
        for i in ans:
            if name == i[1]:
                # name相同询问是否替换
                pd = messagebox.askokcancel(
                    title="替换",
                    message=f"是否替换记录\n{i[0]} {i[1]} {i[2]} {i[3]} {i[4]}")
                # 选择替换则更新数据库中记录
                if pd == True:
                    sql_update_item = f"update information set name='{name}',account='{account}',password='{password}',remark='{remark}' where id={i[0]}"
                    rows = self.db.executeUpdate(sql_update_item)
                    #替换后停止循环
                    break
        # 如果未做替换则进行添加信息
        if pd == False:
            sql_add_item = f"insert into information values({tmp_num+1},'{name}','{account}','{password}','{remark}')"
            rows = self.db.executeUpdate(sql_add_item)
        self.refresh_records_num()
        self.add_info_in_table()
        return

    # 事件:删除项目
    def btn_del_func(self):
        global tree
        # 未选中
        if len(self.tree.selection()) < 1:
            return
        # 询问是否确认删除
        s = '是否要删除以下信息:\n'
        for item in self.tree.selection():
            item_text = self.tree.item(item, "values")
            s += f'{item_text[0]} {item_text[1]} {item_text[2]} {item_text[3]} {item_text[4]}\n'
        pd = messagebox.askokcancel(title="删除", message=s)
        #选择不删除
        if pd == False:
            return
        #选择删除
        sql_delete_item = 'delete from information where id=? and name=? and account=? and password=? and remark=?'
        for item in self.tree.selection():  #逐条删除
            item_text = self.tree.item(item, "values")
            self.db.executeUpdate(sql_delete_item,
                                  (item_text[0], item_text[1], item_text[2],
                                   item_text[3], item_text[4]))
        # 刷新数据库中的数据 重新排序
        indb = []
        ans = self.db.executeQuery("select * from information")
        for i in range(0, len(ans)):
            indb.append((i + 1, ans[i][1], ans[i][2], ans[i][3], ans[i][4]))
        sql_clear_table = 'delete from information'
        self.db.executeUpdate(sql_clear_table)
        sql_add_info = 'insert into information values(?,?,?,?,?)'
        self.db.executeUpdate(sql_add_info, indb)
        # 刷新数据数量 表格内容
        self.refresh_records_num()
        self.add_info_in_table()
        return

    # 事件: 搜索项目
    def btn_searchinfo_func(self):
        # 获取搜索内容
        search = self.entry_search.get().lower()
        # 获取搜索类型
        select = self.combobox_selected.get()
        # 获取所有数据
        ans = self.db.executeQuery("select * from information")
        # 保存查找到的项目
        ans_list = []
        if select == '名称':
            for i in range(0, len(ans)):
                if search in ans[i][1].lower():
                    ans_list.append(str(i + 1))
        elif select == '账号':
            for i in range(0, len(ans)):
                if search in ans[i][2].lower():
                    ans_list.append(str(i + 1))
        elif select == '密码':
            for i in range(0, len(ans)):
                if search in ans[i][3]:
                    ans_list.append(str(i + 1))
        elif select == '备注':
            for i in range(0, len(ans)):
                if search in ans[i][4].lower():
                    ans_list.append(str(i + 1))
        self.tree.selection_set(ans_list)
        # entry_search.delete(0, tk.END)
        # self.tree.selection_add(iid)
        # self.tree.selection_set(["1", "8"])
        return

    # 事件:单击表格中元素的单击事件
    def tree_item_Click(self, event):
        global tree
        # self.tree.selection() 获取到所有选中
        # print(self.tree.selection())
        if len(self.tree.selection()) != 1:
            return
        item = self.tree.selection()[0]
        item_text = self.tree.item(item, "values")
        # 删除原有内容
        self.entry_name.delete(0, tk.END)
        self.entry_remark.delete(0, tk.END)
        self.entry_account.delete(0, tk.END)
        self.entry_password.delete(0, tk.END)
        # 显示点击的内容
        self.entry_name.insert(0, item_text[1])
        self.entry_account.insert(0, item_text[2])
        self.entry_password.insert(0, item_text[3])
        self.entry_remark.insert(0, item_text[4])
        # 复制账号密码到剪切板
        # s = item_text[2]
        # s = s.encode("utf-8")
        # xerox.copy(s)
        # xerox.copy(item_text[3])
        return


test = window()
