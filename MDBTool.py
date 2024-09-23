import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyodbc
import os
import subprocess  # 新增，用于打开文件夹

class MDBReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MDB 文件读取器")

        # 初始化变量
        self.conn = None
        self.cursor = None
        self.table_names = []
        self.columns = []
        self.selected_columns = []
        self.data = []
        self.output_file_path = None  # 新增，存储输出文件的路径

        # 创建 GUI 组件
        self.create_widgets()

    def create_widgets(self):
        # 文件选择部分
        self.file_frame = tk.Frame(self.root)
        self.file_frame.pack(pady=10)

        self.file_label = tk.Label(self.file_frame, text="选择 MDB 文件:")
        self.file_label.pack(side=tk.LEFT)

        self.file_entry = tk.Entry(self.file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)

        self.browse_button = tk.Button(self.file_frame, text="浏览", command=self.browse_file)
        self.browse_button.pack(side=tk.LEFT)

        self.connect_button = tk.Button(self.root, text="连接数据库", command=self.connect_db)
        self.connect_button.pack(pady=5)

        # 表格选择部分
        self.table_frame = tk.Frame(self.root)
        self.table_label = tk.Label(self.table_frame, text="选择表格:")
        self.table_listbox = tk.Listbox(self.table_frame, height=5, exportselection=False)
        self.table_listbox.bind('<<ListboxSelect>>', self.on_table_select)

        # 字段选择部分
        self.column_frame = tk.Frame(self.root)
        self.column_label = tk.Label(self.column_frame, text="选择字段:")
        self.column_listbox = tk.Listbox(self.column_frame, height=10, selectmode=tk.MULTIPLE)

        # 按钮部分
        self.button_frame = tk.Frame(self.root)
        self.show_button = tk.Button(self.button_frame, text="显示数据", command=self.show_data)
        self.export_button = tk.Button(self.button_frame, text="输出到 txt", command=self.export_to_txt)
        self.open_folder_button = tk.Button(self.button_frame, text="打开输出文件夹", command=self.open_output_folder)

        # 数据显示部分
        self.data_frame = tk.Frame(self.root)
        self.data_text = tk.Text(self.data_frame, height=15, width=80)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Access Databases", "*.mdb;*.accdb"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def connect_db(self):
        mdb_file = self.file_entry.get()
        if not os.path.isfile(mdb_file):
            messagebox.showerror("错误", f"文件 {mdb_file} 不存在，请检查路径是否正确。")
            return

        # **关闭之前的数据库连接和游标**
        self.close_connection()

        # 构建连接字符串
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            rf'DBQ={mdb_file};'
        )

        try:
            # 创建数据库连接
            self.conn = pyodbc.connect(conn_str)
            self.cursor = self.conn.cursor()
        except pyodbc.Error as e:
            messagebox.showerror("错误", f"连接到数据库时出错：\n{e}")
            return

        # 获取所有表名
        self.table_names = [table.table_name for table in self.cursor.tables(tableType='TABLE')]

        if not self.table_names:
            messagebox.showinfo("信息", "数据库中没有可用的表。")
            return

        # **重置界面元素和变量**
        self.reset_interface()

        # 显示表格选择部分
        self.table_frame.pack(pady=5)
        self.table_label.pack()
        self.table_listbox.pack()
        for table_name in self.table_names:
            self.table_listbox.insert(tk.END, table_name)

    def on_table_select(self, event):
        selection = event.widget.curselection()
        if not selection:
            return

        index = selection[0]
        selected_table = self.table_names[index]

        # 获取字段名
        try:
            self.cursor.execute(f"SELECT * FROM [{selected_table}]")
            self.columns = [column[0] for column in self.cursor.description]
        except pyodbc.Error as e:
            messagebox.showerror("错误", f"获取字段名时出错：\n{e}")
            return

        # **清空并重置字段列表、数据展示等**
        self.column_listbox.delete(0, tk.END)
        self.selected_columns = []
        self.data = []
        self.data_text.delete(1.0, tk.END)

        # 显示字段选择部分
        self.column_frame.pack(pady=5)
        self.column_label.pack()
        self.column_listbox.pack()
        for column_name in self.columns:
            self.column_listbox.insert(tk.END, column_name)

        # 显示按钮部分
        self.button_frame.pack(pady=5)
        self.show_button.pack(side=tk.LEFT, padx=5)
        self.export_button.pack(side=tk.LEFT, padx=5)
        self.open_folder_button.pack(side=tk.LEFT, padx=5)

    def show_data(self):
        if not self.get_selected_columns():
            return

        # 获取所选表格名
        table_index = self.table_listbox.curselection()[0]
        selected_table = self.table_names[table_index]

        # 执行查询
        query = f"SELECT {', '.join(f'[{col}]' for col in self.selected_columns)} FROM [{selected_table}]"
        try:
            self.cursor.execute(query)
            self.data = self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("错误", f"执行查询时出错：\n{e}")
            return

        # 显示数据
        self.data_frame.pack(pady=5)
        self.data_text.pack()
        self.data_text.delete(1.0, tk.END)
        # 添加表头
        header = ', '.join(self.selected_columns)
        self.data_text.insert(tk.END, header + '\n')
        self.data_text.insert(tk.END, '-' * len(header) + '\n')
        for row in self.data:
            row_str = ', '.join(f"{value}" for value in row)
            self.data_text.insert(tk.END, row_str + '\n')

    def get_selected_columns(self):
        selected_indices = self.column_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("信息", "请至少选择一个字段。")
            return False

        self.selected_columns = [self.columns[i] for i in selected_indices]
        return True

    def export_to_txt(self):
        if not self.get_selected_columns():
            return

        # 获取所选表格名
        table_index = self.table_listbox.curselection()[0]
        selected_table = self.table_names[table_index]

        # 执行查询，获取数据
        query = f"SELECT {', '.join(f'[{col}]' for col in self.selected_columns)} FROM [{selected_table}]"
        try:
            self.cursor.execute(query)
            data = self.cursor.fetchall()
        except pyodbc.Error as e:
            messagebox.showerror("错误", f"执行查询时出错：\n{e}")
            return

        # 选择保存文件的路径
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            initialfile="output.txt"
        )
        if file_path:
            # 存储输出文件的路径
            self.output_file_path = file_path

            # 检查目录是否存在，如果不存在则创建
            directory = os.path.dirname(file_path)
            if not os.path.exists(directory):
                try:
                    os.makedirs(directory)
                except Exception as e:
                    messagebox.showerror("错误", f"无法创建目录 {directory}：\n{e}")
                    return

            try:
                # 写入数据到 txt 文件
                with open(file_path, 'w', encoding='utf-8') as f:
                    # 写入表头
                    f.write(', '.join(self.selected_columns) + '\n')
                    f.write('-' * 50 + '\n')
                    # 写入数据行
                    for row in data:
                        row_str = ', '.join(str(value) for value in row)
                        f.write(row_str + '\n')
                messagebox.showinfo("成功", f"数据已成功导出到 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出到 txt 时出错：\n{e}")

    def open_output_folder(self):
        if self.output_file_path:
            directory = os.path.dirname(self.output_file_path)
            try:
                # 在 Windows 上打开文件夹
                if os.name == 'nt':
                    os.startfile(directory)
                # 在 macOS 上打开文件夹
                elif os.name == 'posix':
                    subprocess.Popen(['open', directory])
                # 在 Linux 上打开文件夹
                else:
                    subprocess.Popen(['xdg-open', directory])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件夹 {directory}：\n{e}")
        else:
            messagebox.showinfo("信息", "还没有导出文件。")

    def close_connection(self):
        """关闭数据库连接和游标"""
        # 尝试关闭游标
        if self.cursor is not None:
            try:
                self.cursor.close()
            except pyodbc.Error as e:
                print(f"关闭游标时出错：{e}")
            self.cursor = None

        # 尝试关闭连接
        if self.conn is not None:
            try:
                self.conn.close()
            except pyodbc.Error as e:
                print(f"关闭连接时出错：{e}")
            self.conn = None

    def reset_interface(self):
        """重置界面元素和相关变量"""
        # 清空表格列表
        self.table_listbox.delete(0, tk.END)
        # 清空字段列表
        self.column_listbox.delete(0, tk.END)
        # 清空数据展示
        self.data_text.delete(1.0, tk.END)
        # 重置变量
        self.columns = []
        self.selected_columns = []
        self.data = []
        self.output_file_path = None

    def close(self):
        # 关闭数据库连接和游标
        self.close_connection()
        # 销毁主窗口
        try:
            self.root.destroy()
        except tk.TclError:
            pass  # 如果窗口已经被销毁，忽略异常

def main():
    root = tk.Tk()
    app = MDBReaderApp(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()

if __name__ == "__main__":
    main()