from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
import xlrd
import sqlite3
import datetime
import smtplib


class Window:
    def __init__(self, title, width=400, height=300, resizable=(False, False), icon=None):
        self.root = Tk()
        self.root.title(title)
        self.root.geometry(f"{width}x{height}+100+20")
        self.root.resizable(resizable[0], resizable[1])
        if icon:
            self.root.iconbitmap(icon)
        # self.label = Label(self.root, text="Статистика по задолженностям в разрезе ССП")

    def run(self):
        self.root.mainloop()

    def draw_label(self, pady, text="Статистика по задолженностям в разрезе ССП"):
        f = Frame(self.root)
        f.pack(pady=pady)
        Label(f, text=text).pack()

    def draw_stats(self):
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute("""
                    SELECT
                        DISTINCT date
                    FROM courses
                """)
        dates = cur.fetchall()
        dates_as_date = [datetime.datetime.strptime(d[0], "%Y-%m-%d").date() for d in dates] # переформат в даты
        last_date = dates_as_date[0].strftime('%Y-%m-%d')
        Label(self.root, text='Дата загрузки отчета из Пульс: '+last_date).pack()
        Label(self.root, text='').pack()

        query = f"""
            SELECT
                boss_mail,
                dep_3_level,
                dep_5_level,
                count(courses.emp_tab)
            FROM courses LEFT JOIN (SELECT * FROM status LEFT JOIN employees ON emp_tab = tab) as s 
            ON
                courses.emp_tab = s.emp_tab
            WHERE courses.date = "{last_date}" AND s.status = "болен"
            GROUP BY boss_mail
        """
        # print(query)
        cur.execute(query)
        result = cur.fetchall()

        for row in result:
            Label(self.root, text='Рук-ль: ' + row[0]).pack(anchor=NW)
            Label(self.root, text='Подразделение: ' + row[1] + '-->' + row[2]).pack(anchor=NW)
            Label(self.root, text='Кол-во незавершенных курсов: ' + str(row[3])).pack(anchor=NW)
            Label(self.root, text='--------------------------------').pack(anchor=NW)

    def draw_menu(self):
        menu_bar = Menu(self.root)
        import_menu = Menu(menu_bar, tearoff=0)
        import_menu.add_command(label="Импорт выгрузки из Пульса", command=self.import_pulse)
        import_menu.add_command(label="Импорт отчета по больным", command=self.import_illness_report)
        import_menu.add_command(label="Импорт ШР", command=self.import_statka)
        menu_bar.add_cascade(label="Импорт", menu=import_menu)

        tools_menu = Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Рассылка уведомлений", command=self.send_mails)
        tools_menu.add_command(label="Рассылка предупреждений")
        menu_bar.add_cascade(label="Инструменты", menu=tools_menu)

        # dict_menu = Menu(menu_bar, tearoff=0)
        # dict_menu.add_command(label="Штатное расписание", command=self.dict_employees)
        # menu_bar.add_cascade(label="Справочники", menu=dict_menu)
        self.root.configure(menu=menu_bar)

    def send_mails(self):
        smtpObj = smtplib.SMTP('smtp.mail.ru', 587)
        smtpObj.starttls()
        smtpObj.login('bb_sales@bk.ru', 'RebL7NiHRiphw6AX1Xqx')
        smtpObj.sendmail("bb_sales@bk.ru", "admitriev211@gmail.com", "Test autosend")
        smtpObj.quit()
        messagebox.showinfo('Внимание', 'Сообщение отправлено')


    wanted_files = (
        ("excel files", "*.xls;*.xlsx"),
    )

    def import_pulse(self, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title="Импорт выгрузки из Пульса", filetypes=wanted_files)
        if file_name:
            xl = xlrd.open_workbook(file_name, on_demand=True)
            sh = xl.sheet_by_index(0)
            course_list = []
            for row in range(4, sh.nrows - 1):
                # if sh.cell(row, 10).value == 'Байкальский банк' \
                #         and sh.cell(row, 13).value == 'Якутское отделение № 8603' \
                #         and sh.cell(row, 14).value == 'Аппарат отделения':
                course = [
                    str(datetime.date.today()),
                    sh.cell(row, 1).value,
                    sh.cell(row, 7).value,
                    sh.cell(row, 22).value,
                    sh.cell(row, 25).value,
                    sh.cell(row, 26).value,
                    sh.cell(row, 28).value,
                ]

                course_list.append(course)

            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS courses(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                emp_tab INT,
                emp_mail TEXT,
                boss_tab INT,
                boss_mail TEXT,
                course_name TEXT,
                deadline TEXT,
                FOREIGN KEY (emp_tab) REFERENCES employees(tab),
                FOREIGN KEY (boss_tab) REFERENCES employees(tab));
                """
            )
            conn.commit()

            cur.executemany("""
            INSERT INTO courses(
                date,
                emp_tab,
                emp_mail,
                boss_tab,
                boss_mail,
                course_name,
                deadline
            ) VALUES(?, ?, ?, ?, ?, ?, ? );
            """, course_list)
            conn.commit()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(course_list)) + ', обновлено записей: ' + str(
                len(course_list)))

    def import_illness_report(self, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title="Импорт отчета о заболевших", filetypes=wanted_files)
        if file_name:
            xl = xlrd.open_workbook(file_name, on_demand=True)
            sh = xl.sheet_by_index(0)
            sik_list = []
            for row in range(4, sh.nrows - 1):
                sik = [
                    str(datetime.date.today()),
                    sh.cell(row, 2).value,
                    sh.cell(row, 24).value,
                ]

                sik_list.append(sik)

            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS status(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                emp_tab INT,
                status TEXT,
                FOREIGN KEY (emp_tab) REFERENCES employees(tab)
                );
                """
            )
            conn.commit()

            cur.executemany("""
            INSERT INTO status(
                date,
                emp_tab,
                status
            ) VALUES(?, ?, ?);
            """, sik_list)
            conn.commit()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(sik_list)) + ', обновлено записей: ' + str(
                len(sik_list)))

    def import_statka(self, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title="Импорт ШР", filetypes=wanted_files)
        if file_name:
            xl = xlrd.open_workbook(file_name, on_demand=True)
            sh = xl.sheet_by_index(0)
            record_list = []
            for row in range(1, sh.nrows - 1):
                if sh.cell(row,12).value != 0.0:
                    try:
                        record = (
                        int(sh.cell(row, 12).value),
                        sh.cell(row, 2).value,
                        sh.cell(row, 4).value,
                        )
                        record_list.append(record)
                    except:
                        print(str(row))

            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS employees(
                tab INT PRIMARY KEY,
                dep_3_level TEXT,
                dep_5_level TEXT);
                """
            )
            conn.commit()
            cur.execute("SELECT * FROM employees;")
            results = cur.fetchall()
            tabs = [r[0] for r in results]
            insert_list = [r for r in record_list if r[0] not in tabs]
            update_list = [r for r in record_list if r[0] in tabs]
            cur.executemany("UPDATE employees set dep_3_level = ?, dep_5_level = ? where tab = ?;", update_list)
            conn.commit()
            cur.executemany("INSERT INTO employees VALUES(?, ?, ?);", insert_list)
            conn.commit()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(insert_list)) + ', обновлено записей: ' + str(len(update_list)))

    def dict_employees(self):
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        # cur.execute("SELECT * FROM employees;")
        # cur.execute("SELECT * FROM courses;")
        cur.execute("SELECT * FROM status;")
        results = cur.fetchall()
        for rec in results:
            Label(self.root, text=','.join([str(r) for r in rec])).pack()
