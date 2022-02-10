from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from child_window import ChildWindow
from email.mime.text import MIMEText
from email.header import Header
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
        self.server = Entry(self.root, width=100)
        self.login = Entry(self.root, width=100)
        self.password = Entry(self.root, width=100)
        self.text_for_letter = None
        self.text_on_screen = None

    def run(self):
        self.root.mainloop()

    def create_child(self, width, height, title, resizable=(False,False), icon=None):
        ChildWindow(self.root, width, height, title, resizable, icon)

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

        query = f"""
            SELECT
                a.boss_mail,
                a.dep_3_level,
                a.dep_5_level,
                b.employees_count,
                b.on_distant,
                a.course_count
            FROM (
            SELECT
                boss_mail,
                dep_3_level,
                dep_5_level,
                count(courses.emp_tab) as course_count
            FROM (SELECT * FROM employees LEFT JOIN status ON emp_tab = tab) as s LEFT JOIN courses  
            ON
                courses.emp_tab = s.emp_tab
            WHERE courses.date = "{last_date}" AND s.status = "болен"
            GROUP BY boss_mail) as a
            LEFT JOIN (SELECT dep_3_level, dep_5_level, count(tab) as employees_count, count(status) as on_distant FROM employees LEFT JOIN (SELECT * FROM status WHERE status = "болен") ON tab = emp_tab GROUP BY dep_3_level, dep_5_level) as b
            ON a.dep_3_level = b.dep_3_level AND a.dep_5_level = b.dep_5_level
        """
        print(query)
        cur.execute(query)
        result = cur.fetchall()
        print(result)



        scroll_bar = Scrollbar(self.root)
        scroll_bar.pack(side=RIGHT, fill=Y)

        text_widget = Listbox(self.root, width=800 , yscrollcommand = scroll_bar.set)

        text_widget.insert(END, 'Дата загрузки отчета из Пульс: ' + last_date +'\n')
        text_widget.insert(END, '' + '\n')

        for row in result:
            text_widget.insert(END, 'Рук-ль: ' + row[0] +'\n')
            text_widget.insert(END, 'Подразделение: ' + row[1] + '-->' + row[2] +'\n')
            text_widget.insert(END, 'Всего сотрудников: ' + str(row[3]) +'\n')
            text_widget.insert(END, 'Находятся дома: ' + str(row[4]) +'\n')
            text_widget.insert(END, 'Кол-во незавершенных курсов у сотрудников, находящихся дома: ' + str(row[5]) +'\n')
            text_widget.insert(END, '---------------------------------------------' + '\n')

        text_widget.pack(side = LEFT, fill=BOTH)
        scroll_bar.config(command=text_widget.yview)

        #
        # for row in result:
        #     Label(frame, text='Рук-ль: ' + row[0]).pack(anchor=NW)
        #     Label(frame, text='Подразделение: ' + row[1] + '-->' + row[2]).pack(anchor=NW)
        #     Label(frame, text='Всего сотрудников: ' + str(row[3])).pack(anchor=NW)
        #     Label(frame, text='Находятся дома: ' + str(row[4])).pack(anchor=NW)
        #     Label(frame, text='Кол-во незавершенных курсов у сотрудников, находящихся дома: ' + str(row[5])).pack(anchor=NW)
        #     Label(frame, text='--------------------------------').pack(anchor=NW)

    def draw_menu(self):
        menu_bar = Menu(self.root)
        import_menu = Menu(menu_bar, tearoff=0)
        import_menu.add_command(label="Импорт выгрузки из Пульса", command=self.import_pulse)
        import_menu.add_command(label="Импорт отчета по больным", command=self.import_illness_report)
        import_menu.add_command(label="Импорт ШР", command=self.import_statka)
        menu_bar.add_cascade(label="Импорт", menu=import_menu)

        tools_menu = Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Рассылка уведомлений", command=self.send_mail_form)
        # tools_menu.add_command(label="Рассылка уведомлений", command=self.send_mails)
        tools_menu.add_command(label="Рассылка предупреждений")
        menu_bar.add_cascade(label="Инструменты", menu=tools_menu)

        # dict_menu = Menu(menu_bar, tearoff=0)
        # dict_menu.add_command(label="Штатное расписание", command=self.dict_employees)
        # menu_bar.add_cascade(label="Справочники", menu=dict_menu)
        self.root.configure(menu=menu_bar)

    def send_mail_form(self):
        send_form = ChildWindow(self.root, 400, 400, "Отправка уведомлений")

        Label(send_form.root, text="Введите сервер").pack()
        self.server = Entry(send_form.root, width=100)
        self.server.pack()
        Label(send_form.root, text="Введите логин").pack()
        self.login = Entry(send_form.root, width=100)
        self.login.pack()
        Label(send_form.root, text="Введите пароль").pack()
        self.password = Entry(send_form.root, width=100)
        self.password.pack()
        Button(send_form.root, text="Файл с текстом письма", command=self.get_text_for_letter).pack()
        Button(send_form.root, text="Разослать уведомления", command=self.send_mails).pack()
        scroll_bar = Scrollbar(send_form.root)
        scroll_bar.pack(side=RIGHT, fill=Y)
        self.text_on_screen = Text(send_form.root, width=400, height=300, wrap = WORD, yscrollcommand = scroll_bar.set)
        self.text_on_screen.pack()
        send_form.grab_focus()

    def send_mails(self):
        if self.text_for_letter:
            msg = MIMEText(self.text_for_letter)
            msg['Subject'] = Header('Пройдите курсы в Пульс!', 'utf-8')
            msg['From'] = self.login.get()
            msg['To'] = "admitriev211@gmail.com"

            smtpObj = smtplib.SMTP(self.server.get(), 587)
            smtpObj.starttls()
            smtpObj.login(self.login.get(), self.password.get())
            try:
                smtpObj.sendmail(self.login.get(), "admitriev211@gmail.com", msg.as_string())
                smtpObj.quit()
                # smtpObj = smtplib.SMTP('smtp.mail.ru', 587)
                # smtpObj.starttls()
                # smtpObj.login('bb_sales@bk.ru', 'DSPQV5c5NFM7G2mY7bPM')
                # smtpObj.sendmail("bb_sales@bk.ru", "admitriev211@gmail.com", "Test autosend")
                # smtpObj.quit()
                messagebox.showinfo('Внимание', 'Сообщение отправлено')
            except Exception as e:
                print(e)
                messagebox.showinfo('Внимание', 'Что-то пошло не так')
        else:
            messagebox.showinfo('Внимание', 'Файл с текстом не прочитан')


    wanted_files = (
        ("excel files", "*.xls;*.xlsx"),
    )

    def get_text_for_letter(self):
        file_name = fd.askopenfilename(title="Выберите файл с текстом письма")
        if file_name:
            with open(file_name, encoding = 'utf-8', mode='r') as f:
                self.text_for_letter = f.read()
                self.text_on_screen.insert(END, self.text_for_letter)
            print(self.text_for_letter)


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
