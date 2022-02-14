from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter.ttk import Combobox
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
        self.top = None
        self.server = Entry(self.root, width=100)
        self.login = Entry(self.root, width=100)
        self.password = Entry(self.root, width=100)
        self.text_for_letter = None
        self.text_on_screen = None
        self.importDate = None
        self.chosen_file=None
        self.list_for_import_pulse = None
        self.list_for_import_drug = None
        self.c = None
        self.scroll_bar = Scrollbar(self.root)
        self.scroll_bar.pack(side=RIGHT, fill=Y)
        self.text_widget = None
        self.last_date_pulse=None
        self.last_date_distant = None
        self.draw_stats()
        # try:
        #     self.draw_stats()
        # except Exception as e:
        #     print(e)
        #     messagebox.showinfo("Ошибка запроса",e)

    def run(self):
        self.root.mainloop()

    def create_child(self, width, height, title, resizable=(False,False), icon=None):
        ChildWindow(self.root, width, height, title, resizable, icon)

    def get_data_for_dash(self):
        def find_last_date(table):
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(f"""
                                SELECT
                                    DISTINCT date
                                FROM {table}
                            """)
            dates = cur.fetchall()
            cur.close()
            dates_as_date = sorted([datetime.datetime.strptime(d[0], "%Y-%m-%d").date() for d in dates], key=lambda x: x,
                                   reverse=True)  # переформат в даты и сортировка
            return dates_as_date[0].strftime('%Y-%m-%d')

        self.last_date_pulse = find_last_date('courses')
        self.last_date_distant = find_last_date('status')

        query = f"""
                    SELECT
                        dep_3_level,
                        dep_5_level,
                        count(tab) as emp_count,
                        count(status.emp_tab) as distant_count,
                        count(c.course_tab) as course_count
                    FROM employees
                    LEFT JOIN status
                    ON tab = status.emp_tab and status.status = "болен"
                    LEFT JOIN (
                        SELECT
                            courses.emp_tab as course_tab,
                            status
                        FROM courses
                        LEFT JOIN status
                        ON courses.emp_tab = status.emp_tab
                        WHERE status = "болен"
                        ) as c
                    ON tab = c.course_tab
                    GROUP BY
                        dep_3_level,
                        dep_5_level
                    ORDER BY course_count DESC, distant_count DESC
                              
                """
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()
        cur.close()
        return result

    def draw_stats(self):
        gotData = self.get_data_for_dash()
        self.text_widget = Listbox(self.root, width=800, yscrollcommand = self.scroll_bar.set)

        self.text_widget.insert(END, 'Дата загрузки отчета из Пульс: ' + self.last_date_pulse +'\n')
        self.text_widget.insert(END, 'Дата загрузки отчета из ДРУГ: ' + self.last_date_distant + '\n')

        self.text_widget.insert(END, '' + '\n')
        item = 3

        for i in range(0, len(gotData)):
            row = gotData[i]
            # self.text_widget.insert(END, 'Рук-ль: ' + row[0] +'\n')
            self.text_widget.insert(END, 'Подразделение: ' + row[0] + '-->' + row[1] +'\n')
            self.text_widget.insert(END, 'Всего сотрудников: ' + str(row[2]) +'\n')
            self.text_widget.insert(END, 'Находятся дома: ' + str(row[3]) +'\n')
            item += 3
            self.text_widget.insert(END, 'Кол-во незавершенных курсов у сотрудников, находящихся дома: ' + str(row[4]) +'\n')
            self.text_widget.insert(END, '---------------------------------------------' + '\n')
            if row[4] > 0:
                print(str(row[4]),str(item))
                self.text_widget.itemconfig(item, bg='red')
            item += 2

        self.text_widget.pack(side = LEFT, fill=BOTH)
        self.scroll_bar.config(command=self.text_widget.yview)

    wanted_files = (
        ("excel files", "*.xls;*.xlsx"),
    )

    def draw_menu(self):
        menu_bar = Menu(self.root)
        # import_menu = Menu(menu_bar, tearoff=0)
        # import_menu.add_command(label="Импорт выгрузки из Пульса", command=self.import_pulse_form)
        # import_menu.add_command(label="Импорт отчета по больным", command=self.import_illness_report)
        # import_menu.add_command(label="Импорт ШР", command=self.import_statka)
        # menu_bar.add_cascade(label="Импорт", menu=import_menu)

        tools_menu = Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Рассылка уведомлений", command=self.send_mail_form)
        # tools_menu.add_command(label="Рассылка предупреждений")
        menu_bar.add_cascade(label="Инструменты", menu=tools_menu)

        reports_menu = Menu(menu_bar, tearoff=0)
        reports_menu.add_command(label="Отчеты из Пульса", command=self.pulse_reports)
        reports_menu.add_command(label="Отчеты из ДРУГа", command=self.drug_reports)
        menu_bar.add_cascade(label="Отчеты", menu=reports_menu)

        # dict_menu = Menu(menu_bar, tearoff=0)
        # dict_menu.add_command(label="Штатное расписание", command=self.dict_employees)
        # menu_bar.add_cascade(label="Справочники", menu=dict_menu)
        self.root.configure(menu=menu_bar)

    def pulse_reports(self):
        pulse_reports_window = ChildWindow(self.root, 300, 200, "Отчеты из Пульс")
        self.top = pulse_reports_window.root
        Label(pulse_reports_window.root, text="Выберите дату отчета").pack()

        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute("""
                            SELECT
                                DISTINCT date
                            FROM courses
                        """)
        dates = cur.fetchall()
        dates_as_date = sorted([datetime.datetime.strptime(d[0], "%Y-%m-%d").date() for d in dates], key=lambda x: x,
                               reverse=True)  # переформат в даты и сортировка
        values = tuple(d.strftime('%Y-%m-%d') for d in dates_as_date)
        self.c = Combobox(pulse_reports_window.root, width=25, values=values)
        self.c.current(0)
        self.c.pack()

        Button(pulse_reports_window.root, width=23, text="Удалить отчет", command=self.del_pulse_confirmation).pack()

        Button(pulse_reports_window.root, width=23, text="Загрузить отчет", command=self.import_pulse_form).pack(pady=10)
        pulse_reports_window.grab_focus()

    def drug_reports(self):
        pulse_reports_window = ChildWindow(self.root, 300, 200, "Отчеты из ДРУГ")
        self.top = pulse_reports_window.root
        Label(pulse_reports_window.root, text="Выберите дату отчета").pack()

        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute("""
                            SELECT
                                DISTINCT date
                            FROM status
                        """)
        dates = cur.fetchall()
        dates_as_date = sorted([datetime.datetime.strptime(d[0], "%Y-%m-%d").date() for d in dates], key=lambda x: x,
                               reverse=True)  # переформат в даты и сортировка
        values = tuple(d.strftime('%Y-%m-%d') for d in dates_as_date)
        self.c = Combobox(pulse_reports_window.root, width=25, values=values)
        self.c.current(0)
        self.c.pack()

        Button(pulse_reports_window.root, width=23, text="Удалить отчет", command=self.del_drug_confirmation).pack()

        Button(pulse_reports_window.root, width=23, text="Загрузить отчет", command=self.import_drug_form).pack(pady=10)
        pulse_reports_window.grab_focus()

    def del_report_course(self):
        date_for_del = self.c.get()
        print(date_for_del)
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(f"""
                                    DELETE FROM courses WHERE date = "{date_for_del}"
                                """)
        conn.commit()
        cur.close()
        self.top.destroy()
        self.top.update()
        self.root.update()
        messagebox.showinfo('Внимание', 'Отчет за ' + date_for_del + ' удален')
        self.text_widget.destroy()
        self.text_widget = None
        self.draw_stats()

    def del_pulse_confirmation(self):
        result = messagebox.askyesno(
            title="Подтверждение удаления",
            message="Вы действительно хотите удалить отчет"
        )
        if result:
            self.del_report_course()
        else:
            exit()

    def del_report_drug(self):
        date_for_del = self.c.get()
        print(date_for_del)
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(f"""
                                    DELETE FROM status WHERE date = "{date_for_del}"
                                """)
        conn.commit()
        cur.close()
        self.top.destroy()
        self.top.update()
        self.root.update()
        messagebox.showinfo('Внимание', 'Отчет за ' + date_for_del + ' удален')
        self.text_widget.destroy()
        self.text_widget = None
        self.draw_stats()

    def del_drug_confirmation(self):
        result = messagebox.askyesno(
            title="Подтверждение удаления",
            message="Вы действительно хотите удалить отчет"
        )
        if result:
            self.del_report_drug()
        else:
            exit()

    def import_pulse_form(self):
        pulse_form = ChildWindow(self.root, 200, 300, "Импорт выгрузки из Пульс")
        self.top = pulse_form.root
        Label(pulse_form.root, text="Дата отчета").pack()
        self.importDate = DateEntry(pulse_form.root, width=25, date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.importDate.pack()
        Button(pulse_form.root, width=25, text="Выбрать файл", command=self.get_file_pulse).pack()
        self.chosen_file = Text(pulse_form.root)
        Button(pulse_form.root, width=25, text="Импортировать", command=self.import_file_pulse).pack(pady=10)
        pulse_form.grab_focus()

    def get_file_pulse(self, wanted_files=wanted_files):
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
                    sh.cell(row, 1).value,
                    sh.cell(row, 7).value,
                    sh.cell(row, 22).value,
                    sh.cell(row, 25).value,
                    sh.cell(row, 26).value,
                    sh.cell(row, 28).value,
                ]

                course_list.append(course)
            self.list_for_import_pulse = course_list
            self.chosen_file.insert(END, file_name)
            self.chosen_file.pack()

    def import_file_pulse(self):
        if self.list_for_import_pulse and self.importDate.get_date():
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

            cur.execute(
                f"""
                DELETE FROM courses WHERE date = "{str(self.importDate.get_date())}"
                """
            )
            conn.commit()

            import_list=[]
            for i in self.list_for_import_pulse:
                row=[str(self.importDate.get_date())]
                for j in i:
                    row.append(j)
                import_list.append(row)

            print(import_list)
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
                        """, import_list)
            conn.commit()
            self.top.destroy()
            self.top.update()
            self.root.update()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(self.list_for_import_pulse)))
            self.text_widget.destroy()
            self.text_widget = None
            self.draw_stats()

    def import_drug_form(self):
        drug_form = ChildWindow(self.root, 200, 300, "Импорт данных из ДРУГ")
        self.top = drug_form.root
        Label(drug_form.root, text="Дата отчета").pack()
        self.importDate = DateEntry(drug_form.root, width=25, date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.importDate.pack()
        Button(drug_form.root, width=25, text="Выбрать файл", command=self.get_file_drug).pack()
        self.chosen_file = Text(drug_form.root)
        Button(drug_form.root, width=25, text="Импортировать", command=self.import_file_drug).pack(pady=10)
        drug_form.grab_focus()

    def get_file_drug(self, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title="Импорт данных из Пульса", filetypes=wanted_files)
        if file_name:
            xl = xlrd.open_workbook(file_name, on_demand=True)
            sh = xl.sheet_by_index(0)
            drug_list = []
            for row in range(4, sh.nrows - 1):
                # if sh.cell(row, 10).value == 'Байкальский банк' \
                #         and sh.cell(row, 13).value == 'Якутское отделение № 8603' \
                #         and sh.cell(row, 14).value == 'Аппарат отделения':
                drug = [
                    sh.cell(row, 2).value,
                    sh.cell(row, 24).value,
                ]

                drug_list.append(drug)
            self.list_for_import_drug = drug_list
            self.chosen_file.insert(END, file_name)
            self.chosen_file.pack()

    def import_file_drug(self):
        if self.list_for_import_drug and self.importDate.get_date():
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

            cur.execute(
                f"""
                DELETE FROM status WHERE date = "{str(self.importDate.get_date())}"
                """
            )
            conn.commit()

            import_list=[]
            for i in self.list_for_import_drug:
                row=[str(self.importDate.get_date())]
                for j in i:
                    row.append(j)
                import_list.append(row)

            print(import_list)
            cur.executemany("""
                INSERT INTO status(
                    date,
                    emp_tab,
                    status
                ) VALUES(?, ?, ?);
                """, import_list)
            conn.commit()
            self.top.destroy()
            self.top.update()
            self.root.update()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(self.list_for_import_drug)))
            self.text_widget.destroy()
            self.text_widget = None
            self.draw_stats()

    def send_mail_form(self):
        send_form = ChildWindow(self.root, 400, 400, "Отправка уведомлений")
        self.top = send_form.root
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
        self.text_on_screen = Text(send_form.root, width=400, height=300, wrap=WORD, yscrollcommand=scroll_bar.set)
        self.text_on_screen.pack()
        send_form.grab_focus()

    def send_mails(self):
        if self.text_for_letter:
            query = f"""
                SELECT
                    emp_mail,
                    boss_mail,
                    course_name,
                    deadline 
                FROM courses
                LEFT JOIN status
                ON courses.emp_tab = status.emp_tab and status.status = "болен"
                WHERE status.status = "болен"
                     """
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(query)
            result = cur.fetchall()
            cur.close()

            to_list = list(set([r[0] for r in result]))

            courses_dict = {
                reciever: [
                    [
                        course[2],
                        course[3],
                        course[1]
                    ] for course in result if course[0] == reciever
                ] for reciever in to_list
            }

            print(courses_dict)

            try:
                for k, v in courses_dict.items():
                    course_list = ''
                    for course in v:
                        course_list += course[0] + '. Срок: ' + course[1] + '\n'
                    msg_text = f'''
                    to {k} \n
                    {self.text_for_letter}:\n
                    {course_list}                    
                    '''
                    smtpObj = smtplib.SMTP(self.server.get(), 587)
                    smtpObj.starttls()
                    smtpObj.login(self.login.get(), self.password.get())

                    msg = MIMEText(msg_text)
                    msg['Subject'] = Header('Пройдите курсы в Пульс!', 'utf-8')
                    msg['From'] = self.login.get()
                    msg['To'] = "nvivanchikov@sberbank.ru"
                    msg['CC'] = "arsadmitriev@sberbank.ru"
                    smtpObj.sendmail(self.login.get(), ['nvivanchikov@sberbank.ru', 'arsadmitriev@sberbank.ru'],
                                     msg.as_string())
                    # msg['To'] = k
                    # msg['CC'] = course[2]
                    # smtpObj.sendmail(self.login.get(), [k, course[2]],msg.as_string())
                    smtpObj.quit()

                    # smtpObj = smtplib.SMTP('smtp.mail.ru', 587)
                    # smtpObj.starttls()
                    # smtpObj.login('bb_sales@bk.ru', 'DSPQV5c5NFM7G2mY7bPM')
                    # smtpObj.sendmail("bb_sales@bk.ru", "admitriev211@gmail.com", "Test autosend")

                self.top.destroy()
                messagebox.showinfo('Внимание', 'Сообщения отправлены')

            except Exception as e:
                print(e)
                self.top.destroy()
                messagebox.showinfo('Внимание', 'Что-то пошло не так')
        else:
            messagebox.showinfo('Внимание', 'Файл с текстом не прочитан')

    def get_text_for_letter(self):
        file_name = fd.askopenfilename(title="Выберите файл с текстом письма")
        if file_name:
            with open(file_name, encoding = 'utf-8', mode='r') as f:
                self.text_for_letter = f.read()
                self.text_on_screen.insert(END, self.text_for_letter)
            print(self.text_for_letter)

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
