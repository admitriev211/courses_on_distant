from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkcalendar import DateEntry
from tkinter.ttk import Combobox
from child_window import ChildWindow
from email.mime.text import MIMEText
from email.header import Header
import xlrd
import xlwt
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
        self.list_for_import = None
        self.list_for_import_pulse = None
        self.list_for_import_drug = None
        self.c = None
        self.scroll_bar = Scrollbar(self.root)
        self.scroll_bar.pack(side=RIGHT, fill=Y)
        self.text_widget = None
        self.last_date_pulse=None
        self.last_date_distant = None
        # self.draw_stats()
        try:
            self.draw_stats()
        except Exception as e:
            print(e)
            messagebox.showinfo("Ошибка запроса",e)

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
        self.last_date_vacations = find_last_date('vacations')

        query = f"""
            SELECT
                t1.dep_3_level,
                t1.dep_5_level,
                t1.tabs,
                t1.distant_count,
                t1.course_count,
                t2.days
            FROM (
                SELECT
                    dep_3_level,
                    dep_5_level,
                    count(DISTINCT tab) as tabs,
                    count(DISTINCT status.emp_tab) as distant_count,
                    count(DISTINCT courses.course_name) as course_count
                FROM employees
                LEFT JOIN status
                ON tab = status.emp_tab and status.date = "{self.last_date_distant}"
                LEFT JOIN courses
                ON status.emp_tab = courses.emp_tab and courses.date = "{self.last_date_pulse}"
                GROUP BY dep_3_level, dep_5_level
                ORDER BY course_count DESC, distant_count DESC
                ) as t1
            LEFT JOIN (
                SELECT
                    dep_3_level,
                    dep_5_level,
                    count(DISTINCT tab),
                    count(DISTINCT status.emp_tab),
                    sum(days_left) as days
                FROM employees
                LEFT JOIN status
                ON tab = status.emp_tab and status.date = "{self.last_date_distant}"
                LEFT JOIN vacations
                ON status.emp_tab = vacations.emp_tab and vacations.date = "{self.last_date_vacations}"
                GROUP BY
                    dep_3_level,
                    dep_5_level
                ) as t2
            ON t1.dep_3_level = t2.dep_3_level and t1.dep_5_level = t2.dep_5_level

        """

        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()
        print(result)
        # print(str(len(result)))
        cur.close()
        return result

    def report_name(self, table):
        names = {
            'courses': 'Отчет из Пульс',
            'status': 'Данные о заболевших',
            'vacations': 'Данные об отпусках'
        }
        return names[table]

    def create_query(self, table):
        if table == 'courses':
            create_query = """
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
        elif table == 'status':
            create_query = """
                CREATE TABLE IF NOT EXISTS status(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                emp_tab INTEGER,
                status TEXT,
                FOREIGN KEY (emp_tab) REFERENCES employees(tab)
                );
            """
        elif table == 'vacations':
            create_query = """
                CREATE TABLE IF NOT EXISTS vacations(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                emp_tab INTEGER,
                days_left REAL,
                FOREIGN KEY (emp_tab) REFERENCES employees(tab)
                );
            """
        return create_query

    def insert_query(self, table):
        if table == 'courses':
            insert_query = """
                INSERT INTO courses(
                    date,
                    emp_tab,
                    emp_mail,
                    boss_tab,
                    boss_mail,
                    course_name,
                    deadline
                ) VALUES(?, ?, ?, ?, ?, ?, ? );
                        """
        elif table == 'status':
            insert_query = """
                INSERT INTO status(
                    date,
                    emp_tab,
                    status
                ) VALUES(?, ?, ?);
            """
        elif table == 'vacations':
            insert_query = """
                            INSERT INTO vacations(
                                date,
                                emp_tab,
                                days_left
                            ) VALUES(?, ?, ?);
                        """
        return insert_query

    def parse_excel(self, table, sh):
        header = {
            'courses': 1,
            'status': 0,
            'vacations': 3
        }
        def find_cols(table):
            cols_list=[]
            fields = {
                'courses': [
                    'ТН',
                    'Внешняя почта',
                    'ТН руководителя',
                    'Внешняя почта руководителя',
                    'Наименование курса',
                    'Контрольная дата прохождения'
                ],
                'status': [
                    'Табельный номер',
                    'Статус'
                ],
                'vacations': [
                    'ТН',
                    'годнакоплено дней'
                ]
            }

            for f in fields[table]:
                for col in range(header[table], sh.ncols):
                    if sh.cell(header[table], col).value == f:
                        cols_list.append(col)

            return cols_list

        row_list = []
        cols = find_cols(table)
        print(cols)
        for row in range(header[table]+1, sh.nrows):
            line = [sh.cell(row, col).value for col in cols if sh.cell(row, cols[0]).value != '']
            if len(line) == len(cols):
                row_list.append(line)
        return row_list

    def draw_stats(self):
        gotData = [d for d in self.get_data_for_dash() if d[3]>0]
        self.text_widget = Listbox(self.root, width=800, yscrollcommand = self.scroll_bar.set)

        self.text_widget.insert(END, 'Дата загрузки отчета из Пульс: ' + self.last_date_pulse +'\n')
        self.text_widget.insert(END, 'Дата загрузки данных о заболевших: ' + self.last_date_distant + '\n')
        self.text_widget.insert(END, 'Дата загрузки данных об отпусках: ' + self.last_date_vacations + '\n')

        self.text_widget.insert(END, '' + '\n')
        item = 4

        for i in range(0, len(gotData)):
            row = gotData[i]
            # self.text_widget.insert(END, 'Рук-ль: ' + row[0] +'\n')
            self.text_widget.insert(END, 'Подразделение: ' + row[0] + '-->' + row[1] +'\n')
            self.text_widget.insert(END, 'Всего сотрудников: ' + str(row[2]) +'\n')
            self.text_widget.insert(END, 'Находятся дома: ' + str(row[3]) +'\n')
            item += 3
            self.text_widget.insert(END, 'Кол-во незавершенных курсов у сотрудников, находящихся дома: ' + str(row[4]) +'\n')
            if row[5]:
                self.text_widget.insert(END, 'Ср.кол-во накопленных дней отпуска на 1 сотрудника, находящегося дома: ' + str(int(row[5]/row[3])))
            else:
                self.text_widget.insert(END,
                                        'Ср.кол-во накопленных дней отпуска на 1 сотрудника, находящегося дома: 0')
            self.text_widget.insert(END, '---------------------------------------------' + '\n')
            if row[4] > 0:
                self.text_widget.itemconfig(item, bg='red')
            item += 3

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
        reports_menu.add_command(label="Отчеты из Пульса", command=lambda: self.reportWindow('courses'))
        reports_menu.add_command(label="Данные о заболевших", command=lambda: self.reportWindow('status'))
        reports_menu.add_command(label="Данные об отпусках", command=lambda: self.reportWindow('vacations'))
        menu_bar.add_cascade(label="Отчеты", menu=reports_menu)

        dict_menu = Menu(menu_bar, tearoff=0)
        dict_menu.add_command(label="Загрузить ШР", command=self.import_statka)
        menu_bar.add_cascade(label="Справочники", menu=dict_menu)

        export_menu = Menu(menu_bar, tearoff=0)
        export_menu.add_command(label="Выгрузить отчет", command=self.export_report)
        menu_bar.add_cascade(label="Экспорт", menu=export_menu)
        self.root.configure(menu=menu_bar)

    def export_report(self):
        try:
            query = f"""
                SELECT
                    t1.dep_3_level,
                    t1.dep_5_level,
                    t1.tab,
                    t1.status,
                    t1.course_name,
                    t1.deadline,
                    t1.emp_mail,
                    t1.boss_mail,
                    t2.days_left
                FROM (
                    SELECT
                        dep_3_level,
                        dep_5_level,
                        tab,
                        status,
                        course_name,
                        deadline,
                        emp_mail,
                        boss_mail
                    FROM employees
                    LEFT JOIN status
                    ON tab = status.emp_tab and status.date = "{self.last_date_distant}"
                    LEFT JOIN courses
                    ON status.emp_tab = courses.emp_tab and courses.date = "{self.last_date_pulse}"
                    WHERE status = "Болен"
                    ) as t1
                LEFT JOIN (
                    SELECT
                        tab,
                        days_left
                    FROM employees
                    LEFT JOIN vacations
                    ON tab = vacations.emp_tab and vacations.date = "{self.last_date_vacations}"
                    ) as t2
                ON t1.tab = t2.tab
            """
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(query)
            result = cur.fetchall()
            result.insert(0,(
                'Подразделение 3',
                'Подразделение 5',
                'Табельный',
                'Статус',
                'Не пройден курс',
                'Контрольный срок',
                'Внешняя почта',
                'Внешняя почта руководителя',
                'Накоплено дней'
            ))

            book = xlwt.Workbook(encoding="utf-8")
            sheet1 = book.add_sheet("Sheet1")
            for x in range(len(result)):
                for y in range(len(result[x])):
                    sheet1.write(x, y, result[x][y])
            book.save(r'out/export.xls')
            messagebox.showinfo('Внимание', 'Экспорт завершен. Файл в папке OUT')
        except Exception as e:
            messagebox.showinfo('Внимание', str(e))

    def reportWindow(self, table):
        reports_window = ChildWindow(self.root, 300, 200, self.report_name(table))
        self.top = reports_window.root
        Label(reports_window.root, text="Выберите дату отчета").pack()
        try:
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(f"""
                                SELECT
                                    DISTINCT date
                                FROM {table}
                            """)
            dates = cur.fetchall()
            dates_as_date = sorted([datetime.datetime.strptime(d[0], "%Y-%m-%d").date() for d in dates],
                                   key=lambda x: x,
                                   reverse=True)  # переформат в даты и сортировка
            values = tuple(d.strftime('%Y-%m-%d') for d in dates_as_date)
            self.c = Combobox(reports_window.root, width=25, values=values)
            self.c.current(0)
            self.c.pack()
            Button(reports_window.root, width=23, text="Удалить отчет",
                   command=lambda: self.del_confirmation(table)).pack()
        except:
            pass

        Button(reports_window.root, width=23, text="Загрузить отчет", command=lambda: self.import_report_form(table)).pack(pady=10)
        reports_window.grab_focus()

    def del_confirmation(self, table):
        result = messagebox.askyesno(
            title="Подтверждение удаления",
            message="Вы действительно хотите удалить отчет"
        )
        if result:
            self.del_report(table)
        else:
            exit()

    def del_report(self, table):
        date_for_del = self.c.get()
        print(date_for_del)
        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(f"""
                                    DELETE FROM {table} WHERE date = "{date_for_del}"
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

    def import_report_form(self, table):
        import_form = ChildWindow(self.root, 350, 200, 'Импорт ' + self.report_name(table))
        self.top.destroy()
        self.top = import_form.root
        Label(import_form.root, text="Дата отчета").pack()
        self.importDate = DateEntry(import_form.root, width=25, date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.importDate.pack()
        Button(import_form.root, width=25, text="Выбрать файл", command=lambda: self.get_file(table)).pack()
        self.chosen_file = Text(import_form.root)
        Button(import_form.root, width=25, text="Импортировать", command=lambda: self.import_file(table)).pack(pady=10)
        import_form.grab_focus()

    def get_file(self, table, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title=self.report_name(table), filetypes=wanted_files)
        if file_name:
            try:
                xl = xlrd.open_workbook(file_name, on_demand=True)
                sh = xl.sheet_by_index(0)
                row_list = self.parse_excel(table, sh)

                self.list_for_import = row_list
                self.chosen_file.insert(END, file_name)
                self.chosen_file.pack()
            except Exception as e:
                messagebox.showinfo('Внимание', 'Ошибка парсинга: ' + str(e))

    def import_file(self, table):

        create_query = self.create_query(table)
        insert_query = self.insert_query(table)
        if self.list_for_import and self.importDate.get_date():
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(
                create_query
            )
            conn.commit()

            cur.execute(
                f"""
                DELETE FROM {table} WHERE date = "{str(self.importDate.get_date())}"
                """
            )
            conn.commit()

            import_list=[]
            for i in self.list_for_import:
                row=[str(self.importDate.get_date())]
                for j in i:
                    row.append(j)
                import_list.append(row)

            print(import_list)
            cur.executemany(insert_query, import_list)
            conn.commit()
            self.top.destroy()
            self.top.update()
            self.root.update()
            messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(self.list_for_import)))
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
                FROM status
                LEFT JOIN courses
                ON courses.emp_tab = status.emp_tab and status.date = "{self.last_date_distant}" and courses.date = "{self.last_date_pulse}" 
                WHERE status.status = "Болен"
                     """
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(query)
            result = cur.fetchall()
            cur.close()

            to_list = list(set([r[0] for r in result if r[0]]))

            courses_dict = {
                reciever: [
                    [
                        course[2],
                        course[3],
                        course[1]
                    ] for course in result if course[0] == reciever
                ] for reciever in to_list
            }

            # print(courses_dict)

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
                    # msg['To'] = "nvivanchikov@sberbank.ru"
                    # msg['CC'] = "arsadmitriev@sberbank.ru"
                    # smtpObj.sendmail(self.login.get(), ['nvivanchikov@sberbank.ru', 'arsadmitriev@sberbank.ru'],
                    #                  msg.as_string())


                    msg['To'] = k
                    msg['CC'] = v[0][2]

                    smtpObj.sendmail(self.login.get(), [k, v[0][2]], msg.as_string())
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
                messagebox.showinfo('Внимание', str(e))
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
            for row in range(1, sh.nrows):
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
