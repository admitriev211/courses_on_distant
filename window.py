from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter.ttk import Combobox
from child_window import ChildWindow
from email.mime.text import MIMEText
from email.header import Header
import xlrd
import xlwt
import sqlite3
import datetime
import smtplib
import math
# from tkcalendar import DateEntry

class Window:
    wanted_files = (
        ("excel files", "*.xls;*.xlsx"),
    )

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
        self.last_date_last_date_status = None
        self.draw_menu()
        self.draw_buttons()
        try:
            try:
                conn = sqlite3.connect(r'dbase/dbase.db')
                cur = conn.cursor()
                cur.execute(f"""
                                                SELECT
                                                    *
                                                FROM employees
                                            """)
                employees = cur.fetchall()
                employees[0]
                cur.close()
            except:
                self.import_statka()
            self.draw_stats()
        except Exception as e:
            print(e)
            messagebox.showinfo("Ошибка запроса", e)

    def run(self):
        self.root.mainloop()

    def create_child(self, width, height, title, resizable=(False,False), icon=None):
        ChildWindow(self.root, width, height, title, resizable, icon)

    def draw_buttons(self):
        # gosb_dict = {
        #     'ТБ': 'tb',
        #     '8586': 'Иркутское отделение № 8586',
        #     '8600': 'Читинское отделение № 8600',
        #     '8601': 'Бурятское отделение № 8601',
        #     '8603': 'Якутское отделение № 8603'
        # }
        def show_stats(gosb='tb'):
            self.text_widget.destroy()
            self.text_widget = None
            self.draw_stats(gosb)

        new_frame = Frame(self.root)
        new_frame.pack(side=TOP)

        Button(
            new_frame,
            width=23,
            text="ТБ",
            command=lambda: show_stats('tb')
        ).pack(side='left', pady=10)
        Button(
            new_frame,
            width=23,
            text="8586",
            command=lambda: show_stats('Иркутское отделение № 8586')
        ).pack(side='left', pady=10)
        Button(
            new_frame,
            width=23,
            text="8600",
            command=lambda: show_stats('Читинское отделение № 8600')
        ).pack(side='left', pady=10)
        Button(
            new_frame,
            width=23,
            text="8601",
            command=lambda: show_stats('Бурятское отделение № 8601')
        ).pack(side='left', pady=10)
        Button(
            new_frame,
            width=23,
            text="8603",
            command=lambda: show_stats('Якутское отделение № 8603')
        ).pack(side='left', pady=10)

    def get_data_for_dash(self, gosb = "tb"):
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

        try:
            self.last_date_status = find_last_date('status')
        except:
            self.reportWindow('status')
        try:
            self.last_date_distance = find_last_date('distance')
        except:
            self.reportWindow('distance')
        try:
            self.last_date_pulse = find_last_date('courses')
        except:
            self.reportWindow('courses')
        try:
            self.last_date_course_dt = find_last_date('course_dt')
        except:
            self.reportWindow('course_dt')

        # try:
        #     self.last_date_vacations = find_last_date('vacations')
        # except:
        #     self.reportWindow('vacations')


        query = f"""
            SELECT
                dep_3_level,
                dep_5_level,
                count(DISTINCT tab) as tabs,
                count(courses.course_name) as course_count,
                count(course_dt.course_name) as course_dt_count
            FROM
                employees
            LEFT JOIN
                distance
            ON tab = distance.emp_tab and distance.date = "{self.last_date_distance}"
            LEFT JOIN
                status
            ON tab = status.emp_tab and status.date = "{self.last_date_status}"
            LEFT JOIN
                courses
            ON courses.emp_tab = tab and courses.date = "{self.last_date_pulse}"
            LEFT join
                course_dt
            ON course_dt.emp_tab = tab and course_dt.date = "{self.last_date_course_dt}"
            WHERE status = "Болен" OR distance = "дистант"          
            GROUP BY dep_3_level, dep_5_level
        """

        if gosb != 'tb':
            query = f"""
            SELECT
                dep_5_level,
                dep_7_level,
                count(DISTINCT tab) as tabs,
                count(courses.course_name) as course_count,
                count(course_dt.course_name) as course_dt_count
            FROM
                employees
            LEFT JOIN
                distance
            ON tab = distance.emp_tab and distance.date = "{self.last_date_distance}"
            LEFT JOIN
                status
            ON tab = status.emp_tab and status.date = "{self.last_date_status}"
            LEFT JOIN
                courses
            ON courses.emp_tab = tab and courses.date = "{self.last_date_pulse}"
            LEFT join
                course_dt
            ON course_dt.emp_tab = tab and course_dt.date = "{self.last_date_course_dt}"
            WHERE (status = "Болен" OR distance = "дистант") and dep_5_level = "{gosb}"      
            GROUP BY dep_5_level, dep_7_level
        """

        print(gosb)
        print(query)
        # query = f"""
        #     SELECT
        #         t1.dep_3_level,
        #         t1.dep_5_level,
        #         t1.tabs,
        #         t1.distance_count,
        #         t1.sick_count,
        #         t1.course_count,
        #         t2.days
        #     FROM (
        #         SELECT
        #             dep_3_level,
        #             dep_5_level,
        #             count(DISTINCT tab) as tabs,
        #             count(DISTINCT distance.emp_tab) as distance_count,
        #             count(DISTINCT status.emp_tab) as sick_count,
        #             count(courses.course_name) as course_count
        #         FROM employees
        #         LEFT JOIN distance
        #         ON tab = distance.emp_tab and distance.date = "{self.last_date_distance}"
        #         LEFT JOIN status
        #         ON tab = status.emp_tab and status.date = "{self.last_date_status}"
        #         LEFT JOIN courses
        #         ON (status.emp_tab = courses.emp_tab or distance.emp_tab = courses.emp_tab) and courses.date = "{self.last_date_pulse}"
        #         GROUP BY dep_3_level, dep_5_level
        #         ORDER BY course_count DESC, distance_count DESC
        #         ) as t1
        #     LEFT JOIN (
        #         SELECT
        #             dep_3_level,
        #             dep_5_level,
        #             count(DISTINCT tab),
        #             count(DISTINCT status.emp_tab),
        #             sum(days_left) as days
        #         FROM employees
        #         LEFT JOIN status
        #         ON tab = status.emp_tab and status.date = "{self.last_date_status}"
        #         LEFT JOIN vacations
        #         ON status.emp_tab = vacations.emp_tab and vacations.date = "{self.last_date_vacations}"
        #         GROUP BY
        #             dep_3_level,
        #             dep_5_level
        #         ) as t2
        #     ON t1.dep_3_level = t2.dep_3_level and t1.dep_5_level = t2.dep_5_level
        #
        # """

        conn = sqlite3.connect(r'dbase/dbase.db')
        cur = conn.cursor()
        cur.execute(query)
        result = cur.fetchall()
        print(result[:10])
        # print(str(len(result)))
        cur.close()
        return result

    def report_name(self, table):
        names = {
            'courses': 'Обученность по обяз.программам',
            'status': 'Данные о заболевших',
            'vacations': 'Данные об отпусках',
            'distance': 'Данные по дистанционке',
            'course_dt': 'Обученность по СЦТ для массовых должностей'
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
        elif table == 'course_dt':
            create_query = """
                CREATE TABLE IF NOT EXISTS course_dt(
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
        elif table == 'distance':
            create_query = """
                CREATE TABLE IF NOT EXISTS distance(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                emp_tab INTEGER,
                distance TEXT DEFAULT "дистант",
                FOREIGN KEY (emp_tab) REFERENCES employees(tab)
                )
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
        elif table == 'course_dt':
            insert_query = """
                INSERT INTO course_dt(
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
        elif table == 'distance':
            insert_query = """
                                        INSERT INTO distance(
                                            date,
                                            emp_tab
                                        ) VALUES(?, ?);
                                    """
        return insert_query

    def parse_excel(self, table, sh):
        header = {
            'courses': 1,
            'course_dt': 7,
            'status': 0,
            'vacations': 3,
            'statka': 6,
            'distance': 0
        }
        def find_cols(table):
            cols_list = []
            fields = {
                'courses': [
                    'ТН',
                    'Внешняя почта',
                    'ТН руководителя',
                    'Внешняя почта руководителя',
                    'Наименование курса',
                    'Контрольная дата прохождения'
                ],
                'course_dt': [
                    'Табельный номер',
                    'Внешняя почта',
                    'ТН руководителя',
                    'Внешняя почта руководителя',
                    'Наименование предмета',
                    'Контрольная дата'
                ],
                'status': [
                    'Табельный номер',
                    'Статус'
                ],
                'vacations': [
                    'ТН',
                    'годнакоплено дней'
                ],
                'statka': [
                    "Сотрудник",
                    "Подразделение 03 ур.",
                    "Подразделение 05 ур.",
                    "Подразделение 06 ур.",
                    "Подразделение 07 ур."

                ],
                'distance': [
                    "I_PERNR_PR",
                ]
            }
            for f in fields[table]:
                for col in range(0, sh.ncols):
                    if sh.cell(header[table], col).value == f:
                        cols_list.append(col)

            return cols_list

        row_list = []
        cols = find_cols(table)
        for row in range(header[table]+1, sh.nrows):
            line = [sh.cell(row, col).value for col in cols if sh.cell(row, cols[0]).value != '' and sh.cell(row, cols[0]).value != '#']
            if len(line) == len(cols):
                row_list.append(line)
        return row_list

    def draw_stats(self, gosb='tb'):
        gotData = [d for d in self.get_data_for_dash() if d[3] + d[4] >0]
        if gosb != 'tb':
            gotData = [d for d in self.get_data_for_dash(gosb) if d[3] + d[4] > 0]
        gotData = sorted(gotData, key=lambda x: x[3] + x[4], reverse=True)
        self.text_widget = Listbox(self.root, width=800, yscrollcommand = self.scroll_bar.set)

        self.text_widget.insert(END, 'Дата загрузки отчета из Пульс: ' + self.last_date_pulse +'\n')
        self.text_widget.insert(END, 'Дата загрузки данных о заболевших: ' + self.last_date_status + '\n')
        # self.text_widget.insert(END, 'Дата загрузки данных об отпусках: ' + self.last_date_vacations + '\n')

        self.text_widget.insert(END, '' + '\n')
        item = 3

        for i in range(0, len(gotData)):
            row = gotData[i]
            # self.text_widget.insert(END, 'Рук-ль: ' + row[0] +'\n')
            self.text_widget.insert(END, 'Подразделение: ' + row[0] + '-->' + row[1] +'\n')
            self.text_widget.insert(END, 'Находятся дома: ' + str(row[2]) +'\n')
            item += 2
            self.text_widget.insert(END, 'Кол-во незавершенных обязательных курсов: ' + str(row[3]) +'\n')
            if row[3] > 0:
                self.text_widget.itemconfig(item, bg='red')
            item += 1
            self.text_widget.insert(END, 'Кол-во незавершенных курсов СТЦ для массовых должностей: ' + str(row[4]) + '\n')
            if row[4] > 0:
                self.text_widget.itemconfig(item, bg='red')
            # if row[6]:
            #     self.text_widget.insert(END, 'Ср.кол-во накопленных дней отпуска на 1 сотрудника, находящегося дома: ' + str(int(row[6]/(row[3] + row[4]))))
            # else:
            #     self.text_widget.insert(END,
            #                             'Ср.кол-во накопленных дней отпуска на 1 сотрудника, находящегося дома: 0')
            self.text_widget.insert(END, '---------------------------------------------' + '\n')
            item += 2


        self.text_widget.pack(side = LEFT, fill=BOTH)
        self.scroll_bar.config(command=self.text_widget.yview)

    def draw_menu(self):
        menu_bar = Menu(self.root)
        # import_menu = Menu(menu_bar, tearoff=0)
        # import_menu.add_command(label="Импорт выгрузки из Пульса", command=self.import_pulse_form)
        # import_menu.add_command(label="Импорт отчета по больным", command=self.import_illness_report)
        # import_menu.add_command(label="Импорт ШР", command=self.import_statka)
        # menu_bar.add_cascade(label="Импорт", menu=import_menu)

        tools_menu = Menu(menu_bar, tearoff=0)
        tools_menu.add_command(label="Рассылка уведомлений заболевшим", command=lambda: self.send_mail_form('sick'))
        tools_menu.add_command(label="Рассылка уведомлений на дистанте", command=lambda: self.send_mail_form('distant'))
        menu_bar.add_cascade(label="Инструменты", menu=tools_menu)

        reports_menu = Menu(menu_bar, tearoff=0)
        reports_menu.add_command(label="Обученность по обязательным программам", command=lambda: self.reportWindow('courses'))
        reports_menu.add_command(label="Обученность по СЦТ",
                                 command=lambda: self.reportWindow('course_dt'))
        reports_menu.add_command(label="Данные о заболевших", command=lambda: self.reportWindow('status'))
        reports_menu.add_command(label="Данные о УРМ", command=lambda: self.reportWindow('distance'))
        # reports_menu.add_command(label="Данные об отпусках", command=lambda: self.reportWindow('vacations'))
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
                dep_3_level,
                dep_5_level,
                dep_6_level,
                dep_7_level,
                tab,
                status,
                distance,
                courses.course_name,
                courses.deadline,
                courses.emp_mail,
                courses.boss_mail,
                course_dt.course_name,
                course_dt.deadline,
                course_dt.emp_mail,
                course_dt.boss_mail           
            FROM
                employees
            LEFT JOIN
                distance
            ON tab = distance.emp_tab and distance.date = "{self.last_date_distance}"
            LEFT JOIN
                status
            ON tab = status.emp_tab and status.date = "{self.last_date_status}"
            LEFT JOIN
                courses
            ON courses.emp_tab = tab and courses.date = "{self.last_date_pulse}"
            LEFT join
                course_dt
            ON course_dt.emp_tab = tab and course_dt.date = "{self.last_date_course_dt}"
            WHERE (status = "Болен" OR distance = "дистант")         
            """
            #AND (courses.course_name != "" OR course_dt.course_name != "")
            # query = f"""
            #     SELECT
            #         t1.dep_3_level,
            #         t1.dep_5_level,
            #         t1.tab,
            #         t1.status,
            #         t1.course_name,
            #         t1.deadline,
            #         t1.emp_mail,
            #         t1.boss_mail,
            #         t2.days_left
            #     FROM (
            #         SELECT
            #             dep_3_level,
            #             dep_5_level,
            #             tab,
            #             status,
            #             course_name,
            #             deadline,
            #             emp_mail,
            #             boss_mail
            #         FROM employees
            #         LEFT JOIN status
            #         ON tab = status.emp_tab and status.date = "{self.last_date_status}"
            #         LEFT JOIN courses
            #         ON status.emp_tab = courses.emp_tab and courses.date = "{self.last_date_pulse}"
            #         WHERE status = "Болен"
            #         ) as t1
            #     LEFT JOIN (
            #         SELECT
            #             tab,
            #             days_left
            #         FROM employees
            #         LEFT JOIN vacations
            #         ON tab = vacations.emp_tab and vacations.date = "{self.last_date_vacations}"
            #         ) as t2
            #     ON t1.tab = t2.tab
            # """
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(query)
            result = cur.fetchall()
            result.insert(0,(
                'Подразделение 3',
                'Подразделение 5',
                'Подразделение 6',
                'Подразделение 7',
                'Табельный',
                'Состояние здоровья',
                'Формат работы',
                'Не пройден обязательным курсам',
                'Контрольный срок',
                'Внешняя почта',
                'Внешняя почта руководителя',
                'Не пройден курсам по СЦТ для массовых должностей',
                'Контрольный срок',
                'Внешняя почта',
                'Внешняя почта руководителя',
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
        Label(import_form.root, text="Дата отчета (ГГГГ-ММ-ДД)").pack()
        # self.importDate = DateEntry(import_form.root, width=25, date_pattern='dd/mm/yyyy', background='darkblue', foreground='white', borderwidth=2)
        self.importDate = Entry(import_form.root, width=25)
        self.importDate.insert(0, datetime.date.today().strftime('%Y-%m-%d'))
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
        # gotDate = self.importDate.get_date()
        try:
            gotDate = self.importDate.get()
            datetime.datetime.strptime(gotDate, "%Y-%m-%d")
            create_query = self.create_query(table)
            insert_query = self.insert_query(table)
            if self.list_for_import and gotDate:
                conn = sqlite3.connect(r'dbase/dbase.db')
                cur = conn.cursor()
                cur.execute(
                    create_query
                )
                conn.commit()

                cur.execute(
                    f"""
                    DELETE FROM {table} WHERE date = "{str(gotDate)}"
                    """
                )
                conn.commit()

                import_list=[]
                for i in self.list_for_import:
                    row=[str(gotDate)]
                    for j in i:
                        row.append(j)
                    import_list.append(row)

                cur.executemany(insert_query, import_list)
                conn.commit()
                self.top.destroy()
                self.top.update()
                self.root.update()
                messagebox.showinfo('Внимание', 'Создано записей: ' + str(len(self.list_for_import)))
                if self.text_widget:
                    self.text_widget.destroy()
                self.text_widget = None
                self.draw_stats()
        except Exception as e:
            messagebox.showinfo('Внимание', str(e))
    def send_mail_form(self, reciever_type):
        header = {
            'sick': "Отправка уведомлений болеющим дома",
            'distant': "Отправка уведомлений на дистанте"
        }
        send_form = ChildWindow(self.root, 400, 400, header[reciever_type])
        self.top = send_form.root
        Label(send_form.root, text="Введите сервер").pack()
        self.server = Entry(send_form.root, width=100)
        self.server.insert(END, 'smtp.mail.ru')
        self.server.pack()
        Label(send_form.root, text="Введите логин").pack()
        self.login = Entry(send_form.root, width=100)
        self.login.insert(END, 'bb_sales@bk.ru')
        self.login.pack()
        Label(send_form.root, text="Введите пароль").pack()
        self.password = Entry(send_form.root, width=100)
        self.password.insert(END, 'DSPQV5c5NFM7G2mY7bPM')
        self.password.pack()
        Button(send_form.root, text="Файл с текстом письма", command=self.get_text_for_letter).pack()
        Button(send_form.root, text="Разослать уведомления", command=lambda: self.send_mails(reciever_type)).pack()
        scroll_bar = Scrollbar(send_form.root)
        scroll_bar.pack(side=RIGHT, fill=Y)
        self.text_on_screen = Text(send_form.root, width=400, height=300, wrap=WORD, yscrollcommand=scroll_bar.set)
        self.text_on_screen.pack()
        send_form.grab_focus()

    def get_text_for_letter(self):
        file_name = fd.askopenfilename(title="Выберите файл с текстом письма")
        if file_name:
            with open(file_name, encoding = 'utf-8', mode='r') as f:
                self.text_for_letter = f.read()
                self.text_on_screen.insert(END, self.text_for_letter)

    def send_mails(self, reciever_type):
        if self.text_for_letter:
            query = {
                'sick': f"""
                SELECT
                    emp_mail,
                    boss_mail,
                    course_name,
                    deadline 
                FROM status
                LEFT JOIN courses
                ON courses.emp_tab = status.emp_tab and status.date = "{self.last_date_status}" and courses.date = "{self.last_date_pulse}" 
                WHERE status.status = "Болен"
                UNION
                SELECT
                    emp_mail,
                    boss_mail,
                    course_name,
                    deadline 
                FROM status
                LEFT JOIN course_dt
                ON course_dt.emp_tab = status.emp_tab and status.date = "{self.last_date_status}" and course_dt.date = "{self.last_date_course_dt}" 
                WHERE status.status = "Болен"
                     """,
                'distant': f"""
                SELECT
                    emp_mail,
                    boss_mail,
                    course_name,
                    deadline 
                FROM distance
                LEFT JOIN courses
                ON courses.emp_tab = distance.emp_tab and distance.date = "{self.last_date_status}" and courses.date = "{self.last_date_pulse}" 
                UNION
                SELECT
                    emp_mail,
                    boss_mail,
                    course_name,
                    deadline 
                FROM distance
                LEFT JOIN course_dt
                ON course_dt.emp_tab = distance.emp_tab and distance.date = "{self.last_date_status}" and course_dt.date = "{self.last_date_course_dt}"
                    """
            }
            conn = sqlite3.connect(r'dbase/dbase.db')
            cur = conn.cursor()
            cur.execute(query[reciever_type])
            result = cur.fetchall()
            print(result)
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
            try:
                def xldate_to_datetime(xldatetime):  # something like 43705.6158241088

                    tempDate = datetime.datetime(1899, 12, 31)
                    (days, portion) = math.modf(xldatetime)

                    deltaDays = datetime.timedelta(days=days)
                    # changing the variable name in the edit
                    secs = int(24 * 60 * 60 * portion)
                    detlaSeconds = datetime.timedelta(seconds=secs)
                    TheTime = (tempDate + deltaDays + detlaSeconds)
                    return TheTime.strftime("%d-%m-%Y")

                for k, v in courses_dict.items():

                    msg_text = f'''
                    to {k} \n
                    {self.text_for_letter}:\n
                    '''
                    for course in v:
                        try:
                            deadline = xldate_to_datetime(float(course[1]))
                        except Exception as e:
                            print(e)
                            deadline = course[1]
                        msg_text += f"""
                        {course[0]}. Срок: {deadline} \n
                        """

                    smtpObj = smtplib.SMTP(self.server.get(), 587)
                    smtpObj.starttls()
                    smtpObj.login(self.login.get(), self.password.get())

                    # to = k
                    # cc = v[0][2]

                    to = "ars-dmitriev@mail.ru"
                    cc = "admitriev211@gmail.com"

                    msg = MIMEText(msg_text)
                    msg['Subject'] = Header('Пройдите курсы в Пульс!', 'utf-8')
                    msg['From'] = self.login.get()
                    msg['To'] = to
                    msg['CC'] = cc
                    smtpObj.sendmail(self.login.get(), [to, cc],
                                     msg.as_string())
                    smtpObj.quit()
                    break
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

    def import_statka(self, wanted_files=wanted_files):
        file_name = fd.askopenfilename(title="Импорт ШР", filetypes=wanted_files)
        if file_name:
            xl = xlrd.open_workbook(file_name, on_demand=True)
            sh = xl.sheet_by_index(0)
            row_list = self.parse_excel('statka', sh)

            print(row_list[:10])
            if row_list:
                conn = sqlite3.connect(r'dbase/dbase.db')
                cur = conn.cursor()
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS employees(
                    tab INT PRIMARY KEY,
                    dep_3_level TEXT,
                    dep_5_level TEXT,
                    dep_6_level TEXT,
                    dep_7_level TEXT);
                    """
                )
                conn.commit()
                cur.execute("SELECT * FROM employees;")
                results = cur.fetchall()
                tabs = [r[0] for r in results]
                print(tabs[:10])

                def kill_doubles(rows):
                    existing_rows=[]
                    new_rows = []
                    for row in rows:
                        if row[0] not in existing_rows:
                            new_rows.append(row)
                            existing_rows.append(row[0])
                    return new_rows

                insert_list = kill_doubles([r for r in row_list if r[0] not in tabs])
                update_list = kill_doubles([r for r in row_list if r[0] in tabs])
                cur.executemany("UPDATE employees set dep_3_level = ?, dep_5_level = ?, dep_6_level = ?, dep_7_level = ?  where tab = ?;", update_list)
                conn.commit()
                cur.executemany("INSERT INTO employees VALUES(?, ?, ?, ?, ?);", insert_list)
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
