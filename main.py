import tkinter
from tkinter import filedialog, messagebox, ttk, PhotoImage, StringVar
import pandas as pd
import os
import sys


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

custom_img = resource_path('sp_icon.png')


class FileEditorApp(tkinter.Tk):
    def __init__(self):
        super().__init__()
        self.title('Редактор сервисных звонков')
        img = PhotoImage(file=custom_img)
        self.tk.call('wm', 'iconphoto', self._w, img)
        self.geometry('1100x1100')
        self.configure(highlightbackground='gray', highlightthickness=1)
        self.resizable(True, True)
        self.file_workplace = tkinter.LabelFrame(
            self, width=250, height=150, text='Работа с файлами')
        self.file_workplace.place(x=10, y=5)
        self.upload_button = tkinter.Button(
            self, width=30, height=1, text='Загрузить',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: get_file_to_upload())
        self.upload_button.place(x=20, y=30)
        self.save_button = tkinter.Button(
            self, width=30, height=1, text='Сохранить',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: save_to_file())
        self.save_button.place(x=20, y=60)
        self.name_var = StringVar()
        self.name_var.set(0)
        self.include_name_button = tkinter.Checkbutton(
            self, text='Включить ФИО в список', 
		    highlightbackground='gray', highlightthickness=1, border=0,
            variable=self.name_var, onvalue=1, offvalue=0)
        self.include_name_button.place(x=20, y=90)

        self.phone_workplace = tkinter.LabelFrame(
            self, width=250, height=150, text='Корректировка номеров')
        self.phone_workplace.place(x=270, y=5)
        self.label_all_numbers = tkinter.Label(
            self, text='')
        self.label_all_numbers.place(x=280, y=25)
        self.label_changed_numbers = tkinter.Label(
            self, text='')
        self.label_changed_numbers.place(x=280, y=50)
        self.label_excluded_numbers = tkinter.Label(
            self, text='')
        self.label_excluded_numbers.place(x=280, y=75)
        self.label_repeat_numbers = tkinter.Label(
            self, text='')
        self.label_repeat_numbers.place(x=280, y=100)
        self.label_correct_numbers = tkinter.Label(
            self, text='')
        self.label_correct_numbers.place(x=280, y=125)

        self.names_list = []
        self.list_numbers_to_save = []
        self.black_list = []

        def save_to_file():
            if self.list_numbers_to_save != []:
                saved_filename = filedialog.asksaveasfilename(
                initialdir='C:\\Users\\{oper}\\Desktop',
                title='Сохранить файл',
                filetype=([('Excel', '*.xlsx')]))
                try:
                    if saved_filename != '':
                        if not '.xlsx' in saved_filename:
                            saved_filename = saved_filename + '.xlsx'
                        else:
                            saved_filename = saved_filename.split('.')[0]
                            saved_filename = saved_filename + '.xlsx'
                        print(self.name_var.get())
                        if self.name_var.get() == '1':
                            df = pd.DataFrame.from_dict(
                                {
                                'Phone': self.list_numbers_to_save, 
                                'Name': self.names_list}
                                )
                            df.to_excel(
                                saved_filename,
                                index=False,
                                header=False,
                                sheet_name='Список номеров и ФИО')
                        else:
                            df = pd.DataFrame.from_dict(
                                {'Phone': self.list_numbers_to_save}
                                )
                            df.to_excel(
                                saved_filename,
                                index=False,
                                header=False,
                                sheet_name='Список номеров')

                except FileNotFoundError:
                    return
                except Exception:
                    return

        def get_file_to_upload():
            filename = filedialog.askopenfilename(
                initialdir='C:\\Users\\{oper}\\Desktop',
                title='Выберите файл',
                filetype=(
                        ('Excel', '*.xlsx'),
                        ('Excel', '*.xls')))
            create_table_data(filename)

        def create_table_data(name: str):
            try:
                if name != '':
                    clear_table()
                    self.list_numbers_to_save.clear()
                    self.names_list.clear()
                    dataframe = pd.read_excel(name)
                    if len(list(dataframe.columns)) == 1:
                        dataframe = dataframe['Телефон']
                    else:
                        dataframe = dataframe[
                            [' Статус', ' Телефон', ' ФИО Персоны', ' Шаблон']
                            ]
                    length_before_clean = len(dataframe)
                    self.table['column'] = list(dataframe.columns)
                    self.table['show'] = 'headings'
                    for column in self.table['columns']:
                        self.table.heading(column, text=column)
                    rows = dataframe.to_numpy().tolist()
                    cleared_counter = 0
                    exlcuded_counter = 0
                    repeat_count = 0
                    for row in rows:
                        row[1], corrected = delete_symbols_from_number(
                            str(row[1]))
                        if len(row[1]) == 10:
                            if 900 <= int(row[1][:3]) <= 999:
                                row[1] = '7' + row[1]
                                if not row[1] in self.list_numbers_to_save:
                                    self.table.insert(
                                        '', 'end', value=row)
                                    self.list_numbers_to_save.append(
                                        row[1])
                                    self.names_list.append(row[2])
                                    cleared_counter += 1
                                else:
                                    repeat_count += 1
                            else:
                                exlcuded_counter += 1
                        elif len(row[1]) == 11:
                            if not row[1].startswith('7'):
                                if 900 <= int(row[1][1:4]) <= 999:
                                    row[1] = '7' + row[1][1:]
                                    if (not row[1] in 
                                        self.list_numbers_to_save):
                                        self.table.insert(
                                            '', 'end', value=row)
                                        self.list_numbers_to_save.append(
                                            row[1])
                                        self.names_list.append(row[2])
                                        cleared_counter += 1
                                    else:
                                        repeat_count += 1
                                else:
                                    exlcuded_counter += 1
                            else:
                                if 900 <= int(row[1][1:4]) <= 999:
                                    if (not row[1] in 
                                        self.list_numbers_to_save):
                                        self.table.insert(
                                            '', 'end', value=row)
                                        self.list_numbers_to_save.append(
                                            row[1])
                                        self.names_list.append(row[2])
                                        cleared_counter += corrected
                                    else:
                                        repeat_count += 1
                                else:
                                    exlcuded_counter += 1
                        else:
                            exlcuded_counter += 1
                        self.label_all_numbers[
                            'text'] = f'Обработано: {length_before_clean}'
                        self.label_changed_numbers[
                            'text'] = f'Исправлено: {cleared_counter}'
                        self.label_excluded_numbers[
                            'text'] = f'Исключено: {exlcuded_counter}'
                        self.label_repeat_numbers[
                            'text'] = f'Повторяющиеся: {repeat_count}'
                        corrected_nums = (length_before_clean 
                                          - exlcuded_counter)
                        self.label_correct_numbers[
                            'text'] = f'Корректные номера: {corrected_nums}'
            except FileNotFoundError:
                return
            except Exception:
                messagebox.showerror('Информация', 'Ошибка при чтении файла')

        self.table_workplace = tkinter.LabelFrame(
            self, height=150, text='Данные таблицы')
        self.table_workplace.place(x=10, y=160, relwidth=.985, relheight=.85)
        self.table = ttk.Treeview(self.table_workplace, show='headings')
        self.table.place(relheight=1, relwidth=1)
        treescrolly = tkinter.Scrollbar(
            self.table, orient='vertical', command=self.table.yview)
        treescrollx = tkinter.Scrollbar(
            self.table, orient='horizontal', command=self.table.xview)
        self.table.configure(
            xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
        treescrollx.pack(side="bottom", fill="x")
        treescrolly.pack(side="right", fill="y")
        
        def delete_symbols_from_number(arg: str):
            is_correct = 0
            avoid_sym = [' ', '-', '/', '+', '_', '*', ',', '(', ')', '.']
            for sym in avoid_sym:
                if sym in arg:
                    arg = ''.join(arg.split(sym))
                    is_correct = 1
            return arg, is_correct

        def clear_table():
            self.table.delete(*self.table.get_children())
            self.table['show'] = ''
            return None


if __name__ == "__main__":
    app = FileEditorApp()
    app.mainloop()
