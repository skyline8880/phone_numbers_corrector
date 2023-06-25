import os
import sys
import tkinter
from tkinter import PhotoImage, StringVar, filedialog, messagebox, ttk
import pandas as pd


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')
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
            self, width=30, height=1, text='Загрузить основной файл',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: get_file_to_upload())
        self.upload_button.place(x=20, y=30)
        self.upload_sec_button = tkinter.Button(
            self, width=30, height=1, text='Загрузить файл сравнения',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: create_sec_data())
        self.upload_sec_button.place(x=20, y=60)
        self.save_button = tkinter.Button(
            self, width=30, height=1, text='Сохранить',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: save_to_file())
        self.save_button.place(x=20, y=100)
        self.name_var = StringVar()
        self.name_var.set(0)
        self.include_name_button = tkinter.Checkbutton(
            self, text='Включить ФИО в список', 
		    highlightbackground='gray', highlightthickness=1, border=0,
            variable=self.name_var, onvalue=1, offvalue=0)
        self.include_name_button.place(x=20, y=130)
        self.include_name_button.config(state='disabled')

        self.phone_workplace = tkinter.LabelFrame(
            self, width=250, height=150, text='Статистика основного файла')
        self.phone_workplace.place(x=270, y=5)
        self.label_all_numbers = tkinter.Label(
            self, text='')
        self.label_all_numbers.place(x=280, y=20)
        self.label_changed_numbers = tkinter.Label(
            self, text='')
        self.label_changed_numbers.place(x=280, y=40)
        self.label_invalid_numbers = tkinter.Label(
            self, text='')
        self.label_invalid_numbers.place(x=280, y=60)
        self.label_repeat_numbers = tkinter.Label(
            self, text='')
        self.label_repeat_numbers.place(x=280, y=80)
        self.label_excluded_numbers = tkinter.Label(
            self, text='')
        self.label_excluded_numbers.place(x=280, y=100)        
        self.label_correct_numbers = tkinter.Label(
            self, text='')
        self.label_correct_numbers.place(x=280, y=120)

        self.sec_phone_workplace = tkinter.LabelFrame(
            self, width=250, height=150, text='Статистика файла сравнения')
        self.sec_phone_workplace.place(x=530, y=5)
        self.sec_label_all_numbers = tkinter.Label(
            self, text='')
        self.sec_label_all_numbers.place(x=540, y=20)
        self.sec_label_changed_numbers = tkinter.Label(
            self, text='')
        self.sec_label_changed_numbers.place(x=540, y=40)
        self.sec_label_invalid_numbers = tkinter.Label(
            self, text='')
        self.sec_label_invalid_numbers.place(x=540, y=60)
        self.sec_label_repeat_numbers = tkinter.Label(
            self, text='')
        self.sec_label_repeat_numbers.place(x=540, y=80)
        self.sec_label_excluded_numbers = tkinter.Label(
            self, text='')
        self.sec_label_excluded_numbers.place(x=540, y=100)        
        self.sec_label_correct_numbers = tkinter.Label(
            self, text='')
        self.sec_label_correct_numbers.place(x=540, y=120)

        self.sec_phone_workplace = tkinter.LabelFrame(
            self, width=300, height=150, text='Обработка файлов')
        self.sec_phone_workplace.place(x=790, y=5)
        self.exclude_sec_button = tkinter.Button(
            self, width=17, height=2, state='disabled',
            text='Удалить номера',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: remove_numbers())
        self.exclude_sec_button.place(x=800, y=30)
        self.cleandata_button = tkinter.Button(
            self, width=17, height=2, state='normal',
            text='Очистить панель',
            highlightbackground='gray', highlightthickness=2, border=1,
            command=lambda: clean_data())
        self.cleandata_button.place(x=950, y=30)
        self.label_worked_numbers = tkinter.Label(
            self, text='')
        self.label_worked_numbers.place(x=800, y=80)        
        self.label_removed_numbers = tkinter.Label(
            self, text='')
        self.label_removed_numbers.place(x=800, y=100)
        self.label_left_numbers = tkinter.Label(
            self, text='')
        self.label_left_numbers.place(x=800, y=120)

        self.names_list = []
        self.list_numbers_to_save = []
        self.sec_list_numbers_to_save = []

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
        treescrollx.pack(side='bottom', fill='x')
        treescrolly.pack(side='right', fill='y')

        def save_to_file():
            if self.list_numbers_to_save != []:
                saved_filename = filedialog.asksaveasfilename(
                initialdir='C:\\Users\\{oper}\\Desktop',
                defaultextension=[('Excel', '*.xlsx'), 
                                   ('CSV', '*.csv'),],
                title='Сохранить файл',
                filetypes=([('Excel', '*.xlsx'),
                            ('CSV', '*.csv'),]))
                try:
                    if saved_filename != '':
                        if '.csv' in saved_filename:
                            if self.name_var.get() == '1':
                                csv_array = []
                                for name in self.names_list:
                                    csv_array.append(''.join(str(name).split(' ')))
                                df = pd.DataFrame.from_dict(
                                        {
                                    'Phone': self.list_numbers_to_save, 
                                    'Name': self.names_list}
                                    )
                                df.to_csv(
                                    saved_filename,
                                    encoding='utf-16',
                                    sep=',',
                                    index=False,
                                    header=False)
                            else:
                                df = pd.DataFrame.from_dict(
                                    {'Phone': self.list_numbers_to_save}
                                    )
                                df.to_csv(
                                    saved_filename,
                                    index=False,
                                    header=False,)
                        else:
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
                except Exception as e:
                    messagebox.showerror('Ошибка', f'Описание: {e}')

        def get_file_to_upload():
            filename = filedialog.askopenfilename(
                initialdir='C:\\Users\\{oper}\\Desktop',
                title='Выберите файл',
                filetypes=(
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

                    def get_file_columns(data):
                        try:
                            data = data[
                                [' Статус', ' Телефон',
                                 ' Персона', ' Шаблон контракта']
                                ]
                            cols = list(data.columns)
                            phone_col_idx = 1
                            return data, cols, phone_col_idx
                        except:
                            try:
                                data = data[[' Телефон']]
                                phone_col_idx = 0
                                return data, None, phone_col_idx
                            except:    
                                try:
                                    col = list(data.columns)[0]
                                    if col.lower().strip() in (
                                        'телефон', 'тел', 'тел.', 'телефон.',  'телеф.', 'телеф'
                                        ):
                                        phone_col_idx = 0
                                        data = data[[col]]
                                        return data, None, phone_col_idx
                                except:
                                    phone_col_idx = 0
                                    return data, phone_col_idx
                        
                    dataframe, columns, idx = get_file_columns(dataframe)
                    if columns is None:
                        columns = 'Телефон'
                    length_before_clean = len(dataframe)
                    self.table['column'] = columns
                    self.table['show'] = 'headings'
                    for column in self.table['columns']:
                        self.table.heading(column, text=column)
                    rows = dataframe.to_numpy().tolist()
                    cleared_counter = 0
                    invalid_counter = 0
                    repeat_count = 0
                    for row in rows:
                        row[idx], corrected = delete_symbols_from_number(
                            str(row[idx]))
                        if len(row[idx]) == 10:
                            if 900 <= int(row[idx][:3]) <= 999:
                                row[idx] = '7' + row[idx]
                                if not row[idx] in self.list_numbers_to_save:
                                    self.table.insert(
                                        '', 'end', value=row)
                                    self.list_numbers_to_save.append(
                                        row[idx])
                                    if idx != 0:
                                        self.names_list.append(row[2])
                                    cleared_counter += 1
                                else:
                                    repeat_count += 1
                            else:
                                invalid_counter += 1
                        elif len(row[idx]) == 11:
                            if not row[idx].startswith('7'):
                                if 900 <= int(row[idx][1:4]) <= 999:
                                    row[idx] = '7' + row[idx][1:]
                                    if (not row[idx] in 
                                        self.list_numbers_to_save):
                                        self.table.insert(
                                            '', 'end', value=row)
                                        self.list_numbers_to_save.append(
                                            row[idx])
                                        if idx != 0:
                                            self.names_list.append(row[2])
                                        cleared_counter += 1
                                    else:
                                        repeat_count += 1
                                else:
                                    invalid_counter += 1
                            else:
                                if 900 <= int(row[idx][1:4]) <= 999:
                                    if (not row[idx] in 
                                        self.list_numbers_to_save):
                                        self.table.insert(
                                            '', 'end', value=row)
                                        self.list_numbers_to_save.append(
                                            row[idx])
                                        if idx != 0:
                                            self.names_list.append(row[2])
                                        cleared_counter += corrected
                                    else:
                                        repeat_count += 1
                                else:
                                    invalid_counter += 1
                        else:
                            invalid_counter += 1
                    corrected_nums = (length_before_clean 
                                        - invalid_counter
                                        - repeat_count)
                    excluded = (invalid_counter + repeat_count)
                    f_name = name.split('/')[-1]
                    self.title(f'Редактор сервисных звонков. Основной файл: {f_name}')
                    self.label_all_numbers[
                        'text'] = f'Обработано: {length_before_clean}'
                    self.label_changed_numbers[
                        'text'] = f'Исправлено: {cleared_counter}'
                    self.label_invalid_numbers[
                        'text'] = f'Неисправно: {invalid_counter}'
                    self.label_repeat_numbers[
                        'text'] = f'Повторяющиеся: {repeat_count}'
                    self.label_excluded_numbers[
                        'text'] = f'Исключено: {excluded}'
                    self.label_correct_numbers[
                        'text'] = f'Корректные номера: {corrected_nums}'
                    if self.names_list != []:
                        self.include_name_button.config(state='normal')
                    if (self.list_numbers_to_save != [] 
                        and self.sec_list_numbers_to_save != []):
                        self.exclude_sec_button.config(state='normal')
            except FileNotFoundError:
                return
            except Exception as e:
                messagebox.showerror('Ошибка', f'Описание: {e}')
        
        def delete_symbols_from_number(arg: str):
            is_correct = 0
            avoid_sym = [' ', '-', '/', '+', '_', '*', ',', '(', ')', '.']
            if len(arg) > 10:
                if arg[-2] == '.':
                    arg = arg.split('.')[0]
            for sym in avoid_sym:
                if sym in arg:
                    arg = ''.join(arg.split(sym))
                    is_correct = 1
            return arg, is_correct

        def create_sec_data():
            filename = filedialog.askopenfilename(
                initialdir='C:\\Users\\{oper}\\Desktop',
                title='Выберите файл',
                filetype=(
                        ('Excel', '*.xlsx'),
                        ('Excel', '*.xls')))
            try:
                if filename != '':
                    self.sec_list_numbers_to_save.clear()
                    dataframe = pd.read_excel(filename)
                    def get_file_columns(data):
                        try:
                            data = data[
                                [' Статус', ' Телефон',
                                 ' Персона', ' Шаблон контракта']
                                ]
                            phone_col_idx = 1
                            return data, phone_col_idx
                        except:
                            try:
                                data = data[[' Телефон']]
                                phone_col_idx = 0
                                return data, phone_col_idx
                            except:
                                try:
                                    col = list(data.columns)[0]
                                    if col.lower().strip() in (
                                        'телефон', 'тел', 'тел.', 'телефон.',  'телеф.', 'телеф'
                                        ):
                                        phone_col_idx = 0
                                        data = data[[col]]
                                        return data, phone_col_idx
                                except:
                                    phone_col_idx = 0
                                    return data, phone_col_idx
                    dataframe, idx = get_file_columns(dataframe)
                    length_before_clean = len(dataframe)
                    rows = dataframe.to_numpy().tolist()
                    cleared_counter = 0
                    invalid_counter = 0
                    repeat_count = 0
                    for row in rows:
                        row[idx], corrected = delete_symbols_from_number(
                            str(row[idx]))
                        if len(row[idx]) == 10:
                            if 900 <= int(row[idx][:3]) <= 999:
                                row[idx] = '7' + row[idx]
                                if (not row[idx]
                                    in self.sec_list_numbers_to_save):
                                    self.sec_list_numbers_to_save.append(
                                        row[idx])
                                    cleared_counter += 1
                                else:
                                    repeat_count += 1
                            else:
                                invalid_counter += 1
                        elif len(row[idx]) == 11:
                            if not row[idx].startswith('7'):
                                if 900 <= int(row[idx][1:4]) <= 999:
                                    row[idx] = '7' + row[idx][1:]
                                    if (not row[idx] in 
                                        self.sec_list_numbers_to_save):
                                        self.sec_list_numbers_to_save.append(
                                            row[idx])
                                        cleared_counter += 1
                                    else:
                                        repeat_count += 1
                                else:
                                    invalid_counter += 1
                            else:
                                if 900 <= int(row[idx][1:4]) <= 999:
                                    if (not row[idx] in 
                                        self.sec_list_numbers_to_save):
                                        self.sec_list_numbers_to_save.append(
                                            row[idx])
                                        cleared_counter += corrected
                                    else:
                                        repeat_count += 1
                                else:
                                    invalid_counter += 1
                        else:
                            invalid_counter += 1
                    corrected_nums = (length_before_clean 
                                        - invalid_counter
                                        - repeat_count)
                    excluded = (invalid_counter + repeat_count)
                    self.sec_label_all_numbers[
                        'text'] = f'Обработано: {length_before_clean}'
                    self.sec_label_changed_numbers[
                        'text'] = f'Исправлено: {cleared_counter}'
                    self.sec_label_invalid_numbers[
                        'text'] = f'Неисправно: {invalid_counter}'
                    self.sec_label_repeat_numbers[
                        'text'] = f'Повторяющиеся: {repeat_count}'
                    self.sec_label_excluded_numbers[
                        'text'] = f'Исключено: {excluded}'
                    self.sec_label_correct_numbers[
                        'text'] = f'Корректные номера: {corrected_nums}'
                    if (self.list_numbers_to_save != [] 
                        and self.sec_list_numbers_to_save != []):
                        self.exclude_sec_button.config(state='normal')
            except FileNotFoundError:
                return
            except Exception as e:
                messagebox.showerror('Информация', f'Описание: {e}')

        def remove_numbers():
            removed = 0
            length = len(self.list_numbers_to_save)
            for num in self.sec_list_numbers_to_save:
                if num in self.list_numbers_to_save:
                    if self.names_list != []:
                        self.names_list.pop(
                            self.list_numbers_to_save.index(num))
                    self.list_numbers_to_save.remove(num)
                    removed += 1
            left = length - removed
            self.label_worked_numbers[
                'text'] = f'Обработано: {length}'
            self.label_removed_numbers[
                'text'] = f'Удалено из списка: {removed}'
            self.label_left_numbers[
                'text'] = f'Осталось в списке: {left}'

        def clean_data():
            labels = [
                self.label_all_numbers,
                self.label_changed_numbers,
                self.label_correct_numbers,
                self.label_excluded_numbers,
                self.label_invalid_numbers,
                self.label_repeat_numbers,
                self.label_left_numbers,
                self.label_worked_numbers,
                self.label_removed_numbers,
                self.sec_label_all_numbers,
                self.sec_label_changed_numbers,
                self.sec_label_correct_numbers,
                self.sec_label_excluded_numbers,
                self.sec_label_invalid_numbers,
                self.sec_label_repeat_numbers,
            ]
            for label in labels:
                label['text'] = ''
            self.list_numbers_to_save.clear()
            self.sec_list_numbers_to_save.clear()
            self.names_list.clear()
            self.title('Редактор сервисных звонков')
            clear_table()
            self.exclude_sec_button.config(state='disabled')
            self.name_var.set(0)
            self.include_name_button.config(state='disabled')

        def clear_table():
            self.table.delete(*self.table.get_children())
            self.table['show'] = ''
            return None


if __name__ == '__main__':
    app = FileEditorApp()
    app.mainloop()
