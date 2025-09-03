import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import pandas as pd
from fpdf import FPDF
import webbrowser

# Создаем класс FPDF с поддержкой кириллицы и шрифтом ChakraPetch
class PDFWithCyrillic(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font('ChakraPetch', '', 'ChakraPetch-Regular.ttf', uni=True)



class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система инвентаризации оборудования")
        self.root.geometry("900x700")

        # Путь к файлу инвентаризации
        self.inventory_file = r"\\fs\SHARE_BH\it\inventory\inventory.json"

        # Загрузка данных
        self.inventory_data = self.load_data()

        # Создание интерфейса
        self.create_widgets()

    def load_data(self):
        """Загрузка данных из JSON файла"""
        try:
            if os.path.exists(self.inventory_file):
                with open(self.inventory_file, 'r', encoding='utf-8') as file:
                    return json.load(file)
            else:
                # Создаем пустой файл, если он не существует
                with open(self.inventory_file, 'w', encoding='utf-8') as file:
                    json.dump([], file)
                return []
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")
            return []

    def save_data(self):
        """Сохранение данных в JSON файл"""
        try:
            # Создаем директорию, если она не существует
            os.makedirs(os.path.dirname(self.inventory_file), exist_ok=True)

            with open(self.inventory_file, 'w', encoding='utf-8') as file:
                json.dump(self.inventory_data, file, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {e}")
            return False

    def create_widgets(self):
        """Создание элементов интерфейса"""
        # Главный фрейм с кнопкой PDF
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Кнопка выгрузки в PDF
        pdf_button = ttk.Button(main_frame, text="📄 Выгрузить полный отчет в PDF",
                                command=self.export_to_pdf, style='Big.TButton')
        pdf_button.pack(pady=10, fill='x')

        # Создаем notebook для вкладок
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True)

        # Вкладка добавления записи
        self.add_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.add_frame, text="Добавить оборудование")

        # Вкладка поиска
        self.search_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.search_frame, text="Поиск оборудования")

        # Вкладка просмотра по сотрудникам
        self.employee_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.employee_frame, text="Оборудование по сотрудникам")

        # Вкладка "Показать всё"
        self.show_all_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.show_all_frame, text="Показать всё")

        # Вкладка "About"
        self.about_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.about_frame, text="About")

        # Заполняем вкладки
        self.create_add_tab()
        self.create_search_tab()
        self.create_employee_tab()
        self.create_show_all_tab()
        self.create_about_tab()

        # Стиль для большой кнопки
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 12, 'bold'))

    def create_add_tab(self):
        """Создание вкладки для добавления оборудования"""
        # Поля для ввода
        fields = [
            ("Тип оборудования", "equipment_type"),
            ("Модель", "model"),
            ("Серийный номер", "serial_number"),
            ("Закрепление", "assignment"),
            ("Дата", "date"),
            ("Комментарии", "comments")
        ]

        self.entries = {}

        for i, (label_text, field_name) in enumerate(fields):
            label = ttk.Label(self.add_frame, text=label_text + ":")
            label.grid(row=i, column=0, sticky='w', padx=10, pady=5)

            if field_name == "comments":
                entry = scrolledtext.ScrolledText(self.add_frame, width=40, height=4)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')
            elif field_name == "date":
                # Поле даты с текущей датой по умолчанию
                entry = ttk.Entry(self.add_frame, width=40)
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')
            else:
                entry = ttk.Entry(self.add_frame, width=40)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')

            self.entries[field_name] = entry

        # Кнопка добавления
        add_button = ttk.Button(self.add_frame, text="Добавить оборудование",
                                command=self.add_equipment)
        add_button.grid(row=len(fields), column=0, columnspan=2, pady=20)

        # Настройка веса колонок для растягивания
        self.add_frame.columnconfigure(1, weight=1)
        self.add_frame.rowconfigure(len(fields), weight=1)

    def create_search_tab(self):
        """Создание вкладки для поиска"""
        # Поле поиска
        ttk.Label(self.search_frame, text="Поиск:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.search_entry = ttk.Entry(self.search_frame, width=40)
        self.search_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.search_entry.bind('<KeyRelease>', self.perform_search)

        # Кнопка очистки
        clear_button = ttk.Button(self.search_frame, text="Очистить", command=self.clear_search)
        clear_button.grid(row=0, column=2, padx=10, pady=5)

        # Таблица результатов
        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.search_tree = ttk.Treeview(self.search_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.search_tree.heading(col, text=col)
            self.search_tree.column(col, width=100)

        # Привязываем двойной клик для редактирования
        self.search_tree.bind('<Double-1>', self.on_tree_double_click)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.search_frame, orient="vertical", command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=scrollbar.set)

        self.search_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        # Настройка веса для растягивания
        self.search_frame.columnconfigure(1, weight=1)
        self.search_frame.rowconfigure(1, weight=1)

    def create_employee_tab(self):
        """Создание вкладки для просмотра оборудования по сотрудникам"""
        # Выбор сотрудника
        ttk.Label(self.employee_frame, text="Сотрудник:").grid(row=0, column=0, sticky='w', padx=10, pady=5)

        # Получаем список всех сотрудников
        employees = sorted(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
        self.employee_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(self.employee_frame, textvariable=self.employee_var,
                                           values=employees, width=30)
        self.employee_combo.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.employee_combo.bind('<<ComboboxSelected>>', self.show_employee_equipment)

        # Кнопка обновления списка сотрудников
        refresh_button = ttk.Button(self.employee_frame, text="Обновить",
                                    command=self.refresh_employee_list)
        refresh_button.grid(row=0, column=2, padx=10, pady=5)

        # Таблица оборудования сотрудника
        columns = ("Тип", "Модель", "Серийный номер", "Дата", "Комментарии")
        self.employee_tree = ttk.Treeview(self.employee_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.employee_tree.heading(col, text=col)
            self.employee_tree.column(col, width=120)

        # Привязываем двойной клик для редактирования
        self.employee_tree.bind('<Double-1>', self.on_tree_double_click)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.employee_frame, orient="vertical", command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        # Настройка веса для растягивания
        self.employee_frame.columnconfigure(1, weight=1)
        self.employee_frame.rowconfigure(1, weight=1)

    def create_show_all_tab(self):
        """Создание вкладки для отображения всех данных"""
        # Кнопка обновления данных
        refresh_button = ttk.Button(self.show_all_frame, text="Обновить данные",
                                    command=self.show_all_data)
        refresh_button.pack(pady=10)

        # Таблица всех данных
        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.all_tree = ttk.Treeview(self.show_all_frame, columns=columns, show='headings', height=20)

        for col in columns:
            self.all_tree.heading(col, text=col)
            self.all_tree.column(col, width=100)

        # Привязываем двойной клик для редактирования
        self.all_tree.bind('<Double-1>', self.on_tree_double_click)

        # Scrollbar для таблицы
        scrollbar = ttk.Scrollbar(self.show_all_frame, orient="vertical", command=self.all_tree.yview)
        self.all_tree.configure(yscrollcommand=scrollbar.set)

        # Упаковка элементов
        self.all_tree.pack(side='left', fill='both', expand=True, padx=10, pady=5)
        scrollbar.pack(side='right', fill='y', pady=5)

        # Загружаем данные при открытии вкладки
        self.show_all_data()

    def create_about_tab(self):
        """Создание вкладки About"""
        # Основной фрейм для центрирования
        center_frame = ttk.Frame(self.about_frame)
        center_frame.pack(expand=True, fill='both')

        # Информация об авторе
        info_text = """
        Система инвентаризации оборудования

        Версия: 0.1
        Разработано: Разин Григорий

        Контактная информация:
        Email: lantester35@gmail.com

        Функционал:
        - Ведение учета оборудования
        - Поиск и фильтрация данных
        - Экспорт отчетов в PDF
        - Управление закреплением за сотрудниками
        """

        # Текстовое поле с информацией
        about_text = scrolledtext.ScrolledText(center_frame, width=60, height=15,
                                               font=('Arial', 11), wrap=tk.WORD)
        about_text.insert('1.0', info_text)
        about_text.config(state='disabled')
        about_text.pack(pady=20, padx=20, expand=True)

        # Кнопка закрытия
        close_button = ttk.Button(center_frame, text="Закрыть", command=self.root.quit)
        close_button.pack(pady=10)

    def add_equipment(self):
        """Добавление нового оборудования"""
        # Получаем данные из полей ввода
        equipment_data = {}

        for field_name, entry in self.entries.items():
            if field_name == "comments":
                equipment_data[field_name] = entry.get("1.0", tk.END).strip()
            else:
                equipment_data[field_name] = entry.get().strip()

        # Проверяем обязательные поля
        if not equipment_data.get('equipment_type') or not equipment_data.get('serial_number'):
            messagebox.showwarning("Предупреждение", "Заполните обязательные поля: Тип оборудования и Серийный номер")
            return

        # Добавляем дату и время создания
        equipment_data['created_datetime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Добавляем в данные
        self.inventory_data.append(equipment_data)

        # Сохраняем данные
        if self.save_data():
            messagebox.showinfo("Успех", "Оборудование успешно добавлено!")
            self.clear_entries()
            self.refresh_employee_list()

    def clear_entries(self):
        """Очистка полей ввода"""
        for field_name, entry in self.entries.items():
            if field_name == "comments":
                entry.delete("1.0", tk.END)
            elif field_name == "date":
                # Очищаем поле даты и устанавливаем текущую дату
                entry.delete(0, tk.END)
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
            else:
                entry.delete(0, tk.END)

    def perform_search(self, event=None):
        """Поиск оборудования"""
        search_text = self.search_entry.get().lower().strip()

        # Очищаем таблицу
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)

        if not search_text:
            return

        # Ищем совпадения
        for item in self.inventory_data:
            # Проверяем все поля на совпадение
            match_found = any(search_text in str(value).lower()
                              for value in item.values() if value)

            if match_found:
                self.search_tree.insert("", "end", values=(
                    item.get('equipment_type', ''),
                    item.get('model', ''),
                    item.get('serial_number', ''),
                    item.get('assignment', ''),
                    item.get('date', ''),
                    item.get('comments', '')[:50] + '...' if item.get('comments') and len(
                        item.get('comments', '')) > 50 else item.get('comments', '')
                ))

    def clear_search(self):
        """Очистка поиска"""
        self.search_entry.delete(0, tk.END)
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)

    def refresh_employee_list(self):
        """Обновление списка сотрудников"""
        employees = sorted(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
        self.employee_combo['values'] = employees

    def show_employee_equipment(self, event=None):
        """Показать оборудование выбранного сотрудника"""
        employee = self.employee_var.get()

        # Очищаем таблицу
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        if not employee:
            return

        # Находим оборудование сотрудника
        for item in self.inventory_data:
            if item.get('assignment') == employee:
                self.employee_tree.insert("", "end", values=(
                    item.get('equipment_type', ''),
                    item.get('model', ''),
                    item.get('serial_number', ''),
                    item.get('date', ''),
                    item.get('comments', '')[:50] + '...' if item.get('comments') and len(
                        item.get('comments', '')) > 50 else item.get('comments', '')
                ))

    def show_all_data(self):
        """Показать все данные"""
        # Очищаем таблицу
        for item in self.all_tree.get_children():
            self.all_tree.delete(item)

        # Перезагружаем данные
        self.inventory_data = self.load_data()

        # Заполняем таблицу всеми данными
        for item in self.inventory_data:
            self.all_tree.insert("", "end", values=(
                item.get('equipment_type', ''),
                item.get('model', ''),
                item.get('serial_number', ''),
                item.get('assignment', ''),
                item.get('date', ''),
                item.get('comments', '')[:50] + '...' if item.get('comments') and len(
                    item.get('comments', '')) > 50 else item.get('comments', '')
            ))

    def on_tree_double_click(self, event):
        """Обработка двойного клика для редактирования ячейки"""
        tree = event.widget
        item = tree.identify('item', event.x, event.y)
        column = tree.identify_column(event.x)

        if not item or item not in tree.get_children():
            return

        # Получаем индекс колонки
        col_index = int(column.replace('#', '')) - 1

        # Получаем текущее значение
        current_values = tree.item(item, 'values')
        current_value = current_values[col_index]

        # Определяем поле для редактирования
        field_names = {
            0: 'equipment_type',
            1: 'model',
            2: 'serial_number',
            3: 'assignment',
            4: 'date',
            5: 'comments'
        }

        field_name = field_names.get(col_index)
        if not field_name:
            return

        # Создаем окно редактирования
        self.edit_cell(tree, item, col_index, field_name, current_value)

    def edit_cell(self, tree, item, col_index, field_name, current_value):
        """Редактирование ячейки"""
        # Получаем bounding box ячейки
        bbox = tree.bbox(item, column=f'#{col_index + 1}')

        # Создаем поле ввода
        entry_edit = ttk.Entry(tree, width=bbox[2])
        entry_edit.insert(0, current_value)
        entry_edit.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry_edit.focus()

        def save_edit(event=None):
            """Сохранение изменений"""
            new_value = entry_edit.get().strip()
            entry_edit.destroy()

            # Обновляем значение в дереве
            current_values = list(tree.item(item, 'values'))
            current_values[col_index] = new_value
            tree.item(item, values=current_values)

            # Обновляем значение в данных
            # Находим индекс элемента в inventory_data
            all_items = list(tree.get_children())
            item_index = all_items.index(item)

            if item_index < len(self.inventory_data):
                self.inventory_data[item_index][field_name] = new_value
                self.save_data()

        def cancel_edit(event=None):
            """Отмена редактирования"""
            entry_edit.destroy()

        entry_edit.bind('<Return>', save_edit)
        entry_edit.bind('<Escape>', cancel_edit)
        entry_edit.bind('<FocusOut>', lambda e: save_edit())

    def export_to_pdf(self):
        """Экспорт всех данных в PDF"""
        if not self.inventory_data:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return
        try:
            if not os.path.exists('ChakraPetch-Regular.ttf'):
                messagebox.showerror("Ошибка", "Отсутствует файл шрифта: ChakraPetch-Regular.ttf")
                return
            pdf = PDFWithCyrillic(orientation='L')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            pdf.set_font("ChakraPetch", '', 16)
            pdf.cell(0, 10, "Полный отчет по инвентаризации оборудования", 0, 1, 'C')
            pdf.ln(10)

            pdf.set_font("ChakraPetch", '', 10)
            pdf.cell(0, 10, f"Отчет сгенерирован: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}", 0, 1, 'R')
            pdf.ln(5)

            total_equipment = len(self.inventory_data)
            unique_employees = len(
                set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
            pdf.set_font("ChakraPetch", '', 12)
            pdf.cell(0, 10, "Общая статистика:", 0, 1)
            pdf.cell(0, 10, f"Всего единиц оборудования: {total_equipment}", 0, 1)
            pdf.cell(0, 10, f"Количество сотрудников с оборудованием: {unique_employees}", 0, 1)
            pdf.ln(10)

            pdf.set_font("ChakraPetch", '', 10)
            columns = ["Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии"]
            col_widths = [30, 35, 40, 35, 25, 45]

            for i, col in enumerate(columns):
                pdf.cell(col_widths[i], 10, col, 1, 0, 'C')
            pdf.ln()

            pdf.set_font("ChakraPetch", '', 8)
            for item in self.inventory_data:
                pdf.cell(col_widths[0], 10, item.get('equipment_type', '')[:20] or '-', 1)
                pdf.cell(col_widths[1], 10, item.get('model', '')[:20] or '-', 1)
                pdf.cell(col_widths[2], 10, item.get('serial_number', '')[:15] or '-', 1)
                pdf.cell(col_widths[3], 10, item.get('assignment', '')[:15] or '-', 1)
                pdf.cell(col_widths[4], 10, item.get('date', '')[:10] or '-', 1)
                pdf.cell(col_widths[5], 10, item.get('comments', '')[:25] or '-', 1)
                pdf.ln()

            # Сохраняем PDF файл на рабочий стол
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"inventory_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            pdf.output(pdf_filename)

            webbrowser.open(pdf_filename)
            messagebox.showinfo("Успех", f"Отчет сохранен на рабочем столе:\n{pdf_filename}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF отчет: {e}")


def main():
    """Основная функция"""
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()