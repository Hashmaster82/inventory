import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime
import webbrowser
from fpdf import FPDF
import tkinter.font as tkFont
import shutil


class PDFWithCyrillic(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font('ChakraPetch', '', 'ChakraPetch-Regular.ttf', uni=True)


class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система инвентаризации оборудования")
        self.root.state('zoomed')

        self.default_font = tkFont.Font(family='Arial', size=14)
        self.root.option_add("*Font", self.default_font)

        # Основной файл инвентаря
        self.inventory_file = r"\\fs\SHARE_BH\it\inventory\inventory.json"
        self.inventory_data = self.load_data()

        # Файл типов оборудования
        self.equipment_types_file = "equipment_types.json"
        self.equipment_types = self.load_equipment_types()

        self.create_widgets()

    # =============== РАБОТА С ТИПАМИ ОБОРУДОВАНИЯ ===============
    def load_equipment_types(self):
        try:
            if os.path.exists(self.equipment_types_file):
                with open(self.equipment_types_file, 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    return data if isinstance(data, list) else []
            else:
                default_types = ["Монитор", "Сисблок", "МФУ", "Клавиатура", "Мышь", "Наушники"]
                self.save_equipment_types(default_types)
                return default_types
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить типы оборудования: {e}")
            return []

    def save_equipment_types(self, types_list):
        try:
            with open(self.equipment_types_file, 'w', encoding='utf-8') as file:
                json.dump(types_list, file, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить типы оборудования: {e}")
            return False

    # =============== ОСНОВНАЯ ЛОГИКА ===============
    def load_data(self):
        try:
            if os.path.exists(self.inventory_file):
                with open(self.inventory_file, 'r', encoding='utf-8') as file:
                    return json.load(file)
            else:
                with open(self.inventory_file, 'w', encoding='utf-8') as file:
                    json.dump([], file)
                return []
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {e}")
            return []

    def save_data(self):
        try:
            os.makedirs(os.path.dirname(self.inventory_file), exist_ok=True)
            with open(self.inventory_file, 'w', encoding='utf-8') as file:
                json.dump(self.inventory_data, file, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить данные: {e}")
            return False

    def create_backup(self):
        """Создание резервной копии JSON файла"""
        try:
            if not os.path.exists(self.inventory_file):
                messagebox.showwarning("Предупреждение", "Файл инвентаризации не существует")
                return

            backup_dir = os.path.join(os.path.dirname(self.inventory_file), "backups")
            os.makedirs(backup_dir, exist_ok=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"inventory_backup_{timestamp}.json"
            backup_path = os.path.join(backup_dir, backup_filename)

            shutil.copy2(self.inventory_file, backup_path)
            messagebox.showinfo("Успех", f"Резервная копия создана:\n{backup_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать резервную копию: {e}")

    def bind_clipboard_events(self, widget):
        def do_copy(event):
            widget.event_generate("<<Copy>>")
            return "break"

        def do_cut(event):
            widget.event_generate("<<Cut>>")
            return "break"

        def do_paste(event):
            widget.event_generate("<<Paste>>")
            return "break"

        widget.bind("<Control-c>", do_copy)
        widget.bind("<Control-x>", do_cut)
        widget.bind("<Control-v>", do_paste)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10, fill='x')

        backup_button = ttk.Button(button_frame, text="📂 Создать резервную копию",
                                   command=self.create_backup, style='Big.TButton')
        backup_button.pack(side='left', padx=5, fill='x', expand=True)

        pdf_button = ttk.Button(button_frame, text="📄 Выгрузить полный отчет в PDF",
                                command=self.export_to_pdf, style='Big.TButton')
        pdf_button.pack(side='left', padx=5, fill='x', expand=True)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True)

        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 18, 'bold'))
        style.configure('TNotebook.Tab', font=('Arial', 16, 'bold'), padding=[20, 10])

        self.add_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.add_frame, text="Добавить оборудование")

        self.search_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.search_frame, text="Поиск оборудования")

        self.employee_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.employee_frame, text="Сотрудники")

        self.show_all_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.show_all_frame, text="Показать всё")

        # НОВАЯ ВКЛАДКА: Оборудование
        self.equipment_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.equipment_frame, text="Оборудование")

        # НОВАЯ ВКЛАДКА: Настройки
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="Настройки")

        self.about_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.about_frame, text="Инфо")

        self.create_add_tab()
        self.create_search_tab()
        self.create_employee_tab()
        self.create_show_all_tab()
        self.create_equipment_tab()      # <-- НОВЫЙ МЕТОД
        self.create_settings_tab()       # <-- НОВЫЙ МЕТОД
        self.create_about_tab()

    # =============== ВКЛАДКА: ОБОРУДОВАНИЕ ===============
    def create_equipment_tab(self):
        frame = self.equipment_frame

        ttk.Label(frame, text="Тип оборудования:", font=self.default_font).grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.equipment_type_entry = ttk.Entry(frame, width=40, font=self.default_font)
        self.equipment_type_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.bind_clipboard_events(self.equipment_type_entry)

        add_btn = ttk.Button(frame, text="➕ Добавить тип", command=self.add_equipment_type, style='Big.TButton')
        add_btn.grid(row=0, column=2, padx=10, pady=5)

        # Список типов
        columns = ("Тип оборудования",)
        self.equipment_tree = ttk.Treeview(frame, columns=columns, show='headings', height=15)
        self.equipment_tree.heading("Тип оборудования", text="Тип оборудования")
        self.equipment_tree.column("Тип оборудования", width=300, anchor='center')

        self.equipment_tree.bind('<Button-3>', self.show_equipment_context_menu)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.equipment_tree.yview)
        self.equipment_tree.configure(yscrollcommand=scrollbar.set)

        self.equipment_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        self.equipment_context_menu = tk.Menu(self.equipment_tree, tearoff=0)
        self.equipment_context_menu.add_command(label="🗑️ Удалить тип", command=self.delete_equipment_type)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)

        self.refresh_equipment_list()

    def refresh_equipment_list(self):
        for item in self.equipment_tree.get_children():
            self.equipment_tree.delete(item)
        for eq_type in sorted(self.equipment_types):
            self.equipment_tree.insert("", "end", values=(eq_type,))

    def add_equipment_type(self):
        new_type = self.equipment_type_entry.get().strip()
        if not new_type:
            messagebox.showwarning("Предупреждение", "Введите название типа оборудования")
            return
        if new_type in self.equipment_types:
            messagebox.showwarning("Предупреждение", "Такой тип уже существует")
            return

        self.equipment_types.append(new_type)
        if self.save_equipment_types(self.equipment_types):
            messagebox.showinfo("Успех", "Тип оборудования добавлен")
            self.equipment_type_entry.delete(0, tk.END)
            self.refresh_equipment_list()
            # Обновляем Combobox в add_tab, если он уже создан
            if hasattr(self, 'equipment_type_combo'):
                self.equipment_type_combo['values'] = sorted(self.equipment_types)

    def show_equipment_context_menu(self, event):
        item = self.equipment_tree.identify_row(event.y)
        if item:
            self.equipment_tree.selection_set(item)
            self.equipment_context_menu.post(event.x_root, event.y_root)

    def delete_equipment_type(self):
        selected_items = self.equipment_tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите тип для удаления")
            return

        selected_type = self.equipment_tree.item(selected_items[0], 'values')[0]

        # Проверка: используется ли тип в инвентаре
        in_use = any(item.get('equipment_type') == selected_type for item in self.inventory_data)
        if in_use:
            if not messagebox.askyesno("Подтверждение",
                                       f"Тип '{selected_type}' используется в записях. Удалить все связанные записи?"):
                return
            # Удаляем все записи с этим типом
            self.inventory_data = [item for item in self.inventory_data if item.get('equipment_type') != selected_type]
            self.save_data()
            self.show_all_data()

        self.equipment_types.remove(selected_type)
        if self.save_equipment_types(self.equipment_types):
            messagebox.showinfo("Успех", "Тип оборудования удален")
            self.refresh_equipment_list()
            if hasattr(self, 'equipment_type_combo'):
                self.equipment_type_combo['values'] = sorted(self.equipment_types)

    # =============== ВКЛАДКА: НАСТРОЙКИ ===============
    def create_settings_tab(self):
        frame = self.settings_frame

        ttk.Label(frame, text="Текущий файл базы:", font=self.default_font).pack(pady=10)
        self.current_path_label = ttk.Label(frame, text=self.inventory_file, font=self.default_font, wraplength=800)
        self.current_path_label.pack(pady=5)

        change_btn = ttk.Button(frame, text="📂 Выбрать другой файл", command=self.change_inventory_file, style='Big.TButton')
        change_btn.pack(pady=20)

        ttk.Label(frame, text="⚠️ После смены файла данные будут перезагружены", font=self.default_font, foreground="red").pack(pady=10)

    def change_inventory_file(self):
        new_file = filedialog.askopenfilename(
            title="Выберите JSON-файл инвентаризации",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not new_file:
            return

        # Проверим, можно ли прочитать файл
        try:
            with open(new_file, 'r', encoding='utf-8') as f:
                test_data = json.load(f)
                if not isinstance(test_data, list):
                    raise ValueError("Файл должен содержать массив объектов")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Невозможно загрузить файл:\n{e}")
            return

        self.inventory_file = new_file
        self.inventory_data = self.load_data()
        self.current_path_label.config(text=self.inventory_file)
        self.show_all_data()
        self.refresh_employee_list()
        messagebox.showinfo("Успех", "Файл базы успешно изменен!")

    # =============== ВКЛАДКА: ДОБАВИТЬ ОБОРУДОВАНИЕ ===============
    def create_add_tab(self):
        fields = [
            ("Модель", "model"),
            ("Серийный номер", "serial_number"),
            ("Закрепление", "assignment"),
            ("Дата", "date"),
            ("Комментарии", "comments")
        ]
        self.entries = {}

        # Поле "Тип оборудования" — теперь Combobox
        ttk.Label(self.add_frame, text="Тип оборудования:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.equipment_type_var = tk.StringVar()
        self.equipment_type_combo = ttk.Combobox(self.add_frame, textvariable=self.equipment_type_var,
                                                 values=sorted(self.equipment_types), width=38, font=self.default_font)
        self.equipment_type_combo.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.bind_clipboard_events(self.equipment_type_combo)
        self.entries['equipment_type'] = self.equipment_type_combo  # для совместимости с add_equipment()

        row_offset = 1
        for i, (label_text, field_name) in enumerate(fields):
            label = ttk.Label(self.add_frame, text=label_text + ":")
            label.grid(row=i + row_offset, column=0, sticky='w', padx=10, pady=5)
            if field_name == "comments":
                entry = scrolledtext.ScrolledText(self.add_frame, width=40, height=4, font=self.default_font)
                entry.grid(row=i + row_offset, column=1, padx=10, pady=5, sticky='we')
                self.bind_clipboard_events(entry)
            elif field_name == "date":
                entry = ttk.Entry(self.add_frame, width=40, font=self.default_font)
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
                entry.grid(row=i + row_offset, column=1, padx=10, pady=5, sticky='we')
                self.bind_clipboard_events(entry)
            else:
                entry = ttk.Entry(self.add_frame, width=40, font=self.default_font)
                entry.grid(row=i + row_offset, column=1, padx=10, pady=5, sticky='we')
                self.bind_clipboard_events(entry)
            self.entries[field_name] = entry

        add_button = ttk.Button(self.add_frame, text="Добавить оборудование",
                                command=self.add_equipment, style='Big.TButton')
        add_button.grid(row=len(fields) + row_offset, column=0, columnspan=2, pady=20)
        self.add_frame.columnconfigure(1, weight=1)
        self.add_frame.rowconfigure(len(fields) + row_offset, weight=1)

    # =============== ВКЛАДКА: ПОИСК ОБОРУДОВАНИЯ ===============
    def create_search_tab(self):
        ttk.Label(self.search_frame, text="Поиск:", font=self.default_font).grid(row=0, column=0, sticky='w', padx=10,
                                                                                 pady=5)
        self.search_entry = ttk.Entry(self.search_frame, width=40, font=self.default_font)
        self.search_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.bind_clipboard_events(self.search_entry)
        self.search_entry.bind('<KeyRelease>', self.perform_search)

        clear_button = ttk.Button(self.search_frame, text="Очистить", command=self.clear_search, style='Big.TButton')
        clear_button.grid(row=0, column=2, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.search_tree = ttk.Treeview(self.search_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.search_tree.heading(col, text=col,
                                     command=lambda _col=col: self.treeview_sort_column(self.search_tree, _col, False))
            self.search_tree.column(col, width=150, anchor='center')

        self.search_tree.bind('<Double-1>', self.on_tree_double_click)
        self.search_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(self.search_frame, orient="vertical", command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=scrollbar.set)

        self.search_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        export_pdf_btn = ttk.Button(self.search_frame, text="📄 Экспорт результатов поиска в PDF",
                                    command=self.export_search_results_to_pdf, style='Big.TButton')
        export_pdf_btn.grid(row=2, column=0, columnspan=3, pady=10, sticky='we')

        self.search_context_menu = tk.Menu(self.search_tree, tearoff=0)
        self.search_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.search_frame.columnconfigure(1, weight=1)
        self.search_frame.rowconfigure(1, weight=1)

    # =============== ВКЛАДКА: СОТРУДНИКИ ===============
    def create_employee_tab(self):
        ttk.Label(self.employee_frame, text="Сотрудник:", font=self.default_font).grid(row=0, column=0, sticky='w',
                                                                                       padx=10, pady=5)
        employees = sorted(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
        self.employee_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(self.employee_frame, textvariable=self.employee_var,
                                           values=employees, width=30, font=self.default_font)
        self.employee_combo.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.bind_clipboard_events(self.employee_combo)
        self.employee_combo.bind('<<ComboboxSelected>>', self.show_employee_equipment)

        refresh_button = ttk.Button(self.employee_frame, text="Обновить",
                                    command=self.refresh_employee_list, style='Big.TButton')
        refresh_button.grid(row=0, column=2, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Дата", "Комментарии")
        self.employee_tree = ttk.Treeview(self.employee_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.employee_tree.heading(col, text=col,
                                       command=lambda _col=col: self.treeview_sort_column(self.employee_tree, _col,
                                                                                          False))
            self.employee_tree.column(col, width=150, anchor='center')

        self.employee_tree.bind('<Double-1>', self.on_tree_double_click)
        self.employee_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(self.employee_frame, orient="vertical", command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        # Кнопка экспорта в PDF для раздела сотрудников
        export_pdf_btn = ttk.Button(self.employee_frame, text="📄 Экспорт в PDF",
                                    command=self.export_employee_results_to_pdf, style='Big.TButton')
        export_pdf_btn.grid(row=2, column=0, columnspan=3, pady=10, sticky='we')

        self.employee_context_menu = tk.Menu(self.employee_tree, tearoff=0)
        self.employee_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.employee_frame.columnconfigure(1, weight=1)
        self.employee_frame.rowconfigure(1, weight=1)

    # =============== ВКЛАДКА: ПОКАЗАТЬ ВСЁ ===============
    def create_show_all_tab(self):
        refresh_button = ttk.Button(self.show_all_frame, text="Обновить данные",
                                    command=self.show_all_data, style='Big.TButton')
        refresh_button.pack(pady=10)

        table_frame = ttk.Frame(self.show_all_frame)
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.all_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

        for col in columns:
            self.all_tree.heading(col, text=col,
                                  command=lambda _col=col: self.treeview_sort_column(self.all_tree, _col, False))
            self.all_tree.column(col, width=150, anchor='center')

        self.all_tree.bind('<Double-1>', self.on_tree_double_click)
        self.all_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.all_tree.yview)
        self.all_tree.configure(yscrollcommand=scrollbar.set)

        self.all_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.all_context_menu = tk.Menu(self.all_tree, tearoff=0)
        self.all_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.show_all_data()

    # =============== ВКЛАДКА: ИНФО ===============
    def create_about_tab(self):
        center_frame = ttk.Frame(self.about_frame)
        center_frame.pack(expand=True, fill='both')

        info_text = """
        Система инвентаризации оборудования
        Версия: 0.7 (обновлено)
        Разработано: Разин Григорий
        Контактная информация:
        Email: lantester35@gmail.com
        Функционал:
        - Ведение учета оборудования
        - Поиск и фильтрация данных
        - Экспорт отчетов в PDF
        - Управление закреплением за сотрудниками
        - Управление типами оборудования
        - Настройка пути к файлу базы
        - Удаление записей (правый клик или кнопка удаления)
        """

        about_text = scrolledtext.ScrolledText(center_frame, width=60, height=15,
                                               font=self.default_font, wrap=tk.WORD)
        about_text.insert('1.0', info_text)
        about_text.config(state='disabled')
        about_text.pack(pady=20, padx=20, expand=True)

        close_button = ttk.Button(center_frame, text="Закрыть", command=self.root.quit, style='Big.TButton')
        close_button.pack(pady=10)

    # =============== ОБЩИЕ МЕТОДЫ ===============
    def treeview_sort_column(self, tree, col, reverse):
        date_columns = ['Дата']
        int_columns = []

        def sort_key(val):
            if col in date_columns:
                try:
                    return datetime.strptime(val, "%d.%m.%Y")
                except:
                    return datetime.min
            elif col in int_columns:
                try:
                    return int(val)
                except:
                    return -1
            else:
                return val.lower() if isinstance(val, str) else val

        data = [(tree.set(k, col), k) for k in tree.get_children('')]
        data.sort(key=lambda t: sort_key(t[0]), reverse=reverse)

        for index, (_, k) in enumerate(data):
            tree.move(k, '', index)

        serial_col = 'Серийный номер' if 'Серийный номер' in tree['columns'] else 'serial_number'

        sorted_serials = [tree.set(k, serial_col) for _, k in data]

        new_inventory = []
        for serial in sorted_serials:
            for item in self.inventory_data:
                if item.get('serial_number') == serial:
                    new_inventory.append(item)
                    break
        self.inventory_data = new_inventory

        tree.heading(col, command=lambda: self.treeview_sort_column(tree, col, not reverse))

    def show_context_menu(self, event):
        tree = event.widget
        item = tree.identify_row(event.y)

        if item:
            tree.selection_set(item)
            if tree == self.search_tree:
                self.search_context_menu.post(event.x_root, event.y_root)
            elif tree == self.employee_tree:
                self.employee_context_menu.post(event.x_root, event.y_root)
            elif tree == self.all_tree:
                self.all_context_menu.post(event.x_root, event.y_root)

    def delete_selected_item(self):
        current_tab = self.notebook.index(self.notebook.select())

        if current_tab == 1:
            tree = self.search_tree
        elif current_tab == 2:
            tree = self.employee_tree
        elif current_tab == 3:
            tree = self.all_tree
        else:
            return

        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Предупреждение", "Выберите запись для удаления")
            return

        if not messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить выбранную запись?"):
            return

        for item in selected_items:
            values = tree.item(item, 'values')
            serial_number = values[2]

            for i, inv_item in enumerate(self.inventory_data):
                if inv_item.get('serial_number') == serial_number:
                    del self.inventory_data[i]
                    tree.delete(item)
                    if self.save_data():
                        messagebox.showinfo("Успех", "Запись успешно удалена")
                    self.refresh_employee_list()
                    if current_tab != 3:
                        self.show_all_data()
                    break

    def add_equipment(self):
        equipment_data = {}

        for field_name, entry in self.entries.items():
            if field_name == "comments":
                equipment_data[field_name] = entry.get("1.0", tk.END).strip()
            else:
                equipment_data[field_name] = entry.get().strip()

        if not equipment_data.get('equipment_type') or not equipment_data.get('serial_number'):
            messagebox.showwarning("Предупреждение", "Заполните обязательные поля: Тип оборудования и Серийный номер")
            return

        equipment_data['created_datetime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.inventory_data.append(equipment_data)

        if self.save_data():
            messagebox.showinfo("Успех", "Оборудование успешно добавлено!")
            self.clear_entries()
            self.refresh_employee_list()
            self.show_all_data()

    def clear_entries(self):
        for field_name, entry in self.entries.items():
            if field_name == "comments":
                entry.delete("1.0", tk.END)
            elif field_name == "date":
                entry.delete(0, tk.END)
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
            else:
                entry.delete(0, tk.END)
        # Также очистим Combobox
        if 'equipment_type' in self.entries:
            self.entries['equipment_type'].set('')

    def perform_search(self, event=None):
        search_text = self.search_entry.get().lower().strip()

        for item in self.search_tree.get_children():
            self.search_tree.delete(item)

        if not search_text:
            return

        for item in self.inventory_data:
            if any(search_text in str(value).lower() for value in item.values() if value):
                self.search_tree.insert("", "end", values=(
                    item.get('equipment_type', ''),
                    item.get('model', ''),
                    item.get('serial_number', ''),
                    item.get('assignment', ''),
                    item.get('date', ''),
                    (item.get('comments', '')[:50] + '...') if item.get('comments') and len(
                        item.get('comments')) > 50 else item.get('comments', '')
                ))

    def clear_search(self):
        self.search_entry.delete(0, tk.END)
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)

    def refresh_employee_list(self):
        employees = sorted(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
        self.employee_combo['values'] = employees

    def show_employee_equipment(self, event=None):
        employee = self.employee_var.get()

        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)

        if not employee:
            return

        for item in self.inventory_data:
            if item.get('assignment') == employee:
                self.employee_tree.insert("", "end", values=(
                    item.get('equipment_type', ''),
                    item.get('model', ''),
                    item.get('serial_number', ''),
                    item.get('date', ''),
                    (item.get('comments', '')[:50] + '...') if item.get('comments') and len(
                        item.get('comments')) > 50 else item.get('comments', '')
                ))

    def show_all_data(self):
        for item in self.all_tree.get_children():
            self.all_tree.delete(item)
        self.inventory_data = self.load_data()
        for item in self.inventory_data:
            self.all_tree.insert("", "end", values=(
                item.get('equipment_type', ''),
                item.get('model', ''),
                item.get('serial_number', ''),
                item.get('assignment', ''),
                item.get('date', ''),
                (item.get('comments', '')[:50] + '...') if item.get('comments') and len(
                    item.get('comments')) > 50 else item.get('comments', '')
            ))

    def on_tree_double_click(self, event):
        tree = event.widget
        item = tree.identify('item', event.x, event.y)
        column = tree.identify_column(event.x)

        if not item or item not in tree.get_children():
            return

        col_index = int(column.replace('#', '')) - 1
        current_values = tree.item(item, 'values')
        current_value = current_values[col_index]

        # Маппинг колонок (учитываем, что в employee_tree меньше колонок)
        if tree == self.employee_tree:
            field_names = {
                0: 'equipment_type',
                1: 'model',
                2: 'serial_number',
                3: 'date',
                4: 'comments'
            }
        else:
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

        self.edit_cell(tree, item, col_index, field_name, current_value)

    def edit_cell(self, tree, item, col_index, field_name, current_value):
        bbox = tree.bbox(item, column=f'#{col_index + 1}')
        if not bbox:  # иногда bbox возвращает пустой кортеж, если элемент не виден
            return

        if field_name == 'comments':
            text_edit = scrolledtext.ScrolledText(tree, width=40, height=4, font=self.default_font)
            text_edit.insert('1.0', current_value)
            text_edit.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3] * 3)
            text_edit.focus()
            self.bind_clipboard_events(text_edit)

            def save_edit(event=None):
                new_value = text_edit.get('1.0', tk.END).strip()
                text_edit.destroy()
                current_values = list(tree.item(item, 'values'))
                current_values[col_index] = new_value
                tree.item(item, values=current_values)
                serial_number = current_values[2] if len(current_values) > 2 else None
                if serial_number:
                    for inventory_item in self.inventory_data:
                        if inventory_item.get('serial_number') == serial_number:
                            inventory_item[field_name] = new_value
                            break
                    self.save_data()

            def cancel_edit(event=None):
                text_edit.destroy()

            text_edit.bind('<Return>', save_edit)
            text_edit.bind('<Escape>', cancel_edit)
            text_edit.bind('<FocusOut>', lambda e: save_edit())
        else:
            entry_edit = ttk.Entry(tree, width=bbox[2] // 8, font=self.default_font)  # ширина в символах
            entry_edit.insert(0, current_value)
            entry_edit.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            entry_edit.focus()
            self.bind_clipboard_events(entry_edit)

            def save_edit(event=None):
                new_value = entry_edit.get().strip()
                entry_edit.destroy()
                current_values = list(tree.item(item, 'values'))
                current_values[col_index] = new_value
                tree.item(item, values=current_values)
                serial_number = current_values[2] if len(current_values) > 2 else None
                if serial_number:
                    for inventory_item in self.inventory_data:
                        if inventory_item.get('serial_number') == serial_number:
                            inventory_item[field_name] = new_value
                            break
                    self.save_data()

            def cancel_edit(event=None):
                entry_edit.destroy()

            entry_edit.bind('<Return>', save_edit)
            entry_edit.bind('<Escape>', cancel_edit)
            entry_edit.bind('<FocusOut>', lambda e: save_edit())

    def export_to_pdf(self):
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

            pdf.set_font("ChakraPetch", '', 18)
            pdf.cell(0, 10, "Полный отчет по инвентаризации оборудования", 0, 1, 'C')
            pdf.ln(10)

            pdf.set_font("ChakraPetch", '', 12)
            pdf.cell(0, 10, f"Отчет сгенерирован: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}", 0, 1, 'R')
            pdf.ln(5)

            total_equipment = len(self.inventory_data)
            unique_employees = len(
                set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
            pdf.set_font("ChakraPetch", '', 14)
            pdf.cell(0, 10, "Общая статистика:", 0, 1)
            pdf.cell(0, 10, f"Всего единиц оборудования: {total_equipment}", 0, 1)
            pdf.cell(0, 10, f"Количество сотрудников с оборудованием: {unique_employees}", 0, 1)
            pdf.ln(10)

            columns = ["Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии"]

            data_rows = []
            for item in self.inventory_data:
                row = [
                    item.get('equipment_type', '') or '-',
                    item.get('model', '') or '-',
                    item.get('serial_number', '') or '-',
                    item.get('assignment', '') or '-',
                    item.get('date', '') or '-',
                    (item.get('comments', '')[:50] + '...') if item.get('comments') and len(
                        item.get('comments')) > 50 else (item.get('comments', '') or '-')
                ]
                data_rows.append(row)

            pdf.set_font("ChakraPetch", '', 12)
            col_widths = []
            for col_index in range(len(columns)):
                max_width = pdf.get_string_width(columns[col_index]) + 6
                for row in data_rows:
                    w = pdf.get_string_width(str(row[col_index])) + 6
                    if w > max_width:
                        max_width = w
                col_widths.append(max_width)

            pdf.set_font("ChakraPetch", '', 14)
            for i, col in enumerate(columns):
                pdf.cell(col_widths[i], 10, col, 1, 0, 'C')
            pdf.ln()

            pdf.set_font("ChakraPetch", '', 12)
            for row in data_rows:
                for i, cell_text in enumerate(row):
                    pdf.cell(col_widths[i], 10, str(cell_text), 1)
                pdf.ln()

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"inventory_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            pdf.output(pdf_filename)

            webbrowser.open(pdf_filename)
            messagebox.showinfo("Успех", f"Отчет сохранен на рабочем столе:\n{pdf_filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF отчет: {e}")

    def export_search_results_to_pdf(self):
        items = self.search_tree.get_children()
        if not items:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return

        try:
            if not os.path.exists('ChakraPetch-Regular.ttf'):
                messagebox.showerror("Ошибка", "Отсутствует файл шрифта: ChakraPetch-Regular.ttf")
                return

            pdf = PDFWithCyrillic(orientation='L')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            pdf.set_font("ChakraPetch", '', 14)
            pdf.cell(0, 10, "Отчет по результатам поиска оборудования", 0, 1, 'C')
            pdf.ln(10)

            columns = ["Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии"]

            data_rows = [self.search_tree.item(item, 'values') for item in items]

            pdf.set_font("ChakraPetch", '', 12)
            col_widths = []
            for col_index in range(len(columns)):
                max_width = pdf.get_string_width(columns[col_index]) + 6
                for row in data_rows:
                    w = pdf.get_string_width(str(row[col_index])) + 6
                    if w > max_width:
                        max_width = w
                col_widths.append(max_width)

            pdf.set_font("ChakraPetch", '', 14)
            for i, col in enumerate(columns):
                pdf.cell(col_widths[i], 10, col, 1, 0, 'C')
            pdf.ln()

            pdf.set_font("ChakraPetch", '', 12)
            for row in data_rows:
                for i, cell_text in enumerate(row):
                    pdf.cell(col_widths[i], 10, str(cell_text), 1)
                pdf.ln()

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            pdf.output(pdf_filename)

            webbrowser.open(pdf_filename)
            messagebox.showinfo("Успех", f"Отчет сохранен на рабочем столе:\n{pdf_filename}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF отчет: {e}")

    def export_employee_results_to_pdf(self):
        items = self.employee_tree.get_children()
        if not items:
            messagebox.showwarning("Предупреждение", "Нет данных для экспорта")
            return

        try:
            if not os.path.exists('ChakraPetch-Regular.ttf'):
                messagebox.showerror("Ошибка", "Отсутствует файл шрифта: ChakraPetch-Regular.ttf")
                return

            pdf = PDFWithCyrillic(orientation='L')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            employee_name = self.employee_var.get()
            pdf.set_font("ChakraPetch", '', 14)
            pdf.cell(0, 10, f"Отчет по оборудованию сотрудника: {employee_name}", 0, 1, 'C')
            pdf.ln(10)

            columns = ["Тип", "Модель", "Серийный номер", "Дата", "Комментарии"]

            data_rows = [self.employee_tree.item(item, 'values') for item in items]

            pdf.set_font("ChakraPetch", '', 12)
            col_widths = []
            for col_index in range(len(columns)):
                max_width = pdf.get_string_width(columns[col_index]) + 6
                for row in data_rows:
                    w = pdf.get_string_width(str(row[col_index])) + 6
                    if w > max_width:
                        max_width = w
                col_widths.append(max_width)

            pdf.set_font("ChakraPetch", '', 14)
            for i, col in enumerate(columns):
                pdf.cell(col_widths[i], 10, col, 1, 0, 'C')
            pdf.ln()

            pdf.set_font("ChakraPetch", '', 12)
            for row in data_rows:
                for i, cell_text in enumerate(row):
                    pdf.cell(col_widths[i], 10, str(cell_text), 1)
                pdf.ln()

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"employee_equipment_{employee_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            pdf.output(pdf_filename)

            webbrowser.open(pdf_filename)
            messagebox.showinfo("Успех", f"Отчет сохранен на рабочем столе:\n{pdf_filename}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF отчет: {e}")


def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()