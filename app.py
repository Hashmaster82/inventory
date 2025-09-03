import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import webbrowser
from fpdf import FPDF  # Импорт FPDF для PDF создания


class PDFWithCyrillic(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.add_font('ChakraPetch', '', 'ChakraPetch-Regular.ttf', uni=True)


class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система инвентаризации оборудования")
        self.root.geometry("900x700")

        self.inventory_file = r"\\fs\SHARE_BH\it\inventory\inventory.json"
        self.inventory_data = self.load_data()
        self.create_widgets()

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

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        pdf_button = ttk.Button(main_frame, text="📄 Выгрузить полный отчет в PDF",
                                command=self.export_to_pdf, style='Big.TButton')
        pdf_button.pack(pady=10, fill='x')

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill='both', expand=True)

        self.add_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.add_frame, text="Добавить оборудование")

        self.search_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.search_frame, text="Поиск оборудования")

        self.employee_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.employee_frame, text="Оборудование по сотрудникам")

        self.show_all_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.show_all_frame, text="Показать всё")

        self.about_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.about_frame, text="About")

        self.create_add_tab()
        self.create_search_tab()
        self.create_employee_tab()
        self.create_show_all_tab()
        self.create_about_tab()

        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 12, 'bold'))

    def create_add_tab(self):
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
                entry = ttk.Entry(self.add_frame, width=40)
                entry.insert(0, datetime.now().strftime("%d.%m.%Y"))
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')
            else:
                entry = ttk.Entry(self.add_frame, width=40)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='we')

            self.entries[field_name] = entry

        add_button = ttk.Button(self.add_frame, text="Добавить оборудование",
                                command=self.add_equipment)
        add_button.grid(row=len(fields), column=0, columnspan=2, pady=20)
        self.add_frame.columnconfigure(1, weight=1)
        self.add_frame.rowconfigure(len(fields), weight=1)

    def create_search_tab(self):
        ttk.Label(self.search_frame, text="Поиск:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        self.search_entry = ttk.Entry(self.search_frame, width=40)
        self.search_entry.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.search_entry.bind('<KeyRelease>', self.perform_search)

        clear_button = ttk.Button(self.search_frame, text="Очистить", command=self.clear_search)
        clear_button.grid(row=0, column=2, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.search_tree = ttk.Treeview(self.search_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.search_tree.heading(col, text=col)
            self.search_tree.column(col, width=100)

        self.search_tree.bind('<Double-1>', self.on_tree_double_click)
        self.search_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(self.search_frame, orient="vertical", command=self.search_tree.yview)
        self.search_tree.configure(yscrollcommand=scrollbar.set)

        self.search_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        # Кнопка экспорта результатов поиска в PDF
        export_pdf_btn = ttk.Button(self.search_frame, text="📄 Экспорт результатов поиска в PDF",
                                    command=self.export_search_results_to_pdf)
        export_pdf_btn.grid(row=2, column=0, columnspan=3, pady=10, sticky='we')

        self.search_context_menu = tk.Menu(self.search_tree, tearoff=0)
        self.search_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.search_frame.columnconfigure(1, weight=1)
        self.search_frame.rowconfigure(1, weight=1)

    def create_employee_tab(self):
        ttk.Label(self.employee_frame, text="Сотрудник:").grid(row=0, column=0, sticky='w', padx=10, pady=5)
        employees = sorted(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
        self.employee_var = tk.StringVar()
        self.employee_combo = ttk.Combobox(self.employee_frame, textvariable=self.employee_var,
                                           values=employees, width=30)
        self.employee_combo.grid(row=0, column=1, padx=10, pady=5, sticky='we')
        self.employee_combo.bind('<<ComboboxSelected>>', self.show_employee_equipment)

        refresh_button = ttk.Button(self.employee_frame, text="Обновить",
                                    command=self.refresh_employee_list)
        refresh_button.grid(row=0, column=2, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Дата", "Комментарии")
        self.employee_tree = ttk.Treeview(self.employee_frame, columns=columns, show='headings', height=15)

        for col in columns:
            self.employee_tree.heading(col, text=col)
            self.employee_tree.column(col, width=120)

        self.employee_tree.bind('<Double-1>', self.on_tree_double_click)
        self.employee_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(self.employee_frame, orient="vertical", command=self.employee_tree.yview)
        self.employee_tree.configure(yscrollcommand=scrollbar.set)

        self.employee_tree.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky='nsew')
        scrollbar.grid(row=1, column=3, sticky='ns', pady=5)

        self.employee_context_menu = tk.Menu(self.employee_tree, tearoff=0)
        self.employee_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.employee_frame.columnconfigure(1, weight=1)
        self.employee_frame.rowconfigure(1, weight=1)

    def create_show_all_tab(self):
        refresh_button = ttk.Button(self.show_all_frame, text="Обновить данные",
                                    command=self.show_all_data)
        refresh_button.pack(pady=10)

        table_frame = ttk.Frame(self.show_all_frame)
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)

        columns = ("Тип", "Модель", "Серийный номер", "Закрепление", "Дата", "Комментарии")
        self.all_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=20)

        for col in columns:
            self.all_tree.heading(col, text=col)
            self.all_tree.column(col, width=100)

        self.all_tree.bind('<Double-1>', self.on_tree_double_click)
        self.all_tree.bind('<Button-3>', self.show_context_menu)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.all_tree.yview)
        self.all_tree.configure(yscrollcommand=scrollbar.set)

        delete_button = ttk.Button(self.show_all_frame, text="Удалить выбранную запись",
                                   command=self.delete_selected_item)
        delete_button.pack(pady=5)

        self.all_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.all_context_menu = tk.Menu(self.all_tree, tearoff=0)
        self.all_context_menu.add_command(label="Удалить запись", command=self.delete_selected_item)

        self.show_all_data()

    def create_about_tab(self):
        center_frame = ttk.Frame(self.about_frame)
        center_frame.pack(expand=True, fill='both')

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
        - Удаление записей (правый клик или кнопка удаления)
        """

        about_text = scrolledtext.ScrolledText(center_frame, width=60, height=15,
                                               font=('Arial', 11), wrap=tk.WORD)
        about_text.insert('1.0', info_text)
        about_text.config(state='disabled')
        about_text.pack(pady=20, padx=20, expand=True)

        close_button = ttk.Button(center_frame, text="Закрыть", command=self.root.quit)
        close_button.pack(pady=10)

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
                    (item.get('comments', '')[:50] + '...') if item.get('comments') and len(item.get('comments')) > 50 else item.get('comments', '')
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
                    (item.get('comments', '')[:50] + '...') if item.get('comments') and len(item.get('comments')) > 50 else item.get('comments', '')
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
                (item.get('comments', '')[:50] + '...') if item.get('comments') and len(item.get('comments')) > 50 else item.get('comments', '')
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
        entry_edit = ttk.Entry(tree, width=bbox[2])
        entry_edit.insert(0, current_value)
        entry_edit.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry_edit.focus()

        def save_edit(event=None):
            new_value = entry_edit.get().strip()
            entry_edit.destroy()

            current_values = list(tree.item(item, 'values'))
            current_values[col_index] = new_value
            tree.item(item, values=current_values)

            serial_number = current_values[2]

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

            pdf.set_font("ChakraPetch", '', 16)
            pdf.cell(0, 10, "Полный отчет по инвентаризации оборудования", 0, 1, 'C')
            pdf.ln(10)

            pdf.set_font("ChakraPetch", '', 10)
            pdf.cell(0, 10, f"Отчет сгенерирован: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}", 0, 1, 'R')
            pdf.ln(5)

            total_equipment = len(self.inventory_data)
            unique_employees = len(set(item.get('assignment', '') for item in self.inventory_data if item.get('assignment')))
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

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"inventory_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            pdf.output(pdf_filename)

            webbrowser.open(pdf_filename)
            messagebox.showinfo("Успех", f"Отчет сохранен на рабочем столе:\n{pdf_filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать PDF отчет: {e}")

    # Новый метод для экспорта PDF из результатов поиска
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
            col_widths = [30, 35, 40, 35, 25, 45]

            for i, col in enumerate(columns):
                pdf.cell(col_widths[i], 10, col, 1, 0, 'C')
            pdf.ln()

            pdf.set_font("ChakraPetch", '', 9)
            for item in items:
                values = self.search_tree.item(item, 'values')
                for i, val in enumerate(values):
                    text = val if val else '-'
                    pdf.cell(col_widths[i], 10, str(text)[:30], 1)
                pdf.ln()

            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            pdf_filename = os.path.join(desktop_path,
                                        f"search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
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
