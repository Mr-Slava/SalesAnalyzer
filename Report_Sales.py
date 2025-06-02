import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import matplotlib.ticker as ticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import filedialog
import matplotlib.dates as mdates
from matplotlib.figure import Figure


file_path = filedialog.askopenfilename(title="Выберите файл Excel", filetypes=[("Excel files", "*.xlsx")])
if not file_path:
    messagebox.showerror("Ошибка", "Файл не выбран")
    exit()

data = pd.read_excel(file_path)

button_options = {
    "width": 40,  # Ширина в символах — подогнана под самую длинную строку
    "bg": "#39b54a",  # Зеленый фон
    "fg": "white",    # Белый текст
    "font": ("Arial", 12),  # По желанию: чуть крупнее и ровнее
    "relief": "raised",     # Эффект объема
    "bd": 2                 # Толщина границы
}


# Основное окно приложения
root = Tk()
root.title("Отчеты")
root.geometry("400x400+700+300")
root.configure(bg="#dff5e1")

# Верхняя полоса
TOP_BAR_HEIGHT = 60
WINDOW_WIDTH = 400

top_bar = Frame(root, bg="#1e1e1e", height=TOP_BAR_HEIGHT)
top_bar.pack(fill='x', side='top')

def resource_path(relative_path):
    """ Возвращает абсолютный путь, даже если .exe """
    try:
        base_path = sys._MEIPASS  # при запуске из .exe
    except Exception:
        base_path = os.path.abspath(".")  # при запуске из .py

    return os.path.join(base_path, relative_path)

# Загрузка и масштабирование PNG под ширину окна
logo_path = resource_path("logo.png")
original_image = Image.open(logo_path)
resized_image = original_image.resize((WINDOW_WIDTH, TOP_BAR_HEIGHT), Image.Resampling.LANCZOS)
logo_photo = ImageTk.PhotoImage(resized_image)

# Отображение изображения
logo_label = Label(top_bar, image=logo_photo, bg="#1e1e1e")
logo_label.place(x=0, y=0, relwidth=1, relheight=1)

# Центральный фрейм с кнопками
center_frame = Frame(root, bg="#e8f5e9")
center_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

# Функция для работы с кнопками для сброса чекбоксов
def toggle_checkboxes(vars_dict, button):
    if button['text'] == "Выбрать все":
        for var in vars_dict.values():
            var.set(1)
        button['text'] = "Сбросить все"
    else:
        for var in vars_dict.values():
            var.set(0)
        button['text'] = "Выбрать все"

# Функция для форматирования оси Y
def millions(x, pos):
    return f'{x * 1e-6:.0f}M'  # Форматирование в миллионы


# Функция для отображения Общей статистики по продажам
def show_plot(data, title):
    data = pd.DataFrame(data)
    data.index = data.index.astype(int)  # Удаляем дробные части у годов

    n_years = len(data)
    fig_width = max(10, n_years * 0.7)

    # Создание окна
    plot_window = Toplevel()
    plot_window.title(title)
    plot_window.geometry("1200x700+300+100")
    plot_window.configure(bg='#f5f5f5')

    # Закрытие по Esc
    plot_window.bind('<Escape>', lambda e: plot_window.destroy())

    # Создание фигуры matplotlib
    fig = Figure(figsize=(fig_width, 6), dpi=100)
    ax = fig.add_subplot(111)

    data.plot(kind='bar', stacked=True, ax=ax, width=0.7)

    ax.set_title(title)
    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    ax.legend(title='Категория')
    ax.tick_params(axis='x', rotation=30)
    ax.grid(axis='y')

    totals = data.sum(axis=1)
    y_max = totals.max() * 1.15
    ax.set_ylim(0, y_max)

    for idx, total in enumerate(totals):
        ax.text(
            idx,
            total + y_max * 0.01,
            f'{total / 1e6:.1f} М',
            ha='center',
            va='bottom',
            fontsize=9,
            color='black',
            rotation=0
        )

    fig.tight_layout()

    # Вставка графика в окно
    canvas = FigureCanvasTkAgg(fig, master=plot_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=BOTH, expand=True, padx=10, pady=10)

    # Кнопка закрытия
    Button(plot_window, text="Закрыть (Esc)", command=plot_window.destroy,
           bg='#ff4444', fg='white').pack(side=BOTTOM, pady=10)
    plt.close(fig)
#---------------------------------

# Функция для отображения Статистики продаж по месяца за последние 5 лет
def show_monthly_sales():
    # Подготовка данных
    data['SalesDate'] = pd.to_datetime(data['SalesDate'])
    last_5_years = data[data['SalesDate'] >= (pd.Timestamp.now() - pd.DateOffset(years=5))].copy()
    last_5_years['YearMonth'] = last_5_years['SalesDate'].dt.to_period('M')

    monthly_sales = last_5_years.groupby('YearMonth')['Total'].sum().reset_index()
    monthly_sales['YearMonth'] = monthly_sales['YearMonth'].dt.to_timestamp()
    monthly_sales['Year'] = monthly_sales['YearMonth'].dt.year

    unique_years = sorted(monthly_sales['Year'].unique())
    n_years = len(unique_years)
    n_rows = (n_years + 1) // 2

    # Окно
    graph_window = Toplevel(root)
    graph_window.title("Статистика продаж по месяцам")
    graph_window.geometry("1300x900+300+50")

    # Главный контейнер
    main_frame = Frame(graph_window, bg='#f5f5f5')
    main_frame.pack(fill=BOTH, expand=True)

    # Область с прокруткой
    border_frame = Frame(main_frame, bg='#e0e0e0')
    border_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    canvas_frame = Canvas(border_frame, bg='white', highlightthickness=0)
    scrollbar = ttk.Scrollbar(border_frame, orient=VERTICAL, command=canvas_frame.yview)
    canvas_frame.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=RIGHT, fill=Y)
    canvas_frame.pack(side=LEFT, fill=BOTH, expand=True)

    # Вложенный фрейм для графиков
    graph_frame = Frame(canvas_frame, bg='white')
    canvas_window = canvas_frame.create_window((0, 0), window=graph_frame, anchor='nw')

    # Привязка прокрутки
    def on_frame_configure(event):
        canvas_frame.configure(scrollregion=canvas_frame.bbox("all"))

    def on_canvas_configure(event):
        canvas_frame.itemconfig(canvas_window, width=event.width)

    graph_frame.bind("<Configure>", on_frame_configure)
    canvas_frame.bind("<Configure>", on_canvas_configure)

    # Подготовка графиков
    fig, axs = plt.subplots(n_rows, 2, figsize=(12, 2 * n_years), dpi=100)  # немного увеличено
    fig.patch.set_edgecolor('#cccccc')
    fig.patch.set_linewidth(2)
    plt.subplots_adjust(hspace=0.8, wspace=0.4)

    if n_years == 1:
        axs = [[axs]]
    elif n_rows == 1:
        axs = [axs]

    max_height = 0
    for i, year in enumerate(unique_years):
        row = i // 2
        col = i % 2
        ax = axs[row][col]

        yearly_data = monthly_sales[monthly_sales['Year'] == year]
        max_height = max(max_height, yearly_data['Total'].max() * 1.15)

        bars = ax.bar(yearly_data['YearMonth'], yearly_data['Total'],
                      width=20, color='#1f77b4', edgecolor='white', linewidth=0.7)

        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height * 1.01,
                    f"{height / 1e6:.1f}M",
                    ha='center', va='bottom', fontsize=9)

        ax.set_title(f"{year} год", fontsize=12, pad=15)
        ax.yaxis.set_major_formatter(FuncFormatter(millions))
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
        ax.tick_params(axis='x', rotation=45, labelsize=9)
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        ax.set_facecolor('#f9f9f9')

    for ax_row in axs:
        for ax in ax_row:
            if hasattr(ax, 'set_ylim'):
                ax.set_ylim(0, max_height)

    for i in range(n_years, n_rows * 2):
        row = i // 2
        col = i % 2
        axs[row][col].axis('off')

    fig.suptitle("Динамика продаж за последние 5 лет", y=0.98, fontsize=14)

    # Вставка графика
    fig_canvas = FigureCanvasTkAgg(fig, master=graph_frame)
    fig_canvas.draw()
    fig_canvas.get_tk_widget().pack(fill=BOTH, expand=True, padx=10, pady=10)

    # Закрываем фигуру matplotlib, чтобы не оставалась "в памяти"
    plt.close(fig)

    # Прокрутка колесиком мыши
    def on_mousewheel(event):
        canvas_frame.yview_scroll(-1 * int(event.delta / 120), "units")
        return "break"

    graph_window.bind_all("<MouseWheel>", on_mousewheel)

    # Кнопка закрытия
    Button(main_frame, text="Закрыть (Esc)", command=graph_window.destroy,
           bg='#ff4444', fg='white').pack(side=BOTTOM, pady=10)
    graph_window.bind('<Escape>', lambda e: graph_window.destroy())

    graph_window.update_idletasks()
    canvas_frame.configure(scrollregion=canvas_frame.bbox("all"))
#---------------------------------

# Функция для отображения статистики продаж по "Product" за последние 5 лет
def show_product_sales():
    data['SalesDate'] = pd.to_datetime(data['SalesDate'])
    last_5_years = data[data['SalesDate'] >= (pd.Timestamp.now() - pd.DateOffset(years=5))].copy()
    last_5_years['Year'] = last_5_years['SalesDate'].dt.year
    product_sales = last_5_years.groupby(['Year', 'Product'])['Total'].sum().reset_index()

    base_x, base_y = 700, 300
    offset = 100
    product_window = Toplevel(root)
    product_window.title("Select Products to Display")
    product_window.geometry(f"500x450+{base_x + offset}+{base_y + offset}")
    product_window.configure(bg="#dff5e1")

    header = Label(product_window, text="Выберите продукт", bg="#4caf50", fg="white", font=("Arial", 14, "bold"))
    header.pack(fill='x')

    frame_checkbox_area = Frame(product_window, bg="#dff5e1")
    frame_checkbox_area.pack(fill="both", expand=True, padx=10, pady=10)

    canvas = Canvas(frame_checkbox_area, bg="#dff5e1", highlightthickness=0)
    scrollbar = Scrollbar(frame_checkbox_area, orient="vertical", command=canvas.yview)
    scroll_frame = Frame(canvas, bg="#dff5e1")

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    global product_vars
    product_vars = {}
    checkbox_widgets = {}
    for product in product_sales['Product'].unique():
        var = IntVar(value=0)
        product_vars[product] = var
        cb = Checkbutton(scroll_frame, text=product, variable=var, bg="#dff5e1", anchor='w')
        cb.pack(anchor='w')
        checkbox_widgets[product] = cb

    frame_buttons = Frame(product_window, bg="#dff5e1")
    frame_buttons.pack(pady=5)

    def toggle_checkboxes(vars_dict, button):
        if button['text'] == "Стереть все":
            for var in vars_dict.values():
                var.set(0)
            button['text'] = "Выбрать все"
        else:
            for var in vars_dict.values():
                var.set(1)
            button['text'] = "Стереть все"

    toggle_btn = Button(frame_buttons, text="Выбрать все", command=lambda: toggle_checkboxes(product_vars, toggle_btn), **button_options)
    toggle_btn.pack(pady=5)

    Button(frame_buttons, text="Построить график", command=lambda: plot_selected_products(product_vars, product_sales), **button_options).pack(pady=5)


def plot_selected_products(product_vars, product_sales):
    selected_products = [product for product, var in product_vars.items() if var.get() == 1]
    if not selected_products:
        return

    filtered_data = product_sales[product_sales['Product'].isin(selected_products)]
    pivot_data = filtered_data.pivot(index='Year', columns='Product', values='Total').fillna(0)

    # Создаем окно Tkinter
    table_window = tk.Toplevel()
    table_window.title("Статистика продаж по продуктам")
    table_window.geometry("1500x950+200+50")
    table_window.bind("<Escape>", lambda event: table_window.destroy())

    # --- ГРАФИК ---
    fig = plt.figure(figsize=(14, 6))
    gs = fig.add_gridspec(1, 1, left=0.1, right=0.7)
    ax = fig.add_subplot(gs[0])

    colors = plt.cm.get_cmap('tab20', len(selected_products))
    color_dict = {product: colors(i) for i, product in enumerate(selected_products)}

    bars = pivot_data.plot(kind='bar', ax=ax, width=0.8, color=[color_dict[col] for col in pivot_data.columns])

    ax.set_title("Статистика продаж по 'Product' за последние 5 лет", fontsize=14, pad=20)
    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    ax.grid(True, linestyle='--', alpha=0.7)

    leg = ax.legend(
        title='Продукты',
        bbox_to_anchor=(1.05, 1),
        loc='upper left',
        borderaxespad=0.,
        prop={'size': 9}
    )
    leg.get_title().set_fontsize(10)

    canvas = FigureCanvasTkAgg(fig, master=table_window)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # --- ТАБЛИЦА С ИНТЕРАКТИВНОЙ СОРТИРОВКОЙ ---
    table_container = ttk.Frame(table_window)
    table_container.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    # Добавляем столбец с годами и итоговую строку
    totals = pivot_data.sum().to_frame('Итого').T
    table_data = pd.concat([pivot_data, totals])
    table_data.insert(0, 'Год', table_data.index)

    # Создаем Treeview
    tree = ttk.Treeview(table_container,
                        columns=['Год'] + list(pivot_data.columns),
                        show="headings")

    # Настраиваем столбцы
    tree.column('Год', width=80, anchor="center")
    for col in pivot_data.columns:
        tree.column(col, width=120, anchor="center", stretch=tk.YES)
        tree.heading(col, text=col,
                     command=lambda c=col: sort_column(tree, c, pivot_data.columns))
    tree.heading('Год', text='Год')

    # Заполняем данными
    for idx, row in table_data.iterrows():
        values = [idx] + list(row[pivot_data.columns].round(2))
        tree.insert("", "end", values=values, tags=("total" if idx == "Итого" else ""))

    # Выделяем итоговую строку
    tree.tag_configure("total", background="#f0f0f0", font=('Arial', 9, 'bold'))

    # Словарь для хранения состояния сортировки
    sort_states = {col: False for col in pivot_data.columns}  # False = по убыванию, True = по возрастанию

    def sort_column(treeview, column, all_columns):
        """Функция сортировки при клике на заголовок"""
        nonlocal sort_states

        # Получаем все данные из таблицы
        items = [(treeview.set(child, column), child) for child in treeview.get_children('')]

        # Сортируем по значениям (игнорируем итоговую строку)
        items = [x for x in items if x[1] != 'Итого']

        try:
            # Пробуем преобразовать к числам
            items.sort(key=lambda x: float(x[0]), reverse=sort_states[column])
        except ValueError:
            # Если не числа - сортируем как строки
            items.sort(key=lambda x: x[0], reverse=sort_states[column])

        # Перемещаем элементы в отсортированном порядке
        for index, (_, child) in enumerate(items):
            treeview.move(child, '', index)

        # Добавляем итоговую строку в конец
        total_items = [child for child in treeview.get_children('')
                       if treeview.item(child)['tags'] and 'total' in treeview.item(child)['tags']]
        for child in total_items:
            treeview.move(child, '', 'end')

        # Меняем состояние сортировки
        sort_states[column] = not sort_states[column]

        # Обновляем график согласно сортировке
        update_plot(column, sort_states[column])

    def update_plot(sorted_column, ascending):
        """Обновляем график согласно выбранной сортировке"""
        # Получаем порядок столбцов по выбранной сортировке
        sorted_cols = pivot_data.sum().sort_values(ascending=ascending).index.tolist()

        # Перестраиваем график
        ax.clear()
        pivot_data[sorted_cols].plot(kind='bar', ax=ax, width=0.8, color=[color_dict[col] for col in sorted_cols])

        ax.set_title(f"Статистика продаж (сортировка: {sorted_column} {'↑' if ascending else '↓'})",
                     fontsize=14, pad=20)
        ax.yaxis.set_major_formatter(FuncFormatter(millions))
        ax.grid(True, linestyle='--', alpha=0.7)
        ax.legend(title='Продукты', bbox_to_anchor=(1.05, 1), loc='upper left')

        canvas.draw()

    # Прокрутка
    y_scroll = ttk.Scrollbar(table_container, orient="vertical", command=tree.yview)
    x_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

    tree.grid(row=0, column=0, sticky="nsew")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")

    table_container.grid_rowconfigure(0, weight=1)
    table_container.grid_columnconfigure(0, weight=1)

    Button(table_window, text="Закрыть [Esc]", command=table_window.destroy,
           font=("Arial", 10), bg="#f44336", fg="white").pack(pady=10)

    try:
        plt.close(fig)
        table_window.mainloop()
    except Exception as e:
        print(f"Ошибка: {e}")
    plt.close(fig)
#---------------------------------

# Функция для отображения статистики продаж по "InviceType" за последние 5 лет
def show_invoice_type_sales():
    from tkinter import Canvas, Frame, Scrollbar, StringVar

    data['SalesDate'] = pd.to_datetime(data['SalesDate'])
    last_5_years = data[data['SalesDate'] >= (pd.Timestamp.now() - pd.DateOffset(years=5))].copy()
    last_5_years['Year'] = last_5_years['SalesDate'].dt.year
    invoice_sales = last_5_years.groupby(['Year', 'InvoiceType'])['Total'].sum().reset_index()

    base_x, base_y = 700, 300
    offset = 120
    invoice_window = Toplevel(root)
    invoice_window.title("Выбор типов счетов для отображения")
    invoice_window.geometry(f"400x450+{base_x + offset}+{base_y + offset}")
    invoice_window.configure(bg="#dff5e1")

    header = Label(invoice_window, text="Выберите типы счета", bg="#4caf50", fg="white", font=("Arial", 14, "bold"))
    header.pack(fill='x')

    frame_checkbox_area = Frame(invoice_window, bg="#dff5e1")
    frame_checkbox_area.pack(fill="both", expand=True, padx=10, pady=10)

    canvas = Canvas(frame_checkbox_area, bg="#dff5e1", highlightthickness=0)
    scrollbar = Scrollbar(frame_checkbox_area, orient="vertical", command=canvas.yview)
    scroll_frame = Frame(canvas, bg="#dff5e1")

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    invoice_vars = {}
    checkbox_widgets = {}
    for invoice_type in invoice_sales['InvoiceType'].unique():
        var = IntVar(value=0)
        invoice_vars[invoice_type] = var
        cb = Checkbutton(scroll_frame, text=invoice_type, variable=var, bg="#dff5e1", anchor='w')
        cb.pack(anchor='w')
        checkbox_widgets[invoice_type] = cb

    frame_buttons = Frame(invoice_window, bg="#dff5e1")
    frame_buttons.pack(pady=5)

    toggle_btn = Button(frame_buttons, text="Выбрать все", command=lambda: toggle_checkboxes(invoice_vars, toggle_btn), **button_options)
    toggle_btn.pack(pady=5)

    Button(frame_buttons, text="Построить график", command=lambda: plot_selected_invoices(invoice_vars, invoice_sales), **button_options).pack(pady=5)


def plot_selected_invoices(invoice_vars, invoice_sales):
    selected_invoices = [invoice_type for invoice_type, var in invoice_vars.items() if var.get() == 1]
    if not selected_invoices:
        return

    filtered_data = invoice_sales[invoice_sales['InvoiceType'].isin(selected_invoices)]
    pivot_data = filtered_data.pivot(index='Year', columns='InvoiceType', values='Total').fillna(0)

    chart_window = Toplevel(root)
    chart_window.title("График по типам счетов")
    chart_window.geometry("1300x800+300+100")
    chart_window.configure(bg="#ffffff")
    chart_window.bind("<Escape>", lambda event: chart_window.destroy())

    fig, ax = plt.subplots(figsize=(12, 5))
    pivot_data.plot(kind='bar', ax=ax)

    ax.set_title("Статистика продаж по 'InvoiceType' за последние 5 лет", fontsize=14, color='#1e1e1e')
    ax.set_xlabel("Год", fontsize=12, color='#1e1e1e')
    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(title='Типы счетов')
    ax.tick_params(axis='x', rotation=0, labelcolor='#1e1e1e')
    ax.tick_params(axis='y', labelcolor='#1e1e1e')
    plt.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=chart_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="both", expand=True)

    # --- Таблица под графиком ---
    table_frame = ttk.Frame(chart_window)
    table_frame.pack(fill=tk.BOTH, expand=False, pady=10)

    # Добавляем столбец с годами и итоговую строку
    totals = pivot_data.sum().to_frame('Итого').T
    table_data = pd.concat([pivot_data, totals])
    table_data.insert(0, 'Год', table_data.index)

    # Создаем Treeview
    tree = ttk.Treeview(table_frame,
                        columns=['Год'] + list(pivot_data.columns),
                        show="headings")

    tree.column('Год', width=80, anchor="center")
    for col in pivot_data.columns:
        tree.column(col, width=120, anchor="center")
        tree.heading(col, text=col, command=lambda c=col: sort_column(c))
    tree.heading('Год', text='Год', command=lambda: sort_column("Год"))

    # Словарь для хранения состояния сортировки
    sort_states = {col: False for col in pivot_data.columns}
    sort_states['Год'] = False

    def insert_table_data():
        # Удаляем старые данные
        for i in tree.get_children():
            tree.delete(i)

        for idx, row in table_data.iterrows():
            values = [idx] + list(row[pivot_data.columns].round(2))
            tag = "total" if idx == "Итого" else ""
            tree.insert("", "end", values=values, tags=(tag,))

        # Обновляем стиль итоговой строки
        tree.tag_configure("total", background="#f0f0f0", font=('Arial', 9, 'bold'))

    def sort_column(col):
        nonlocal table_data
        ascending = sort_states[col]

        if col == "Год":
            sorted_part = table_data.loc[table_data.index != "Итого"].copy()
            sorted_part['Год'] = sorted_part['Год'].astype(str)
            sorted_part = sorted_part.sort_values(by='Год', ascending=ascending)
        else:
            sorted_part = table_data.loc[table_data.index != "Итого"].copy()
            sorted_part = sorted_part.sort_values(by=col, ascending=ascending)

        total_row = table_data.loc[table_data.index == "Итого"]
        table_data_sorted = pd.concat([sorted_part, total_row])
        table_data = table_data_sorted
        insert_table_data()
        sort_states[col] = not ascending

    insert_table_data()
    tree.pack(fill="x")

    Button(chart_window, text="Закрыть [Esc]", command=chart_window.destroy,
           font=("Arial", 10), bg="#f44336", fg="white").pack(pady=5)

    plt.close(fig)

#---------------------------------

# Функция для отображения статистики продаж по "SWType" за последние 5 лет
def show_swtype_sales():
    data['SalesDate'] = pd.to_datetime(data['SalesDate'])
    last_5_years = data[data['SalesDate'] >= (pd.Timestamp.now() - pd.DateOffset(years=5))].copy()
    last_5_years['Year'] = last_5_years['SalesDate'].dt.year

    global swtype_sales
    swtype_sales = last_5_years.groupby(['Year', 'SWType'])['Total'].sum().reset_index()

    swtype_window = Toplevel(root)
    swtype_window.title("Выбор типов ПО для отображения")
    swtype_window.geometry("400x450+750+350")
    swtype_window.configure(bg="#dff5e1")

    header = Label(swtype_window, text="Выберите типы ПО", bg="#4caf50", fg="white", font=("Arial", 14, "bold"))
    header.pack(fill='x')

    check_frame = Frame(swtype_window, bg="#dff5e1")
    check_frame.pack(padx=10, pady=10, fill='both', expand=True)

    canvas = Canvas(check_frame, bg="#dff5e1")
    scrollbar = Scrollbar(check_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas, bg="#dff5e1")

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    global swtype_vars
    swtype_vars = {}

    for swtype in swtype_sales['SWType'].unique():
        var = IntVar(value=0)
        swtype_vars[swtype] = var
        chk = Checkbutton(scrollable_frame, text=swtype, variable=var, bg="#dff5e1")
        chk.pack(anchor='w')

    toggle_button = Button(swtype_window, text="Выбрать все", command=lambda: toggle_checkboxes(swtype_vars, toggle_button), **button_options)
    toggle_button.pack(pady=5)

    Button(swtype_window, text="Построить график", command=plot_selected_swtypes, **button_options).pack(pady=10)


def plot_selected_swtypes():
    selected_types = [swtype for swtype, var in swtype_vars.items() if var.get() == 1]
    if not selected_types:
        messagebox.showwarning("Предупреждение", "Выберите хотя бы один тип ПО.")
        return

    filtered_data = swtype_sales[swtype_sales['SWType'].isin(selected_types)]
    pivot_data = filtered_data.pivot(index='Year', columns='SWType', values='Total').fillna(0)

    chart_window = Toplevel(root)
    chart_window.title("График по типам ПО")
    chart_window.geometry("1300x800+300+100")
    chart_window.bind("<Escape>", lambda event: chart_window.destroy())

    fig, ax = plt.subplots(figsize=(14, 5))
    pivot_data.plot(kind='bar', ax=ax)

    ax.set_title("Статистика продаж по 'SWType' за последние 5 лет", fontsize=14)
    ax.set_xlabel("Год", fontsize=12)
    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    ax.grid(True, linestyle='--', alpha=0.7)
    ax.legend(title='Типы ПО')
    ax.tick_params(axis='x', rotation=0)

    fig.tight_layout(rect=[0, 0, 1, 0.95])

    canvas = FigureCanvasTkAgg(fig, master=chart_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)

    # --- Таблица под графиком ---
    table_frame = ttk.Frame(chart_window)
    table_frame.pack(fill=tk.BOTH, expand=False, pady=10)

    totals = pivot_data.sum().to_frame('Итого').T
    table_data = pd.concat([pivot_data, totals])
    table_data.insert(0, 'Год', table_data.index)

    tree = ttk.Treeview(table_frame, columns=['Год'] + list(pivot_data.columns), show='headings')

    tree.column('Год', width=80, anchor='center')
    for col in pivot_data.columns:
        tree.column(col, width=120, anchor='center')
        tree.heading(col, text=col, command=lambda c=col: sort_column(c))
    tree.heading('Год', text='Год', command=lambda: sort_column("Год"))

    sort_states = {col: False for col in pivot_data.columns}
    sort_states['Год'] = False

    def insert_table_data():
        for i in tree.get_children():
            tree.delete(i)

        for idx, row in table_data.iterrows():
            values = [idx] + list(row[pivot_data.columns].round(2))
            tag = "total" if idx == "Итого" else ""
            tree.insert("", "end", values=values, tags=(tag,))

        tree.tag_configure("total", background="#f0f0f0", font=('Arial', 9, 'bold'))

    def sort_column(col):
        nonlocal table_data
        ascending = sort_states[col]

        if col == "Год":
            sorted_part = table_data.loc[table_data.index != "Итого"].copy()
            sorted_part['Год'] = sorted_part['Год'].astype(str)
            sorted_part = sorted_part.sort_values(by='Год', ascending=ascending)
        else:
            sorted_part = table_data.loc[table_data.index != "Итого"].copy()
            sorted_part = sorted_part.sort_values(by=col, ascending=ascending)

        total_row = table_data.loc[table_data.index == "Итого"]
        table_data = pd.concat([sorted_part, total_row])
        insert_table_data()
        sort_states[col] = not ascending

    insert_table_data()
    tree.pack(fill="x")

    Button(chart_window, text="Закрыть [Esc]", command=chart_window.destroy,
           font=("Arial", 10), bg="#f44336", fg="white").pack(pady=5)

    plt.close(fig)

#---------------------------------

# Преобразование столбца "Dealer" в строковый тип
data["Dealer"] = data["Dealer"].astype(str)

# Получаем уникальные годы, города и клиенты
unique_years = sorted(data['Year'].dropna().unique())
unique_cities = sorted(data['City'].dropna().unique())
unique_clients = sorted(data['Client'].dropna().unique())
unique_dealer = sorted([d for d in data['Dealer'].unique() if d != 'nan'])

# Функция для создания чекбоксов с прокруткой
def create_scrollable_checkboxes(frame, items, vars_dict, label_text):
    scrollable_frame = Frame(frame)
    scrollable_frame.pack(side='top', padx=20, pady=5)

    canvas = Canvas(scrollable_frame, width=200, height=200)
    scrollbar = Scrollbar(scrollable_frame, orient=VERTICAL, command=canvas.yview)
    scrollbar.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    canvas.configure(yscrollcommand=scrollbar.set)

    check_frame = Frame(canvas)
    canvas.create_window((0, 0), window=check_frame, anchor='nw')

    Label(check_frame, text=label_text).pack()
    for item in items:
        var = IntVar()
        vars_dict[item] = var
        Checkbutton(check_frame, text=str(item), variable=var).pack(anchor='w')

    check_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

    return vars_dict

#---------------------------------
#ФУНКЦИЯ ДЛЯ ФОРМИРОВАНИЯ ТАБЛИЦ
def plot_table(year_vars, other_vars, group_by):
    selected_years = [year for year, var in year_vars.items() if var.get() == 1]
    selected_others = [other for other, var in other_vars.items() if var.get() == 1]

    if not selected_years or not selected_others:
        messagebox.showwarning("Предупреждение", "Выберите хотя бы один год и один элемент.")
        return

    filtered_data = data[data['Year'].isin(selected_years) & data[group_by].isin(selected_others)]

    if 'Total' not in filtered_data.columns:
        messagebox.showerror("Ошибка", "Столбец 'Total' не найден в данных.")
        return

    filtered_data['Total'] = pd.to_numeric(filtered_data['Total'], errors='coerce')
    filtered_data = filtered_data.dropna(subset=['Total'])
    filtered_data['Year'] = filtered_data['Year'].astype(int)

    pivot_table = filtered_data.pivot_table(
        index=group_by,
        columns='Year',
        values='Total',
        aggfunc='sum',
        fill_value=0
    )

    pivot_table['Итого'] = pivot_table.sum(axis=1)

    if group_by == 'Client' and 'City' in filtered_data.columns:
        client_city_map = filtered_data.drop_duplicates(subset=['Client'])[['Client', 'City']].set_index('Client')['City'].to_dict()
        pivot_table.insert(1, 'Город клиента', pivot_table.index.map(client_city_map))

    if group_by == 'Dealer' and 'City' in filtered_data.columns:
        dealer_city_map = filtered_data.drop_duplicates(subset=['Dealer'])[['Dealer', 'City']].set_index('Dealer')['City'].to_dict()
        pivot_table.insert(1, 'Город дилера', pivot_table.index.map(dealer_city_map))

    columns = [group_by]
    if 'Город клиента' in pivot_table.columns:
        columns.append('Город клиента')
    if 'Город дилера' in pivot_table.columns:
        columns.append('Город дилера')
    columns += [str(year) for year in pivot_table.columns if str(year).isdigit()]
    columns.append('Итого')

    table_window = Toplevel(root)
    table_window.title("Сводная таблица")
    table_window.geometry("1000x700")

    main_frame = Frame(table_window)
    main_frame.pack(expand=True, fill='both', padx=10, pady=10)

    search_frame = Frame(main_frame)
    search_frame.pack(fill='x', pady=(0, 10))
    Label(search_frame, text="Поиск по ключевым словам:").pack(side='left')
    search_var = StringVar()
    search_entry = Entry(search_frame, textvariable=search_var, font=('Arial', 10))
    search_entry.pack(side='left', fill='x', expand=True, padx=5)
    Button(search_frame, text="Очистить", command=lambda: search_var.set("")).pack(side='left')

    tree_frame = Frame(main_frame)
    tree_frame.pack(expand=True, fill='both')

    tree_scroll_y = Scrollbar(tree_frame, orient="vertical")
    tree_scroll_x = Scrollbar(tree_frame, orient="horizontal")

    style = ttk.Style()
    style.configure("Treeview.Heading", font=("Arial", 10, "bold"), background="#d9d9d9", foreground="black")

    tree = ttk.Treeview(
        tree_frame,
        columns=columns,
        yscrollcommand=tree_scroll_y.set,
        xscrollcommand=tree_scroll_x.set,
        selectmode='browse'
    )

    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)

    tree.column("#0", width=0, stretch=tk.NO)
    tree.column(group_by, anchor="w", width=200)
    if 'Город клиента' in columns:
        tree.column('Город клиента', anchor="w", width=150)
    if 'Город дилера' in columns:
        tree.column('Город дилера', anchor="w", width=150)

    for year in columns:
        if year not in [group_by, 'Город клиента', 'Город дилера']:
            tree.column(year, anchor="e", width=100 if year != 'Итого' else 120)

    sort_states = {col: False for col in columns}

    def sort_treeview(tv, col, is_numeric):
        items = [(tv.set(k, col), k) for k in tv.get_children('')]
        if is_numeric:
            items.sort(key=lambda t: float(t[0].replace(',', '')) if t[0] else 0, reverse=sort_states[col])
        else:
            items.sort(key=lambda t: t[0], reverse=sort_states[col])

        for index, (_, k) in enumerate(items):
            tv.move(k, '', index)

        sort_states[col] = not sort_states[col]

        direction = " ↓" if sort_states[col] else " ↑"
        for c in tv["columns"]:
            base_text = c.capitalize() if c == group_by else str(c)
            tv.heading(c, text=base_text, anchor="center")
        base_text = col.capitalize() if col == group_by else str(col)
        tv.heading(col, text=base_text + direction, anchor="center")

    tree.heading(group_by, text=group_by.capitalize(), anchor="w", command=lambda: sort_treeview(tree, group_by, False))
    if 'Город клиента' in columns:
        tree.heading('Город клиента', text='Город клиента', anchor="w", command=lambda: sort_treeview(tree, 'Город клиента', False))
    if 'Город дилера' in columns:
        tree.heading('Город дилера', text='Город дилера', anchor="w", command=lambda: sort_treeview(tree, 'Город дилера', False))

    for year in columns:
        if year not in [group_by, 'Город клиента', 'Город дилера']:
            tree.heading(year, text=str(year), anchor="center", command=lambda y=year: sort_treeview(tree, y, True))

    tree.tag_configure('evenrow', background='#f2f2f2', font=('Arial', 10))
    tree.tag_configure('oddrow', background='#ffffff', font=('Arial', 10))
    tree.tag_configure('bold_total', font=('Arial', 10, 'bold'))

    for i, item in enumerate(pivot_table.index):
        row = pivot_table.loc[item]
        base_data = [item]
        if 'Город клиента' in columns:
            base_data.append(row['Город клиента'])
        if 'Город дилера' in columns:
            base_data.append(row['Город дилера'])

        year_values = [f"{row[year]:,.0f}" for year in pivot_table.columns if year not in ['Итого', 'Город клиента', 'Город дилера']]
        total_value = f"{row['Итого']:,.0f}"
        row_data = base_data + year_values + [total_value]

        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        tree.insert("", "end", values=row_data, tags=(tag, 'bold_total'))

    tree.grid(row=0, column=0, sticky="nsew")
    tree_scroll_y.grid(row=0, column=1, sticky="ns")
    tree_scroll_x.grid(row=1, column=0, sticky="ew")

    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)

    original_items = list(tree.get_children())

    def filter_table(*args):
        search_terms = search_var.get().lower().strip().split()
        for child in original_items:
            tree.reattach(child, '', 'end')

        if search_terms:
            for child in original_items:
                values = [tree.set(child, col).lower() for col in columns]
                combined_text = " ".join(values)
                if not all(term in combined_text for term in search_terms):
                    tree.detach(child)

        for col in columns:
            base_text = col.capitalize() if col == group_by else str(col)
            tree.heading(col, text=base_text, anchor="center")

    search_var.trace_add("write", filter_table)

#---------------------------------

# Функция для отображения отчета по территории продаж
def show_sales_by_region():
    region_window = Toplevel(root)
    region_window.title("Отчет по территории продаж")
    region_window.geometry("420x600+750+300")
    region_window.configure(bg="#dff5e1")

    Label(region_window, text="Выберите годы и города", bg="#4caf50", fg="white", font=("Arial", 14, "bold")).pack(fill='x')

    selection_frame = Frame(region_window, bg="#dff5e1")
    selection_frame.pack(padx=10, pady=10, fill='both', expand=True)

    year_vars = {}
    city_vars = {}

    # БЛОК С ГОДАМИ
    Label(selection_frame, text="Годы", bg="#dff5e1", font=("Arial", 12, "bold")).pack(anchor='w')

    year_frame = Frame(selection_frame, bg="#dff5e1")
    year_frame.pack(fill='x', pady=(0, 10))

    year_canvas = Canvas(year_frame, height=120, bg="#dff5e1", highlightthickness=0)
    year_scrollbar = Scrollbar(year_frame, orient="vertical", command=year_canvas.yview)
    year_inner_frame = Frame(year_canvas, bg="#dff5e1")
    year_inner_frame.bind("<Configure>", lambda e: year_canvas.configure(scrollregion=year_canvas.bbox("all")))
    year_canvas.create_window((0, 0), window=year_inner_frame, anchor="nw")
    year_canvas.configure(yscrollcommand=year_scrollbar.set)

    year_canvas.pack(side="left", fill="both", expand=True)
    year_scrollbar.pack(side="right", fill="y")

    for year in unique_years:
        year_int = int(year)
        var = IntVar(value=0)
        year_vars[year_int] = var
        Checkbutton(year_inner_frame, text=str(year_int), variable=var, bg="#dff5e1").pack(anchor='w')

    toggle_years_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(year_vars, toggle_years_button), **button_options)
    toggle_years_button.pack(pady=(0, 15))

    # БЛОК С ГОРОДАМИ
    Label(selection_frame, text="Города", bg="#dff5e1", font=("Arial", 12, "bold")).pack(anchor='w')

    # Рамка и канвас для списка городов
    city_frame = Frame(selection_frame, bg="#dff5e1")
    city_frame.pack(fill='x')

    city_canvas = Canvas(city_frame, height=150, bg="#dff5e1", highlightthickness=0)
    city_scrollbar = Scrollbar(city_frame, orient="vertical", command=city_canvas.yview)
    city_inner_frame = Frame(city_canvas, bg="#dff5e1")
    city_inner_frame.bind("<Configure>", lambda e: city_canvas.configure(scrollregion=city_canvas.bbox("all")))
    city_canvas.create_window((0, 0), window=city_inner_frame, anchor="nw")
    city_canvas.configure(yscrollcommand=city_scrollbar.set)

    city_canvas.pack(side="left", fill="both", expand=True)
    city_scrollbar.pack(side="right", fill="y")

    # Поисковая строка
    search_frame = Frame(selection_frame, bg="#dff5e1")
    search_frame.pack(fill='x', pady=(5, 5))

    search_var = StringVar()

    Entry(search_frame, textvariable=search_var, font=("Arial", 10)).pack(side='left', fill='x', expand=True,
                                                                          padx=(0, 5))

    Button(
        search_frame,
        text="Стереть",
        command=lambda: search_var.set(""),
        width=8,
        bg="#dcdcdc",
        fg="black",
        activebackground="#c0c0c0",
        activeforeground="black",
        relief="flat",
        font=("Arial", 9)
    ).pack(side='left')

    # Глобальный список чекбоксов городов
    def update_city_checkboxes(*args):
        for widget in city_inner_frame.winfo_children():
            widget.destroy()

        # Получаем города, удовлетворяющие поиску
        filtered = [c for c in unique_cities if search_var.get().lower() in c.lower()]

        for city in filtered:
            if city not in city_vars:
                city_vars[city] = IntVar(value=0)
            Checkbutton(city_inner_frame, text=city, variable=city_vars[city], bg="#dff5e1").pack(anchor='w')

    search_var.trace_add("write", update_city_checkboxes)
    update_city_checkboxes()

    toggle_cities_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(city_vars, toggle_cities_button), **button_options)
    toggle_cities_button.pack(pady=(5, 10))

    # Кнопка подготовки таблицы
    Button(region_window, text="Подготовить таблицу", command=lambda: plot_table(year_vars, city_vars, 'City'), **button_options).pack(pady=15)

#---------------------------------

# Функция для отображения Отчета по клиентам
def show_sales_by_client():
    client_window = Toplevel(root)
    client_window.title("Отчет по клиентам")
    client_window.geometry("420x600+750+300")
    client_window.configure(bg="#dff5e1")

    Label(client_window, text="Выберите годы и клиентов", bg="#4caf50", fg="white", font=("Arial", 14, "bold")).pack(fill='x')

    selection_frame = Frame(client_window, bg="#dff5e1")
    selection_frame.pack(padx=10, pady=10, fill='both', expand=True)

    year_vars = {}
    client_vars = {}

    # ГОДЫ
    Label(selection_frame, text="Годы", bg="#dff5e1", font=("Arial", 12, "bold")).pack(anchor='w')

    year_frame = Frame(selection_frame, bg="#dff5e1")
    year_frame.pack(fill='x', pady=(0, 10))

    year_canvas = Canvas(year_frame, height=120, bg="#dff5e1", highlightthickness=0)
    year_scrollbar = Scrollbar(year_frame, orient="vertical", command=year_canvas.yview)
    year_inner_frame = Frame(year_canvas, bg="#dff5e1")
    year_inner_frame.bind("<Configure>", lambda e: year_canvas.configure(scrollregion=year_canvas.bbox("all")))
    year_canvas.create_window((0, 0), window=year_inner_frame, anchor="nw")
    year_canvas.configure(yscrollcommand=year_scrollbar.set)

    year_canvas.pack(side="left", fill="both", expand=True)
    year_scrollbar.pack(side="right", fill="y")

    for year in unique_years:
        year_int = int(year)
        var = IntVar(value=0)
        year_vars[year_int] = var
        Checkbutton(year_inner_frame, text=str(year_int), variable=var, bg="#dff5e1").pack(anchor='w')

    toggle_years_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(year_vars, toggle_years_button), **button_options)
    toggle_years_button.pack(pady=(0, 15))

    # КЛИЕНТЫ
    Label(selection_frame, text="Клиенты", bg="#dff5e1", font=("Arial", 12, "bold")).pack(anchor='w')

    client_frame = Frame(selection_frame, bg="#dff5e1")
    client_frame.pack(fill='x')

    client_canvas = Canvas(client_frame, height=150, bg="#dff5e1", highlightthickness=0)
    client_scrollbar = Scrollbar(client_frame, orient="vertical", command=client_canvas.yview)
    client_inner_frame = Frame(client_canvas, bg="#dff5e1")
    client_inner_frame.bind("<Configure>", lambda e: client_canvas.configure(scrollregion=client_canvas.bbox("all")))
    client_canvas.create_window((0, 0), window=client_inner_frame, anchor="nw")
    client_canvas.configure(yscrollcommand=client_scrollbar.set)

    client_canvas.pack(side="left", fill="both", expand=True)
    client_scrollbar.pack(side="right", fill="y")

    # Поисковая строка под списком
    search_frame = Frame(selection_frame, bg="#dff5e1")
    search_frame.pack(fill='x', pady=(5, 5))

    search_var = StringVar()
    Entry(search_frame, textvariable=search_var, font=("Arial", 10)).pack(side='left', fill='x', expand=True, padx=(0, 5))
    Button(
        search_frame,
        text="Стереть",
        command=lambda: search_var.set(""),
        width=8,
        bg="#dcdcdc",
        fg="black",
        activebackground="#c0c0c0",
        activeforeground="black",
        relief="flat",
        font=("Arial", 9)
    ).pack(side='left')

    def update_client_checkboxes(*args):
        for widget in client_inner_frame.winfo_children():
            widget.destroy()
        filtered = [c for c in unique_clients if search_var.get().lower() in c.lower()]
        for client in filtered:
            if client not in client_vars:
                client_vars[client] = IntVar(value=0)
            Checkbutton(client_inner_frame, text=client, variable=client_vars[client], bg="#dff5e1").pack(anchor='w')

    search_var.trace_add("write", update_client_checkboxes)
    update_client_checkboxes()

    toggle_clients_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(client_vars, toggle_clients_button), **button_options)
    toggle_clients_button.pack(pady=(5, 10))

    Button(client_window, text="Подготовить таблицу", command=lambda: plot_table(year_vars, client_vars, 'Client'), **button_options).pack(pady=15)

#---------------------------------

# Функция для отображения Отчета по дилерам
def show_sales_by_dealer():
    dealer_window = Toplevel(root)
    dealer_window.title("Отчет по дилерам")
    dealer_window.geometry("420x600+750+300")
    dealer_window.configure(bg="#dff5e1")

    Label(dealer_window, text="Выберите годы и дилеров", bg="#4caf50", fg="white", font=("Arial", 14, "bold")).pack(fill='x')

    selection_frame = Frame(dealer_window, bg="#dff5e1")
    selection_frame.pack(padx=10, pady=10, fill='both', expand=True)

    year_vars = {}
    dealer_vars = {}

    # ГОДЫ
    year_frame = Frame(selection_frame, bg="#dff5e1")
    year_frame.pack(fill='x', pady=(0, 10))

    year_canvas = Canvas(year_frame, height=120, bg="#dff5e1", highlightthickness=0)
    year_scrollbar = Scrollbar(year_frame, orient="vertical", command=year_canvas.yview)
    year_inner_frame = Frame(year_canvas, bg="#dff5e1")
    year_inner_frame.bind("<Configure>", lambda e: year_canvas.configure(scrollregion=year_canvas.bbox("all")))
    year_canvas.create_window((0, 0), window=year_inner_frame, anchor="nw")
    year_canvas.configure(yscrollcommand=year_scrollbar.set)

    year_canvas.pack(side="left", fill="both", expand=True)
    year_scrollbar.pack(side="right", fill="y")

    for year in unique_years:
        year_int = int(year)
        var = IntVar(value=0)
        year_vars[year_int] = var
        Checkbutton(year_inner_frame, text=str(year_int), variable=var, bg="#dff5e1").pack(anchor='w')

    toggle_years_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(year_vars, toggle_years_button), **button_options)
    toggle_years_button.pack(pady=(0, 15))

    # ДИЛЕРЫ
    Label(selection_frame, text="Дилеры", bg="#dff5e1", font=("Arial", 12, "bold")).pack(anchor='w')

    dealer_frame = Frame(selection_frame, bg="#dff5e1")
    dealer_frame.pack(fill='x')

    dealer_canvas = Canvas(dealer_frame, height=150, bg="#dff5e1", highlightthickness=0)
    dealer_scrollbar = Scrollbar(dealer_frame, orient="vertical", command=dealer_canvas.yview)
    dealer_inner_frame = Frame(dealer_canvas, bg="#dff5e1")
    dealer_inner_frame.bind("<Configure>", lambda e: dealer_canvas.configure(scrollregion=dealer_canvas.bbox("all")))
    dealer_canvas.create_window((0, 0), window=dealer_inner_frame, anchor="nw")
    dealer_canvas.configure(yscrollcommand=dealer_scrollbar.set)

    dealer_canvas.pack(side="left", fill="both", expand=True)
    dealer_scrollbar.pack(side="right", fill="y")

    # Поисковая строка под списком
    search_frame = Frame(selection_frame, bg="#dff5e1")
    search_frame.pack(fill='x', pady=(5, 5))

    search_var = StringVar()
    Entry(search_frame, textvariable=search_var, font=("Arial", 10)).pack(side='left', fill='x', expand=True, padx=(0, 5))
    Button(
        search_frame,
        text="Стереть",
        command=lambda: search_var.set(""),
        width=8,
        bg="#dcdcdc",
        fg="black",
        activebackground="#c0c0c0",
        activeforeground="black",
        relief="flat",
        font=("Arial", 9)
    ).pack(side='left')

    def update_dealer_checkboxes(*args):
        for widget in dealer_inner_frame.winfo_children():
            widget.destroy()
        filtered = [d for d in unique_dealer if search_var.get().lower() in d.lower()]
        for dealer in filtered:
            if dealer not in dealer_vars:
                dealer_vars[dealer] = IntVar(value=0)
            Checkbutton(dealer_inner_frame, text=dealer, variable=dealer_vars[dealer], bg="#dff5e1").pack(anchor='w')

    search_var.trace_add("write", update_dealer_checkboxes)
    update_dealer_checkboxes()

    toggle_dealers_button = Button(selection_frame, text="Выбрать все", command=lambda: toggle_checkboxes(dealer_vars, toggle_dealers_button), **button_options)
    toggle_dealers_button.pack(pady=(5, 10))

    Button(dealer_window, text="Подготовить таблицу", command=lambda: plot_table(year_vars, dealer_vars, 'Dealer'), **button_options).pack(pady=15)

#---------------------------------

# Функция для отображения Отчета по веткам ПО
def generate_branch_report():
    # Группировка данных по годам и веткам
    report_data_gen = data.groupby(['Year', 'Gen'])['Total'].sum().unstack(fill_value=0)

    # --- Создание окна ---
    table_window = Toplevel()
    table_window.title("Отчёт по поколениям")
    table_window.geometry("1000x800")
    table_window.bind("<Escape>", lambda event: table_window.destroy())

    # --- Построение графика ---
    fig, ax = plt.subplots(figsize=(8, 5))
    report_data_gen.plot(kind='bar', stacked=True, ax=ax)

    plt.title('Сумма покупок по поколениям')

    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    ax.set_xticklabels([str(int(float(label.get_text()))) for label in ax.get_xticklabels()])

    plt.legend(title='Поколение')
    plt.tight_layout()

    # --- Вставка графика в tkinter ---
    canvas = FigureCanvasTkAgg(fig, master=table_window)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # --- Сводная таблица ---
    pivot_data = report_data_gen.copy()
    table_container = ttk.Frame(table_window)
    table_container.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    # Добавляем итоговую строку
    totals = pivot_data.sum().to_frame('Итого').T
    table_data = pd.concat([pivot_data, totals])
    table_data.insert(0, 'Год', table_data.index)


    # Создаем Treeview
    columns = ['Год'] + list(pivot_data.columns)
    tree = ttk.Treeview(table_container, columns=columns, show="headings")
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Тэг для жирной итоговой строки
    tree.tag_configure('total', font=('TkDefaultFont', 10, 'bold'))

    # Скролл
    scroll_y = ttk.Scrollbar(table_container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll_y.set)
    scroll_y.pack(side=tk.RIGHT, fill='y')

    # Функция сортировки
    def sort_column(tree, col, columns, reverse=False):
        children = tree.get_children()
        data = [(tree.set(k, col), k) for k in children if tree.set(k, 'Год') != 'Итого']

        try:
            data.sort(key=lambda t: float(t[0].replace(',', '')), reverse=reverse)
        except ValueError:
            data.sort(key=lambda t: t[0], reverse=reverse)

        for index, (val, k) in enumerate(data):
            tree.move(k, '', index)

        # Перемещаем строку "Итого" вниз
        for k in children:
            if tree.set(k, 'Год') == 'Итого':
                tree.move(k, '', 'end')

        # Обновляем команду на обратную сортировку
        tree.heading(col, command=lambda: sort_column(tree, col, columns, not reverse))

    # Заголовки и колонки
    for col in columns:
        width = 100 if col == 'Год' else 120
        tree.heading(col, text=col, command=lambda c=col: sort_column(tree, c, columns))
        tree.column(col, anchor="center", width=width)

    # Заполнение строк
    for i, row in table_data.iterrows():
        formatted_row = []
        for col in columns:
            val = row[col]
            if col == 'Год':
                val = str(int(val)) if isinstance(val, (int, float)) and not isinstance(val, bool) else str(val)
            elif isinstance(val, (int, float)):
                val = f"{val:,.0f}"
            formatted_row.append(val)

        tag = 'total' if row['Год'] == 'Итого' else ''
        tree.insert('', 'end', values=formatted_row, tags=(tag,))
    exit_button = Button(table_window, text="Закрыть [Esc]", command=table_window.destroy,
                         font=("Arial", 10), bg="#f44336", fg="white")
    exit_button.pack(pady=10)
    plt.close(fig)
#---------------------------------

# Функция для генерации отчета по поколениям ПО
def generate_pricelist_report():
    # --- Фильтрация нужных прайс-листов и сортировка по заданному порядку ---
    filtered_data = data[data['Pricelist'].isin(['Basic', 'Standard', 'Professional'])].copy()
    order = ['Basic', 'Standard', 'Professional']
    filtered_data['Pricelist'] = pd.Categorical(filtered_data['Pricelist'], categories=order, ordered=True)

    # --- Группировка данных ---
    report_data_pricelist = filtered_data.groupby(['Year', 'Pricelist'])['Total'].sum().unstack(fill_value=0)
    report_data_pricelist = report_data_pricelist[order]  # порядок столбцов

    # --- Создание окна ---
    table_window = Toplevel()
    table_window.title("Отчёт по Веткам")
    table_window.geometry("1000x800")
    table_window.bind("<Escape>", lambda event: table_window.destroy())

    # --- Построение графика ---
    fig, ax = plt.subplots(figsize=(8, 5))
    report_data_pricelist.plot(kind='bar', stacked=True, ax=ax)
    plt.title('Сумма покупок по Веткам')

    ax.yaxis.set_major_formatter(FuncFormatter(millions))

    # Убираем .0 у годов
    ax.set_xticklabels([str(int(float(label.get_text()))) for label in ax.get_xticklabels()])

    plt.legend(title='Ветка')
    plt.tight_layout()

    # --- Вставка графика в tkinter ---
    canvas = FigureCanvasTkAgg(fig, master=table_window)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # --- Сводная таблица ---
    pivot_data = report_data_pricelist.copy()
    table_container = ttk.Frame(table_window)
    table_container.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    # Добавляем итоговую строку
    totals = pivot_data.sum().to_frame('Итого').T
    totals.index = ['Итого']
    table_data = pd.concat([pivot_data, totals])
    table_data.insert(0, 'Год', table_data.index)

    # Исправляем года
    table_data['Год'] = table_data['Год'].apply(lambda x: str(int(x)) if isinstance(x, (int, float)) else x)

    # --- Создаем Treeview ---
    columns = ['Год'] + list(pivot_data.columns)
    tree = ttk.Treeview(table_container, columns=columns, show="headings")
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Scroll
    scroll_y = ttk.Scrollbar(table_container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scroll_y.set)
    scroll_y.pack(side=tk.RIGHT, fill='y')

    # Заголовки и колонки
    for col in columns:
        width = 100 if col == 'Год' else 120
        tree.heading(col, text=col, command=lambda c=col: sort_column_with_totals(tree, c, columns))
        tree.column(col, anchor="center", width=width)

    # Тэг для жирной итоговой строки
    tree.tag_configure('total', font=('TkDefaultFont', 10, 'bold'))

    # Заполнение строк
    for i, row in table_data.iterrows():
        values = [row[col] if pd.notnull(row[col]) else '' for col in columns]
        values = [f"{val:,.0f}" if isinstance(val, (int, float)) else val for val in values]
        if row['Год'] == 'Итого':
            tree.insert('', 'end', values=values, tags=('total',))
        else:
            tree.insert('', 'end', values=values)
    exit_button = Button(table_window, text="Закрыть [Esc]", command=table_window.destroy,
                         font=("Arial", 10), bg="#f44336", fg="white")
    exit_button.pack(pady=10)
    plt.close(fig)

    def sort_column_with_totals(tree, col, columns, reverse=False):
        data = [(tree.set(child, col), tree.item(child)["values"], tree.item(child).get("tags", ()))
                for child in tree.get_children()]

        data_normal = [item for item in data if item[1][0] != 'Итого']
        data_total = [item for item in data if item[1][0] == 'Итого']

        try:
            data_normal.sort(key=lambda x: float(str(x[0]).replace(',', '').replace(' ', '')), reverse=reverse)
        except ValueError:
            data_normal.sort(key=lambda x: x[0], reverse=reverse)

        for item in tree.get_children():
            tree.delete(item)

        # Сначала обычные строки, потом "Итого"
        for _, values, tags in data_normal + data_total:
            tree.insert('', 'end', values=values, tags=tags)

        # Переназначение сортировки при следующем клике
        tree.heading(col, command=lambda: sort_column_with_totals(tree, col, columns, not reverse))

# Функция для открытия нового окна с кнопками
def open_report_window():
    report_window = Toplevel(root)
    report_window.title("Выбор отчета")
    report_window.geometry("400x300+750+300")
    report_window.configure(bg="#dff5e1")

    Label(report_window, text="Выберите отчет", bg="#4caf50", fg="white", font=("Arial", 14, "bold")).pack(fill='x')

    # Фрейм для кнопок
    button_frame = Frame(report_window, bg="#dff5e1")
    button_frame.pack(padx=10, pady=20, fill='both', expand=True)

    # Кнопка для статистики по веткам
    Button(button_frame, text="Статистика по поколениям", command=generate_branch_report, **button_options).pack(pady=10)

    # Кнопка для статистики по поколениям
    Button(button_frame, text="Статистика по веткам", command=generate_pricelist_report, **button_options).pack(pady=10)

    report_window.mainloop()

#---------------------------------

# Функция для обработки нажатия кнопки "Общая статистика по продажам"
def show_total_sales():
    # Группируем данные по годам и суммируем значения в столбце "Total"
    total_sales_by_year = data.groupby('Year')['Total'].sum()
    show_plot(total_sales_by_year, "Общая статистика по продажам (год, сумма)")

# Функция для генерации графика по выбранным продуктам
def generate_product_report(selected_products):
    filtered_data = data[data['Product'].isin(selected_products)]
    report_data = filtered_data.groupby('Position')['Total'].sum()

    ax = report_data.plot(kind='bar')
    plt.title('Сумма покупок по позициям')
    plt.xlabel('Позиция')
    plt.ylabel('Сумма покупок (в миллионах)')
    ax.yaxis.set_major_formatter(FuncFormatter(millions))
    plt.tight_layout()


# Функция для загрузки таблицы
def load_table(selected_years):
    if not selected_years:
        messagebox.showwarning("Предупреждение", "Пожалуйста, выберите хотя бы один год.")
        return

    # Фильтрация данных по выбранным годам
    filtered_data = data[data['Year'].isin(selected_years)]

    # Получение уникальных позиций и продуктов
    positions = filtered_data['Position'].unique()
    products = filtered_data['Product'].unique()

    # Создание сводной таблицы
    pivot_table = pd.DataFrame(columns=products, index=positions).fillna(0)

    for position in positions:
        for product in products:
            sales_data = filtered_data[filtered_data['Position'] == position]
            product_sales = sales_data[sales_data['Product'] == product]
            if not product_sales.empty:
                quantity = product_sales.shape[0]
                total_sales = product_sales['Total'].sum()
                pivot_table.at[position, product] = f"{quantity} (Сумма: {total_sales})"

    # Отображение таблицы
    display_table(pivot_table)

# Функция для отображения таблицы
def display_table(pivot_table):
    table_window = Toplevel(root)
    table_window.title("Таблица продаж")

    # Добавляем индекс в отдельный столбец
    pivot_table = pivot_table.copy()
    pivot_table.insert(0, "Позиция", pivot_table.index)

    # Создание таблицы
    tree = ttk.Treeview(table_window, columns=list(pivot_table.columns), show='headings')
    for col in pivot_table.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='center', width=150)

    # Добавление строк в таблицу
    for index, row in pivot_table.iterrows():
        tree.insert("", "end", values=list(row))

    tree.pack(expand=True, fill='both')

#---------------------------------

def open_year_selection_window():
    year_window = Toplevel(root)
    year_window.title("Выбор года и построение графиков")
    year_window.geometry("420x500+750+300")
    year_window.configure(bg="#dff5e1")

    Label(year_window, text="Выберите годы", bg="#4caf50", fg="white", font=("Arial", 14, "bold")).pack(fill='x')

    selection_frame = Frame(year_window, bg="#dff5e1")
    selection_frame.pack(padx=10, pady=10, fill='both', expand=True)

    # === ГОДЫ ===
    year_frame = Frame(selection_frame, bg="#dff5e1")
    year_frame.pack(fill='x', pady=(0, 20))  # Увеличен отступ снизу

    year_canvas = Canvas(year_frame, height=200, bg="#dff5e1", highlightthickness=0)  # Больше высота
    year_scrollbar = Scrollbar(year_frame, orient="vertical", command=year_canvas.yview)
    year_inner_frame = Frame(year_canvas, bg="#dff5e1")
    year_inner_frame.bind("<Configure>", lambda e: year_canvas.configure(scrollregion=year_canvas.bbox("all")))
    year_canvas.create_window((0, 0), window=year_inner_frame, anchor="nw")
    year_canvas.configure(yscrollcommand=year_scrollbar.set)

    year_canvas.pack(side="left", fill="both", expand=True)
    year_scrollbar.pack(side="right", fill="y")

    year_vars = {}
    try:
        unique_years = sorted([int(y) for y in data['Year'].dropna().unique()])
    except KeyError:
        messagebox.showerror("Ошибка", "Столбец 'Year' не найден в данных.")
        return

    for year in unique_years:
        var = IntVar()
        year_vars[year] = var
        Checkbutton(year_inner_frame, text=str(year), variable=var, bg="#dff5e1").pack(anchor='w')

    # Кнопка "Выбрать все" / "Сбросить все"
    def toggle_select_all():
        if select_all_button['text'] == "Выбрать все":
            for var in year_vars.values():
                var.set(1)
            select_all_button['text'] = "Сбросить все"
        else:
            for var in year_vars.values():
                var.set(0)
            select_all_button['text'] = "Выбрать все"

    select_all_button = Button(selection_frame, text="Выбрать все", command=toggle_select_all, **button_options)
    select_all_button.pack(pady=(5, 15))

    # Кнопки для построения графиков (перенесены внутрь selection_frame и подняты)
    Button(selection_frame, text="Построить график Digispot",
           command=lambda: prepare_charts(year_vars, "Digispot"), **button_options).pack(pady=(55, 5))

    Button(selection_frame, text="Построить график Synapse",
           command=lambda: prepare_charts(year_vars, "Synapse"), **button_options).pack(pady=5)


def prepare_charts(year_vars, keyword):
    selected_years = [int(year) for year, var in year_vars.items() if var.get() == 1]
    if not selected_years:
        messagebox.showwarning("Предупреждение", "Пожалуйста, выберите хотя бы один год.")
        return

    filtered = data[data['Year'].isin(selected_years)]
    filtered = filtered[filtered['Product'].str.contains(keyword, na=False)]

    if filtered.empty:
        messagebox.showinfo("Информация", f"Нет данных для {keyword} в выбранные года.")
        return

    qty_data = filtered.groupby('Year')['Qty'].sum()
    total_data = filtered.groupby('Year')['Total'].sum()

    chart_window = Toplevel(root)
    chart_window.title(f"Графики для {keyword}")
    chart_window.geometry("1300x800+300+100")  # немного увеличена высота

    # === Закрытие по ESC ===
    chart_window.bind("<Escape>", lambda event: chart_window.destroy())

    fig, axs = plt.subplots(1, 2, figsize=(14, 6))  # Размер графиков
    fig.suptitle(f"{keyword}", fontsize=16)

    bar_width = 0.7
    font_tick = 9
    font_value = 8

    # === Количество ===
    years_str = qty_data.index.astype(int).astype(str)
    bars1 = axs[0].bar(years_str, qty_data.values, color="#4caf50", width=bar_width)
    axs[0].set_title("Количество")
    axs[0].tick_params(axis='x', labelrotation=45, labelsize=font_tick)
    axs[0].tick_params(axis='y', labelsize=font_tick)

    for bar in bars1:
        height = bar.get_height()
        axs[0].text(
            bar.get_x() + bar.get_width() / 2,
            height + 0.01 * height,
            f"{height:.0f}",
            ha='center', va='bottom', rotation=0,
            fontsize=font_value
        )

    # === Сумма ===
    years_str_total = total_data.index.astype(int).astype(str)
    bars2 = axs[1].bar(years_str_total, total_data.values, color="#2196f3", width=bar_width)
    axs[1].set_title("Сумма")
    axs[1].yaxis.set_major_formatter(FuncFormatter(millions))
    axs[1].tick_params(axis='x', labelrotation=45, labelsize=font_tick)
    axs[1].tick_params(axis='y', labelsize=font_tick)

    for bar in bars2:
        height = bar.get_height()
        axs[1].text(
            bar.get_x() + bar.get_width() / 2,
            height + 0.01 * height,
            f"{height / 1e6:.1f}M",
            ha='center', va='bottom', rotation=0,
            fontsize=font_value
        )

    fig.tight_layout(rect=[0, 0, 1, 0.92])
    canvas = FigureCanvasTkAgg(fig, master=chart_window)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)

    # === Кнопка "ESC" ===
    exit_button = Button(chart_window, text="Закрыть [Esc]", command=chart_window.destroy,
                         font=("Arial", 10), bg="#f44336", fg="white")
    exit_button.pack(pady=10)
    plt.close(fig)


def open_money_report():
    # Создаём новое окно
    money_window = Toplevel(root)
    money_window.title("Отчет по финансам")

    # Получаем координаты главного окна и задаём смещение
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    offset_x = 50
    offset_y = 50
    money_window.geometry(f"500x400+{root_x + offset_x}+{root_y + offset_y}")

    # Устанавливаем цвет фона
    money_window.configure(bg="#dff5e1")

    # Центрируем всё содержимое в отдельный фрейм
    frame = Frame(money_window, bg="#dff5e1")
    frame.pack(expand=True)

    # Определяем настройки кнопок
    button_options = {
        "width": 40,  # Ширина в символах — подогнана под самую длинную строку
        "bg": "#39b54a",  # Зеленый фон
        "fg": "white",    # Белый текст
        "font": ("Arial", 10),  # Чуть крупнее и ровнее
        "relief": "raised",     # Эффект объема
        "bd": 2                 # Толщина границы
    }

    # Кнопки для отчетов по деньгам
    Button(frame, text="Общая статистика по продажам (год, сумма)",
           command=lambda: show_plot(data['Total'].groupby(data['Year']).sum(), "Общая статистика по продажам"),
           **button_options).pack(pady=10)

    Button(frame, text="Статистика продаж по месяцам за последние 5 лет",
           command=show_monthly_sales,
           **button_options).pack(pady=10)

    Button(frame, text="Статистика продаж по 'Product' за последние 5 лет",
           command=show_product_sales,
           **button_options).pack(pady=10)

    Button(frame, text="Статистика продаж по 'InvoiceType' за последние 5 лет",
           command=show_invoice_type_sales,
           **button_options).pack(pady=10)

    Button(frame, text="Статистика продаж по 'SWType' за последние 5 лет",
           command=show_swtype_sales,
           **button_options).pack(pady=10)



Button(center_frame, text="Отчет по финансам", command=open_money_report, **button_options).pack(pady=5)
Button(center_frame, text="Отчет по территории продаж", command=show_sales_by_region, **button_options).pack(pady=5)
Button(center_frame, text="Отчет по клиентам", command=show_sales_by_client, **button_options).pack(pady=5)
Button(center_frame, text="Отчет по дилерам", command=show_sales_by_dealer, **button_options).pack(pady=5)
Button(center_frame, text="Отчет по веткам и поколениям ПО", command=open_report_window, **button_options).pack(pady=5)
Button(center_frame, text="Отчет по приложениям", command=open_year_selection_window, **button_options).pack(pady=5)


root.mainloop()