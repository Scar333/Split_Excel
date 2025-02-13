import os
import tkinter as tk
from multiprocessing import Process, Queue, freeze_support, set_start_method
from tkinter import filedialog, messagebox

from job_with_excel import creat_new_files


def run_split_process(excel_file: str, num_files: int) -> None:
    """Запуск процесса деления в отдельном потоке

    :param excel_file: путь до excel файла
    :param num_files: количество файлов, на которое необходимо разделить excel файл
    """
    creat_new_files(excel_file, num_files)


def create_gui() -> None:
    """Main окно"""
    root = tk.Tk()
    root.title("Программа для разделения Excel файла")
    root.geometry(
        "600x150+{}+{}".format(int(root.winfo_screenwidth() / 2 - 300), int(root.winfo_screenheight() / 2 - 150)))

    icon_path = os.path.join(os.path.dirname(__file__), 'excel.ico')
    root.iconbitmap(icon_path)

    root.resizable(False, False)

    tk.Label(root, text="Введите количество файлов:", anchor="e").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    num_files_entry = tk.Entry(root, width=30)
    num_files_entry.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(root, text="Выберите путь до файла:       ", anchor="e").grid(row=1, column=0, padx=10, pady=10,
                                                                           sticky="e")

    def select_excel_file() -> None:
        """Выбор Excel файла"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])
        if file_path:
            excel_file_entry.delete(0, tk.END)
            excel_file_entry.insert(0, file_path)

    excel_file_entry = tk.Entry(root, width=30)
    excel_file_entry.grid(row=1, column=1, padx=10, pady=10)
    select_excel_button = tk.Button(root, text="Выбрать Excel файл", command=select_excel_file)
    select_excel_button.grid(row=1, column=2, padx=10, pady=10)

    def split_excel():
        """Кнопка разделить"""
        if validate_input():
            num_files = int(num_files_entry.get())
            excel_file = excel_file_entry.get()

            split_button.config(state=tk.DISABLED)

            progress_queue = Queue()

            proc = Process(target=run_split_process, args=(excel_file, num_files))
            proc.start()

            # Проверяем завершение процесса и обновляем интерфейс
            root.after(100, check_process_completion, proc, progress_queue, num_files_entry, excel_file_entry,
                       split_button, num_files)

    def check_process_completion(proc, progress_queue, num_files_entry, excel_file_entry, split_button, num_files):
        """Проверка завершения процесса"""
        if not proc.is_alive():
            messagebox.showinfo("Успешная операция", f"Excel файл был успешно разделен!")
            num_files_entry.delete(0, tk.END)
            excel_file_entry.delete(0, tk.END)
            split_button.config(state=tk.NORMAL)
        else:
            root.after(100, check_process_completion, proc, progress_queue, num_files_entry, excel_file_entry,
                       split_button, num_files)

    split_button = tk.Button(root, text="Разделить", command=split_excel)
    split_button.grid(row=2, column=1, padx=10, pady=10)

    def validate_input():
        """Проверка ввода данных"""
        num_files = num_files_entry.get()
        excel_file = excel_file_entry.get()

        try:
            int(num_files)
        except ValueError:
            messagebox.showerror("Ошибка", "Пожалуйста, введите число в поле 'Введите количество файлов'.")
            return False
        if int(num_files) > 1:
            pass
        else:
            messagebox.showerror("Ошибка", "Пожалуйста, введите число больше 1")
            return False
        if excel_file and (excel_file.endswith('.xlsx') or excel_file.endswith('.xls')):
            return True
        else:
            messagebox.showerror("Ошибка", "Пожалуйста, прикрепите Excel файл.")
            return False

    root.mainloop()


if __name__ == "__main__":
    freeze_support()
    set_start_method('spawn')
    create_gui()
