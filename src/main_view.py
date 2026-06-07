import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk


class MainView:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(
            "Программа для формирования приложения 1 "
            "к договору о практической подготовке"
        )
        self.root.geometry("700x500")
        self.root.configure(bg="#ffffff")

        self.on_load_click = None
        self.on_download_click = None

        self.create_widgets()

    def create_widgets(self) -> None:
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main_frame,
            text="СибГУТИ",
            font=("Arial", 24, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
        ).pack(pady=(0, 30))

        info_frame = tk.Frame(main_frame, bg="#f0f0f0", relief=tk.GROOVE, bd=2)
        info_frame.pack(fill=tk.X, pady=10)

        self.records_count_label = tk.Label(
            info_frame,
            text="Загружено направлений: 0",
            font=("Arial", 11),
            bg="#f0f0f0",
            fg="#34495e",
        )
        self.records_count_label.pack(side=tk.LEFT, padx=10, pady=5)

        self.load_btn = tk.Button(
            main_frame,
            text="Загрузить УП",
            font=("Arial", 12),
            bg="#47ff8d",
            fg="black",
            activebackground="#00CC55",
            cursor="hand2",
            padx=20,
            pady=10,
            command=lambda: self.on_load_click() if self.on_load_click else None,
        )
        self.load_btn.pack(pady=20)

        self.download_btn = tk.Button(
            main_frame,
            text="Скачать приложение",
            font=("Arial", 12),
            bg="#0077ff",
            fg="white",
            activebackground="#0058CC",
            cursor="hand2",
            padx=20,
            pady=10,
            command=lambda: (
                self.on_download_click() if self.on_download_click else None
            ),
            state=tk.DISABLED,
        )
        self.download_btn.pack(pady=10)

        status_frame = tk.LabelFrame(
            main_frame,
            text="Статус обработки",
            font=("Arial", 11, "bold"),
            bg="#f0f0f0",
            fg="#2c3e50",
            padx=10,
            pady=10,
        )
        status_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))

        self.status_text = tk.Text(
            status_frame,
            height=10,
            font=("Consolas", 9),
            bg="#d6d6d6",
            wrap=tk.WORD,
            state=tk.DISABLED,
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.status_text, command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        style = ttk.Style()

        style.configure(
            "Color.Horizontal.TProgressbar",
            background="#00ff6a",  # Цвет заполнения
            troughcolor="#707070",  # Цвет фона
            bordercolor="#27ae60",  # Цвет границы
            lightcolor="#2ecc71",  # Светлый оттенок
            darkcolor="#27ae60",  # Тёмный оттенок
        )

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
            style="Color.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

    def add_status_message(self, message: str) -> None:
        self.status_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")

        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()

    def update_progress(self, value: int, maximum: int = 100) -> None:
        progress = (value / maximum) * 100 if maximum > 0 else 0
        self.progress_var.set(progress)
        self.root.update_idletasks()

    def update_stats(self, records_count: int) -> None:
        self.records_count_label.configure(
            text=f"Загружено учебных планов: {records_count}"
        )

    def set_download_button_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.download_btn.configure(state=state)  # type: ignore

    def set_load_button_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        text = "Загрузить УП" if enabled else "Обработка..."
        self.load_btn.configure(state=state, text=text)  # type: ignore

    def ask_for_folder(self) -> str:
        return filedialog.askdirectory(
            title="Выберите корневую папку с учебными планами",
            initialdir=os.path.expanduser("~"),
        )

    def ask_for_save_path(self) -> str:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        default_name = (
            f"Приложение_1_к_договору_о_практической_подготовке_{timestamp}.docx"
        )

        return filedialog.asksaveasfilename(
            title="Сохранить приложение № 1 как",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
            initialfile=default_name,
        )

    def show_success_message(self, file_path: str) -> None:
        messagebox.showinfo(
            "Успех",
            f"Приложение №1 успешно сформировано!\n\nФайл сохранён:\n{file_path}",
        )

    def show_error_message(self, error_text: str) -> None:
        messagebox.showerror(
            "Ошибка", f"Не удалось сформировать документ:\n\n{error_text}"
        )

    def clear_status(self) -> None:
        self.status_text.configure(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.configure(state=tk.DISABLED)
