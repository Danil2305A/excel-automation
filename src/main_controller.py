import threading

from main_view import MainView
from models.main_model import MainModel


class MainController:
    def __init__(self, model: MainModel, view: MainView):
        self.model = model
        self.view = view

        self.model.add_observer(self)

        self.view.on_load_click = self.handle_load  # type: ignore
        self.view.on_download_click = self.handle_download  # type: ignore

        self.is_loading = False

    def on_progress(self, current: int, total: int, message: str) -> None:
        self.view.update_progress(current, total)
        if message:
            self.view.add_status_message(message)

    def on_file_processed(self, message: str) -> None:
        self.view.add_status_message(message)

        stats = self.model.get_summary_stats()
        self.view.update_stats(stats["total_records"])

    def handle_load(self) -> None:
        if self.is_loading:
            self.view.add_status_message("Загрузка уже выполняется...")
            return

        folder = self.view.ask_for_folder()
        if not folder:
            self.view.add_status_message("Выбор папки отменён")
            return

        self.is_loading = True
        self.view.set_load_button_state(False)
        self.view.set_download_button_state(False)
        self.view.update_progress(0, 100)
        self.view.clear_status()

        thread = threading.Thread(target=self._load_thread, args=(folder,))
        thread.daemon = True
        thread.start()

    def _load_thread(self, folder: str) -> None:
        try:
            records = self.model.load_all_files(folder)
            success_count = len(records)

            if success_count > 0:
                self.view.add_status_message(f"Обработано файлов: {success_count}.")
                self.view.set_download_button_state(True)
            else:
                self.view.add_status_message(
                    "Не удалось загрузить ни одного корректного файла. "
                    "Проверьте структуру папок и формат Excel-файлов.",
                )

        except Exception as e:
            self.view.add_status_message(f"Ошибка при загрузке: {str(e)}")
        finally:
            self.is_loading = False
            self.view.set_load_button_state(True)

    def handle_download(self) -> None:
        if not self.model.records:
            self.view.add_status_message(
                "Нет загруженных данных. Сначала загрузите учебные планы."
            )
            return

        save_path = self.view.ask_for_save_path()
        if not save_path:
            self.view.add_status_message("Сохранение отменено")
            return

        try:
            self.view.add_status_message("Формирование Word-документа...")
            self.view.set_download_button_state(False)

            success = self.model.generate_word_document(save_path)

            if success:
                self.view.show_success_message(save_path)
                self.view.add_status_message(f"Документ сохранён: {save_path}")
            else:
                self.view.show_error_message("Не удалось сформировать документ")

        except Exception as e:
            self.view.show_error_message(str(e))
            self.view.add_status_message(f"Ошибка: {str(e)}")
        finally:
            self.view.set_download_button_state(True)
