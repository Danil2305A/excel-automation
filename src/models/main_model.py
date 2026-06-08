import os
from typing import Any

from models.data_processor import DataProcessor
from models.word_writer import Record


class MainModel:
    def __init__(self) -> None:
        self.selected_folder: str = ""
        self.records: list[Record] = []
        self.processor = DataProcessor()
        self._observers: list = []

    def add_observer(self, observer: Any) -> None:
        self._observers.append(observer)

    def notify_progress(self, current: int, total: int, message: str = "") -> None:
        for observer in self._observers:
            observer.on_progress(current, total, message)

    def notify_file_processed(self, message: str = "") -> None:
        for observer in self._observers:
            observer.on_file_processed(message)

    def load_all_files(self, folder_path: str) -> list[Record]:
        self.selected_folder = folder_path
        self.processor.selected_folder = folder_path

        excel_files = self.processor.scan_files()

        if not excel_files:
            self.notify_progress(0, 0, "Excel-файлы не найдены")
            return []

        total = len(excel_files)
        self.records = []

        for idx, file_path in enumerate(excel_files, 1):
            self.notify_progress(
                idx, total, f"Обработка: {os.path.basename(file_path)}"
            )

            try:
                record = self.processor.process_excel_file(file_path)
                self.records.append(record)
                self.notify_file_processed(f"  {record.specialization}")
            except Exception:
                self.notify_file_processed(
                    f"Ошибка: {file_path} не пригоден для извлечения данных",
                )

        self.notify_progress(total, total, "Загрузка завершена")
        return self.records

    def generate_word_document(self, output_path: str) -> bool:
        if not self.records:
            return False

        try:
            self.processor.process_word_file(output_path, self.records)
            return True
        except Exception as e:
            raise e

    def get_summary_stats(self) -> dict[str, Any]:
        if not self.records:
            return {"total_records": 0}

        return {
            "total_records": len(self.records),
            "specializations": [r.specialization for r in self.records],
        }

    def clear_data(self) -> None:
        self.records = []
        self.selected_folder = ""
