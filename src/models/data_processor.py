import os
from dataclasses import dataclass

from models.excel_parser import ExcelParser, Practice, Specialization
from models.word_writer import PracticeData, Record, WordWriter


@dataclass
class DataProcessor:
    selected_folder: str = ""

    def scan_files(self) -> list[str]:
        excel_files = []
        for root, _, files in os.walk(self.selected_folder):
            for file in files:
                if file.endswith((".xlsx", ".xls")):
                    excel_files.append(os.path.join(root, file))
        return excel_files

    def process_excel_file(self, filepath: str) -> Record:
        excel_parser: ExcelParser = ExcelParser(filepath)

        try:
            specialization: Specialization = excel_parser.get_specialization()
            practices: list[Practice] = excel_parser.get_practices()

            practice_data_list = []
            for practice in practices:
                practice_data_list.append(
                    PracticeData(practice=str(practice), course=practice.course)
                )

            return Record(
                specialization=str(specialization), practices=practice_data_list
            )
        finally:
            excel_parser.wb.close()

    def process_word_file(self, filepath: str, records: list[Record]) -> None:
        word_writer: WordWriter = WordWriter(filepath, records)
        word_writer.generate_document()
