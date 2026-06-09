import datetime
from dataclasses import dataclass

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt
from docx.table import Table

COLUMN_COUNT = 5


@dataclass
class PracticeData:
    practice: str
    course: str


@dataclass
class Record:
    specialization: str
    practices: list[PracticeData]
    max_student_count: str = "15"
    practice_format: str = "Очно, с применением ДОТ"


class WordWriter:
    def __init__(self, word_filepath: str, records: list[Record]):
        self.word_filepath = word_filepath
        self.records = records
        self.doc = Document()

    def set_column_widths(self, table: Table) -> None:
        widths = [Inches(3.5) for _ in range(COLUMN_COUNT)]

        for col_idx, width in enumerate(widths):
            for cell in table.columns[col_idx].cells:
                cell.width = width

    def enable_autofit(self, table: Table) -> None:
        tbl = table._element
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement("w:tblPr")
        tblLayout = OxmlElement("w:tblLayout")
        tblLayout.set(qn("w:type"), "fixed")
        tblPr.append(tblLayout)

    def generate_document(self) -> None:
        for section in self.doc.sections:
            section.top_margin = Cm(1.79)
            section.bottom_margin = Cm(1.79)
            section.left_margin = Cm(1.79)
            section.right_margin = Cm(1.39)

        self.set_text_part()
        self.set_table_part()
        self.doc.save(self.word_filepath)

    def set_text_part(self) -> None:
        style = self.doc.styles["Normal"]
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(f"""Приложение № 1
        к договору №____
        от «____»____________ {datetime.datetime.now().year}г.
        о практической подготовке обучающихся""")
        run.font.name = "XO Thames"
        run.font.size = Pt(12)
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "XO Thames")  # type: ignore

        self.doc.add_paragraph()

        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(
            "Образовательная программа (программы),\n\nкомпоненты образовательной "
            "программы, при реализации которых "
            "организуется практическая подготовка, количество обучающихся, "
            "осваивающих соответствующие компоненты "
            "образовательной программы, сроки организации практической подготовки"
        )
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.font.bold = False
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")  # type: ignore

    def set_header(self, table: Table) -> None:
        headers = [
            "Код, направление подготовки (профиль),\nосновные профессиональные\n"
            "образовательные программы (ОПОП)",
            "Наименование компонента\nобразовательной программы\n"
            "по учебному плану\n(дисциплина и вид(ы) "
            "учебных занятий\n(лекц., практич., лаб.)\nи (или) вид и тип практики)",
            "Максимальное количество\nобучающихся, осваивающих\nкомпонент "
            "образовательной\nпрограммы (за учебный год)",
            "Сроки организации\nпрактической подготовки\n(№ курса)",
            "Формат организации\nпрактической подготовки\n(очно/с применением ДОТ)",
        ]

        for col_idx, header_text in enumerate(headers):
            cell = table.rows[0].cells[col_idx]
            cell.text = header_text

            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(11)
                    run.font.bold = False
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")

            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def fill_data(self, table: Table) -> None:
        row_idx = 1

        for record in self.records:
            for practice_idx, practice_data in enumerate(record.practices):
                # Если строк в таблице меньше, чем нужно - добавляем
                if row_idx >= len(table.rows):
                    table.add_row()

                row = table.rows[row_idx]

                # Заполняем специализацию только для первой практики
                if practice_idx == 0:
                    row.cells[0].text = record.specialization
                else:
                    row.cells[0].text = ""

                row.cells[1].text = practice_data.practice
                row.cells[2].text = record.max_student_count
                row.cells[3].text = practice_data.course
                row.cells[4].text = record.practice_format

                # Настройка выравнивания
                for col_idx in range(COLUMN_COUNT):
                    cell = row.cells[col_idx]
                    for paragraph in cell.paragraphs:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        for run in paragraph.runs:
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(11)
                            run.font.bold = False
                            run._element.rPr.rFonts.set(
                                qn("w:eastAsia"), "Times New Roman"
                            )

                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                row_idx += 1

        self.remove_extra_rows(table, row_idx)

        self.merge_specialization_cells(table)

    def remove_extra_rows(self, table: Table, filled_rows: int) -> None:
        """Удаляет лишние пустые строки после заполненных"""
        rows = list(table.rows)
        # Удаляем строки, начиная с filled_rows и до конца
        for i in range(len(rows) - 1, filled_rows - 1, -1):
            if i >= filled_rows:
                row = rows[i]
                row.get_or_add_tr().getparent().remove(row._element)

    def merge_specialization_cells(self, table: Table) -> None:
        row_count = len(table.rows)

        if row_count <= 1:
            return

        row = 1
        while row < row_count:
            current_spec = table.cell(row, 0).text
            if not current_spec:
                row += 1
                continue

            end_row = row
            while end_row + 1 < row_count and not table.cell(end_row + 1, 0).text:
                end_row += 1

            if end_row > row:
                start_cell = table.cell(row, 0)
                end_cell = table.cell(end_row, 0)
                start_cell.merge(end_cell)

                merged_cell = table.cell(row, 0)
                merged_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                if len(merged_cell.paragraphs) > 1:
                    for p in merged_cell.paragraphs[1:]:
                        p.clear()

            row = end_row + 1

    def set_table_part(self) -> None:
        table = self.doc.add_table(rows=1, cols=COLUMN_COUNT)
        table.style = "Table Grid"

        self.enable_autofit(table)
        self.set_column_widths(table)
        self.set_header(table)
        self.fill_data(table)
