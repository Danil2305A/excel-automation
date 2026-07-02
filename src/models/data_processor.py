from models.word_writer import Record


class DataProcessor:
    def __parse_practice(self, practice: str) -> tuple[str, str]:
        if " - " in practice:
            type, name = practice.split(" - ", 1)
            return type, name
        return practice, ""

    def __get_type_order(self, type: str) -> int:
        order = {
            "Учебная практика": 1,
            "Производственная практика": 2,
            "Преддипломная практика": 3,
        }
        return order.get(type, 999)

    def process_records(self, records: list[Record]) -> None:
        if not records:
            return

        records.sort(key=lambda record: record.specialization)

        i = 1
        while i < len(records):
            current_record = records[i]
            previous_record = records[i - 1]

            if current_record.specialization == previous_record.specialization:
                previous_record.practices.extend(current_record.practices)
                records.pop(i)
            else:
                i += 1

        for record in records:
            if len(record.practices) > 1:
                seen: set[tuple[str, str]] = set()
                unique_practices = []

                for practice in record.practices:
                    key = (practice.practice, practice.course)
                    if key not in seen:
                        seen.add(key)
                        unique_practices.append(practice)

                unique_practices.sort(
                    key=lambda p: (
                        self.__get_type_order(self.__parse_practice(p.practice)[0]),
                        p.course,
                    )
                )

                record.practices = unique_practices
