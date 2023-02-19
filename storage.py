from abc import ABC, abstractmethod

from openpyxl import load_workbook

from entities import Measurements
from exceptions import NotExcelFile


class Storage(ABC):

    @abstractmethod
    def add(self, data: Measurements) -> bool:
        pass

    @abstractmethod
    def read(self) -> list[Measurements]:
        pass

    @abstractmethod
    def update(self, pk: int, data: Measurements) -> bool:
        pass

    @abstractmethod
    def clear(self) -> bool:
        pass

    @abstractmethod
    def delete(self, pk: int) -> bool:
        pass


class DBStorage(Storage):

    @abstractmethod
    def get_db_settings(self):
        pass


class ExcelParser(ABC):

    @abstractmethod
    def get_table(self) -> list[list]:
        pass


class ExcelOpenpyxlParser(ExcelParser):
    worksheet = None
    filename: str
    table: list[list]

    def __init__(self, filename: str):
        self.filename = filename

    def _connect(self):
        """Подключается к активному листу в файле."""
        workbook = load_workbook(self.filename, read_only=True, data_only=True)
        self.worksheet = workbook.active

    def _read_table(self):
        """Сохраняет данные из файла в table и освобождает память worksheet."""
        self.table = [[value for value in row] for row in self.worksheet.values]
        self.worksheet = None

    def _strip_rows(self):
        """Удаляет снизу фантомные пустые строки."""
        while len(self.table) > 0 and self.table[-1][0] is None:
            self.table.pop()

    def _strip_columns(self):
        """Удаляет справа фантомные пустые столбцы."""
        guaranteed_filled_row_index = 2
        while (
                len(self.table[guaranteed_filled_row_index]) > 0
                and
                self.table[guaranteed_filled_row_index][-1] is None
        ):
            for row_index in range(len(self.table)):
                self.table[row_index].pop()

    def _strip_table(self):
        """Удаляет фантомные пустые ячейки вокруг данных в таблице."""
        self._strip_rows()
        self._strip_columns()

    def _normalize_header_vertical(self, header_height: int):
        """"
        Находит название столбца в хэдере для вертикально объединенных ячеек
        и заполняет название столбца в объединенных ячейках в хэдере
        """
        column_name = [None, None]
        for column_index in range(len(column_name)):
            # Найти название столбца в хэдере
            for row_index in range(header_height):
                if self.table[row_index][0] is not None:
                    column_name[column_index] = self.table[row_index][column_index]
                    break
            # Заполнить название столбца в объединенных ячейках в хэдере
            for row_index in range(header_height):
                self.table[row_index][column_index] = column_name[column_index]

    def _normalize_header_horizontal(self, header_height: int):
        """"
        Находит название столбца в хэдере для горизонтально объединенных ячеек
        и заполняет название столбца в объединенных ячейках в хэдере
        """
        start_header_column_index = 2
        columns_in_row = len(self.table[0])
        for row_index in range(header_height):
            for column_index in range(
                    start_header_column_index + 1,
                    columns_in_row
            ):
                self.table[row_index][column_index] = (
                    self.table[row_index][column_index - 1]
                    if self.table[row_index][column_index] is None else
                    self.table[row_index][column_index]
                )

    def _normalize_header(self):
        """Исправляет объединенные ячейки в заголовке таблицы."""
        header_height = 3
        self._normalize_header_vertical(header_height)
        self._normalize_header_horizontal(header_height)

    def get_table(self) -> list[list]:
        """
        Загружает из файла лист Excel в матрицу,
        чистит и подготавливает матрицу для дальнейшей работы.
        """
        self._connect()
        self._read_table()
        self._strip_table()
        self._normalize_header()
        return self.table


class ExcelStorage(Storage):
    parser = ExcelOpenpyxlParser

    def __init__(self, filename):
        self.filename = self._check_file(filename)

    @staticmethod
    def _check_file(filename):
        import os
        expected_extensions = ('xls', 'xlsx')
        _, file_extension = os.path.splitext(filename)
        if file_extension[1:].lower() not in expected_extensions:
            raise NotExcelFile(
                f'Файл {filename} не формата excel, '
                f'его расширение не является {" или ".join(expected_extensions)}'
            )
        if not os.path.isfile(filename):
            raise FileExistsError(
                f'Файл {filename} не существует.'
            )
        return filename

    @staticmethod
    def _table_to_measurements(table) -> list[Measurements]:
        """Конвертирует таблицу в сущность Measurements."""
        print(*table, sep='\n')  # !!!!!!!!!!!!
        start_row_index = 3
        start_column_index = 2
        columns_in_row = len(table[0])
        rows_in_columns = len(table)
        data_cells = []
        for row_index in range(start_row_index, rows_in_columns):
            for column_index in range(start_column_index, columns_in_row):
                data_cells.append(
                    Measurements(
                        pk=self.validate_pk(self.get_pk(row_index))
                    )
                )
                validate_quantity(table[row_index][column_index])

        return data_cells

    def add(self, data: Measurements) -> bool:
        raise NotImplementedError

    def read(self) -> list[Measurements]:
        """Читает из Excel-файла и возвращает сущность Measurements."""
        parser = self.parser(self.filename)
        return self._table_to_measurements(parser.get_table())

    def update(self, pk: int, data: Measurements) -> bool:
        raise NotImplementedError

    def clear(self) -> bool:
        raise NotImplementedError

    def delete(self, pk) -> bool:
        raise NotImplementedError


db = ExcelStorage('data.xlsx')
measurements = db.read()

# print(type(db), ExcelStorage.__mro__)
# extensions = ('xls', 'xlsx')
# print(', '.join(extensions))
