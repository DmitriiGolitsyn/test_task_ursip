import datetime
import sqlite3
from abc import ABC, abstractmethod

from openpyxl import load_workbook

from entities import (
    DayMeasurements, Measurements, Companies, FactForecasts, Substances, Datas, FSDs
)
from exceptions import NotExcelFile, ExcelValidationError, MeasurementsAbsentError


class StorageInterface(ABC):

    @abstractmethod
    def add(self, data: list[DayMeasurements]) -> bool:
        pass

    @abstractmethod
    def read(self) -> list[DayMeasurements]:
        pass

    @abstractmethod
    def update(self, data: list[DayMeasurements]) -> bool:
        pass

    @abstractmethod
    def clear(self) -> bool:
        pass

    @abstractmethod
    def delete(self, pk: int) -> bool:
        pass


class DBStorageMixin(ABC):

    @abstractmethod
    def _get_db_settings(self):
        pass


class DBStorageInterface(StorageInterface, DBStorageMixin):

    @abstractmethod
    def _get_db_settings(self):
        pass

    @abstractmethod
    def add(self, data: list[DayMeasurements]) -> bool:
        pass

    @abstractmethod
    def read(self) -> list[DayMeasurements]:
        pass

    @abstractmethod
    def update(self, data: list[DayMeasurements]) -> bool:
        pass

    @abstractmethod
    def clear(self) -> bool:
        pass

    @abstractmethod
    def delete(self, pk: int) -> bool:
        pass


class ExcelParserInterface(ABC):

    @abstractmethod
    def get_table(self) -> list[list]:
        pass


class ExcelOpenpyxlParser(ExcelParserInterface):
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


class ExcelStorage(StorageInterface):
    parser = ExcelOpenpyxlParser

    measurements: list[DayMeasurements] = []
    companies: list[Companies] = []
    fact_forecasts: list[FactForecasts] = []
    substances: list[Substances] = []
    datas: list[Datas] = []
    fsds: list[FSDs] = []
    table: list[list[str]] = []

    company_column_index = 1
    fact_forecast_row_index = 0
    substance_row_index = 1
    data_row_index = 2
    # Начало данных quantity:
    start_row_index = 3
    start_column_index = 2

    def __init__(self, filename):
        self.filename = self._check_excel_file(filename)

    @staticmethod
    def _check_excel_file(filename):
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
    def __get_fake_date(row_index: int) -> datetime.date:
        """Заглушка. 30 - максимум."""
        header_height = 3
        if row_index > 30:
            raise ValueError('Значение не должно превышать 30')
        return datetime.date(2023, 1, row_index - header_height + 1)

    def _get_or_create_company(self, row_index) -> Companies:
        company_name = self.table[row_index][self.company_column_index].strip()
        for company in self.companies:
            if company.name == company_name:
                return company
        self.companies.append(Companies(db_pk=None, name=company_name))
        return self.companies[-1]

    def _get_or_create_fsd_commons(self, entity_class, instances_list, entity_row_index, column_index):
        instance_name = self.table[entity_row_index][column_index].strip()
        for instance in instances_list:
            if instance.name == instance_name:
                return instance
        instances_list.append(entity_class(db_pk=None, name=instance_name))
        return instances_list[-1]

    def _get_or_create_fact_forecast(self, column_index):
        return self._get_or_create_fsd_commons(
            entity_class=FactForecasts,
            instances_list=self.fact_forecasts,
            entity_row_index=self.fact_forecast_row_index,
            column_index=column_index
        )

    def _get_or_create_substance(self, column_index):
        return self._get_or_create_fsd_commons(
            entity_class=Substances,
            instances_list=self.substances,
            entity_row_index=self.substance_row_index,
            column_index=column_index
        )

    def _get_or_create_data(self, column_index):
        return self._get_or_create_fsd_commons(
            entity_class=Datas,
            instances_list=self.datas,
            entity_row_index=self.data_row_index,
            column_index=column_index
        )

    def _get_or_create_fsd(self, fact_forecast, substance, data):
        for fsd in self.fsds:
            if (
                fsd.fact_forecasts is fact_forecast
                and
                fsd.substance is substance
                and
                fsd.data is data
            ):
                return fsd
        self.fsds.append(FSDs(
            db_pk=None,
            fact_forecasts=fact_forecast,
            substance=substance,
            data=data
        ))
        return self.fsds[-1]

    def _get_day_measurements(self, row_index) -> list[Measurements]:
        day_measurements: list[Measurements] = []
        for column_index in range(self.start_column_index, len(self.table[0])):
            try:
                quantity = float(self.table[row_index][column_index])
            except TypeError:
                raise ExcelValidationError(
                    'Количество должно быть числом. '
                    f'{self.table[row_index][column_index]} - это не число.'
                )
            fact_forecast = self._get_or_create_fact_forecast(column_index)
            substance = self._get_or_create_substance(column_index)
            data = self._get_or_create_data(column_index)
            fsd = self._get_or_create_fsd(fact_forecast, substance, data)
            day_measurements.append(Measurements(db_pk=None, fsd=fsd, quantity=quantity))
        return day_measurements

    def _table_to_measurements(self):
        """Создает список ежедневных измерений Measurements."""
        rows_in_columns = len(self.table)
        for row_index in range(self.start_row_index, rows_in_columns):
            self.measurements.append(
                DayMeasurements(
                    db_pk=None,
                    date=self.__get_fake_date(row_index),
                    company=self._get_or_create_company(row_index),
                    day_measurements=self._get_day_measurements(row_index)
                )
            )

    def add(self, data: list[DayMeasurements]) -> bool:
        raise NotImplementedError

    def read(self) -> list[DayMeasurements]:
        """Читает из Excel-файла и возвращает сущность Measurements."""
        parser = self.parser(self.filename)
        self.table = parser.get_table()
        self._table_to_measurements()
        return self.measurements

    def update(self, data: list[DayMeasurements]) -> bool:
        raise NotImplementedError

    def clear(self) -> bool:
        raise NotImplementedError

    def delete(self, pk: int) -> bool:
        raise NotImplementedError


class DBStorageSQLite(DBStorageInterface):
    db_name: str
    conn: sqlite3.Connection
    cur: sqlite3.Cursor
    data: list[DayMeasurements] | None

    def _init_connection(self, foreign_keys: bool):
        self.conn = sqlite3.connect(
            self.db_name,
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES
        )
        switcher = 'ON' if foreign_keys else 'OFF'
        self.conn.execute(f'PRAGMA foreign_keys = {switcher}')
        self.cur = self.conn.cursor()

    def __init__(self, db_name: str):
        self.db_name = db_name
        self._init_connection(foreign_keys=True)

    def _get_db_settings(self):
        """Не реализовано. Импорт настроек БД из конфига settings.py."""
        raise NotImplementedError

    def _create_table_companies(self):
        sql = f'CREATE TABLE IF NOT EXISTS {Companies.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'name TEXT'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_fact_forecasts(self):
        sql = f'CREATE TABLE IF NOT EXISTS {FactForecasts.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'name TEXT'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_substances(self):
        sql = f'CREATE TABLE IF NOT EXISTS {Substances.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'name TEXT'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_datas(self):
        sql = f'CREATE TABLE IF NOT EXISTS {Datas.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'name TEXT'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_fsd_s(self):
        sql = f'CREATE TABLE IF NOT EXISTS {FSDs.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'fact_forecasts_id INTEGER,'
        sql += 'substance_id INTEGER,'
        sql += 'data_id INTEGER,'
        sql += f'FOREIGN KEY(fact_forecasts_id) REFERENCES {FactForecasts.get_db_name()}(id),'
        sql += f'FOREIGN KEY(substance_id) REFERENCES {Substances.get_db_name()}(id),'
        sql += f'FOREIGN KEY(data_id) REFERENCES {Datas.get_db_name()}(id)'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_measurements(self):
        sql = f'CREATE TABLE IF NOT EXISTS {Measurements.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'quantity REAL,'
        sql += 'fsd_s_id INTEGER,'
        sql += 'day_measurements_id INTEGER,'
        sql += f'FOREIGN KEY(fsd_s_id) REFERENCES {FSDs.get_db_name()}(id),'
        sql += f'FOREIGN KEY(day_measurements_id) REFERENCES {DayMeasurements.get_db_name()}(id)'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_table_day_measurements(self):
        sql = f'CREATE TABLE IF NOT EXISTS {DayMeasurements.get_db_name()}('
        sql += 'id INTEGER PRIMARY KEY AUTOINCREMENT,'
        sql += 'date DATE,'
        sql += 'company_id INTEGER,'
        sql += f'FOREIGN KEY(company_id) REFERENCES {Companies.get_db_name()}(id)'
        sql += ');'
        self.cur.execute(sql)
        self.conn.commit()

    def _create_tables(self):
        self._create_table_companies()
        self._create_table_fact_forecasts()
        self._create_table_substances()
        self._create_table_datas()
        self._create_table_fsd_s()
        self._create_table_measurements()
        self._create_table_day_measurements()

    def _get_or_create_fact_forecasts(self, fact_forecasts: FactForecasts) -> FactForecasts:
        if fact_forecasts.db_pk:
            return fact_forecasts
        sql = f'INSERT INTO {FactForecasts.get_db_name()}(name) VALUES(?);'
        self.cur.execute(sql, (fact_forecasts.name, ))
        fact_forecasts_id = self.cur.lastrowid
        self.conn.commit()
        fact_forecasts.db_pk = fact_forecasts_id
        return fact_forecasts

    def _get_or_create_substances(self, substance: Substances) -> Substances:
        if substance.db_pk:
            return substance
        sql = f'INSERT INTO {Substances.get_db_name()}(name) VALUES(?);'
        self.cur.execute(sql, (substance.name, ))
        substance_id = self.cur.lastrowid
        self.conn.commit()
        substance.db_pk = substance_id
        return substance

    def _get_or_create_datas(self, data: Datas) -> Datas:
        if data.db_pk:
            return data
        sql = f'INSERT INTO {Datas.get_db_name()}(name) VALUES(?);'
        self.cur.execute(sql, (data.name, ))
        data_id = self.cur.lastrowid
        self.conn.commit()
        data.db_pk = data_id
        return data

    def _get_or_create_fsd(self, fsd: FSDs) -> FSDs:
        if fsd.db_pk:
            return fsd
        fact_forecasts = self._get_or_create_fact_forecasts(fsd.fact_forecasts)
        substance = self._get_or_create_substances(fsd.substance)
        data = self._get_or_create_datas(fsd.data)
        sql = f'INSERT INTO {FSDs.get_db_name()}('
        sql += 'fact_forecasts_id, substance_id, data_id'
        sql += ') VALUES(?, ?, ?);'
        self.cur.execute(sql, (fact_forecasts.db_pk, substance.db_pk, data.db_pk))
        fsd_id = self.cur.lastrowid
        self.conn.commit()
        fsd.db_pk = fsd_id
        return fsd

    def _get_or_create_company(self, day_index) -> Companies:
        if self.data[day_index].company.db_pk:
            return self.data[day_index].company
        company_name = self.data[day_index].company.name
        sql = f'INSERT INTO {Companies.get_db_name()}(name) VALUES(?);'
        self.cur.execute(sql, (company_name, ))
        company_id = self.cur.lastrowid
        self.conn.commit()
        self.data[day_index].company.db_pk = company_id
        return self.data[day_index].company

    def _add_measurement_data(self, measurement: Measurements, day_measurements_id: int):
        fsd = self._get_or_create_fsd(measurement.fsd)
        sql = f'INSERT INTO {Measurements.get_db_name()}('
        sql += 'quantity, fsd_s_id, day_measurements_id'
        sql += ') VALUES(?, ?, ?);'
        self.cur.execute(sql, (
            measurement.quantity,
            fsd.db_pk,
            day_measurements_id
        ))
        measurement_id = self.cur.lastrowid
        self.conn.commit()
        measurement.db_pk = measurement_id

    def _add_day_measurements_data(self, day_index: int):
        date = self.data[day_index].date
        company = self._get_or_create_company(day_index)
        sql = f'INSERT INTO {DayMeasurements.get_db_name()}(date, company_id) VALUES(?, ?);'
        self.cur.execute(sql, (date, company.db_pk))
        day_measurements_id = self.cur.lastrowid
        self.conn.commit()
        self.data[day_index].db_pk = day_measurements_id
        for measurement in self.data[day_index].day_measurements:
            self._add_measurement_data(measurement, day_measurements_id)

    def _fill_tables_with_data(self):
        for day_index in range(len(self.data)):
            self._add_day_measurements_data(day_index)

    def add(self, data: None | list[DayMeasurements]) -> bool:
        if not data:
            raise MeasurementsAbsentError('Нет измерений для добавления в БД.')
        self.data = data
        self._create_tables()
        self._fill_tables_with_data()
        return True

    def read(self) -> list[DayMeasurements]:
        raise NotImplementedError

    def update(self, data: list[DayMeasurements]) -> bool:
        raise NotImplementedError

    def clear(self) -> bool:
        """Удаляет таблицы в базе данных"""
        self.conn.close()
        self._init_connection(foreign_keys=False)
        sql = 'SELECT name FROM sqlite_schema WHERE type="table";'
        self.cur.execute(sql)
        tables = self.cur.fetchall()
        for table in tables:
            table_name = table[0]
            if 'SQLITE'.lower() in table_name.lower():
                continue
            sql = f"DROP TABLE IF EXISTS {table_name};"
            self.cur.execute(sql)
            self.conn.commit()
        self.conn.close()
        self._init_connection(foreign_keys=True)
        return True

    def delete(self, pk: int) -> bool:
        raise NotImplementedError


class DBStoragePostgreSQL(DBStorageInterface):
    """Не реализовано. Пример реализации интерфейса работы с другой БД."""

    def _get_db_settings(self):
        """Не реализовано. Импорт настроек БД из конфига settings.py."""
        raise NotImplementedError

    def add(self, data: None | list[DayMeasurements]) -> bool:
        raise NotImplementedError

    def read(self) -> list[DayMeasurements]:
        raise NotImplementedError

    def update(self, data: list[DayMeasurements]) -> bool:
        raise NotImplementedError

    def clear(self) -> bool:
        """Удаляет таблицы в базе данных"""
        raise NotImplementedError

    def delete(self, pk: int) -> bool:
        raise NotImplementedError
