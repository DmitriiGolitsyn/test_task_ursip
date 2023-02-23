import datetime
from abc import ABC, abstractmethod
from dataclasses import dataclass

from exceptions import MeasurementsAbsentError
from entities import DayMeasurements, FactForecasts, Substances, Companies


class ReportInterface(ABC):

    @abstractmethod
    def make_report(self):
        pass

    @abstractmethod
    def get_report(self):
        pass


@dataclass
class FSs:
    fact_forecasts: FactForecasts
    substance: Substances


@dataclass
class MeasurementsDataSum:
    fs: FSs
    quantity: float | None

    def __str__(self):
        return self.quantity


@dataclass
class DayMeasurementsDataSum:
    date: datetime.date
    company: Companies
    day_measurements: list[MeasurementsDataSum]

    def __str__(self):
        line = f'{self.date} \t {self.company.name} \t\t'
        line += '\t\t'.join([str(m.quantity) for m in self.day_measurements])
        return line


class DataSumReport(ReportInterface):
    measurements: list[DayMeasurements]
    report: list[DayMeasurementsDataSum] = []

    def __init__(self, measurements: list[DayMeasurements]):
        if not measurements:
            raise MeasurementsAbsentError(
                'Отсутствуют данные для анализа.'
            )
        self.measurements = measurements

    def _get_measurements_data_sum(self, day_measurements) -> list[MeasurementsDataSum]:
        measurements_data = []
        m = day_measurements
        data_sum = m[0].quantity
        for index in range(0, len(day_measurements) - 1):
            if (
                    (m[index].fsd.fact_forecasts, m[index].fsd.substance) ==
                    (m[index + 1].fsd.fact_forecasts, m[index + 1].fsd.substance)
            ):
                data_sum += m[index + 1].quantity
                if index + 1 == len(day_measurements) - 1:
                    measurements_data.append(MeasurementsDataSum(
                        fs=FSs(
                            m[index].fsd.fact_forecasts,
                            m[index].fsd.substance
                        ),
                        quantity=data_sum
                    ))
            else:
                measurements_data.append(MeasurementsDataSum(
                    fs=FSs(
                        m[index].fsd.fact_forecasts,
                        m[index].fsd.substance
                    ),
                    quantity=data_sum
                ))
                data_sum = m[index + 1].quantity
        return measurements_data

    def make_report(self):
        for day_measurements in self.measurements:
            self.report.append(DayMeasurementsDataSum(
                date=day_measurements.date,
                company=day_measurements.company,
                day_measurements=self._get_measurements_data_sum(day_measurements.day_measurements)
            ))

    def get_report(self):
        return self.report or None

    def _print_to_terminal(self):
        """Тестовый вывод"""
        header = [
            ['', '', ''],
            ['#', 'date', 'company']
        ]
        for m in self.report[0].day_measurements:
            header[0] += [m.fs.fact_forecasts.name]
            header[1] += [m.fs.substance.name]
        body = [None] * len(self.report)
        for row_index in range(len(self.report)):
            body[row_index] = []
            body[row_index].append(str(row_index + 1))
            body[row_index].append(self.report[row_index].date.strftime('%d/%m/%Y'))
            body[row_index].append(self.report[row_index].company.name)
            for column_index in range(len(self.report[row_index].day_measurements)):
                body[row_index].append(
                    str(self.report[row_index].day_measurements[column_index].quantity)
                )
        width = 10
        table = header + body
        for row_data in table:
            line = ' \t'.join(map(lambda cell: cell.rjust(width), row_data))
            print(line)
