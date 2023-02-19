import datetime
from dataclasses import dataclass


@dataclass
class Entities:
    pk: int

    @classmethod
    def db_name(cls):
        return cls.__name__.lower()


@dataclass
class EntityNameMixin:
    name: str


@dataclass
class Companies(Entities, EntityNameMixin):
    ...


@dataclass
class FactForecasts(Entities, EntityNameMixin):
    ...


@dataclass
class Substances(Entities, EntityNameMixin):
    ...


@dataclass
class Datas(Entities, EntityNameMixin):
    ...


@dataclass
class Measurements(Entities):
    date: datetime.date
    company: Companies
    fact_forecast: FactForecasts
    substance: Substances
    data: Datas
    quantity: float


# @dataclass
# class Measure(Entities):
#     fact_forecast: FactForecasts
#     substance: Substances
#     data: Datas
#     quantity: float
#
#
# @dataclass
# class Measurements(Entities):
#     date: datetime.date
#     company: Companies
#     day_measure: list[Measure]


# c = Companies(pk=1, name='C1')
# print(Companies.db_name())
# print(c.name, c.pk)
# print(Measurements.db_name())
# print(type(Companies))
# print(isinstance(c, Entities))
# print(issubclass(Companies, Entities))
# print(Companies.__mro__)

