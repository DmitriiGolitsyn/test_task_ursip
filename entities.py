import datetime
from dataclasses import dataclass


@dataclass
class Entities:
    db_pk: int | None

    @classmethod
    def get_db_name(cls):
        return cls.__name__.lower()


@dataclass
class EntityNameMixin:
    name: str

    def __str__(self):
        return self.name


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
class FSDs(Entities):
    fact_forecasts: FactForecasts
    substance: Substances
    data: Datas

    def __str__(self):
        return '_'.join((self.fact_forecasts.name, self.substance.name, self.data.name))


@dataclass
class Measurements(Entities):
    fsd: FSDs
    quantity: float | None


@dataclass
class DayMeasurements(Entities):
    date: datetime.date
    company: Companies
    day_measurements: list[Measurements]
