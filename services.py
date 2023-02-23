from storage import ExcelStorage, DBStorageSQLite
from analytics import DataSumReport


def run_test_task():
    excel = ExcelStorage('data.xlsx')
    measurements = excel.read()

    db = DBStorageSQLite('sqlite.db')
    db.clear()
    db.add(measurements)

    report_data_sum = DataSumReport(measurements)
    report_data_sum.make_report()
    report_data_sum._print_to_terminal()
