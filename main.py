import traceback
from services import run_test_task


def main():
    run_test_task()


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print('Возникла ошибка при выполнении:', error)
        traceback.print_exception(error)
    else:
        print('Выполнено успешно.')
