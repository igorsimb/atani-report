import pytest

from data_locators import OzonFile
from collect_info import (
    get_column_index_by_cell_value,
    overview, conclusions,
    plan_for_next_week,
    raw_data,
    get_cell_by_value,
    final_report,
)


@pytest.fixture
def ozon_file():
    return OzonFile("test_file.xlsx")


def test_get_column_index_by_cell_value(ozon_file):
    sheet = ozon_file.report_wb.worksheets[0]
    cell_value = "Выводы"
    assert get_column_index_by_cell_value(sheet, cell_value) == 1


def test_overview(ozon_file):
    expected_output = "Высылаем отчет за 24.04-30.04\n\n\n" \
                      "🛒 Заказано в шт 2 000 на сумму 500 000 руб.\n" \
                      "✅ Продажи в шт 200 на сумму 200 034 руб.\n" \
                      "Средняя цена продажи 500.00 руб.\n" \
                      "CTR 3.84%\n" \
                      "ДРР 10.45%\n"
    assert overview(ozon_file) == expected_output


def test_conclusions(ozon_file):
    expected_output = "❕Мы хорошо поработали\n" \
                      "❕Отличный рост заказов\n" \
                      "❕Надо всем дать премию"
    assert conclusions(ozon_file) == expected_output


def test_plan_for_next_week(ozon_file):
    expected_output = "\n✅ Задачи на неделю:\n\n" \
                      "1. Проработать работу:    регулярно\n" \
                      "2. Запустить тест:    регулярно\n" \
                      "3. Бла бла бла:    в работе\n" \
                      "4. тест тест тест:    регулярно"
    assert plan_for_next_week(ozon_file) == expected_output


def test_raw_data(ozon_file):
    expected_output = "Sheet name: svod_shop\n\n" \
                      "Заказано товаров, шт.:  2 000\n" \
                      "Заказано на сумму, руб.:  500 000\n" \
                      "Продажи, шт.:  200\n" \
                      "Продажи, руб.:  200 034\n" \
                      "Средняя цена продажи, руб.:  500.00\n" \
                      "CTR, %:  3.84\n" \
                      "ДРР, %:  10.45"
    assert raw_data(ozon_file) == expected_output


def test_get_cell_by_value(ozon_file):
    cell_value = "Заказано товаров, шт."
    assert get_cell_by_value(ozon_file, cell_value) == "A5"


def test_final_report(ozon_file):
    expected_output = "Высылаем отчет за 24.04-30.04\n\n\n" \
                      "🛒 Заказано в шт 2 000 на сумму 500 000 руб.\n" \
                      "✅ Продажи в шт 200 на сумму 200 034 руб.\n" \
                      "Средняя цена продажи 500.00 руб.\nCTR 3.84%\n" \
                      "ДРР 10.45%\n\n" \
                      "❕Мы хорошо поработали\n" \
                      "❕Отличный рост заказов\n" \
                      "❕Надо всем дать премию\n\n" \
                      "✅ Задачи на неделю:\n\n" \
                      "1. Проработать работу:    регулярно\n" \
                      "2. Запустить тест:    регулярно\n" \
                      "3. Бла бла бла:    в работе\n" \
                      "4. тест тест тест:    регулярно"
    assert final_report(ozon_file) == expected_output
