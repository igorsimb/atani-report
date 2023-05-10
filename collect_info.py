from data_locators import OzonFile, WbFile

current_marketplace = ""


def get_column_index_by_cell_value(sheet, cell_value):
    result_column = 0
    for column in sheet.iter_rows(min_row=1, max_row=1, max_col=10):
        for cell in column:
            if str(cell.value).lower() == cell_value.lower():
                result_column = cell.column
                break
        else:
            continue
        break
    return result_column


def overview(file):
    overview = f"Высылаем отчет за {file.date_range}\n" \
               f"{current_marketplace.upper()}\n\n" \
               f"🛒 Заказано в шт {file.ordered_items_value} на сумму {file.ordered_sum_value} руб.\n" \
               f"✅ Продажи в шт {file.sold_items_value} на сумму {file.sold_sum_value} руб.\n" \
               f"Средняя цена продажи {file.ave_sum_value} руб.\n" \
               f"CTR {file.ctr_value}%\n" \
               f"ДРР {file.drr_value}%\n"
    return overview


def conclusions(file):
    # ВЫВОДЫ
    conclusions_sheet = file.report_wb.worksheets[0]

    conclusions_column = get_column_index_by_cell_value(sheet=conclusions_sheet, cell_value="Выводы")
    conc_text = conclusions_sheet.iter_rows(min_row=2, min_col=conclusions_column, max_col=conclusions_column)

    conc_text_list = []

    # create a list of values from conc_text generator
    for row in conc_text:
        conc_text_list.extend(f"❕{cell.value}" for cell in row if cell.value is not None)

    # turn list into string with separate lines
    return "\n".join(conc_text_list)


def plan_for_next_week(file):
    # ПЛАН
    plan_sheet = file.report_wb.worksheets[1]
    plan_text = plan_sheet.iter_rows(min_row=2, min_col=2, max_col=2)  # column B
    plan_status = plan_sheet.iter_rows(min_row=2, min_col=3, max_col=3)  # column C

    # create lists from plan_text and plan_status generators
    plan_text_list = []
    plan_status_list = []

    for index, row in enumerate(plan_text, start=1):
        plan_text_list.extend(cell.value for cell in row if cell.value is not None)
    for index, row in enumerate(plan_status, start=1):
        plan_status_list.extend(cell.value for cell in row if cell.value is not None)

    # join the lists into one and enumerate the items
    plan_list = ["\n✅ Задачи на неделю:\n"]
    for index, value in enumerate(zip(plan_text_list, plan_status_list), start=1):
        plan_list.append(f"{index}. {value[0]}:    {value[1]}")

    # turn list into string with separate lines
    return "\n".join(plan_list)


def raw_data(file):
    if current_marketplace == "ozon":
        file = OzonFile(file)
    elif current_marketplace == "wb":
        file = WbFile(file)

    log_list = [
        f"Sheet name: {file.svod_shop.title}\n",
        f"{file.ordered_items_text}:  {file.ordered_items_value}",
        f"{file.ordered_sum_text}:  {file.ordered_sum_value}",
        f"{file.sold_items_text}:  {file.sold_items_value}",
        f"{file.sold_sum_text}:  {file.sold_sum_value}",
        f"{file.ave_sum_text}:  {file.ave_sum_value}",
        f"{file.ctr_text}:  {file.ctr_value}",
        f"{file.drr_text}:  {file.drr_value}",
    ]
    return "\n".join(log_list)


def get_cell_by_value(file, cell_value):
    if current_marketplace == "ozon":
        file = OzonFile(file)
    elif current_marketplace == "wb":
        file = WbFile(file)
    result_cell = ""
    for column in file.svod_shop.iter_rows(min_row=1, max_row=50, max_col=30):
        for cell in column:
            if str(cell.value).lower() == cell_value.lower():
                result_cell = cell.coordinate
                break
        else:
            continue
        break
    return result_cell


def final_report(file):
    # instantiating class File so that generate_report function only needs the file path value
    if current_marketplace == "ozon":
        file = OzonFile(file)
    elif current_marketplace == "wb":
        file = WbFile(file)
    return f"{overview(file)}\n{conclusions(file)}\n{plan_for_next_week(file)}"
