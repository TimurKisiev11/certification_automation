import openpyxl


def read_from_file(file_path, sheet_name, start_row, end_row, column_index):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        values_list = []
        for row in range(start_row, end_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            values_list.append(cell.value)
        return values_list
    except Exception as e:
        print(f"Error while reading from file: {e}")
        return []


def convert_to_float(values_list):
    try:
        float_values_list = []
        for value in values_list:
            if value is not None:
                try:
                    float_value = float(value)
                    float_values_list.append(float_value)
                except Exception as e:
                    if value == "Запрашивает Даша Информацию от РО" or value == "":
                        float_values_list.append(0.0)
                    else:
                        print(f"Error ONE while converting to float: {e}")
                        return []
            else:
                float_values_list.append(0.0)
        return list(map(lambda x: round(x, 1), float_values_list))
    except Exception as e:
        print(f"Error TWO while converting to float: {e}")
        return []


def compare(average, target, confidence_level):
    coincidences = 0
    for (avr, tg) in zip(average, target):
        if tg <= avr:
            coincidences += 1
    if coincidences / len(target) >= confidence_level:
        return True
    else:
        return False


def certificate(file_path_1, sheet_name_1, file_path_2, sheet_name_2, ni_fi=False):
    test_result = {}
    self_esteem = convert_to_float(read_from_file(file_path_1, sheet_name_1, 11, 26, 9))
    lead_esteem = convert_to_float(read_from_file(file_path_1, sheet_name_1, 11, 26, 10))
    average = list(map(lambda x, y: round((x + y) / 2, 1), self_esteem, lead_esteem))
    if self_esteem and lead_esteem:
        for i in range(3, 10):
            target_scores = convert_to_float(read_from_file(file_path_2, sheet_name_2, 2, 18, i))
            if (ni_fi):
                target_scores.pop(10)
            else:
                target_scores.pop(11)
            prof_level = (read_from_file(file_path_2, sheet_name_2, 1, 1, i))[0]
            test_result.update({prof_level: compare(average, target_scores, 0.8)})
    print(test_result)
    return test_result


file_path = 'Оценка_DE_сам.оценка+TL оценка.xlsx'
sheet_name = 'Оценка компетенций'
file_path_2 = 'Компетенции_по_шкале_DE.xlsx'
sheet_name_2 = 'Целевые значения'
certificate(file_path, sheet_name, file_path_2, sheet_name_2, True)
