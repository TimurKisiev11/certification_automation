import openpyxl


def read_and_convert_to_float(file_path, sheet_name, start_row, end_row, column_index):
    try:
        values_list = []
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        for row in range(start_row, end_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value is not None:
                try:
                    float_value = float(cell.value)
                    values_list.append(float_value)
                except ValueError:
                    if cell.value == "Запрашивает Даша Информацию от РО" or cell.value == "":
                        values_list.append(0.0)
                        print(f"Value at row {row}, column {column_index} is not convertible to float")
            else:
                values_list.append(0.0)
        return list(map(lambda x: round(x,1), values_list))
    except Exception as e:
        print(f"An error occurred: {e}")
        return []



file_path_1 = 'Оценка_DE_сам.оценка+TL оценка.xlsx'
sheet_name_1 = 'Оценка компетенций'
start_row_1 = 11
end_row_1 = 30
column_index_1 = 10
float_values_1 = read_and_convert_to_float(file_path_1, sheet_name_1, start_row_1, end_row_1, column_index_1)
print(list(map(lambda x: round(x,1), float_values_1)))

file_path_2 = 'Компетенции_по_шкале_DE.xlsx'
sheet_name_2 = 'Целевые значения'
start_row_2 = 2
end_row_2 = 22
column_index_2 = 8
float_values_2 = read_and_convert_to_float(file_path_2, sheet_name_2, start_row_2, end_row_2, column_index_2)
print(float_values_2)



