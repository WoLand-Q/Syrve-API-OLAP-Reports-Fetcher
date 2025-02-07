import requests
import json
import pandas as pd
import textwrap

# Настройки
API_KEY = ''  # Замените на ваш ТОКЕН ПОСЛЕ АВТОРИЗАЦИИ
BASE_URL = '/resto/api/v2'

def get_olap_columns(report_type='SALES'):
    url = f"{BASE_URL}/reports/olap/columns"
    params = {
        'key': API_KEY,
        'reportType': report_type
    }
    headers = {
        'Content-Type': 'application/json; charset=utf-8'
    }

    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Ошибка при получении полей отчета: {response.status_code}")
        print(response.text)
        return None

def get_olap_report(report_type='SALES', build_summary=True,
                    group_by_row_fields=None, group_by_col_fields=None,
                    aggregate_fields=None, filters=None):
    url = f"{BASE_URL}/reports/olap"
    params = {'key': API_KEY}
    headers = {'Content-Type': 'application/json; charset=utf-8'}

    if group_by_row_fields is None:
        group_by_row_fields = []
    if group_by_col_fields is None:
        group_by_col_fields = []
    if aggregate_fields is None:
        aggregate_fields = []
    if filters is None:
        filters = {}

    request_body = {
        "reportType": report_type,
        "buildSummary": str(build_summary).lower(),
        "groupByRowFields": group_by_row_fields,
        "groupByColFields": group_by_col_fields,
        "aggregateFields": aggregate_fields,
        "filters": filters
    }

    response = requests.post(url, params=params, headers=headers,
                             data=json.dumps(request_body, ensure_ascii=False))

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Ошибка при получении отчета: {response.status_code}")
        print(response.text)
        return None

def main():
    # Получение доступных полей
    columns = get_olap_columns()
    if not columns:
        return

    # Настройка отчета
    report_type = 'SALES'
    build_summary = True
    group_by_row_fields = [
        'Delivery.CustomerPhone',
        'Delivery.CustomerCreatedDateTyped',
        'Delivery.CustomerName',
        'Delivery.Email',
        'Delivery.CustomerComment',
        'OpenDate.Typed',
        'ExternalNumber'
    ]
    group_by_col_fields = []
    aggregate_fields = [
        'GuestNum',
        'DishSumInt',
        'DishDiscountSumInt',
        'UniqOrderId'
    ]

    # Настройка фильтров
    filters = {
        "OpenDate.Typed": {
            "filterType": "DateRange",
            "periodType": "CUSTOM",
            "from": "2023-10-01",
            "to": "2023-11-01",
            "includeLow": True,
            "includeHigh": False
        },
        "DeletedWithWriteoff": {
            "filterType": "IncludeValues",
            "values": ["NOT_DELETED"]
        },
        "OrderDeleted": {
            "filterType": "IncludeValues",
            "values": ["NOT_DELETED"]
        }
    }

    # Получение отчета
    report = get_olap_report(
        report_type=report_type,
        build_summary=build_summary,
        group_by_row_fields=group_by_row_fields,
        group_by_col_fields=group_by_col_fields,
        aggregate_fields=aggregate_fields,
        filters=filters
    )

    if report:
        data = report.get('data', [])
        summary = report.get('summary', [])

        if data:
            # Попытка создать DataFrame напрямую
            try:
                df = pd.DataFrame(data)
            except ValueError as e:
                print(f"Ошибка при создании DataFrame напрямую: {e}")
                # Плоское преобразование данных
                flat_data = []
                for index, item in enumerate(data):
                    combined_dict = {}
                    if isinstance(item, dict):
                        combined_dict.update(item)
                    elif isinstance(item, list):
                        for sub_dict in item:
                            if isinstance(sub_dict, dict):
                                combined_dict.update(sub_dict)
                            else:
                                print(f"Неожиданный тип sub_dict в data[{index}]: {type(sub_dict)}")
                    else:
                        print(f"Неожиданный тип item в data[{index}]: {type(item)}")
                    flat_data.append(combined_dict)
                df = pd.DataFrame(flat_data)

            # Переименование столбцов
            rename_dict = {
                'Delivery.CustomerPhone': 'Телефон клиента',
                'Delivery.CustomerCreatedDateTyped': 'Дата регистрации клиента',
                'Delivery.CustomerName': 'Имя клиента',
                'Delivery.Email': 'Email',
                'Delivery.CustomerComment': 'Комментарий клиента',
                'OpenDate.Typed': 'Дата заказа',
                'ExternalNumber': 'Внешний номер заказа',  # Добавлено новое поле
                'GuestNum': 'Количество гостей',
                'DishSumInt': 'Сумма заказа',
                'DishDiscountSumInt': 'Сумма со скидкой',
                'UniqOrderId': 'Уникальный ID заказа'
            }
            df.rename(columns=rename_dict, inplace=True)

            # Заполняем пропущенные значения пустыми строками для лучшего отображения
            df.fillna('', inplace=True)

            # Настройки отображения
            pd.set_option('display.max_columns', None)
            pd.set_option('display.width', 1000)
            pd.set_option('display.max_colwidth', 50)
            pd.set_option('display.colheader_justify', 'left')
            pd.set_option('display.unicode.east_asian_width', True)
            pd.set_option('display.expand_frame_repr', False)

            # Перенос текста в длинных колонках
            wrap_cols = ['Комментарий клиента', 'Имя клиента', 'Email']
            for col in wrap_cols:
                df[col] = df[col].apply(lambda x: '\n'.join(textwrap.wrap(x, width=30)))

            # Вывод таблицы со всеми столбцами
            print("\n===== Отформатированные данные отчета =====")
            print(df.to_string(index=False))

            # Сохранение в файл Excel с автонастройкой ширины колонок
            with pd.ExcelWriter('olap_report_sales.xlsx', engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Отчет', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Отчет']

                # Автонастройка ширины колонок
                for idx, col in enumerate(df.columns):
                    max_width = df[col].astype(str).apply(len).max()
                    max_width = min(max_width, 50)
                    worksheet.set_column(idx, idx, max_width + 2)

            print("\nОтчет сохранен в 'olap_report_sales.xlsx'")
        else:
            print("\nДанные отчета отсутствуют.")

        if summary:
            print("\n===== Итоги отчета =====")
            summary_df = pd.DataFrame([summary])
            print(summary_df.to_string(index=False))
        else:
            print("\nИтоги отчета отсутствуют.")
    else:
        print("Не удалось получить отчет.")

if __name__ == "__main__":
    main()
