import pandas as pd
import os
import sqlite3

#reads xlsx alters column values adds them to df, and returns
def read_xlsx(sales_file):

    data_current = pd.read_excel(sales_file, sheet_name = "current")
    data_previous = pd.read_excel(sales_file, sheet_name = "previous")

    list_city_current = []
    list_storage_current = []
    for i in data_current['Склад']:
        if 'Ташкентская' in i:
            list_city_current.append('Алматы')
            if 'дисконт' in i:
                list_storage_current.append("Ташкентская 496В (дисконт)")
            else:
                list_storage_current.append("Ташкентская 496В")
        if 'Рыскулова' in i:
            list_city_current.append('Алматы')
            if 'дисконт' in i:
                list_storage_current.append("Рыскулова 232 (дисконт)")
            else:
                list_storage_current.append("Рыскулова 232")
        if 'Town' in i:
            list_city_current.append('Алматы')
            if 'дисконт' in i:
                list_storage_current.append("КарТаун (дисконт)")
            else:
                list_storage_current.append("КарТаун")
        if 'Шемякина' in i:
            list_city_current.append('Алматы')
            if 'дисконт' in i:
                list_storage_current.append("Шемякина (дисконт)")
            else:
                list_storage_current.append("Шемякина")
        if 'Ломова' in i:
            list_city_current.append('Павлодар')
            if 'дисконт' in i:
                list_storage_current.append("Ломова 162 (дисконт)")
            else:
                list_storage_current.append("Ломова 162")
        if 'Естая' in i:
            list_city_current.append('Павлодар')
            if 'дисконт' in i:
                list_storage_current.append("Естая 83 (дисконт)")
            else:
                list_storage_current.append("Естая 83")
        if 'Макатаева' in i:
            list_city_current.append('Алматы')
            if 'дисконт' in i:
                list_storage_current.append("Макатаева 127 (дисконт)")
            else:
                list_storage_current.append("Макатаева 127")
        if 'Нурмаганбетова' in i:
            list_city_current.append('Павлодар')
            if 'дисконт' in i:
                list_storage_current.append("Нурмаганбетова (дисконт)")
            else:
                list_storage_current.append("Нурмаганбетова")

    data_current['Город'] = list_city_current
    data_current['Склад'] = list_storage_current

    list_city_previous = []
    list_storage_previous = []
    for i in data_previous['Склад']:
        if 'Ташкентская' in i:
            list_city_previous.append('Алматы')
            if 'дисконт' in i:
                list_storage_previous.append("Ташкентская 496В (дисконт)")
            else:
                list_storage_previous.append("Ташкентская 496В")
        if 'Рыскулова' in i:
            list_city_previous.append('Алматы')
            if 'дисконт' in i:
                list_storage_previous.append("Рыскулова 232 (дисконт)")
            else:
                list_storage_previous.append("Рыскулова 232")
        if 'Town' in i:
            list_city_previous.append('Алматы')
            if 'дисконт' in i:
                list_storage_previous.append("КарТаун (дисконт)")
            else:
                list_storage_previous.append("КарТаун")
        if 'Шемякина' in i:
            list_city_previous.append('Алматы')
            if 'дисконт' in i:
                list_storage_previous.append("Шемякина (дисконт)")
            else:
                list_storage_previous.append("Шемякина")
        if 'Ломова' in i:
            list_city_previous.append('Павлодар')
            if 'дисконт' in i:
                list_storage_previous.append("Ломова 162 (дисконт)")
            else:
                list_storage_previous.append("Ломова 162")
        if 'Естая' in i:
            list_city_previous.append('Павлодар')
            if 'дисконт' in i:
                list_storage_previous.append("Естая 83 (дисконт)")
            else:
                list_storage_previous.append("Естая 83")
        if 'Макатаева' in i:
            list_city_previous.append('Алматы')
            if 'дисконт' in i:
                list_storage_previous.append("Макатаева 127 (дисконт)")
            else:
                list_storage_previous.append("Макатаева 127")
        if 'Нурмаганбетова' in i:
            list_city_previous.append('Павлодар')
            if 'дисконт' in i:
                list_storage_previous.append("Нурмаганбетова (дисконт)")
            else:
                list_storage_previous.append("Нурмаганбетова")

    data_previous['Город'] = list_city_previous
    data_previous['Склад'] = list_storage_previous

    df_current = pd.DataFrame(data_current, columns = ['Город', 'Склад', 'Номенклатура', 'Группа', 'Применимость', 'Моноблок', 'Менеджер', 'Количество', 'Сумма', 'Дата'])
    df_current_db = [(i['Город'], i['Склад'], i['Номенклатура'], i['Группа'], i['Применимость'], i['Моноблок'], i['Менеджер'], i['Количество'], i['Сумма'], i['Дата']) for index, i in df_current.iterrows()]

    df_previous = pd.DataFrame(data_previous, columns = ['Город', 'Склад', 'Номенклатура', 'Группа', 'Применимость', 'Моноблок', 'Менеджер', 'Количество', 'Сумма', 'Дата'])
    df_previous_db = [(i['Город'], i['Склад'], i['Номенклатура'], i['Группа'], i['Применимость'], i['Моноблок'], i['Менеджер'], i['Количество'], i['Сумма'], i['Дата']) for index, i in data_previous.iterrows()]

    return (df_current_db, df_previous_db)

def create_db_file(df_current_db, df_previous_db):
    #check if db file exists
    if os.path.exists("sales.db"):
        os.remove("sales.db")
    else:
        print("file does not exist")

    con = sqlite3.connect("sales.db")
    cur = con.cursor()
    cur.execute(
        """
        CREATE TABLE current (
            City TEXT,
            Storage TEXT,
            Product TEXT,
            Brand TEXT,
            Vehicle TEXT,
            Bodytype TEXT,
            Manager TEXT,
            Quantity REAL,
            Price REAL,
            Date TEXT
        );
        """
    )

    cur.executemany(f"INSERT INTO current (City, Storage, Product, Brand, Vehicle, Bodytype, Manager, Quantity, Price, Date) VALUES (?, ?, ? ,?, ?, ?, ?, ?, ?, ?)", df_current_db)
    
    cur.execute(
        """
        CREATE TABLE previous (
            City TEXT,
            Storage TEXT,
            Product TEXT,
            Brand TEXT,
            Vehicle TEXT,
            Bodytype TEXT,
            Manager TEXT,
            Quantity REAL,
            Price REAL,
            Date TEXT
        );
        """
    )

    cur.executemany(f"INSERT INTO previous (City, Storage, Product, Brand, Vehicle, Bodytype, Manager, Quantity, Price, Date) VALUES (?, ?, ? ,?, ?, ?, ?, ?, ?, ?)", df_previous_db)
    
    con.commit()
    con.close()

def main():
    sales_file = "sales.xlsx"

    data = read_xlsx(sales_file)
    
    create_db_file(data[0], data[1])

if __name__ ==  '__main__':
    main()