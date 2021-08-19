import pandas as pd
import os
import sqlite3

#read data sheet in xlsx file
def read_xlsx(sales_file):
    data = pd.read_excel(sales_file, sheet_name="data")

    df = pd.DataFrame(data, columns = ['Магазин', 'Город', 'Дата', 'Колличество', 'Бренд', 'Класс'])
    df_db = [(i['Магазин'], i['Город'], i['Дата'], i['Колличество'], i['Бренд'], i['Класс']) for index, i in df.iterrows()]

    return(df_db)

#create db file
def create_db_file(df_db):
    #check if db file exists
    if os.path.exists("braki.db"):
        os.remove("braki.db")
    else:
        print("file does not exist")

    con = sqlite3.connect("braki.db")
    cur = con.cursor()
    cur.execute(
        """
        CREATE TABLE braki (
            Shop TEXT,
            City TEXT,
            Date TIMESTAMP,
            Quantity REAL,
            Brand TEXT,
            Class TEXT
        );
        """
    )

    cur.executemany(f"INSERT INTO braki (Shop, City, Date, Quantity, Brand, Class) VALUES (?, ?, ? ,?, ?, ?)", df_db)
    
    con.commit()
    con.close()

#main
def main():
    sales_file = "braki.xlsx"
    data = read_xlsx(sales_file)
    create_db_file(data)


if __name__ == '__main__':
    main()