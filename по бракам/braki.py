import sqlite3
from numpy import quantile
from openpyxl import Workbook, load_workbook
import datetime as dt

import xlsx_to_db

file_path = "report.xlsx"

shops = []
cities = []
brands = []
qualities = []

#months = AND Date BETWEEN "2020-08-01" AND "2020-08-20"

def connect_to_db(db_file):
    conn = 0
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)

    return conn

def shop(conn):
    cur = conn.cursor()
    cur.execute('SELECT DISTINCT Shop FROM braki')
    shops_raw = cur.fetchall()
    
    for i in shops_raw:
        shops.append(i[0])

    print(shops)

def city(conn):
    cur = conn.cursor()
    cur.execute('SELECT DISTINCT City FROM braki')
    cities_raw = cur.fetchall()
    
    for i in cities_raw:
        cities.append(i[0])

    print(cities)

def brand(conn):
    cur = conn.cursor()
    cur.execute('SELECT DISTINCT Brand FROM braki')
    brands_raw = cur.fetchall()
    
    for i in brands_raw:
        brands.append(i[0])

    print(brands)

def quality(conn):
    cur = conn.cursor()
    cur.execute('SELECT DISTINCT Class FROM braki')
    qualities_raw = cur.fetchall()
    
    for i in qualities_raw:
        qualities.append(i[0])

    print(qualities)

def by_shop(conn):
    print("++++++++++++++++++++++++++++")
    cur = conn.cursor()
    for i in shops:
        cur.execute(f'SELECT SUM(Quantity) FROM braki WHERE Shop = "{i}"')

        data = cur.fetchall()
        quantity = data[0][0] if data[0][0] != None else 0
        print(i, quantity)
    print("============================")

def by_city(conn):
    print("++++++++++++++++++++++++++++")
    cur = conn.cursor()
    for i in cities:
        cur.execute(f'SELECT SUM(Quantity) FROM braki WHERE City = "{i}"')

        data = cur.fetchall()
        quantity = data[0][0] if data[0][0] != None else 0
        print(i, quantity)
    print("============================")

def by_brand(conn):
    print("++++++++++++++++++++++++++++")
    cur = conn.cursor()
    for i in brands:
        cur.execute(f'SELECT SUM(Quantity) FROM braki WHERE Brand = "{i}"')

        data = cur.fetchall()
        quantity = data[0][0] if data[0][0] != None else 0
        print(i, quantity)
    print("============================")

def by_brand_quality(conn):
    print("++++++++++++++++++++++++++++")
    cur = conn.cursor()
    for i in brands:
        for j in qualities:
            cur.execute(f'SELECT SUM(Quantity) FROM braki WHERE Brand = "{i}" AND Class = "{j}"')

            data = cur.fetchall()
            quantity = data[0][0] if data[0][0] != None else 0
            if quantity != 0:
                print(i, j, quantity)
    print("============================")

def by_period(conn):
    print("++++++++++++++++++++++++++++")
    cur = conn.cursor()
    for year in range(0, 1):
        for month in range(1, 13):
            for day in range(1, 32):
                a = month if month > 9 else f"0{month}"
                b = day if day > 9 else f"0{day}"
                print(f"202{year}-{a}-{b}")
    print("============================")

def main():
    database = "braki.db"
    conn = connect_to_db(database)

    with conn:
        shop(conn)
        city(conn)
        brand(conn)
        quality(conn)
        by_shop(conn)
        by_city(conn)
        by_brand(conn)
        by_brand_quality(conn)
        #by_period(conn)

if __name__ == '__main__':
    main()