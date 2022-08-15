import sqlite3
import pandas as pd
import time
import random as rand
from sqlite3 import Error
import os

import main


def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print(sqlite3.version)
    except Error as e:
        print(e)

    return conn


conn = create_connection('test.db')
c = conn.cursor()


def trans_mod_conver_cal(prior, current):
    sheet = "PATransModConverCal_Ext"
    columns = ["`Transition Category`", "`Years Since Acquisition`", "`Lower Bound`", "`Trans Mod Level`", "Factor"]
    prior_sheet = pd.read_excel(prior, sheet_name="PATransModConverCal_Ext", header=9)
    current_sheet = pd.read_excel(current, sheet_name="PATransModConverCal_Ext", header=9)
    start_time = time.time()
    random_num = rand.randrange(1, 999999999)
    random_num2 = rand.randrange(1, 999999999)
    csv_name = 'Comparison' + sheet + str(random_num) + '.csv'
    prior_table_name = "prior" + str(random_num)
    current_table_name = "current" + str(random_num)
    table_name = "Table" + str(random_num)
    table_name2 = "Table" + str(random_num2)
    prior_sheet.to_sql(prior_table_name, conn)
    current_sheet.to_sql(current_table_name, conn)

    c.execute("create table if not exists " + table_name + """(`Transition Category` TEXT, 
    `Years Since Acquisition` INTEGER, `Lower Bound` REAL, Factor Real, `Trans Mod Level` Text, `Prior Factor` REAL, 
    `Percent Change` REAL, Change REAL)""")
    conn.commit()

    column_data = []

    for column in columns:
        c.execute("select " + column + " from " + current_table_name)
        data = c.fetchall()
        column_data.append(data)

    factors = column_data[-1]
    column_data.remove(column_data[-1])
    column_data.append(factors)

    string = "INSERT INTO " + table_name + " ("
    for column in columns:
        string = string + column + ", "
    string = string + "`Prior Factor`, `Percent Change`, `Change`) " + "VALUES ("

    j = 0
    tuple1 = ()
    string2 = string
    for i in range(len(column_data[0])):
        while j < len(column_data):
            if j == len(column_data) - 1:
                tuple1 += column_data[j][i]
                for item in tuple1:
                    if item == tuple1[3]:
                        no_commas = str(item).replace(",", "\,")
                        # no_commas = no_commas.replace("+","plus")
                        # no_commas = no_commas.replace("-","minus")
                        # no_commas = no_commas.replace("_", "U")
                        no_spaces = no_commas.replace(" ", "'")
                        no_spaces = '"' + no_spaces + '"'
                        string2 += no_spaces + ", "
                    else:
                        no_commas = str(item).replace(",", "\,")
                        # no_commas = no_commas.replace("+","plus")
                        # no_commas = no_commas.replace("-","minus")
                        no_spaces = no_commas.replace(" ", "'")
                        no_spaces = '"' + no_spaces + '"'
                        string2 += no_spaces + ", "

                string2 += "NULL, NULL, NULL)"
                print(string2)
                c.execute(string2)
                conn.commit()
                string2 = string
                tuple1 = ()
                j += 1

            else:
                tuple1 += column_data[j][i]
                j += 1
        j = 0
    c.execute("SELECT * from " + table_name)
    c.execute("create table if not exists " + table_name2 + " (`Prior Factor` Real, `Trans Mod Level` Text)")
    c.execute("select Factor, `Trans Mod Level` from " + prior_table_name)
    prior_factors = c.fetchall()
    for i in range(len(prior_factors)):
        fix_string = str(prior_factors[i][1])
        # fix_string = fix_string.replace("_", "U")
        # fix_string = fix_string.replace("-", "minus")
        # fix_string = fix_string.replace("+", "plus")
        string = "update " + table_name + " set `Prior Factor`= " + str(prior_factors[i][0]) + " where `Trans Mod Level` = " + "'" + fix_string + "'"
        c.execute(string)
        conn.commit()

    c.execute("select Factor, `Prior Factor` from " + table_name)
    changes = c.fetchall()
    for i in range(len(changes)):
        string = "update " + table_name + " set `Percent Change` = "
        try:
            per_change = changes[i][0]/changes[i][1] -1

        except (ZeroDivisionError, TypeError):
            if changes[i][0] == changes[i][1]:
                per_change = 0
            else:
                per_change = "NULL"
        string += str(per_change)
        print(string)
        c.execute(string)

    for i in range(len(changes)):
        string = "update " + table_name + " set `Change` = "
        raw_change = changes[i][0] - changes[i][1]
        string += str(raw_change)
        c.execute(string)
        conn.commit()

    c.execute("select * from " + table_name)

    c.execute("select Factor from " + current_table_name)
    current_length = c.fetchall()
    c.execute("select Factor from " + prior_table_name)
    prior_length = c.fetchall()
    if len(current_length) - len(prior_length) == 0:
        pass
    else:
        c.execute("ALTER TABLE " + table_name + " ADD COLUMN `Change in Length`")
        conn.commit()
        c.execute("UPDATE " + table_name + " SET `Change in Length` = " + str(
            len(current_length) - len(prior_length)) + " WHERE `index` = 0")

    sql = "SELECT * FROM " + table_name
    df = pd.read_sql_query(sql, conn)
    df.to_csv(csv_name)

    read_file = pd.read_csv(r'Comparison' + sheet + str(random_num) + '.csv')
    read_file.to_excel(r'Comparison' + sheet + str(random_num) + '.xlsx')
    os.remove('Comparison' + sheet + str(random_num) + '.csv')
    end_time = time.time()
    print(end_time - start_time)
    print(random_num)


if __name__ == "__main__":
    pri = "Homeowner WV Conversion.xlsx"
    cur = "Homeowner WV Conversion.xlsx"

    trans_mod_conver_cal(pri, cur)
    main.create_connection("test.db")