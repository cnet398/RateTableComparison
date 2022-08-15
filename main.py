# Package imports
import sqlite3
import pandas as pd
import time
import random as rand
from sqlite3 import Error
import os
import xlsxwriter.exceptions

# Script imports
import PAtransMod
import PAtransModConverCalc
import excelConsolidation
import transMod
import transModConverCalc

"""Creates the connection to the sqlite database"""


def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        print(sqlite3.version)
    except Error as e:
        print(e)

    return conn


"""Iterates the list of sheet names and runs the main function for each sheet that has a matching name"""

def all_sheets():
    for x in current_sheets_factor:  # Uses the list of sheets that begin with the abbreviation for the coverage type
        if x == "HOTransMod_Ext" or x == "HOTransModConverCal_Ext" or x == "PATransMod_Ext" or x == "PATransModConverCal_Ext":
            current_sheets_factor.remove(x)  # Make the trans mod tables run last
            current_sheets_factor.append(x)
    for x in current_sheets_factor:
        for y in prior_sheets_factor:
            if x == y:  # Takes each sheet in current and finds it's match in prior
                main(x)  # Run the sheets with matching names
            else:
                pass



"""
Makes a string based off the columns retrieved in the data_entry function it then uses this string to create the 
table that eventually gets output. Prior Factor, Percent Change, and Change columns are hard coded and need to remain, 
the rest of the columns are done dynamically based off the columns in the current rate table
"""


def create_table(table, current):
    columns = data_entry(current)  # Retrieves the list of formatted column names from the data_entry function
    string = "create table if not exists " + table + " (`Key` TEXT, "  # Initializes the string that will create an SQL Table
    for column in columns:
        c.execute("select " + column + " from " + current + " limit 1")
        values = c.fetchall()  # Retrieves the first value of every column

        # Appends the string with the name of the column and the data type of the first value in the column
        if type(values[0][0]) is str:
            string += column + " TEXT,"

        elif type(values[0][0]) is int:
            string += column + " INTEGER,"

        elif type(values[0][0]) is float:
            string += column + " REAL,"

        elif type(values[0][0]) is bool:
            string += column + " BIT,"

        else:
            string += column + " TEXT,"
    string += " `Prior Factor` REAL, `Percent Change` REAL, `Change` REAL)"  # Appends the string with the last columns

    c.execute(string)  # Runs the string as an SQL query and creates the table
    conn.commit()  # This is just essentially hitting the save button


"""
Retrieves the columns from the current table and appends them to a list
"""


def data_entry(current):
    data = c.execute("select * from " + current)
    columns = []
    # .description retrieves a list of the columns in the table
    for item in data.description:
        columns.append(item[0])
    # The tildes are required for columns names that are not correctly formatted for sql, the list was always in
    # backwards order, this is why the columns.reverse() is there.
    for item in columns:
        n = 0
        columns.remove(item)
        item = "`" + item + "`"
        columns.insert(n, item)
        n += 1
    columns.reverse()

    return columns


"""
Creates another string that is used to insert the data one row at a time
"""


def create_query(columns, table, current):
    string = "INSERT INTO " + table + " ("
    # not sure why I did this one differently, but it creates a query based on the columns in the current file
    # The columns are taken as an argument, but they also come from the data_entry function
    for column in columns:
        string = string + column + ", "
    string = string + "`Prior Factor`, `Percent Change`, `Change`) " + "VALUES ("
    column_data = []
    # this generates a list of lists of the data in each column the data in sql is stored as tuples, this is why
    # there's a bit of reformatting being done
    for column in columns:
        c.execute("select " + column + " from " + current)
        data = c.fetchall()
        column_data.append(data)
    string2 = string
    j = 0
    tuple1 = ()
    # this corrects the data for any formatting issues such as commas or spaces. It then executes the insert function in
    # sql, the NULL values are there to be updated later on.
    for i in range(len(column_data[0])):
        while j < len(column_data):
            if j == len(column_data)-1:
                tuple1 += column_data[j][i]
                for item in tuple1:
                    no_commas = str(item).replace(",", "")
                    no_spaces = no_commas.replace(" ", "")
                    no_spaces = '"' + no_spaces + '"'
                    string2 += no_spaces + ", "
                string2 += "NULL, NULL, NULL)"
                print(string2)  # This print statement is unnecessary, it's only here so the user knows the program is running correctly.
                                # Feel free to remove.
                c.execute(string2)  # Runs the string as an SQL query
                conn.commit()
                string2 = string  # Resets the string to make room for the next row of values
                tuple1 = ()
                j += 1

            else:
                tuple1 += column_data[j][i]
                j += 1
        j = 0
    return string


"""
This one is pretty self explanatory, it takes the factors from the prior table and inserts them into the generated
one by updating the dummy values in the Prior Factor column.
"""


def insert_prior_factor(table, prior, table2):
    keys = create_prior_key(prior, table2)
    c.execute("select Factor from " + prior)
    factors = c.fetchall()
    for i in range(len(factors)):
        c.execute("UPDATE " + table + " SET `Prior Factor` = " + str(factors[i][0]) + " WHERE `Key` = " + "'" + keys[i] + "'")
    c.execute("select * from " + table)


"""
Converts both the Factor and Prior Factor columns into lists and then finds the percent change and updates the dummy
values in the Percent Change column
"""


def percent_change(table):
    c.execute("select Factor from " + table)
    stuff = c.fetchall()  # Retrieves a list of all the factors from the current table
    factors = []  # Initializes empty lists to store factors
    prior_factors = []
    c.execute("select `Prior Factor` from " + table)
    things = c.fetchall()  # Retrieves a list of all the factors from the prior table
    for i in range(len(stuff)):
        factors.append(stuff[i][0])  # appends the actual numbers to the list of factors
                                     # SQL stores data as tuples and we need the number.
    for i in range(len(things)):
        prior_factors.append(things[i][0])
    delta = []
    for i in range(len(things)):
        # The try except is in case of a divide by zero error, if the factor started as 0 and remained 0 it will output
        # a 0, but if it changed from 0 it will output None for percent change.
        try:
            deltax = factors[i]/prior_factors[i]-1
            delta.append(round(deltax, 4))
        except (TypeError, ZeroDivisionError):
            if factors[i] == prior_factors[i]:
                deltax = 0.0
                delta.append(deltax)
            else:
                deltax = "NULL"
                delta.append(deltax)
    for i in range(len(delta)):
        dummy_string = "UPDATE " + table + " SET `Percent Change` = " + str(delta[i]) + " WHERE `index` = " + str(i)
        c.execute(dummy_string)
        print(dummy_string)  # Another unnecessary print statement, feel free to remove.


"""
Converts the Factor and Prior Factor columns into lists and finds the raw change between the two then updates the dummy
values in the Change column
"""


def raw_change(table):
    c.execute("select Factor from " + table)
    stuff = c.fetchall()  # Retrieves a list of factors from the current table
    factors = []  # Initializes an empty list
    prior_factors = []
    c.execute("select `Prior Factor` from " + table)
    things = c.fetchall()  # Retrieves the list of prior factors
    for i in range(len(stuff)):
        factors.append(stuff[i][0])  # Appends the list with the factor
    for i in range(len(things)):
        prior_factors.append(things[i][0])
    delta = []
    for i in range(len(things)):
        try:
            deltax = factors[i] - prior_factors[i]  # Does the calculation
            delta.append(round(deltax, 4))
        except TypeError:
            deltax = "NULL"
            delta.append(deltax)
    for i in range(len(delta)):
        c.execute("UPDATE " + table + " SET `Change` = " + str(delta[i]) + " WHERE `index` = " + str(i))


"""Creates a unique primary key for every row to allow for tables of differing sizes"""


def create_primary_key(current, table):
    c.execute("select `index` from " + table)
    length = c.fetchall()
    for i in range(len(length)):
        cat_string = ""
        string = "UPDATE " + table + " SET `KEY` = "
        c.execute("select * from " + current + " where `index` = " + str(i))
        concat = c.fetchall()
        print(concat)
        if type(concat[0][-1]) is not float:
            concat[0] = list(concat[0])
            concat[0][-1] = float(concat[0][-1])
            concat[0] = tuple(concat[0])
        for n in range(len(concat[0])):
            try:
                if concat[0][n+1] == concat[0][-1] and type(concat[0][n+1]) is float:
                    cat_string = "'" + cat_string + "'"
                    cat_string += " WHERE `index` = " + str(i)
                    string = string + cat_string
                    print(string)
                    c.execute(string)
                    conn.commit()
                elif concat[0][n+1] == concat[0][1]:
                    no_commas = str(concat[0][n + 1]).replace(",", "")
                    no_spaces = no_commas.replace(" ", "")
                    no_period = no_spaces.replace(".", "")
                    no_underscores = no_period.replace("_", "")
                    cat_string += no_underscores
                else:
                    no_commas = str(concat[0][n + 1]).replace(",", "")
                    no_spaces = no_commas.replace(" ", "")
                    no_period = no_spaces.replace(".", "")
                    no_underscores = no_period.replace("_", "")
                    cat_string += no_underscores
            except IndexError:
                pass


def create_prior_key(prior, table):
    c.execute("create table if not exists " + table + " (`index` INTEGER, Key Text, Factor Real)")
    conn.commit()
    c.execute("select Factor from " + prior)
    factors = c.fetchall()
    for i in range(len(factors)):
        c.execute("insert into " + table + " (`index`)" + " VALUES (" + str(i) + ")")
    keys = []
    key = ""
    stuff = []
    for i in range(len(factors)):
        c.execute("select * from " + prior + " where `index` = " + str(i))
        stuff.append(c.fetchall())
    n = 0
    while n < len(stuff):
        for i in range(len(stuff[0][0])):
            if type(stuff[n][0][-1]) is not float:
                stuff[n][0] = list(stuff[n][0])
                stuff[n][0][-1] = float(stuff[n][0][-1])
                stuff[n][0] = tuple(stuff[n][0])
            try:
                if stuff[n][0][i+1] == stuff[n][0][-1] and type(stuff[n][0][i+1]) is float:
                    keys.append(key)
                    key = ""
                else:
                    string = str(stuff[n][0][i+1]).replace(",", "")
                    string = string.replace(" ", "")
                    string = string.replace("_", "")
                    string = string.replace(".", "")
                    key += string
            except IndexError:
                n += 1

    for i in range(len(factors)):
        dummy = "update " + table + " set Factor = '" + str(factors[i][0]) + "', Key = '" + keys[i] + "' where " + "`index` = " + str(i)
        c.execute(dummy)
        print(dummy)
        conn.commit()
    return keys


def change_in_length(prior, current, table):
    c.execute("select Factor from " + current)
    current_length = c.fetchall()
    c.execute("select Factor from " + prior)
    prior_length = c.fetchall()
    if len(current_length) - len(prior_length) == 0:
        return
    else:
        c.execute("ALTER TABLE " + table + " ADD COLUMN `Change in Length`")
        conn.commit()
        c.execute("UPDATE " + table + " SET `Change in Length` = " + str(len(current_length) - len(prior_length)) + " WHERE `index` = 0")


def main(sheet):
    prior_data = pd.read_excel(prior_book, sheet_name=sheet, header=9)
    current_data = pd.read_excel(current_book, sheet_name=sheet, header=9)
    if sheet == "HOTransMod_Ext":
        conn.close()
        transMod.trans_mod(prior_book, current_book)
        return
    elif sheet == "HOTransModConverCal_Ext":
        conn.close()
        transModConverCalc.trans_mod_conver_cal(prior_book, current_book)
        return
    elif sheet == "PATransMod_Ext":
        conn.close()
        PAtransMod.trans_mod(prior_book, current_book)
        return
    elif sheet == "PATransModConverCal_Ext":
        conn.close()
        PAtransModConverCalc.trans_mod_conver_cal(prior_book, current_book)
        return

    start_time = time.time()
    random_num = rand.randrange(1, 999999999)
    random_num2 = rand.randrange(1, 999999999)
    csv_name = 'Comparison' + sheet + str(random_num) + '.csv'
    prior_table_name = "prior" + str(random_num)
    current_table_name = "current" + str(random_num)
    table_name = "Table" + str(random_num)
    table_name2 = "Table" + str(random_num2)
    prior_data.to_sql(prior_table_name, conn)
    current_data.to_sql(current_table_name, conn)

    c.execute("select * from " + current_table_name + " limit 1")
    print(sheet)

    if c.description[-1][0] == "Each Limit" and c.description[-2][0] == "Factor":
        c.execute("ALTER TABLE " + current_table_name + " DROP COLUMN `Each Limit` ")
        c.execute("ALTER TABLE " + prior_table_name + " DROP COLUMN `Each Limit` ")
    elif c.description[-1][0] != "Factor":
        return print("Improper formatting, Factor must be the last column in the table.")

    create_table(table_name, current_table_name)
    create_query(data_entry(current_table_name), table_name, current_table_name)
    try:
        create_primary_key(current_table_name, table_name)
    except Error:
        print("Formatting Error")
        return

    create_prior_key(prior_table_name, table_name2)
    insert_prior_factor(table_name, prior_table_name, table_name2)
    percent_change(table_name)
    raw_change(table_name)
    change_in_length(prior_table_name, current_table_name, table_name)
    c.execute("select * from " + table_name)

    sql = "SELECT * FROM " + table_name
    df = pd.read_sql_query(sql, conn)
    df.to_csv(csv_name)

    read_file = pd.read_csv(r'Comparison' + sheet + str(random_num) + '.csv')
    try:
        read_file.to_excel(r'Comparison' + sheet + str(random_num) + '.xlsx', sheet_name='Comparison' + sheet)
    except xlsxwriter.exceptions.InvalidWorksheetName:
        read_file.to_excel(r'Comparison' + sheet + str(random_num) + '.xlsx', sheet_name='Comparison' + sheet[:21])
    os.remove('Comparison' + sheet + str(random_num) + '.csv')
    end_time = time.time()
    print(end_time-start_time)
    print(random_num)


if __name__ == "__main__":
    conn = create_connection('test.db')
    c = conn.cursor()

    prior_book = input("Prior Workbook File: ")
    current_book = input("Current Workbook File: ")
    prior_book = prior_book.replace("%20", " ")
    current_book = current_book.replace("%20", " ")
    prior_sheets = pd.ExcelFile(prior_book)
    current_sheets = pd.ExcelFile(current_book)

    current_sheets = current_sheets.sheet_names
    prior_sheets = prior_sheets.sheet_names

    current_sheets_factor = []
    prior_sheets_factor = []

    for thing in prior_sheets:
        if thing[:2] == "HO" or thing[:2] == "PA" or thing[:2] == "DP":
            prior_sheets_factor.append(thing)
        else:
            pass

    for thing in current_sheets:
        if thing[:2] == "HO" or thing[:2] == "PA" or thing[:2] == "DP":
            current_sheets_factor.append(thing)
        else:
            pass

    run_sheets = input("Run all sheets? (y/n): ")

    if run_sheets == "y":
        remover = True
        remove_sheet = input("Do you want to remove a sheet? (sheet name/n): ")
        while remover:
            if remove_sheet == "n":
                remover = False
            else:
                try:
                    current_sheets_factor.remove(remove_sheet)
                    remove_sheet = input("Another? (sheet name/n)")
                except ValueError:
                    remove_sheet = input("Input a valid sheet name (sheet name/n): ")
        all_sheets()
    elif run_sheets == "n":
        sheet_name = input("Sheet Name: ")

        main(sheet_name)
    else:
        run_sheets = input("y/n?")

    excelConsolidation.consolidate()
    print("Complete.")