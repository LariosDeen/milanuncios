import math
import sqlite3 as sq
from datetime import datetime, date

date_now =date.today().strftime('%Y-%m-%d')


def get_last_data(database_name: str, table_name: str, col_name: str) -> str:
    """Returns last row <column_name> data from <database_name> <table_name>."""
    with sq.connect(database_name) as con:
        cur = con.cursor()
        cur.execute(f'''SELECT {col_name} 
                        FROM {table_name} 
                        ORDER BY rowid 
                        DESC 
                        LIMIT 1
        ''')
        last_col_data = cur.fetchone()[0]
        return last_col_data


def insert_entry(database_name: str, table_name: str, lst: list):
    """Add entry into <database_name> <table_name> from list."""
    with sq.connect(database_name) as con:
        cur = con.cursor()
        # cur.execute('INSERT INTO prices VALUES %r;' % (tuple(params),))
        cur.execute(f'INSERT INTO {table_name} VALUES(?,?,?,?,?,?,?,?,?,?,?)', lst)


def diff_dates(start_day: str, stop_day: str) -> int:
    """Returns difference in days between two dates. Dates format is 'yyyy-mm-dd'."""
    start = datetime.strptime(start_day, '%Y-%m-%d').date()
    stop = datetime.strptime(stop_day, '%Y-%m-%d').date()
    difference = (stop - start).days
    return difference


def create_table(database_name: str, table_name: str):
    """Creates <table_name> in <database_name>."""
    with sq.connect(database_name) as con:
        cur = con.cursor()
        cur.execute(f'''CREATE TABLE IF NOT EXISTS {table_name}(
        date DATE,
        malaga600_699 INTEGER NOT NULL,
        malaga700_799 INTEGER NOT NULL,
        malaga800_899 INTEGER NOT NULL,
        malaga900_999 INTEGER NOT NULL,
        torremolinos600_699 INTEGER NOT NULL,
        torremolinos700_799 INTEGER NOT NULL,
        torremolinos800_899 INTEGER NOT NULL,
        torremolinos900_999 INTEGER NOT NULL,
        malaga_4d INTEGER NOT NULL,
        torremolinos_4d INTEGER NOT NULL
        )
        ''')


def delete_table(database_name: str, table_name: str):
    """Deletes <table_name> from <database_name>."""
    with sq.connect(database_name) as con:
        cur = con.cursor()
        cur.execute(f'DROP TABLE IF EXISTS {table_name}')


def estimated_values(lst_1: list, lst_2: list, parts: int) -> list:
    """Returns the suggested entries as a list of lists of values for the missing days."""
    result_lst = []
    difference_lst = []
    entry = []
    for i in range(len(lst_1)):
        difference_lst.append(lst_2[i] - lst_1[i])
    for i in range(parts-1):
        for j in range(len(lst_1)):
            entry.append(lst_1[j] + math.trunc(difference_lst[j] / parts * (i+1)))
        result_lst.append(entry)
        entry = []
    return result_lst
