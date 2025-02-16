import sqlite3

connection = sqlite3.connect('movie.db')

cursor = connection.cursor()

cursor.execute('''CREATE TABLE IF NOT EXISTS
    (Title TEXT, Director TEXT, Year INT)''')

connection.commit()

connection.close()