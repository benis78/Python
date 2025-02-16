import sqlite3

connection = sqlite3.connect('movie.db')

cursor = connection.cursor()

cursor.execute('''CREATE TABLE IF NOT EXISTS Movies (Title TEXT, Location TEXT, Number INT)''')

connection.commit()

connection.close()