import sqlite3
import sys
from datetime import datetime


class Database:

    def __init__(self):
        self.con = sqlite3.connect('time.db')
        self.cursor = self.con.cursor()
        self.create_tables()

    def create_tables(self):
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS Tasks (id INTEGER PRIMARY KEY,task TEXT,date DateTime)''')
        self.con.commit()

    def create_task(self, task):
        self.cursor.execute("INSERT INTO Tasks(task,date) VALUES (?,?)", (task, datetime.now()))
        self.close()

    def get_all_tasks(self):
        data = self.cursor.execute("SELECT * FROM TASKS")
        return data

    @staticmethod
    def get_last_task():
        con = sqlite3.connect('time.db')
        cursor = con.cursor()
        data = cursor.execute("SELECT * FROM TASKS ORDER BY ID DESC LIMIT 1")
        data = data.fetchall()
        con.close()
        return data

    def close(self):
        self.con.commit()
        self.con.close()


sys.path.append(".")

