import sys
import json
import os
import psycopg2
from psycopg2.extras import RealDictCursor


class Database_Manager():    
    def __init__(self, database, db_credentials):
        self.credential = db_credentials#json.load(open(str(os.path.dirname(__file__)) + '/db_credentials.json'))
        # self.credential = self.db_credentials[server]
        self.ip_server = self.credential['ip_server']
        self.username = self.credential['username']
        self.password = self.credential['password']
        self.database = database

        try:
            self._cnxn = psycopg2.connect(
                host=self.ip_server,
                dbname=self.database,
                user=self.username,
                password=self.password,
                options='-c search_path=public'
            )
            self._cnxn.autocommit = True
            self.db = self._cnxn.cursor()
            print(f'***** {self.ip_server} {self.database} Database connection established *****')
        except psycopg2.Error as ex:
            print(ex)
            sys.exit('***** Cannot connect to the PostgreSQL database *****')

    def query_with_exception(self, query):
        return self.db.execute(query)

    def query_exec(self, query):
        try:
            print(query)
            return self.db.execute(query)
        except psycopg2.Error as ex:
            print(ex)

    def query_exec_1(self, query):
        try:
            print(query)
            return self.db.execute(query)
        except psycopg2.Error as ex:
            print(ex)

    def query(self, query):
        try:
            return self.db.execute(query)
        except psycopg2.Error as ex:
            print(ex, query)

    def query_test(self, query):
        self.db.execute(query)
        self._cnxn.commit()
        return self.db.rowcount

    def query_executemany(self, query, batch):
        try:
            return self.db.executemany(query, batch)
        except psycopg2.Error as ex:
            print(ex, query)

    def fetchArray(self, query):
        self.db.execute(query)
        rows = self.db.fetchall()
        return rows

    def getKeys(self, query):
        self.db.execute(query)
        columns = [desc[0] for desc in self.db.description]
        return columns

    def fetchArray_withKey(self, query):
        with self._cnxn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(query)
            return cur.fetchall()

    def numrows(self, query):
        self.db.execute(query)
        rows = self.db.fetchall()
        return len(rows)

    def __del__(self):
        try:
            self.db.close()
            self._cnxn.close()
            print(f'***** {self.server} {self.database} Database connection destructed *****')
        except AttributeError:
            pass
        except psycopg2.Error:
            pass
