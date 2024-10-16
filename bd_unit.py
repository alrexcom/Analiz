from datetime import datetime
import sqlite3
from pathlib import Path

import pandas as pd

DB_NAME = "test.db"


class DatabaseManager:
    def __init__(self):

        self.db_name = Path('.').absolute().joinpath('BD').joinpath(DB_NAME)

    # def create_table(self):
    #     with sqlite3.connect(self.db_name) as conn:
    #         drop_table = "DROP TABLE IF EXISTS tab_fte;"
    #         conn.execute(drop_table)
    #         sql_request = """
    #                         CREATE TABLE tab_fte (
    #                             ID INTEGER PRIMARY KEY AUTOINCREMENT,
    #                             JOB_DAYS INTEGER NOT NULL,
    #                             MONTH_NAME DATE NOT NULL
    #                                 CONSTRAINT tab_fte_pk UNIQUE ON CONFLICT ROLLBACK,
    #                             FTE INTEGER GENERATED ALWAYS AS (JOB_DAYS * 8) VIRTUAL
    #                         );
    #                       """
    #         conn.execute(sql_request)

    def delete_record(self, date_month_name):
        """
        Удаление одной записи
        :param date_month_name:
        :return:
        """
        with sqlite3.connect(self.db_name) as conn:
            sql_delete_table = "DELETE FROM main.tab_fte where MONTH_NAME = ?;"
            conn.execute(sql_delete_table, (date_month_name,))

    def delete_num_query(self, num_query):
        """
        Удаление одной записи
        :param date_month_name:
        :return:
        """
        num_task = self.get_task_number(num_query)
        with sqlite3.connect(self.db_name) as conn:
            sql_delete_table = "DELETE FROM main.tab_lukoil where num_query = ?;"
            conn.execute(sql_delete_table, (num_query,))

        self.set_sum_number_query_on_delete(num_query, num_task)

    def insert_query(self, list_params):
        with sqlite3.connect(self.db_name) as conn:
            ins_sql = (
                "INSERT INTO main.tab_lukoil (num_query, query_hours, quoter, date_registration, month_date,"
                "description, num_task, first_input) VALUES(?,?,?,?,?,?,?,?)")

            for par in list_params:
                conn.execute(ins_sql, par)
            conn.commit()

    def insert_data(self, list_params):
        """
        Вставка в таблицу
        :param list_params: [(,),(,)]
        :return: None
        """
        with sqlite3.connect(self.db_name) as conn:
            ins_sql = "INSERT INTO main.tab_fte (JOB_DAYS, MONTH_NAME) VALUES(?,?)"
            for par in list_params:
                conn.execute(ins_sql, par)
            conn.commit()

    def read_one_rec(self, current_date):
        # current_date = datetime.now()

        current_date = datetime.strptime(current_date, '%d-%m-%Y').date()
        first_data = current_date.replace(day=1).strftime('%Y-%m-%d')

        with sqlite3.connect(self.db_name) as conn:
            sql_read_table = "SELECT FTE FROM tab_fte WHERE MONTH_NAME = ?;"
            cursor = conn.execute(sql_read_table, (first_data,))
            return cursor.fetchall()

    def read_sum_lukoil(self, ds, dp):
        project = 'С0134-КИС "Производственный учет и отчетность"'
        user = 'Тапехин Алексей Александрович'
        with sqlite3.connect(self.db_name) as conn:
            #  = , [Пользователь] = ,
            sql_read_table = (
                f"SELECT '{project}' as [Проект], '{user}' as [Пользователь], month_date as [Дата], "
                f"sum(query_hours) as [Лукойл, час.] , count(*) as [Заявок лукойл] "
                f"FROM tab_lukoil "
                f"where date_registration >= ? and date_registration <= ?  and first_input > 0 "
                f"group by month_date")
        cursor = conn.execute(sql_read_table, (ds, dp))
        rows = cursor.fetchall()
        return pd.DataFrame(rows, columns=['Проект', 'Пользователь', 'Дата', 'Лукойл, час.', 'Заявок лукойл'])
        # return rows

    def read_all_lukoil(self):
        with sqlite3.connect(self.db_name) as conn:
            sql_read_table = (
                "SELECT num_query, num_task, query_hours, date_registration, quoter,  month_date, description "
                "FROM tab_lukoil order by quoter desc")
        cursor = conn.execute(sql_read_table)
        rows = cursor.fetchall()
        return rows

    def read_all_table(self):
        with sqlite3.connect(self.db_name) as conn:
            sql_read_table = "SELECT  JOB_DAYS, MONTH_NAME, FTE FROM tab_fte order by MONTH_NAME;"
            cursor = conn.execute(sql_read_table)
            rows = cursor.fetchall()
            return rows

    def get_summaryon_numbquery(self, num_query):
        with sqlite3.connect(self.db_name) as conn:
            # sql_read_table = ("SELECT sum(first_input) + sum(query_hours) "
            #                   "FROM tab_lukoil "
            #                   "WHERE num_query = ? or num_task = ?;")
            sql_read_table = ("select sum(case when num_task= ? then query_hours else 0 end) "
                              "+  sum(case when num_query= ? then first_input else 0 end) as qh_ "
                              "from tab_lukoil t where  ? in (num_task,num_query)")

            cursor = conn.execute(sql_read_table, (num_query, num_query, num_query,))
            row = cursor.fetchone()
            return row[0] if row else 0

    def set_sum_numbquery(self, num_query, query_hours):
        sum_query = query_hours + self.get_summaryon_numbquery(num_query)
        with sqlite3.connect(self.db_name) as conn:
            sql = "update main.tab_lukoil set query_hours = ? where num_query = ?;"
            params = (sum_query, num_query)
            conn.execute(sql, params)
            conn.commit()

    def get_task_number(self, num_query):
        with sqlite3.connect(self.db_name) as conn:
            sql_read_table = "SELECT num_task FROM tab_lukoil WHERE num_query=?"
            cursor = conn.execute(sql_read_table, (num_query,))
            row = cursor.fetchone()  # Получаем только одну строку
        return row[0] if row else None

    def set_sum_number_query_on_delete(self, num_query, num_task):

        if num_task:
            num_query = num_task

        sum_query = self.get_summaryon_numbquery(num_query)

        if sum_query:
            with sqlite3.connect(self.db_name) as conn:
                if num_task:
                    sql = "update main.tab_lukoil set query_hours = ? where num_query = ?;"
                else:
                    sql = "update main.tab_lukoil set query_hours = first_input where num_query = ?;"
                params = (sum_query, num_query)
                conn.execute(sql, params)
                conn.commit()

