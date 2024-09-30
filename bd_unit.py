from datetime import datetime
import sqlite3

DB_NAME = "test.db"


class DatabaseManager:
    def __init__(self, db_name):
        self.db_name = db_name

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
        with sqlite3.connect(self.db_name) as conn:
            sql_delete_table = "DELETE FROM main.tab_lukoil where num_query = ?;"
            conn.execute(sql_delete_table, (num_query,))

    def insert_query(self, list_params):
        with sqlite3.connect(self.db_name) as conn:
            ins_sql = "INSERT INTO main.tab_lukoil (num_query, query_hours, quoter, date_registration) VALUES(?,?,?,?)"

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

    # def read_one_rec(self,**params):
    #     # current_date = datetime.now()
    #     table_name = params['table_name']
    #     current_date=params['current_date']
    #     current_date=datetime.strptime(current_date,  '%d-%m-%Y').date()
    #     first_data = current_date.replace(day=1).strftime('%Y-%m-%d')
    #
    #     with sqlite3.connect(self.db_name) as conn:
    #         if table_name == 'tab_fte':
    #             sql_read_table = "SELECT FTE FROM tab_fte WHERE MONTH_NAME = ?;"
    #         elif table_name == 'tab_lukoil':
    #             sql_read_table = f"SELECT FTE FROM {table_name} WHERE num_query = ?;"
    #         cursor = conn.execute(sql_read_table, (first_data,))
    #         return cursor.fetchall()
    # rows = cursor.fetchall()
    # for row in rows:
    #     print(row)

    def read_all_table(self, table_name):
        with sqlite3.connect(self.db_name) as conn:
            if table_name == 'tab_fte':
                sql_read_table = "SELECT  JOB_DAYS, MONTH_NAME, FTE FROM tab_fte order by MONTH_NAME;"
            elif table_name == 'tab_lukoil':
                sql_read_table = ("SELECT num_query,query_hours, date_registration, quoter "
                                  "FROM tab_lukoil order by quoter desc")
            cursor = conn.execute(sql_read_table)
            rows = cursor.fetchall()
            return rows

    # def insert_new_data(self):
    #     self.delete_table()
    #     params = [
    #         (19, '2024-06-01'),
    #         (23, '2024-07-01')
    #     ]
    #     self.insert_data(params)

# if __name__ == '__main__':
#     db_manager = DatabaseManager(DB_NAME)
#     db_manager.create_table()
# db_manager.insert_new_data()
# db_manager.read_table()
