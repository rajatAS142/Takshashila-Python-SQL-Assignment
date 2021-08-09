#installing the libraries psycopg2, Workbook, pandas
import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd

class employee_info:

    def emp(self):
        try:
            # to connect to the PostgreSQL database server in the Python program using the psycopg database adapter.
            con = psycopg2.connect(
                database="python-sql",
                user="postgres",
                password="timestone4me")
            # Creating a cursor object using the cursor() method
            cursor = con.cursor()
            # Reading table which we imported using connection through query
            query_data_command = """
                    select dept.deptno, dept_name, sum(total_compensation) from Compensation, dept
                    where Compensation.dept_name=dept.dname
                    group by dept_name, dept.deptno
                    """

            cursor.execute(query_data_command)
            #iterating inside description
            columns = [desc[0] for desc in cursor.description]
            data = cursor.fetchall()
            dataframe = pd.DataFrame(list(data), columns=columns)
            #storing values inside excel
            writer = pd.ExcelWriter('ques_4.xlsx')
            # converting data frame to excel
            dataframe.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("Something went wrong", e)

        finally:

            if conn is not None:
                cursor.close()
                conn.close()


if __name__ == '__main__':
    conn = None
    cur = None
    employee = employee_info()
    employee.emp()