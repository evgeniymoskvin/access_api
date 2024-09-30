import os

import uvicorn
import json
import pyodbc
import datetime
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from dotenv import load_dotenv, dotenv_values

load_dotenv()

app = FastAPI()

FILE_ = os.getenv('FILE_')
FULL_PATH_TO_FILE = os.path.abspath(FILE_)
sql_connection = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" + f"DBQ={FULL_PATH_TO_FILE};"
CONNECTION = pyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" + \
                            f"DBQ={FULL_PATH_TO_FILE};", autocommit=True)
CONNECTION.autocommit = True


@app.get("/")
async def root():
    c = CONNECTION.cursor()
    c.execute(
        'SELECT TOP 1 ИнвНомер.*, ИнвНомер.КодИнвНомер, ИнвНомер.ДатаВыдИнвНом FROM ИнвНомер '
        'WHERE (((ИнвНомер.ДатаВыдИнвНом) Is NOT Null) AND ((ИнвНомер.ГодИнвНом)=2024) AND ((ИнвНомер.КодТипИнвНомер)=1))'
        'ORDER BY ИнвНомер.КодИнвНомер DESC')
    item = c.fetchall()
    c.close()
    print(item)
    # print(json.dumps(item))
    return {"message": f"{item}"}


@app.post("/post")
def say_hello(employee, order, cpe, department, phone: str, type_of_inventory_number: int):
    c = CONNECTION.cursor()
    day_now = datetime.date.today().year
    c.execute(
        'SELECT TOP 1 ИнвНомер.*, ИнвНомер.КодИнвНомер, ИнвНомер.ДатаВыдИнвНом FROM ИнвНомер '
        'WHERE (((ИнвНомер.ДатаВыдИнвНом) Is NOT Null) AND ((ИнвНомер.ГодИнвНом)=(?)) AND ((ИнвНомер.КодТипИнвНомер)=(?)))'
        'ORDER BY ИнвНомер.КодИнвНомер DESC', (day_now, type_of_inventory_number))
    item = int(c.fetchall()[0][0]) + 1
    print(item)
    c_to_insert = CONNECTION.cursor()
    c_to_insert.execute(
        """UPDATE ИнвНомер 
        SET ИнвНомер.ФИОИсполнит = (?), ИнвНомер.ДатаВыдИнвНом = (?), 
        ИнвНомер.ЦКДИ = (?), ИнвНомер.ФИОГИП = (?), ИнвНомер.НомОтдела = (?), ИнвНомер.Телефон = (?) 
        WHERE ((ИнвНомер.КодИнвНомер)=(?));""", (employee, datetime.date.today(), order, cpe, department, phone, item))
    CONNECTION.commit()
    c.execute(
        'SELECT TOP 1 ИнвНомер.*, ИнвНомер.КодИнвНомер, ИнвНомер.ДатаВыдИнвНом FROM ИнвНомер '
        'WHERE (ИнвНомер.КодИнвНомер=(?))', (item))
    # print(c.fetchall())
    new_inventory = c.fetchall()[0][2]
    c.close()
    return {"new_inventory": f"{new_inventory}"}


if __name__ == "__main__":
    port = int(os.getenv('PORT'))
    host = str(os.getenv('HOST'))
    uvicorn.run(app, host=host, port=port)
