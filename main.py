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
async def get_last():
    c = CONNECTION.cursor()
    day_now = datetime.date.today().year
    c.execute(
        'SELECT TOP 1 ИнвНомер.*, ИнвНомер.КодИнвНомер, ИнвНомер.ДатаВыдИнвНом FROM ИнвНомер '
        'WHERE (((ИнвНомер.ДатаВыдИнвНом) Is NOT Null) AND ((ИнвНомер.ГодИнвНом)=(?)) AND ((ИнвНомер.КодТипИнвНомер)=1))'
        'ORDER BY ИнвНомер.КодИнвНомер DESC', (day_now,))
    item = c.fetchall()
    c.close()
    print(item)
    # print(json.dumps(item))
    return {"message": f"{item}"}


@app.post("/post")
def write_new_inventory(employee, order, cpe, department, phone: str, type_of_inventory_number: int):
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
        ИнвНомер.ЦКДИ = (?), ИнвНомер.ФИОГИП = (?), ИнвНомер.НомОтдела = (?), ИнвНомер.Телефон = (?), ИнвНомер.Примечание = (?) 
        WHERE ((ИнвНомер.КодИнвНомер)=(?));""", (employee, datetime.date.today(), order, cpe, department, phone, 'Сайт', item))
    CONNECTION.commit()
    c.execute(
        'SELECT TOP 1 ИнвНомер.*, ИнвНомер.КодИнвНомер, ИнвНомер.ДатаВыдИнвНом FROM ИнвНомер '
        'WHERE (ИнвНомер.КодИнвНомер=(?))', (item, ))
    # print(c.fetchall())
    new_inventory = c.fetchall()[0][2]
    c.close()
    return {"new_inventory": f"{new_inventory}"}




@app.post("/permission")
def write_new_permission(employee, order, cpe, department, phone: str, type_of_permission_number: int):
    c = CONNECTION.cursor()
    day_now = datetime.date.today().year
    c.execute(
        'SELECT TOP 1 ДокИзменения.* FROM ДокИзменения '
        'WHERE (((ДокИзменения.ДатаВыдДокИзм) Is NOT Null) AND ((ДокИзменения.ГодДокИзм)=(?)) AND ((ДокИзменения.КодТипДокИзм)=(?)))'
        'ORDER BY ДокИзменения.КодДокИзм DESC', (day_now, type_of_permission_number))
    item = int(c.fetchall()[0][0]) + 1
    print(item)
    c_to_insert = CONNECTION.cursor()
    c_to_insert.execute(
        """UPDATE ДокИзменения
        SET ДокИзменения.ФИОИсполнит = (?), ДокИзменения.ДатаВыдДокИзм = (?),
        ДокИзменения.ЦКДИ = (?), ДокИзменения.ГИП = (?), ДокИзменения.НомОтдел = (?), ДокИзменения.Телефон = (?)
        WHERE ((ДокИзменения.КодДокИзм)=(?));""", (employee, datetime.date.today(), order, cpe, department, phone, item))
    CONNECTION.commit()
    c.execute(
        'SELECT TOP 1 ДокИзменения.* FROM ДокИзменения '
        'WHERE (ДокИзменения.КодДокИзм=(?))', (item, ))
    new_permission = c.fetchall()[0][1]
    # print(item)
    # print(c.fetchall())
    c.close()
    return {"new_permission": f"{new_permission}"}
    # return {"new_permission": f"ok"}


if __name__ == "__main__":
    port = int(os.getenv('PORT'))
    host = str(os.getenv('HOST'))
    uvicorn.run(app, host=host, port=port)
