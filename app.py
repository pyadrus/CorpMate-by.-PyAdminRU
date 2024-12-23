from datetime import datetime

from fastapi import FastAPI, Form, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from loguru import logger

from database import import_excel_to_db, read_from_db, clear_database
from employment_contracts.filling_a_shortened_work_week import creation_contracts_downtime_week
from employment_contracts.filling_data import creation_contracts, format_date
from employment_contracts.filling_plant_downtime import creation_contracts_downtime
from get import Employee
from parsing_comparison_file import parsing_document_1, compare_and_rewrite_professions

app = FastAPI()
templates = Jinja2Templates(directory="templates")
progress_messages = []  # список сообщений, которые будут отображаться в progress


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/import_excel_form", response_class=HTMLResponse)
async def import_excel_form(request: Request):
    return templates.TemplateResponse("import_excel_form.html", {"request": request})


@app.post("/import_excel")
async def import_excel(min_row: int = Form(...), max_row: int = Form(...)):
    try:
        logger.info(f"Запуск импорта данных с {min_row} по {max_row} строки.")
        await import_excel_to_db(min_row=min_row, max_row=max_row)
        return RedirectResponse(url="/", status_code=303)
    except Exception as e:
        logger.exception("Ошибка при импорте данных.")
        raise HTTPException(status_code=500, detail="Произошла ошибка при импорте данных.")


def search_employee_by_tab_number(tab_number):
    """Ищем данные сотрудника по табельному номеру"""
    try:
        customer = Employee.get(Employee.a4_табельный_номер == tab_number)
        return customer
    except Employee.DoesNotExist:
        return None


@app.get("/get_contract", response_class=HTMLResponse)
async def get_contract_form(request: Request):
    return templates.TemplateResponse("get_contract.html", {"request": request})


@app.post("/get_contract")
async def get_contract(tab_number: str = Form(...)):
    logger.info(f"Введенный табельный номер: {tab_number}")

    if tab_number:
        data = search_employee_by_tab_number(tab_number)
        if data:
            return {
                "message": f"Данные для табельного номера {tab_number}",
                "data": {'a0': data.a0, 'a1': data.a1, 'a2': data.a2, 'a3': data.a3,
                         'a4_табельный_номер': data.a4_табельный_номер, 'a5': data.a5, 'a6': data.a6, 'a7': data.a7,
                         'a8': data.a8, 'a9': data.a9, 'a10': data.a10, 'a11': data.a11, 'a12': data.a12,
                         'a13': data.a13, 'a14': data.a14, 'a15': data.a15, 'a16': data.a16, 'a17': data.a17,
                         'a18': data.a18, 'a19': data.a19, 'a20': data.a20, 'a21': data.a21, 'a22': data.a22,
                         'a23': data.a23, 'a24': data.a24, 'a25_номер_договора': data.a25_номер_договора,
                         'a26': data.a26, 'a27': data.a27, 'a28': data.a28, 'a29': data.a29, 'a30': data.a30,
                         'a31': data.a31, 'a32': data.a32, 'a33': data.a33, 'a34': data.a34, },
            }
        else:
            return {"message": f"Данные для табельного номера {tab_number} не найдены."}
    raise HTTPException(status_code=400, detail="Табельный номер не указан.")


@app.post("/action", response_class=HTMLResponse)
async def action(request: Request, user_input: str = Form(...)):
    try:
        user_input = int(user_input)

        if user_input == 1:
            await parsing_document_1(min_row=5, max_row=1084, column=5, column_1=8)

        elif user_input == 2:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        elif user_input == 3:
            await compare_and_rewrite_professions()


        elif user_input == 4:

            return RedirectResponse(url="/import_excel_form", status_code=303)


        elif user_input == 5:
            return RedirectResponse(url="/get_contract", status_code=303)


        elif user_input == 7:
            try:
                await clear_database()
                logger.info("База данных успешно очищена.")
                return templates.TemplateResponse("database_cleanup.html",
                                                  {"request": request, "message": "База данных успешно очищена!"})
            except Exception as e:
                logger.exception("Ошибка при очистке базы данных.")
                return templates.TemplateResponse("database_cleanup.html",
                                                  {"request": request, "message": f"Ошибка: {e}"})

        elif user_input == 8:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts_downtime(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        elif user_input == 9:
            start = datetime.now()
            logger.info(f"Время старта: {start}")
            data = await read_from_db()
            for row in data:
                logger.info(row)
                ending = "ый" if row.a11 == "Мужчина" else "ая"
                await creation_contracts_downtime_week(row, await format_date(row.a7), ending)
            finish = datetime.now()
            logger.info(f"Время окончания: {finish}")
            logger.info(f"Время работы: {finish - start}")

        return RedirectResponse(url="/", status_code=303)

    except Exception as e:
        logger.exception(e)
        raise HTTPException(status_code=500, detail="Произошла ошибка.")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
