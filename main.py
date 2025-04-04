from fastapi import FastAPI
from fastapi.responses import FileResponse, Response
import pandas as pd
from io import BytesIO

app = FastAPI()

@app.get("/")
async def get_excel():
    # Создаем тестовый DataFrame (можно заменить на свои данные)
    data = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie"],
        "Age": [25, 30, 35],
        "Department": ["HR", "IT", "Finance"]
    })
    
    # Создаем Excel-файл в памяти
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        data.to_excel(writer, index=False, sheet_name='Employees')
    output.seek(0)
    
    # Сохраняем временный файл (альтернативный вариант)
    # excel_path = "/tmp/report.xlsx"
    # data.to_excel(excel_path, index=False)
    # return FileResponse(excel_path, filename="report.xlsx")
    
    # Возвращаем файл напрямую из памяти
    headers = {
        "Content-Disposition": "attachment; filename=report.xlsx"
    }
    return Response(content=output.getvalue(), 
                   media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   headers=headers)
