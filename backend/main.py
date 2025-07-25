from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import os
import uuid
from pathlib import Path
import logging
import time
import win32com.client
import pythoncom

# Настройка логирования
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# Разрешаем CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000",
    "http://176.123.163.173",
    "http://176.123.163.173:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PURCHASES_FILE = os.path.join(BASE_DIR, "Файл с закупками.xlsx")
MACROS_FILTER = os.path.join(BASE_DIR, "Фильтрация строк.txt")
MACROS_PROFIT = os.path.join(BASE_DIR, "ИтогПрибыли_2.txt")

# Создаём папки
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def run_macros(input_path: str):
    """Запускает макросы в Excel с полным подавлением предупреждений"""
    pythoncom.CoInitialize()
    excel = None
    wb = None
    purchases_wb = None
    
    try:
        # Проверка файлов
        if not os.path.exists(PURCHASES_FILE):
            raise FileNotFoundError(f"Файл закупок не найден: {PURCHASES_FILE}")
        
        abs_input_path = os.path.abspath(input_path)
        abs_purchases_path = os.path.abspath(PURCHASES_FILE)
        
        logger.info(f"Обрабатываем файл: {abs_input_path}")
        logger.info(f"Используем файл закупок: {abs_purchases_path}")

        # Настройка Excel
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow
        excel.AlertBeforeOverwriting = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        # Открываем файлы
        purchases_wb = excel.Workbooks.Open(
            Filename=abs_purchases_path,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=1
        )
        time.sleep(1)
        
        wb = excel.Workbooks.Open(
            Filename=abs_input_path,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            CorruptLoad=1
        )
        time.sleep(1)

        # Загружаем макросы
        with open(MACROS_FILTER, 'r', encoding='utf-8') as f:
            filter_macro = f.read()
        with open(MACROS_PROFIT, 'r', encoding='utf-8') as f:
            profit_macro = f.read()

        # Создаем макрос для отключения предупреждений
        disable_warnings_macro = """
Sub DisableAllWarnings()
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.AlertBeforeOverwriting = False
    ActiveWorkbook.RemovePersonalInformation = True
    Application.AutomationSecurity = 1
    Application.AskToUpdateLinks = False
    ThisWorkbook.RemoveDocumentInformation (1)  ' xlRDIDocumentProperties
End Sub
"""

        # Добавляем все макросы в один модуль
        vb_component = wb.VBProject.VBComponents.Add(1)
        vb_component.Name = "TempMacros"
        vb_component.CodeModule.AddFromString(
            disable_warnings_macro + "\n" + 
            filter_macro + "\n" + 
            profit_macro
        )
        time.sleep(1)

        # Выполняем макрос для отключения предупреждений
        excel.Application.Run("DisableAllWarnings")
        time.sleep(1)
        
        # Выполняем основные макросы
        excel.Application.Run("ФильтрацияСтрок")
        time.sleep(1)
        excel.Application.Run("ИтогПрибыли")
        time.sleep(1)

        # Сохраняем результат
        output_filename = f"processed_{os.path.basename(input_path)}"
        output_path = os.path.join(PROCESSED_FOLDER, output_filename)
        
        # Перед сохранением снова отключаем предупреждения
        excel.Application.Run("DisableAllWarnings")
        wb.SaveAs(
            Filename=os.path.abspath(output_path),
            FileFormat=51,  # xlOpenXMLWorkbook
            ConflictResolution=2,  # xlLocalSessionChanges
            Local=True
        )
        
        return output_path

    except Exception as e:
        logger.error(f"Ошибка при обработке: {e}")
        raise
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
            if purchases_wb: purchases_wb.Close(SaveChanges=False)
            if excel: 
                excel.DisplayAlerts = False
                excel.Quit()
        except Exception as e:
            logger.error(f"Ошибка при закрытии Excel: {e}")
        finally:
            pythoncom.CoUninitialize()

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...)):
    input_path = None
    try:
        if not file.filename.lower().endswith((".xls", ".xlsx")):
            raise HTTPException(400, "Только файлы Excel (.xls, .xlsx)")

        file_id = str(uuid.uuid4())
        input_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_{file.filename}")
        
        with open(input_path, "wb") as buffer:
            buffer.write(await file.read())

        output_path = run_macros(input_path)

        return FileResponse(
            output_path,
            filename=f"processed_{file.filename}",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        raise HTTPException(500, str(e))
    finally:
        try:
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
        except Exception as e:
            logger.error(f"Ошибка удаления временного файла: {e}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
