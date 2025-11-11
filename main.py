import httpx
import openpyxl

API_URL = "https://admin.metal-expo.ru/api/public_site/participants/?ex_prefix=ME25&ex_type=ME&lang_code=ru"

r = httpx.get(API_URL)
data = r.json()

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Участники"

ws.append([
    "Название компании",
    "Стенд",
    "Адрес",
    "Телефон",
    "Веб сайт",
    "Email",
    "Описание"
])

for letter, companies in data.items():
    for company in companies:
        ws.append([
            company.get("full_title", ""),
            ", ".join(company.get("stands", [])),
            company.get("address", ""),
            company.get("phone_number", ""),
            company.get("website", ""),
            company.get("email", ""),
            company.get("text", "")
        ])

wb.save("metalexpo.xlsx")

print("✅ Файл metalexpo.xlsx успешно создан!")
input()

