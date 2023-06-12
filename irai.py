import openai
import openpyxl
import asyncio

KEY = "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"
openai.api_key = KEY

wb = openpyxl.load_workbook("./1.xlsx")
sheet = wb["Sheet1"]
value = sheet.cell(row=1, column=1).value
wb.close()

for row in sheet:
    for cell in filter(lambda c: bool(c.value), row):
        messages = (
            {"role": "system", "content": "あなたは、私を助けてくれるアシスタントです。"},
            {"role": "user", "content": f"{cell.value}についての面白い話をして。"}
        )
        
        completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
        )

        print(completion["choices"][0]["message"]["content"])