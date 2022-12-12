import requests
from openpyxl import Workbook

arquivo_excel = Workbook()
planilha = arquivo_excel.create_sheet("Ganhos")


request = requests.get("https://cdn.apicep.com/file/apicep/32606-582.json")
print(request.content)

planilha1['A1'] = request.content