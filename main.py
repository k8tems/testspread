import gspread


if __name__ == '__main__':
    SHEET_NAME = 'StableCascade損失計算シート'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    client = gspread.service_account()
    sheet = client.open(SHEET_NAME).sheet1
    row = 1
    values = ['Cell 1', 'Cell 2']
    sheet.update(f'A{row}:B{row}', [values])
