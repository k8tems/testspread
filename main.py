import gspread


if __name__ == '__main__':
    SHEET_NAME = 'StableCascade Loss Vis'
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    client = gspread.service_account()
    sheet = client.open(SHEET_NAME).sheet1
    row = 1
    values = ['Cell 1', 'Cell 2']
    sheet.update([values], f'A{row}:B{row}')
