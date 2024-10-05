import openpyxl

def create_excel_file():
    # Create a new workbook and get the active sheet
    wb = openpyxl.Workbook()
    sheet = wb.active

    # Add headers
    sheet.append(["PIN", "Balance", "Transaction History"])

    # Add sample data
    data = [
        (1234, 1000, "Deposit: 500, Withdraw: 200"),
        (5678, 500, "Withdraw: 100"),
        (9876, 1500, "Deposit: 700"),
    ]

    # Populate the sheet with data
    for row in data:
        sheet.append(row)

    # Save the workbook as "database.xlsx"
    wb.save("database.xlsx")

if __name__ == "__main__":
    create_excel_file()