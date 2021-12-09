# pip install pandas
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment


def csv_to_excel(fileIn, fileOut):
    # 2) read csv file
    csv = pd.read_csv(fileIn)

    # 3) create excel writer
    excelWriter = pd.ExcelWriter(fileOut)

    # 4) convert csv to excel
    csv.to_excel(
        excelWriter,
        index_label='ABC',
        index=False,
        float_format='%.2f',
        # header = False,
        freeze_panes=(3, 1),
        sheet_name='grades from csv'
    )
    # 5) save Excel file
    excelWriter.save()


def ExtractTxt(canvasText, UPI, Score):
    counterX = 4
    # Step 2 Extract data from txt file
    with open(canvasText) as fp:
        contents = fp.read()
        try:
            for entry in contents.split('\n'):
                UPI.append(int(entry.split(' ')[1]))
                Score.append(int(entry.split(' ')[2]))
        except Exception as e:
            print(e)
        counterX += 1


# Step 3 Add grades to the Excel sheet
def AutoMarking(wb, ws, row_count, UPI, Score, fileOut):
    # Check if upi matches
    counter = 4
    for cols in range(4, row_count):
        try:
            if UPI[counter] == ws['C' + str(cols)].value:
                # Add score
                ws['F' + str(cols)].value = Score[counter]
                # Align left
                ws['F' + str(cols)].alignment = Alignment(horizontal='left')

                print('\nUPI: ' + str(UPI[counter]) + ', Score: ' + str(Score[counter]))
                print('Checking UPI: ' + str(UPI[counter]) + ' for C' + str(cols))
                print('Adding' + ' score ' + str(Score[counter]) + ' for F' + str(cols) + '\n')

                # Increment counter
                counter += 1

        except Exception as e:
            print(e)
            break

    wb.save(fileOut)
    print('Marking completed for ' + str(counter - 4) + ' student(s)')


def main():
    # Constants
    fileOut = 'grades.xlsx'
    fileIn = 'grades.csv'
    canvasText = 'canvasMarking.txt'
    # 4 Columns are empty so they are None
    UPI = [None, None, None, None]
    Score = [None, None, None, None]

    # Convert csv to xlsx
    csv_to_excel(fileIn, fileOut)

    # load excel file
    wb = load_workbook(fileOut)
    ws = wb.active

    # Get total rows and columns
    sheet = wb.worksheets[0]
    row_count = sheet.max_row
    column_count = sheet.max_column
    print('row: ' + str(row_count))
    print('col: ' + str(column_count))

    # Extract data from txt file
    ExtractTxt(canvasText, UPI, Score)
    print(UPI)
    print(Score)

    #  Add grades to the Excel sheet
    AutoMarking(wb, ws, row_count, UPI, Score, fileOut)


if __name__ == '__main__':
    main()
