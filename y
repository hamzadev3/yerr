import openpyxl
import os
import re


def FindFiles(directory):
    patterns = [
        r'\(\s*(\d{4,5})\s*\)',  # Parentheses with 4 or 5 digits
        r'(\d{5})(?=[A-Za-z])',  # Five digits followed by a letter
        r'#\s*(\d{5})\b',        # Hashtag with optional spaces
        r'_?(\d{5})_?',          # Five digits wrapped in underscores or surrounded by other delimiters
        r'(?<!\d)\d{5}(?!\d)'   # Standalone five digits
    ]

    for file in os.scandir(directory):
        if file.is_file() or file.is_dir():
            match = None
            for pattern in patterns:
                match = re.search(pattern, file.name)
                if match:
                    break

            if match:
                keyword = match.group(1) # keyword is extracted

                if isExcludedFile(file.name):
                    print("Excluding file due to exclusion pattern: ", file.path) 
                    continue

                if isInaccurateMatch(file.name, keyword):
                    print("Skipping file due to inaccurate match: ", file.path)
                    continue

                # if not excluded / inaccurate 
                EditExcel(keyword, file.path, 3, "U:/dataexport/copyofalldata3.xlsx")
                continue

        if file.is_dir():
            FindFiles(file.path)

def isExcludedFile(filename):

    exclusionPatterns = [
        r'\bPDA\d{8}\b',      # PDA followed by 8 digits
        r'\bDA\d{8}\b',       # DA followed by 8 digits
        r'\bCPC\b',           # CPC keyword
        r'\bcp\b',            # cp keyword
        r'\bPLANS\b',         # PLANS keyword
        r'\bRAD\b'            # RAD keyword
    ]
    for pattern in exclusionPatterns:
        if re.search(pattern, filename, re.IGNORECASE):
            return True # excluded
    return False # not

def isInaccurateMatch(filename, keyword):
    datePatterns = [
        r'\b(0[1-9]|[12][0-9]|3[01])[01][0-9]\d{4}\b',  # MMDDYYYY or DDMMYYYY
        r'\b(19|20)\d{2}[01][0-9][0-3][0-9]\b'            # YYYYMMDD
    ]
    for pattern in datePatterns:
        if re.search(pattern, filename):
            if keyword in re.findall(r'\d{5}', filename):
                return True # make sure keyword isnt from 01012010 

    sixDigitPatterns = [r'\b\d{6}\b']  # make sure keyword not from 100033 
    for pattern in sixDigitPatterns:
        matches = re.findall(pattern, filename)
        for match in matches:
            if keyword in match:
                return True

    return False # accurate match

def EditExcel(number, address, col, workbook):
    # Open file and sheets
    try:
        wb = openpyxl.load_workbook(workbook, data_only=False)
        ws = wb["Sheet1"]
        ws2 = wb['FilesNotInExcel']  # Different sheet on same file
    except Exception as e:
        print(f"File '{workbook}' could not be opened: {e}")
        return

    numberComp = False  

    # Loop through rows
    for row in ws.iter_rows(min_col=1, max_col=1):
        cell = row[0]
        cellValueAsString = str(cell.value).strip()
        numberAsString = str(number).strip()

        if cellValueAsString == numberAsString:  # Check if the number is in Column A
            # Check to see if address already in row
            for c in ws[row[0].row][col-1:]:
                if c.value == address:
                    print(f"Exact hyperlink already exists in row {cell.row}. Skipping addition.")
                    return  # Skip adding the hyperlink if it already exists

            if 'TechAffairsDeterminations' in address:
                ShiftHyperlinksForTechAffairsPriority(ws, cell.row, col)

            writeColumn = FindLastNonHyperlinkColumn(ws, col, cell.row)
            cellFunction(ws, cell.row, writeColumn, address)
            numberComp = True
            print(f"Successfully added {address} to {number}")
            break

    # If the number is not in Column A, add it to another sheet
    if not numberComp:
        row2 = ws2.max_row + 1
        cellFunction(ws2, row2, col, address)
        addControlNumber(ws2, row2, number)
        print(f"{number}: Number does not exist in the excel, added to Another Excel")

    # Save the workbook directly
    try:
        wb.save(workbook)
    except Exception as e:
        print(f"Failed to save workbook {workbook}: {e}")

def ShiftHyperlinksForTechAffairsPriority(sheet, row, col):
    editCol = col
    while sheet.cell(row=row, column=editCol).value is not None:
        editCol += 2

    # Shift the cells to the right by two columns
    while editCol > col:
        sheet.cell(row=row, column=editCol).value = sheet.cell(row=row, column=editCol - 2).value
        sheet.cell(row=row, column=editCol).hyperlink = sheet.cell(row=row, column=editCol - 2).hyperlink
        editCol -= 2

    # Clear the initial cell
    sheet.cell(row=row, column=editCol).value = None
    sheet.cell(row=row, column=editCol).hyperlink = None

def FindLastNonHyperlinkColumn(sheet, col, rowC):
    editThisCol = col

    # Finds Column with Row 1 labeled "Hyperlinks"
    for cell in sheet[1]:
        if cell.value == "Hyperlinks":
            editThisCol = cell.column
            break

    # Checks if the corresponding cell is empty
    while True:
        cell = sheet.cell(row=rowC, column=editThisCol)
        if cell.hyperlink is None and cell.value is None:
            return editThisCol
        editThisCol += 2

    return col

def cellFunction(sheet, rowC, col, address):
    correctedAddress = address.replace('/', '\\')
    cell = sheet.cell(row=rowC, column=col)
    cell.value = correctedAddress
    if not CheckForHashtag(correctedAddress):  # Adds hyperlink if it does not have hashtag
        cell.hyperlink = address
        sheet.cell(row=rowC, column=col).style = "Hyperlink"

def addControlNumber(sheet, rowC, number):
    if number:
        sheet.cell(row=rowC, column=1).value = number

def CheckForHashtag(path):
    return '#' in path

def main():
    directories = [
        "I:/TechAffairsDeterminations",
        "I:/DETERMINATIONS",
        "I:/Technical Affairs Projects"
    ]
    for directory in directories:
        FindFiles(directory)

if __name__ == "__main__":
    main()
