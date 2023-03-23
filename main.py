import win32com.client

def oddToNormalCsv(inputFileName, outputFileName, fileFormat = 6):
    '''This program uses Excel to convert the odd file to a normal file'''
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    workbook = excel.Workbooks.Open(inputFileName)
    workbook.SaveAs(outputFileName, fileFormat)
    workbook.Close()
    excel.Quit()

if __name__ == '__main__':
    oddToNormalCsv(input("Odd CSV file Path (use '\\' as deliminator):\n"),input("Normal CSV file Path (use '\\' as deliminator):\n"))