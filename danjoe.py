import random
import xlsxwriter

workbook = xlsxwriter.Workbook('DanjoeDivision4.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

division = workbook.add_format({'num_format': '÷#,##0'})

divisor = []
dividend = []
answer = []
while len(divisor) <1000:
    dsr = random.randint(5,11)
    dvd= random.randint(500, 9000)
    if (dvd%dsr) == 0:
        ans = dvd/dsr
        divisor.append(dsr)
        dividend.append(dvd)
        answer.append(ans)

collection = zip(dividend, divisor, answer)

row = 1
col = 0
linelength = 0
num = 1

# Iterate over the data and write it out row by row.
for divd,divr,ansr in (collection):
    worksheet.write(row, col, num)
    worksheet.write(row, col+1,     divd)
    worksheet.write(row+1, col+1, divr, division)
    worksheet.write(row+2, col+1, ansr, bold)
    worksheet.write(row + 2, col , "ANS", bold)
    num +=1
    col += 2
    linelength+=1
    if linelength >4:
        linelength = 0
        row +=4
        col = 0



workbook.close()



