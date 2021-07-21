import string
import urllib.request
from PIL import Image
import win32com.client as win32

import xlsxwriter
import csv

with open('F:\Logos\Logos\Logos.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 1
    cols = list(string.ascii_uppercase)

    workbook = xlsxwriter.Workbook('images.xlsx')
    worksheet = workbook.add_worksheet()

    for row in csv_reader:
        if line_count == 1:
            cell_format = workbook.add_format({'bold': True})
            worksheet.write("A1", row[0], cell_format);
            worksheet.write("B1", row[1], cell_format);
            worksheet.write("C1", row[2], cell_format);
            worksheet.write("D1", row[3], cell_format);
            line_count += 1
        else:
            abbrevation = row[1]
            companyName = row[2]
            typeOfProducts = row[3]
            url= row[0]
            url = url[1:-1]
            url = url.split(", ")
            fileNames =[]
            c=0
            for i in url:

                try:
                    urltouse = i[1:-1]
                    data = urllib.request.urlopen(urltouse).read()
                    filename = cols[c] + str(line_count) + '.jpg'
                    file = open(filename, "wb")
                    file.write(data)
                    file.close()
                    fileNames.append(filename)
                    c+=1

                except Exception as e:
                    print(e)
                    print(line_count)

            if len(fileNames) > 0 :
                images = [Image.open(x) for x in fileNames]
                widths, heights = zip(*(i.size for i in images))
                total_width = sum(widths)+45
                max_height = max(heights)

                new_im = Image.new('RGB', (total_width, max_height),'white')

                x_offset = 0
                for im in images:
                    new_im.paste(im, (x_offset, 0))
                    x_offset += im.size[0]+15

                c=0
                filename = str(line_count)+'.jpg'
                new_im.save(filename)
                with Image.open(filename) as img:
                    image_width = img.width
                    image_height = img.height
                worksheet.insert_image(cols[c] + str(line_count), filename)
                worksheet.set_row_pixels(line_count - 1, image_height + 10)
                worksheet.set_column_pixels(line_count - 1, image_width + 30)

                c += 1
                worksheet.write(cols[c] + str(line_count), abbrevation);
                c += 1
                worksheet.write(cols[c] + str(line_count), companyName);
                c += 1
                worksheet.write(cols[c] + str(line_count), typeOfProducts);
                c += 1

            line_count += 1

    workbook.close()


    print(f'Processed {line_count} lines.')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open('C:\\Users\\dteja\\PycharmProjects\\pythonProject1\\images.xlsx')
    ws = wb.Worksheets("Sheet1")
    ws.Columns.AutoFit()
    wb.SaveAs("C:\\Users\\dteja\\PycharmProjects\\pythonProject1\\imagesfinal.xlsx")
    wb.Close()
