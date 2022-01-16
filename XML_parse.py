import xlsxwriter
from xml.etree import ElementTree
from lxml import etree
import datetime


def parseXML(xmlFile,exelFile):
    a = datetime.datetime.now()
    with open(xmlFile, 'rb') as f:
        e = etree.iterparse(f, tag="RECORD")

        workbook = xlsxwriter.Workbook(exelFile, {'constant_memory': True})
        worksheet = workbook.add_worksheet()

        row = 0
        col = 0
        for event, element in e:

            if row>1000000:
                worksheet = workbook.add_worksheet()
                print("new sheet")
                row=1
            for c in element:

                if row == 1:
                    worksheet.write(row, col, c.tag)
                    col += 1
                if row > 1:
                    worksheet.write(row, col, c.text)
                    col += 1
            element.clear()
            row += 1
            col = 0

            if row % 100000 == 0:
                print(row)

        workbook.close()
    b = datetime.datetime.now()
    print(b - a)


if __name__ == "__main__":
    input_file=input(r"Введите путь и имя входящего файла, например-- С:\\user\\download\\example.xml: ")
    output_file=input(r"Введите путь и имя исходящего файла, например  С:\\user\\download\\example.xlsx: ")
    parseXML(input_file,output_file)