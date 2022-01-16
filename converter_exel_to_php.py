from openpyxl import load_workbook
path=r"/home/office-hpbinotel-588/"
exel= path+"trans.xlx"

dc={}
language={'col':2,'col2':3,'col3':4,'col4':5,'col5':6}
def parse():
    global dc
    global exel
    wb = load_workbook(r"/patch.xlsx")
    ws = wb.active
    sheet = wb['sheet1']
    for k,v in language.items():

        phptrans = open(path + k+'.php', 'w+')
        phptrans.write('<?php' + '\n\n' + '$i18nNoConflict = array()' + '\n\n' + '$i18nNoConflict[\''+ k + '\'] = array(' + '\n')


        for row in sheet.rows:
            if row[0].value==None:
                print('=>')
            elif  row[0].value=='Фраза/слово':
                print('Фраза/слово для: '+k)
            elif row[0].value=='Нет в php документе':
                print('=>')

            else:
                phptrans.write('        '+'\''+str(row[0].value).replace("\\","").replace(r"'",r'\'')+ '\''+' => '+'\''+str(row[int(v)].value).replace("\\","").replace(r"'",r'\'')+'\','+'\n')

        phptrans.write(');')
        phptrans.close()
    wb.close()
if __name__ == "__main__":
    parse()

