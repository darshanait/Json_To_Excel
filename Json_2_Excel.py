import json
import xlsxwriter
# with open('E:/Disguise Programs/5. Json _to_Excel/dummy1.json', 'r') as myfile:
jsonString=open('C:/Users/darshan/Desktop/Json_To_excel/Json_files/Fruit_json.json').read()
    


d=json.loads(jsonString)



workbook = xlsxwriter.Workbook('Output/Fruit.xlsx')
worksheet = workbook.add_worksheet()

rowNum=-1
for i in json.loads(jsonString):
    # print('filename in',i)
    # print(i)
    data=d[i]
    print('Filename: ',data['filename'])
    rowNum=rowNum+1
                        # the row number for excel to write
    worksheet.write(rowNum,0,data['filename'])
    # print('Regions:',data['regions'])
    listData=data['regions']

    # print(len(listData))
    print(listData)
    for i in range(len(listData)):
        # print(i)
        # print(len(listData[i]))
        dictData=listData[i]
        # print('dictionary data',dictData)
        print('dar',dictData['region_attributes']['Fruit_type'])
        Fruit=dictData['region_attributes']['Fruit_type']
   
        if Fruit=='Banana':
            worksheet.write(rowNum,1,Fruit)
        if Fruit=='Mango':
            worksheet.write(rowNum,2,Fruit)
        if Fruit=='Pineapple':
            worksheet.write(rowNum,3,Fruit)
        
#worksheet.save('Fruit.xlsx')
    # for k in i:
        
workbook.close()