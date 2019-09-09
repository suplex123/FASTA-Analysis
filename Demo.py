#=========================================================================
#此函数用于打印J，K列指示的碱基序列
def counting(str_in , str_1 , str_2):
    f = open(str_in)
    ls = []
    for line in f:
        if not line.startswith('>'):
            ls.append(line.replace('\n', ''))  # 去掉行尾的换行符
    f.close()
    _str = ''
    for i in range(0, len(ls)):
        _str += ls[i]#_str是fa文件的完整碱基序列
    arr_1 = getRange(str_1)
    arr_2 = getRange(str_2)

    for i in range(len(arr_1)):
        #print((len(_str),arr_1[i],arr_2[i]))
        print("第" + str(i + 1) + "个碱基序列是")
        print(_str[arr_1[i]:arr_2[i]])
        print("\n其中一共有:")
        countATGC(_str[arr_1[i]:arr_2[i]])


#=========================================================================
#此函数用于统计碱基序列中的A，T，G，C个数
def countATGC(str_in):
    A = C = T = G = 0
    A += str_in.count('A')
    C += str_in.count('C')
    T += str_in.count('T')
    G += str_in.count('G')
    print('A:',A,'\n','C:',C,'\n','T:',T,'\n','G:',G,'\n')


# ========================================================================
#此函数用于去掉J,K列中的‘，’并转换成整型数组

def getRange(str_in):
    arr = str_in.split(',')
    arr = [int(val) for val in arr[0:-1]]
    return arr


#=========================================================================
#
import xlrd

FileContactList = "data.xlsx"
FileName = FileContactList

while True:
    KeyStr = input("请输入要搜索的基因名称: ")
    FileObj = xlrd.open_workbook(FileName)
    sheet = FileObj.sheets()[0]  # 获取第一个工作表
    row_count = sheet.nrows  # 行数
    col_count = sheet.ncols  # 列数    # 搜索关键字符串
    i = 1
    for element in range(row_count):
        if KeyStr.lower() in (str(sheet.row_values(element))).lower():

            str_1 = str(sheet.cell(element, 9).value.encode('utf-8')).lstrip('b\'').rstrip('\'')
            str_2 = str(sheet.cell(element, 10).value.encode('utf-8')).lstrip('b\'').rstrip('\'')
            str_Path = "./chromFa/" + str(sheet.cell(element, 2).value.encode('utf-8')).lstrip('b\'').rstrip(
                '\'') + ".fa"
            #print(str_Path)
            #print(str_1)
            #print(str_2)
            print("==>第" + str(i) + "个" + KeyStr)
            print("剪切变体是:")
            print(str(sheet.cell(element,2).value.encode('utf-8')).lstrip('b\'').rstrip('\''))
            counting(str_Path,str_1,str_2)
            i += 1
# =========================================================================
