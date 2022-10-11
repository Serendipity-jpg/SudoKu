import xlrd
import xlsxwriter
from homework1 import SudoKu


def readExcel(filename, row, col):
    '''
    读取excel文件并转为list返回
    :param filename: 文件名
    :param row:
    :param col:
    :return:
    '''
    workbook = xlrd.open_workbook(filename)
    sheetData = workbook.sheet_by_name('Sheet1')
    # print(sheetData)
    ls = []
    for i in range(row, row + 21):
        lt = []
        for j in range(col, col + 21):
            if sheetData.cell(i, j).value == 0.0:
                lt.append('')

            elif sheetData.cell(i, j).value == '':
                lt.append('*')
            else:
                lt.append(int(sheetData.cell(i, j).value))
        ls.append(lt)
    return ls


def showList(ls):
    """
    打印列表ls
    :param ls:列表ls
    :return:
    """
    print('[')
    for row in ls:
        print(row)
    print(']')


def repalceList(ls, row, col):
    """
    从指定的ls列表取出从（row，col）取出一个9*9的子列表先完成数独并更新
    :param ls: 21*21的列表
    :param row: row坐标
    :param col: col坐标
    :return:
    """
    left_top = []
    for i in range(row, row + 9):
        left_top.append(ls[i][col:col + 9])
    # if row == 12 and col == 12:
    #     showList(left_top)
    # 得到计算好的9*9标准数独
    left_top = SudoKu.SudoKu(left_top).get_result()
    # 判断生成的9*9标准数独是否有效
    # print(SudoKu.judge_sudo_ku_is_legal(left_top))
    # 试探出的9*9标准数独有效，则更新21*21列表的对应模块
    if SudoKu.judge_sudo_ku_is_legal(left_top):
        for i in range(row, row + 9):
            ls[i][col:col + 9] = left_top[i - row]


def writeExcel(filename, ls):
    """
        读取excel文件并转为list返回
        :param ls:
        :param filename: 文件名
        :return:
    """
    workbook = xlsxwriter.Workbook(filename)  # 创建工作簿
    worksheet = workbook.add_worksheet("Sheet1")  # 创建子表
    worksheet.activate()  # 激活工作表
    i = 1  # 从第1行开始写入数据
    while i <= len(ls):
        # 每行的第一列的位置
        row = 'A' + str(i)
        lt = ls[i - 1]
        for j in range(len(lt)):
            if lt[j] == '*':
                lt[j] = ''
        worksheet.write_row(row, lt)
        i += 1
    workbook.close()  # 关闭excel文件


if __name__ == '__main__':
    # 从Excel文件读取出21*21的矩阵
    ls = readExcel("data/input1.xlsx", 0, 0)
    print("input:")
    showList(ls)
    # print(showList(ls))
    repalceList(ls, 0, 0)
    repalceList(ls, 0, 12)
    repalceList(ls, 6, 6)
    repalceList(ls, 12, 0)
    repalceList(ls, 12, 12)
    print("output:")
    showList(ls)
    #结果写入Excel
    # writeExcel("data/output2.xlsx", ls)

