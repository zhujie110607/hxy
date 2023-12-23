import pandas as pd
import xlwings as xw
import common

path1 = common.select_excel_file('请选择考勤表文件')
if path1 is None:
    print('未选择考勤表文件')
    while True:
        if input('是否退出程序？(y/n)').lower() == 'y':
            break
    exit()
path2 = common.select_excel_file('加班单文件')
if path2 is None:
    print('未选择加班单文件')
    while True:
        if input('是否退出程序？(y/n)').lower() == 'y':
            break
    exit()
print('程序正在执行，请暂停其它操作，等待程序提示执行完成......')
try:
    df = pd.read_excel(path1, sheet_name='考勤表模板', skiprows=5,
                       usecols=[2] + list(range(10, 41)))
    df = pd.read_excel(path1, skiprows=5,
                       usecols=[2] + list(range(10, 41)))
    df.drop(0, inplace=True)  # 删除第一行
    df.drop(len(df), inplace=True)  # 删除最后一行

    df.dropna(axis=1, how='all', inplace=True)  # 删除所有空列
    df.drop(df.index[df['姓名'].last_valid_index() + 1:], inplace=True)  # 获取“姓名”列中最后一个有效(非空)值的索引,然后删除从该索引开始的行
    df.set_index('姓名', inplace=True)  # 设置索引
    for i in range(0, len(df) - 1, 2):
        df.iloc[i] = df.iloc[i + 1]
    df.reset_index(level='姓名', inplace=True)  # 重置索引
    df.dropna(subset=['姓名'], inplace=True)  # 删除姓名列中所有为空的空行

    # 用xlwings获取text.xlsx文件
    wb = xw.Book(path2)
    no_Name = []
    flag = False
    for name in df['姓名']:
        ds = df[df['姓名'] == name]  # 选取王亚杰的考勤表
        ds = ds.unstack().reset_index(level=1, drop=True)  # 转置
        ds.drop('姓名', inplace=True)
        # 删除值为空的行
        ds.dropna(inplace=True)
        # 转换为DataFrame
        ds = pd.DataFrame(ds).reset_index()
        ds.columns = ['日期', '加班工时']
        flag = False
        # 把数据写入excel
        for sht in wb.sheets:
            if sht.name == name:
                flag = True
                sht.range("B17:B47").value = ''
                sht.range("D17:D47").value = ''
                # 把日期写入A列，把加班工时写入C列
                for i in range(len(ds)):
                    sht.range('A' + str(i + 17)).value = name
                    sht.range('B' + str(i + 17)).value = ds['加班工时'][i]
                    sht.range('C' + str(i + 17)).value = sht.range('C17').value
                    sht.range('D' + str(i + 17)).value = ds['日期'][i]
                if i < 30:
                    sht.range('A' + str(i + 18) + ':D47').value = ''
                break
        if flag == False:
            no_Name.append(name)

    if len(no_Name) > 0:
        print('未找到以下姓名在加班单中对应的考勤表：\n' + '\n'.join(no_Name))
    print('程序已执行完毕')
    wb.save()
    # 等待用户输入，如果用户输入y，则退出程序，否则继续
    while True:
        if input('是否退出程序？(y/n)').lower() == 'y':
            break
except Exception as e:
    print(e)
    while True:
        if input('是否退出程序？(y/n)').lower() == 'y':
            break
