# -*- coding: utf-8 -*-
import os
import pandas as pd
import config


def autochecklist():
    model = pd.ExcelFile(config.checklist)
    select = pd.ExcelFile(config.menu)

    sheetname = select.parse('menu')
    a = sheetname.loc[:, ['功能点1', '功能点2']].values
    antidict = {}
    for i in a:
        antidict[i[1]] = i[0]
    dict = {}
    for i in a:
        if i[1] in model.sheet_names:
            temp = model.parse(i[1])
        else:
            print('请检查用例切页是否命名正确')
            return 0
        dict[i[1]] = temp.values
    No = []  # 编号列
    func1 = []  # 功能点一
    func2 = []  # 功能点二
    condition = []  # 前提条件
    operation = []  # 操作步骤
    expect = []  # 期待结果
    severity = []  # 优先级
    self = []  # 自测标记
    listtype = []  # 用例类型
    desc = []  # 备注

    cal = 0
    for p2 in dict:
        p1 = antidict[p2]
        for i in dict[p2]:
            cal += 1
            No.append(cal)
            func1.append(p1)
            func2.append(p2)
            condition.append(i[0])
            operation.append(i[1])
            expect.append(i[2])
            severity.append('')
            self.append('')
            listtype.append('')
            desc.append('')

    dic1 = {
        '编号': No,
        '功能点1': func1,
        '功能点2': func2,
        '前提条件': condition,
        '操作步骤': operation,
        '期望结果': expect,
        '优先级': severity,
        '项目自测': self,
        '用例类型': listtype,
        '备注': desc,
    }
    df = pd.DataFrame(dic1)
    df.to_excel(config.output + config.name, index=False)

if __name__ == '__main__':
    autochecklist()
