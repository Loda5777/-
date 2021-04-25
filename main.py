"""
本程序旨在从 'QTP汇总-电测组' 中提取出 '测试系统展示名称' 和  'PDM系统测试项目名称' 的一一对应关系
方便对日后每日展示数据的分类做一个自动化地写入
"""
import pandas as pd
import datetime


class DataTransformation:
    def __init__(self, path):
        dataframe = pd.read_excel(io=r'C:\Users\yucheng2.zhou\Desktop\QTP汇总-电测组.xlsx', sheet_name='测试项目', )
        dataframe = dataframe.fillna(method='ffill', axis=0, )
        # 将index改成从2开始
        dataframe.index = range(2, len(dataframe) + 2)
        self.dataframe = dataframe
        # TODO 还需加一个自动寻取今日文件名的函数
        self.data = pd.read_excel(io=path, sheet_name='Sheet1', )
        # print(self.data)
        self.dic = {}

    # 查询电测组QTP表中第i行的数据，且i>=2
    def index_ori_loc(self, i):
        if 2 <= i <= (len(self.dataframe) + 2):
            print(self.dataframe.loc[i])
        else:
            print("请确认输入行数无误")

    # 查询数据库导出表中第i行的数据，且i>=2
    def index_loc(self, i):
        if 2 <= i <= (len(self.data) + 2):
            print(self.data.loc[i])
        else:
            print("请确认输入行数无误")

    # 用字典定义一个对应关系
    def my_mapping(self, ):
        n = '测试项目编码'
        m = '测试系统展示名称'
        for i in range(2, len(self.dataframe) + 2):
            self.dic[self.dataframe.loc[i, n]] = self.dataframe.loc[i, m]

    # 打印字典存储的对应关系
    def print_map(self, ):
        for key in self.dic:
            print(key, self.dic[key])

    # 数据预处理，去除无效数据和数据的重命名
    def data_preprocessing(self):
        self.data.drop(['TRUSTAPPLYID', 'TRUSTAPPLYNO', 'LABGROUPNAME', 'LABITEMID', 'PROJMANAGER', 'STRUCTTYPE', 'QTY',
                        'EXPERIMENTTASKID',
                       'PLANEND', 'DELETEFLAG', 'CHGCOUNT', ], axis=1, inplace=True, )
        self.data.rename(
            columns={'SOURCEPROJCODE': '项目编号', 'TESTPHASENAME': '项目阶段', 'LABITEMNAME': '测试项目', 'SOURCEPROJNAME': '机型',
                     'CORENAME': '机芯', 'PANELMODEL': '屏型号', 'POWER': '电源', 'HARDLEADER': '硬件Leader',
                     'PJMSIGNER': 'PJM会签',
                     'QTPMAKER': 'QTP制作', 'TESTHOURS': '需求测试时长', 'PLANSTART': '计划开始时间', 'PLANENDTIME': '计划结束时间',
                     'NEWTESTRESULT': '最新测试结果', 'REQUIRECOMPLETEDATE': '需求完成时间', 'DATAENTERMAN': '测试人员',
                     'TESTCHECKTOR': '审核人',
                     'ENDTIME': '实际完成时间', 'TESTRESULT': '测试结果', 'TASKREGISTERNO': '任务编号', 'BATCHNO': '批次号',
                     'UNQUALIFIEDDESC': '备注', 'TASKSTAT': '任务状态', 'LABITEMCODE': '测试编号', }, inplace=True, )

        self.data.drop(['屏型号', '电源', '硬件Leader', 'PJM会签', 'QTP制作', '审核人', '任务编号', '测试结果', '需求测试时长', '需求完成时间',
                        '实际完成时间', ], axis=1, inplace=True, )
        # TODO 列名的重新排序
        self.data = self.data[['测试项目', '测试人员', '项目编号', '项目阶段', '批次号', '机型', '机芯', '计划开始时间', '计划结束时间', '最新测试结果',
                               '备注', '任务状态', '测试编号', ]]
        self.data = self.data.sort_values('测试项目')
        # TODO index按项目数索引
        # self.data

    # TODO 需要验证是否有因为测试编号和测试项目名称不规范造成测试项误删的情况
    # 对测试项目名称进行一个复写
    def classification(self, ):
        for i in range(0, len(self.data)):
            to_labitemcode = self.data.loc[i, '测试编号']
            to_labitemname = self.data.loc[i, '测试项目']
            # TODO 测试项目名的获取
            # if (to_labitemcode in self.dic) and (to_labitemname in self.dic):
            if to_labitemcode in self.dic:
                self.data.loc[i, '测试项目'] = self.dic[to_labitemcode]
            else:
                self.data.drop(labels=i, axis=0, inplace=True, )
        self.data.drop('测试编号', axis=1, inplace=True,)

    # 将处理的数据写入新表中
    def excel_output(self):
        address_str = r"D:\data_for_PowerBi\\"
        time_str = str(datetime.date.today())
        suffix = ".xlsx"
        string = address_str + time_str + suffix
        print('已生成', string)
        self.data.to_excel(string, sheet_name=time_str, index=False, )

    def excel_output_for_powerbi(self):
        address_str = r"D:\data_for_PowerBi\\"
        new_str = "今日测试数据"
        suffix = ".xlsx"
        string = address_str + new_str + suffix
        print('已生成', string)
        self.data.to_excel(string, sheet_name=new_str, index=False, )


if __name__ == '__main__':
    df = DataTransformation(path=r'D:\oracle_csv\2021-04-25.xlsx', )
    # df.index_ori_loc(91)
    df.my_mapping()
    # df.index_loc(20)
    df.data_preprocessing()
    df.classification()
    df.excel_output()
    df.excel_output_for_powerbi()
