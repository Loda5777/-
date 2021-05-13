"""
本程序旨在从 'QTP汇总-电测组' 中提取出 '测试系统展示名称' 和  'PDM系统测试项目名称' 的一一对应关系
方便对日后每日展示数据的分类做一个自动化地写入
"""
import pandas as pd
import datetime
from datetime import timedelta

'''
dataframe 用来读取对应关系
data      用来读取数据库提取出来的数据
dic       做存放对应关系的字典
'''


class DataTransformation:
    def __init__(self, ):
        self.dataframe = DataTransformation.diance()
        self.data = pd.read_excel(io=DataTransformation.path_name(), sheet_name='Sheet1', )
        self.dic = {}
        self.my_mapping()
        self.data_preprocessing()
        self.classification()
        self.excel_output()
        self.excel_output_for_powerbi()
        DataTransformation.latest_refresh_time()

    # 确认电测文件路径
    @staticmethod
    def diance():
        diance_path = r'C:\Users\yucheng2.zhou\Desktop\QTP汇总-电测组.xlsx'
        dataframe = pd.read_excel(io=diance_path, sheet_name='测试项目', )
        dataframe = dataframe.fillna(method='ffill', axis=0, )
        # 将index改成从2开始
        dataframe.index = range(2, len(dataframe) + 2)
        return dataframe

    # 自动读取今日份的数据,需确认数据存放路径是否一致
    @staticmethod
    def path_name():
        path_head = r"D:\oracle_xlsx\\"
        time_today = str(DataTransformation.date_today())
        path_suffix = ".xlsx"
        pathname = path_head + time_today + path_suffix
        print('数据读取路径已生成', pathname)
        return pathname

    # 今日时间
    @staticmethod
    def date_today():
        return datetime.date.today()

    # 明日时间
    @staticmethod
    def date_tomorrow():
        tomorrow = datetime.date.today() + timedelta(days=+1)
        return tomorrow

    # 第三日时间
    @staticmethod
    def data_theDayAfterTomorrow():
        theDayAfterTomorrow = datetime.date.today() + timedelta(days=+2)
        return theDayAfterTomorrow

    # 当前时间
    @staticmethod
    def latest_refresh_time():
        refresh_time = datetime.datetime.now().replace(microsecond=0)
        print("当前时间", refresh_time)
        return refresh_time

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
    # TODO 给一个可更新字典储存
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
        self.data = self.data[['测试项目', '测试人员', '项目编号', '项目阶段', '批次号', '机型', '机芯', '计划开始时间', '计划结束时间', '最新测试结果',
                               '备注', '任务状态', '测试编号', ]]
        self.data = self.data.sort_values('测试项目')

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
        time_str = str(DataTransformation.date_today())
        suffix = ".xlsx"
        string = address_str + time_str + suffix
        print('已生成', string)
        self.data.to_excel(string, sheet_name=time_str, index=False, )

    def excel_output_for_powerbi(self):
        address_str = r"D:\data_for_PowerBi\\"
        new_str = "近三日测试数据"
        suffix = ".xlsx"
        string = address_str + new_str + suffix
        df1 = self.data[self.data['计划结束时间'] == (str(DataTransformation.date_today()))]
        df2 = self.data[self.data['计划结束时间'] == (str(DataTransformation.date_tomorrow()))]
        df3 = self.data[self.data['计划结束时间'] == (str(DataTransformation.data_theDayAfterTomorrow()))]
        with pd.ExcelWriter(string) as writer:
            df1.to_excel(writer, sheet_name='今日数据', index=False)
            df2.to_excel(writer, sheet_name='明日数据', index=False)
            df3.to_excel(writer, sheet_name='后天数据', index=False)
        print('已生成', string)


if __name__ == '__main__':
    df = DataTransformation()
