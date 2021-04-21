'''
本程序旨在从 'QTP汇总-电测组' 中提取出 '测试系统展示名称' 和  'PDM系统测试项目名称' 的一一对应关系
方便对日后每日展示数据的分类做一个自动化地写入
'''
import pandas as pd


class DataTransformation:
    def __init__(self, path):
        dataframe = pd.read_excel(io=path, sheet_name='测试项目', )
        dataframe = dataframe.fillna(method='ffill', axis=0, )
        # 将index改成从2开始
        dataframe.index = range(2, len(dataframe) + 2)
        self.dataframe = dataframe

    # 查询excel表中第i行的数据,且i>=2
    def index_loc(self, i):
        print(self.dataframe.loc[i])

    def my_mapping(self):
        n = '测试项目编码'
        m = 'PDM系统测试项目名称'
        for i in range(2, len(self.dataframe) + 2):
            labitemcode = self.dataframe.loc[i, n]
            labitemname = self.dataframe.loc[i, m]

    def classification(self, LABITEMCODE, LABITENAME, ):
        if():
            pass

if __name__ == '__main__':
    Path = r'C:\Users\yucheng2.zhou\Desktop\QTP汇总-电测组.xlsx'
    df = DataTransformation(Path)
    df.classification()
