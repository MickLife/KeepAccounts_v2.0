import os

import pandas as pd
import numpy as np
from openpyxl.utils import get_column_letter

from src.Annotator import Annotator


class AccountBook:
    def __init__(self, cfg):
        self.cfg = cfg
        self.save_path = cfg.DATABASE_PATH
        self.dtype = {'日期时间': 'datetime64', '来源': str, '收支': str, '支付状态': str, '交易对方': str, '商品': str,
                      '金额': float, '修订金额': float, '分类1': str, '分类2': str}
        self.column_name = list(self.dtype.keys())
        database = self.read_database(cfg.DATABASE_PATH)
        database = self._format_arrange(database)
        self.database = database
        self.data_to_append = pd.DataFrame()
        self.annotator = Annotator()

    def __str__(self):
        print('AccountBook:')
        pd.set_option('display.unicode.ambiguous_as_wide', True)
        pd.set_option('display.unicode.east_asian_width', True)
        pd.set_option('display.max_columns', 1000)
        pd.set_option('display.max_rows', None)
        pd.set_option('max_colwidth', 30)
        pd.set_option('display.width', 1000)
        print(self.database)
        return "The AccountBook has {} rows and {} columns.".format(self.database.shape[0], self.database.shape[1])

    def read_database(self, path: str) -> pd.DataFrame:
        cols = list(range(0, len(self.dtype.keys())))
        df = pd.read_excel(path, sheet_name='Records', usecols=cols)
        df.astype(self.dtype)
        return df

    def save_database(self, path: str):
        writer = pd.ExcelWriter(path)  # ！覆盖写文件！
        self.database.to_excel(writer, sheet_name='Records', index=False)
        widths = np.array([22, 10, 10, 20, 40, 40, 15, 15, 20, 25])  # 设置excel单元格列宽
        worksheet = writer.sheets['Records']
        for i, width in enumerate(widths, 1):
            worksheet.column_dimensions[get_column_letter(i)].width = width
        writer.save()
        print('Saved AccountBook to dir: {}'.format(path))

    def append(self):
        self.database = pd.concat([self.database, self.data_to_append], axis=0, ignore_index=True)

    def build_alipay_records(self, path: str):
        print('Reading file: {}'.format(path))
        df_zfb = pd.read_csv(path, header=4, skipfooter=7, encoding='gbk', engine='python')
        cols = [2, 10, 11, 7, 8, 9]
        df_zfb = df_zfb.iloc[:, cols]
        df_zfb.insert(1, '来源', '支付宝', allow_duplicates=True)
        df_zfb.insert(7, '修订金额', 0.0, allow_duplicates=True)
        df_zfb.insert(8, '分类1', '', allow_duplicates=True)
        df_zfb.insert(9, '分类2', '', allow_duplicates=True)
        df_zfb.columns = self.column_name
        df_zfb = self._format_arrange(df_zfb)
        print('There are {} AliPay records have been loaded.'.format(df_zfb.shape[0]))
        return df_zfb

    def build_wechat_records(self, path: str):
        print('Reading file: {}'.format(path))
        df_wx = pd.read_csv(path, header=16, skipfooter=0, encoding='utf-8')
        cols = [0, 4, 7, 2, 3, 5]
        df_wx = df_wx.iloc[:, cols]
        df_wx = df_wx.applymap(lambda x: x.strip().strip('¥') if isinstance(x, str) else x)  # 去除金额前的¥符号
        df_wx.insert(1, '来源', '微信', allow_duplicates=True)
        df_wx.insert(7, '修订金额', 0.0, allow_duplicates=True)
        df_wx.insert(8, '分类1', '', allow_duplicates=True)
        df_wx.insert(9, '分类2', '', allow_duplicates=True)
        df_wx.columns = self.column_name
        df_wx = df_wx.drop(df_wx[df_wx['收支'] == '/'].index)
        df_wx = self._format_arrange(df_wx)
        print('There are {} WeChat records have been loaded.'.format(df_wx.shape[0]))
        return df_wx

    def _format_arrange(self, df):
        df.astype(self.dtype)
        df['日期时间'] = df['日期时间'].astype(self.dtype['日期时间'])
        df['支付状态'] = df['支付状态'].astype(self.dtype['支付状态'])
        df['交易对方'] = df['交易对方'].astype(self.dtype['交易对方'])
        df['金额'] = pd.to_numeric(df['金额'])
        df.round({'金额': 2})
        df['修订金额'] = pd.to_numeric(df['修订金额'])
        df.round({'修订金额': 2})
        return df

    def data_filtering(self):
        df = self.data_to_append
        df.loc[(df['商品'].str.contains('余额宝')) & (df['商品'].str.contains('收益发放')), ['收支']] = '收入'  # 纳入余额宝收益
        # 关键词滤除
        for column_name, keywords in self.cfg.KEYWORD_FILTER.items():
            for word in keywords:
                df = df[~df[column_name].str.contains(word)]  # ~表示不包含
        # 数值滤除
        df = df[df['金额'] != np.nan]
        df = df[df['金额'] - 0 > 1e-3]
        self.data_to_append = df

    def amount_confirm(self):
        df = self.data_to_append
        df.loc[df['收支'].str.contains('收入'), ['修订金额']] = df['金额']
        df.loc[df['收支'].str.contains('支出'), ['修订金额']] = df['金额'] * -1
        self.data_to_append = df

    def _get_new_record_file_names(self):
        path = self.cfg.RECORDS_DIRECTORY
        file_list = os.listdir(path)
        file_list_alipay = []
        file_list_wechat = []
        for file in file_list:
            if 'alipay_record' in file:
                file_list_alipay.append(os.path.join(path, file))
            elif '微信支付账单' in file:
                file_list_wechat.append(os.path.join(path, file))
            elif file == '.old_records':
                continue
            else:
                raise ValueError('请检查账单文件夹中的文件，确保只有1个.old_records目录，以及要导入的微信/支付宝账单！')
        if (len(file_list_alipay) == 0) and (len(file_list_wechat) == 0):
            raise ValueError('请检查账单文件夹RECORDS_DIRECTORY，没有账单要添加！')
        return file_list_alipay, file_list_wechat

    def read_new_records(self):
        csv_path_alipay, csv_path_wechat = self._get_new_record_file_names()
        for alipay_records in csv_path_alipay:
            data_wechat = self.build_alipay_records(alipay_records)
            self.data_to_append = pd.concat([self.data_to_append, data_wechat], axis=0, ignore_index=True)
        for wechat_records in csv_path_wechat:
            data_alipay = self.build_wechat_records(wechat_records)
            self.data_to_append = pd.concat([self.data_to_append, data_alipay], axis=0, ignore_index=True)
