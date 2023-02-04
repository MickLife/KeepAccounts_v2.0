"""
Author:     MickLife
Project:    https://github.com/MickLife/KeepAccounts_v2.0
versions:   4.0  2023/01/15: 框架重整，新特性：分类自动标注
            3.0  2022/08/03: 增加分类关键词sheet，能够检测关键词给每笔账分类
            2.1  2021/12/29: 修复支付宝理财收支逻辑bug
"""
import time

from src.config import GlobalConfig
from src.AccountBook import AccountBook


def main():
    config = GlobalConfig()
    data = AccountBook(config)

    # 追加数据
    data.read_new_records()
    data.data_filtering()
    data.amount_confirm()

    # 标注
    data.data_to_append = data.annotator.do_annotation(data.data_to_append)
    # todo: 确认金额、分类结果，不正确的实现人工订正
    # todo: 人工/追加其他来源数据

    # 数据存储
    data.append()
    if config.OVERWRITE:
        data.save_database(data.save_path)
    else:
        new_path = f"{config.DATABASE_PATH[:-5]}_{time.strftime('%Y%m%d_%H%M%S', time.localtime())}.xlsx"
        data.save_database(new_path)

    # todo: 数据统计/可视化(另写py)


if __name__ == '__main__':
    main()
