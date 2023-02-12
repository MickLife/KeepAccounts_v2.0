import os
import numpy as np
from enum import Enum


class GlobalConfig:
    # DATABASE_PATH 账本路径
    DATABASE_PATH = os.path.join('AccountBooks', 'KeepAccountDataBase.xlsx')

    # RECORDS_DIRECTORY 账单文件夹
    RECORDS_DIRECTORY = 'records'

    # OVERWRITE: True--覆盖写入，False--另存为新账本
    OVERWRITE = True

    # KEYWORD_FILTER: 自定义关键词过滤
    KEYWORD_FILTER = {
        '支付状态': ['退款', '交易关闭', '提现'],
        '收支': ['其他'],  # 舍弃支付宝理财流水（余额宝除外）
        # '交易对方': [],
        # '商品': [],
    }


class MajorLabels(Enum):
    food = '餐饮'
    daily = '日常'
    shopping = '购物'
    fashion = '时尚'
    transport = '交通'
    house = '住房'
    health = '健康'
    study = '学习'
    relation = '人情往来'
    amusement = '娱乐'
    child = '养育子女'
    other_expenses = '其他支出'
    entry = '主要收入'
    finance = '理财'
    gift = '礼金礼物'
    other_incoming = '其他收入'


sep = '_'


class BillLabels(Enum):
    # expend
    food_simple = MajorLabels.food.value + sep + '简餐(29-)'
    food_big = MajorLabels.food.value + sep + '奢华(30~60)'
    food_luxurious = MajorLabels.food.value + sep + '饕餮(60+)'
    food_fruit = MajorLabels.food.value + sep + '水果'
    food_snacks = MajorLabels.food.value + sep + '零食'
    food_drinks = MajorLabels.food.value + sep + '饮料'
    food_meal_replacement = MajorLabels.food.value + sep + '代餐'
    food_cook = MajorLabels.food.value + sep + '做饭开销'
    food_others = MajorLabels.food.value + sep + '其他'

    daily_utilities = MajorLabels.daily.value + sep + '水电燃、网费、电话费'
    daily_use = MajorLabels.daily.value + sep + '生活用品'
    daily_clean = MajorLabels.daily.value + sep + '美容美发、洗澡、洗衣'
    daily_delivery = MajorLabels.daily.value + sep + '快递物流'
    daily_others = MajorLabels.daily.value + sep + '其他'

    shopping_office = MajorLabels.shopping.value + sep + '桌面办公'
    shopping_electronics = MajorLabels.shopping.value + sep + '电子产品'
    shopping_cooking_utensil = MajorLabels.shopping.value + sep + '厨具'
    shopping_furniture = MajorLabels.shopping.value + sep + '家具'
    shopping_appliance = MajorLabels.shopping.value + sep + '家电'
    shopping_bedding = MajorLabels.shopping.value + sep + '床品、纺织'
    shopping_tools = MajorLabels.shopping.value + sep + '工具、线缆、收纳、支架'
    shopping_toys = MajorLabels.shopping.value + sep + '玩具、摆件'
    shopping_others = MajorLabels.shopping.value + sep + '其他'

    fashion_clothes = MajorLabels.fashion.value + sep + '衣服、裤子'
    fashion_underwear = MajorLabels.fashion.value + sep + '内衣、袜子'
    fashion_bag_hats = MajorLabels.fashion.value + sep + '鞋帽箱包'
    fashion_ornament = MajorLabels.fashion.value + sep + '饰品'
    fashion_beauty = MajorLabels.fashion.value + sep + '护肤美妆'
    fashion_others = MajorLabels.fashion.value + sep + '其他'

    transport_public = MajorLabels.transport.value + sep + '公共交通'
    transport_taxi = MajorLabels.transport.value + sep + '打车、租车'
    transport_train_plane = MajorLabels.transport.value + sep + '客车、火车、飞机'
    transport_fuel_toll_park = MajorLabels.transport.value + sep + '车辆燃油、过路费、停车费'
    transport_car_repair = MajorLabels.transport.value + sep + '车辆保险、保养、维修'
    transport_parking_space = MajorLabels.transport.value + sep + '车位'
    transport_others = MajorLabels.transport.value + sep + '其他'

    house_rent = MajorLabels.house.value + sep + '房租'
    house_loan = MajorLabels.house.value + sep + '房贷'
    house_decoration = MajorLabels.house.value + sep + '房屋装修、维修'
    house_move = MajorLabels.house.value + sep + '搬家'
    house_others = MajorLabels.house.value + sep + '其他'

    health_hospital = MajorLabels.health.value + sep + '就医'
    health_medical_product = MajorLabels.health.value + sep + '医疗产品'
    health_insurance = MajorLabels.health.value + sep + '保险'
    health_fitness_sports = MajorLabels.health.value + sep + '健身运动'
    health_care = MajorLabels.health.value + sep + '保健品、推拿、按摩'
    health_others = MajorLabels.health.value + sep + '其他'

    study_tuition = MajorLabels.study.value + sep + '学费、知识付费、课程付费'
    study_copyright_pay = MajorLabels.study.value + sep + '版权费、著作费、软件费'
    study_book = MajorLabels.study.value + sep + '书籍'
    study_stationery = MajorLabels.study.value + sep + '文体用具、材料打印'
    study_instrument = MajorLabels.study.value + sep + '乐器'
    study_others = MajorLabels.study.value + sep + '其他'

    amusement_audio_video = MajorLabels.amusement.value + sep + '影音付费'
    amusement_game = MajorLabels.amusement.value + sep + '游戏付费'
    amusement_app = MajorLabels.amusement.value + sep + 'APP付费'
    amusement_travel = MajorLabels.amusement.value + sep + '旅游：景点门票、纪念品'
    amusement_gathering = MajorLabels.amusement.value + sep + '聚会：KTV、酒吧、桌游'
    amusement_relax = MajorLabels.amusement.value + sep + '休闲：展览、音乐会、演唱会、LiveHouse'
    amusement_others = MajorLabels.amusement.value + sep + '其他'

    relation_lover = MajorLabels.relation.value + sep + '恋人'
    relation_parents = MajorLabels.relation.value + sep + '父母'
    relation_brother_sister = MajorLabels.relation.value + sep + '兄弟姐妹'
    relation_friend = MajorLabels.relation.value + sep + '朋友'
    relation_relative = MajorLabels.relation.value + sep + '亲戚'
    relation_colleague = MajorLabels.relation.value + sep + '同事'
    relation_charity = MajorLabels.relation.value + sep + '慈善捐助'
    relation_others = MajorLabels.relation.value + sep + '其他'

    child_education = MajorLabels.child.value + sep + '子女教育'
    child_others = MajorLabels.child.value + sep + '其他'

    other_expenses_work_expenses = MajorLabels.other_expenses.value + sep + '无法报销的工作开支'
    other_expenses_5_insurance_1_fond = MajorLabels.other_expenses.value + sep + '五险一金'
    other_expenses_individual_income_tax = MajorLabels.other_expenses.value + sep + '个人所得税'
    other_expenses_CPC_membership = MajorLabels.other_expenses.value + sep + '党费'
    other_expenses_others = MajorLabels.other_expenses.value + sep + '其他'

    # income
    entry_salary = MajorLabels.entry.value + sep + '工资'
    entry_part_time = MajorLabels.entry.value + sep + '兼职、副业'
    entry_bonus = MajorLabels.entry.value + sep + '奖金'
    entry_tax_rebate = MajorLabels.entry.value + sep + '退税'
    entry_scholarship = MajorLabels.entry.value + sep + '奖学金、助学金、补助'
    entry_others = MajorLabels.entry.value + sep + '其他'

    finance_interest_on_deposit = MajorLabels.finance.value + sep + '货币基金、存款利息'
    finance_low_risk = MajorLabels.finance.value + sep + '低风险投资'
    finance_high_risk = MajorLabels.finance.value + sep + '高风险投资'
    finance_gold = MajorLabels.finance.value + sep + '黄金'
    finance_stock = MajorLabels.finance.value + sep + '股票'
    finance_insurance = MajorLabels.finance.value + sep + '保险'
    finance_others = MajorLabels.finance.value + sep + '其他'

    gift_from_parents = MajorLabels.gift.value + sep + '父母'
    gift_from_friends = MajorLabels.gift.value + sep + '朋友'
    gift_lucky_money = MajorLabels.gift.value + sep + '压岁钱'
    gift_red_packet = MajorLabels.gift.value + sep + '红包'
    gift_from_social_feedback = MajorLabels.gift.value + sep + '社会回馈'
    gift_others = MajorLabels.gift.value + sep + '其他'

    other_incoming_reimburse = MajorLabels.other_incoming.value + sep + '报销'
    other_incoming_refund = MajorLabels.other_incoming.value + sep + '押金退款'
    other_incoming_others = MajorLabels.other_incoming.value + sep + '其他'


class AutoLabelRules:
    def __init__(self):
        """自动标注规则
        self.rules: Dict{BillLabels: rule}
        rule: List[Condition1, Condition2, ...]  Condition 全部满足, rule才生效
        Condition: Dict{'列名': regex表达式 or 数值范围tuple(min, max)}  只要有一个列的数值匹配, condition就生效

        模板
        self.rule = {
            BillLabels.food_cook:
                [
                    {'收支': '支出|收入'},
                    {'来源': '支付宝|微信'},
                    {'交易对方': '', '商品': ''},
                    {'金额': (0, np.inf)}
                ],
        }
        """
        food_brand = '麦当劳|肯德基|海底捞|和府捞面|无名缘米粉|满座儿|牛肉面|美食|外婆家|西北莜面村|必胜客|潘多拉|' \
                     '火锅|串串|水煮鱼|自助餐|家常菜|川菜|烤肉|韩国料理'

        self.rules = {
            BillLabels.food_simple:
                [
                    {'收支': '支出'},
                    {'交易对方': food_brand,
                     '商品': '北京-环保园M地块|满座儿|美团订单'},
                    {'金额': (0, 30)}
                ],
            BillLabels.food_big:
                [
                    {'收支': '支出'},
                    {'交易对方': food_brand,
                     '商品': '北京-环保园M地块|满座儿|美团订单'},
                    {'金额': (30, 60)}
                ],
            BillLabels.food_luxurious:
                [
                    {'收支': '支出'},
                    {'交易对方': food_brand,
                     '商品': '北京-环保园M地块|满座儿'},
                    {'金额': (60, np.inf)}
                ],
            BillLabels.food_drinks:
                [
                    {'交易对方': '奈雪的茶|luckin|喜茶|茶太良品|Blueglass|一点点',
                     '商品': 'CoCo|星巴克|奶茶|咖啡|coffee|酸奶'}
                ],
            BillLabels.food_fruit:
                [
                    {'收支': '支出'},
                    {'交易对方': '水果'},
                ],
            BillLabels.food_cook:
                [
                    {'收支': '支出'},
                    {'交易对方': '惠民果蔬|盒马'},
                ],

            BillLabels.daily_delivery:
                [
                    {'收支': '支出'},
                    {'交易对方': '顺丰速运|中国邮政|EMS|中通|圆通|申通|韵达'},
                ],
            BillLabels.daily_utilities:
                [
                    {'收支': '支出'},
                    {'交易对方': '电力公司|燃气公司',
                     '商品': '话费|电费|燃气费|网费|天然气'}
                ],

            BillLabels.house_rent:
                [
                    {'收支': '支出'},
                    {'交易对方': '自如'}
                ],

            BillLabels.transport_public:
                [
                    {'收支': '支出'},
                    {'交易对方': '公交|公共交通|地铁|轨道交通|共享单车|哈啰出行', '商品': '公交|公共交通|地铁|轨道交通|单车骑行'},
                ],
            BillLabels.transport_taxi:
                [
                    {'收支': '支出'},
                    {'交易对方': '滴滴出行|高德', '商品': '打车|快车'},
                ],
            BillLabels.transport_train_plane:
                [
                    {'收支': '支出'},
                    {'商品': '火车票|飞机票|12306'},
                ],

            BillLabels.finance_interest_on_deposit:
                [
                    {'收支': '收入'},
                    {'来源': '支付宝'},
                    {'商品': '余额宝.+收益发放'},
                ],
        }
