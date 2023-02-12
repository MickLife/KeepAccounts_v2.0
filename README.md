# KeepAccounts_v4
从现在起，养成记账的好习惯！

### Update log: v4.0 (2023 Feb.)

时隔多年，作者被迫转行，加入程序员队伍，想起了更新他初学python的第一个项目……
* 调整了账单分类，支持完全自定义的记账分类
* 实现了半自动化分类标注，支持自定义分类标注规则


# Getting Start

1. 配置环境
   * 开发版本: python 3.9
   * `pip install -r requirements.txt`
2. 配置 src/config.py
   * GlobalConfig: 全局路径、设置
   * MajorLabels: 分类标签-大类
   * BillLabels: 分类标签-小类
   * AutoLabelRules: 自动标注规则
3. 下载账单

   * 微信账单
      1. 进入手机版微信，选择 “我”
      2. 点击 “服务”；
      3. 点击 “钱包”， 
      4. 点击右上角的 “账单” 按钮；
      5. 点击右上角“常见问题”
      6. 点击“下载账单”->“用于个人对账”；
      4. 自定义账单时间，然后点击 “下一步”；
      5. 填写要导出的邮箱（微信会把账单发送到你填写的邮箱），点击 “下一步”；
      6. 输入支付密码，提示申请已提交，微信官方会给你发送一条消息，里面有账单的解压码；
      8. 前往你的邮箱下载得到压缩包，用解压码解压得到 .csv 格式微信账单，导出成功。
   
   * 支付宝账单
      1. 打开支付宝服务大厅 https://cshall.alipay.com/lab/selfHelp.htm
      2. 在“交易服务”中点击“交易记录”一项；
      4. 扫码登录；
      5. 选择交易时间，勾选 excel 格式，下载得到 .zip 压缩包，解压得到 .csv 格式的账单。
      6. 备注：商家用户请勿从商家中心导出，否则数据格式不同无法使用本程序导入账单。请按以上步骤或切换至个人版页面导出。
   
4. 按目录结构存放账单：
   ```
    │  .gitignore
    │  main.py
    │  README.md
    │  requirements.txt
    ├─AccountBooks
    │  │   KeepAccountDataBase.xlsx
    │  │   visualize.xlsx
    ├─records
    │  │  alipay_record_20230115_1641_1.csv (可存放多个csv)
    │  │  微信支付账单(20221216-20230115).csv (可存放多个csv)
    │  └─.old_records (请勿删除此目录)
    └─src
        │  AccountBook.py
        │  Annotator.py
        │  config.py
    ```

5. `python3 main.py`
6. KeepAccountDataBase.xlsx中手动订正、追加records
7. Visualize.xlsx中查看可视化图表 (此部分作者暂未提供模板，请手动链接数据、绘制数据透视图表)

Enjoy!

***

Author: MickLife

Github: https://github.com/MickLife/KeepAccounts_v2.0