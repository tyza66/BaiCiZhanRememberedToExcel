import sqlite3
import xlwt
import sys

#音标里有 IPA 字符（ˈ ɑ 等），中文 Windows 的控制台默认是 GBK，直接 print 会
#抛 UnicodeEncodeError。让输出遇到编不出的字符就替换成 ?，别因为打印把导出打断。
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")

#初始化
print("百词斩已背单词导出工具v1.0.1 By:tyza66")
print("主页地址https://github.com/tyza66")
print("1./data/media/0/Android/data/com.jiongji.andriod.card/files/baicizhan")
print("2.拷出手机百词斩文件夹中的两个数据库：baicizhantopicproblem.db（包含已背的单词）和lookup.db（包含所有单词）")
baicizhantopicproblem = input("请输入本地的baicizhantopicproblem.db路径：")
lookup = input("请输入本地的lookup.db路径：")
a = input("是否随机顺序（0不随机，1随机）：")
name = input("请输入输出文件名（不带后缀）：")
print("开始处理。。。")
#从baicizhantopicproblem表读出百词斩已背单词
conn1 = sqlite3.connect(baicizhantopicproblem)
#产生游标
cur = conn1.cursor()
#获得已背单词id
#百词斩换过几次版本，这张表的后缀也跟着变（563 -> 11 -> ...），按名字前缀找，
#换版本时不再需要改代码。
table = None
for row in cur.execute(
    "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'ts_learn_offline_dotopic_sync_ids_%'"
):
    table = row[0]
if table is None:
    sys.exit("在 baicizhantopicproblem.db 里找不到 ts_learn_offline_dotopic_sync_ids_ 开头的表，请确认拷对了数据库。")
if a == 1:
    cur.execute("SELECT topic_id FROM " + table + " ORDER BY random()")
else:
    cur.execute("SELECT topic_id FROM " + table)
wordId = cur.fetchall()
#关闭链接
conn1.close()
#从lookup表中寻找已经背的单词表
conn2 = sqlite3.connect(lookup)
cur = conn2.cursor()
count = 0
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('百词斩已背单词')
for word in wordId:
    cur.execute("SELECT word,accent,mean_cn FROM dict_bcz WHERE topic_id = " + str(word[0]))
    word = cur.fetchall()
    if not word:
        #topic_id 在 lookup.db 里查不到释义（词库没同步全），跳过而不是崩在这里
        continue
    print(word[0])
    worksheet.write(count, 0, label=word[0][0])
    worksheet.write(count, 1, label=word[0][1])
    worksheet.write(count, 2, label=word[0][2])
    count = count + 1
workbook.save(name + '.xls')
conn2.close()
print("共" + str(count) + "个单词，Excel已导出到：" + name + '.xls')
