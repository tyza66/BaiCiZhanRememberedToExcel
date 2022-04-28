import sqlite3;
import xlwt;

#初始化
print("百词斩已背单词导出工具v1.0.0 By:tyza66")
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
if a == 1:
    cur.execute("SELECT topic_id FROM ts_learn_offline_dotopic_sync_ids_563" + " ORDER BY random()")
else:
    cur.execute("SELECT topic_id FROM ts_learn_offline_dotopic_sync_ids_563")
wordId = cur.fetchall()
#关闭链接
conn1.close()
#从lookup表中寻找已经背的单词表
conn2 = sqlite3.connect(lookup)
cur = conn2.cursor()
sum = 0
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('百词斩已背单词')
for word in wordId:
    cur.execute("SELECT word,accent,mean_cn FROM dict_bcz WHERE topic_id = " + str(word[0]))
    word = cur.fetchall()
    print(word[0])
    worksheet.write(sum, 0, label=word[0][0])
    worksheet.write(sum, 1, label=word[0][1])
    worksheet.write(sum, 2, label=word[0][2])
    sum = sum + 1
workbook.save(name + '.xls')
conn2.close()
print("共" + sum + "个单词，Excel已导出到：" + name + '.xls')



