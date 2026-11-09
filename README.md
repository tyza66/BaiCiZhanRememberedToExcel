# BaiCiZhanRememberedToExcel
### 导出百词斩已背单词表到Excel
#### 【2026.11.9】v1.0.1
- 依赖库：sqlite3，xlwt
- 需要手动从手机导出/data/media/0/Android/data/com.jiongji.andriod.card/files/baicizhan"内的baicizhantopicproblem.db（包含已背的单词）和lookup.db（包含所有单词）到电脑
- 按照提示进行输入baicizhantopicproblem路径，lookup路径，是否随机顺序，输出文件名即可
- 测试使用的百词斩版本：7.2.0
- 本版起不再需要关心已背单词表的表名后缀：脚本自己按 `ts_learn_offline_dotopic_sync_ids_` 前缀找表，百词斩换版本改了后缀也能直接跑

By:tyza66(洮羱芝闇)
