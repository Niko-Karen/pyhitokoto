# Pyhitokoto
---
## 简介

#### 一句话概括: **一个爬取一言数据并写入Excel的迷你爬虫**

这是我第一次正式发布我的Python作品，虽有些简陋，但会间断维护加功能。

## 正文
---
### 命令行使用
使用方法: ```python hitokoto.py -c <爬取句子数量> -e <保存的 Excel 文档名>```

或者是: ```python hitokoto.py --count <爬取句子数量> --excel <保存的 Excel 文档名>```

本“模块”已经加装[Alive-Progress](https://pypi.org/project/alive-progress/)进度条，请根据`requirement.txt`自行安装。

---
### 代码内使用。

很简单。
将仓库Clone下来，在仓库目录运行: `pip3 install -r requirement.txt`，然后再你的`.py`文件中导入使用。

1. 直接爬取数据写入Excel
```python
from pyhitokoto.hitokoto import Hito

hito = Hito()

hito.count = 20 #爬取的句子数量
hito.excel_name = 'example.xlsx' #输出的文件名(须为Excel后缀名)

# hito.run() 现更新CSV格式后，不建议直接使用run函数，因为run函数会分别写入Excel和CSV，需分别输入Excel和CSV文件名
hito.write_to_excel() # 可直接调用write_to_excel函数写入

```

2. 只获取数据，返回三个列表
```python
from pyhitokoto.hitokoto import Hito

hito = Hito()

hito.count = 20 #爬取的句子数量 (只获取数据可以不需要Excel名)

hito_list, from_list, who_list = hito.get_list() 
#上行变量含义分别是: 一言列表、来源列表、作者列表
```
3. 写入进`CSV`文件中
```python
from pyhitokoto.hitokoto import Hito

hito = Hito()

hito.count = 20 
hito.csv_name = 'example.csv' #CSV 输出文件名（需为.csv文件后缀）
hito.write_to_csv() # 写入CSV

```

### P.S: 导出的作者列表如果为 None ，我都直接替换为 未知 了。

## 最后

**感谢大家对我的支持😙！如有不足之处，大佬们请指点出来，小弟我会尽快更改😇。（我准备小升初了，可能改Bug、更新不及时）**

P.S: 可能会发布到Pypi上，待我择日发布 (doge
