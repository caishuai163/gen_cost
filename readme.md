# 使用说明
## 目录结构
1. source 存储源文件
   1. user_tb.xlsx 存储用户基础信息
   2. cost 文件夹下存储逾期丢失文件 
   3. ab 存储A+B文件
   4. p2p 存储点对点文件
   5. c 存储C类7%文件
   6. bound 存储奖金文件
2. local.db 数据临时存储位置 sqlite数据库
3. config.json 配置文件
4. gen_cost.py 程序入口
5. output 输出目录
6. config dao service util 程序文件

## 用户匹配规则
规则一：  
1. 匹配用户时，会按姓名优先匹配。
2. 匹配到多个时，会按职位匹配。
3. 再次匹配到多个时，会按平台经销商名称匹配。大区经理的平台经销商名称匹配按单个单元格中多行数据的第一行数据进行匹配。

规则二：  
1. 按姓名优先匹配。   
2. 匹配到多个时，会按职位匹配。  
3. 再次匹配到多个时，会按大区+大区经理名进行匹配。      
4. 再次匹配到多个时，会按分区进行匹配。  

## config.json 配置
#### title_template
表头的模版文件，带大括号中间带数字的尽量不要动，是占位符，其他文字按需可更改。  
#### title_bound_from_day
表头中占位符`{4}`的变量  
#### table_row_find_day 
扣款数据部分 找回月份的表头内容  例如 2023年3月9日找回  
#### title_other_cost_from_day
其他扣款部分的表头月份，目前只支持一个，且此部分数据皆为空，暂未实现  
#### cost
扣款文件配置。值为数组对象，每个对象中包含需要读取的文件配置。程序会按此处数组对象配置的先后顺序进行排序。  
读取单个文件时，会按固定列读取数据，然后读取月份列。  
固定列读取规则：  
1. B D E 姓名 经销商 职位  
2. "盘点后丢失电池数量       "找回数量	"确定丢失电池数量 所在列 G H I  
3. 扣款金额	"扣款月数"	"月扣款金额" 所在列 R S T  
4. 从 U列开始读取月份数据 （格式为时间或自定义） 当格式不再是月份的格式时，停止读取  

匹配用户时，按规则一匹配。  
```json lines
{
    "file_name": "要读取的文件名，带扩展名",
    "read_sheet_name": "要读取的sheet名",
    "row_title": "输出文件时，表格最上方单行文字配置，无需序号会自动生成",
    "row_table_title": "输出文件时，表格的第一行表头文字配置",
    "sum_row_text": "汇总扣款表上的列名"
}
```

#### p2p
点对点文件配置。值为数组对象，每个对象中包含需要读取的文件配置。程序会按此处数组对象配置的先后顺序进行排序。  
匹配用户时，按固定列读取值。  
匹配用户时，按规则二匹配。

```json lines
{
   "file_name": "要读取的文件名，带扩展名",
   "read_sheet_name": "要读取的sheet名", 
   "file_month": "文件所对应的数据月份，需按格式书写. e.g. 2023年12月份",
   "sum_row_text": "汇总扣款表上的列名",
   "read_give_row":"实际奖励金额 列的标题，会检索表头读取这一列",
   "read_cost_row":"实际罚款金额 列的标题，会检索表头读取这一列",
   "read_cost_month_row":"扣款月份 列的标题，会检索表头读取这一列"
}
```

#### ab
A+B文件配置。值为数组对象，每个对象中包含需要读取的文件配置。程序会按此处数组对象配置的先后顺序进行排序。  
匹配用户时，按固定列读取值。  
匹配用户时，按规则二匹配。
```json lines
{
   "file_name": "要读取的文件名，带扩展名",
   "file_month": "文件所对应的数据月份，需按格式书写. e.g. 2023年12月份",
   "read_pre_sheet_name": "2023年8月及以前注册AB类推广员 的sheet名",
   "read_after_sheet_name": "2023年9月及以后注册AB类推广员  的sheet名",
   "read_bound_row": "业务员应得合计  列的标题，会检索表头读取这一列",
   "read_region_row": "大区 大区字段所在的中文名 ，目前检索大区所在列",
   "read_part_row": "分区  字段所在的中文名 ，目前检索所在列",
   "read_name_row": "业务员 字段所在的中文名 ，目前检索所在列",
   "read_work_row": "岗位名称 字段所在的中文名 ，目前检索所在列",
   "read_start_row": 2 // 这个是数字 读取列的所在行
}
```
注意数据是从第5行开始读取真实的数据，
read_pre_sheet_name 是一个历史的字段 ，后续可能没有这个，需要创建一个带正常表头的空sheet，让程序去读取


#### c
C类7%文件配置。值为数组对象，每个对象中包含需要读取的文件配置。程序会按此处数组对象配置的先后顺序进行排序。  
匹配用户时，按固定列读取值。  
匹配用户时，按规则二匹配。
```json lines
{
   
   "file_name": "要读取的文件名，带扩展名",
   "file_month": "文件所对应的数据月份，需按格式书写. e.g. 2023年12月份",
   "read_sheet_name": "要读取的sheet名",
   "read_bound_row": "业务员应得合计  列的标题，会检索表头读取这一列",
   "read_region_row": "大区 大区字段所在的中文名 ，目前检索大区所在列",
   "read_part_row": "分区  字段所在的中文名 ，目前检索所在列",
   "read_name_row": "业务员 字段所在的中文名 ，目前检索所在列",
   "read_work_row": "岗位名称 字段所在的中文名 ，目前检索所在列",
   "read_start_row": 2 // 这个是数字 读取列的所在行

}
```
注意数据是从第3行开始读取真实的数据  

#### spec
业务员名下电池罚款特殊扣款文件配置。值为数组对象，每个对象中包含需要读取的文件配置。程序会按此处数组对象配置的先后顺序进行排序。  
匹配用户时，按固定列读取值。  
匹配用户时，按规则二匹配。
```json lines
{
   "file_month": "2024年4月份",
   "file_name": "2023年11月业务员名下电池罚款特殊扣款.xlsx",
   "read_sheet_name": "4月业务员名下电池扣款",
   "sum_row_text": "4月业务员名下电池扣款 输出到Excel的列名字",
   "region_row":"城市 大区字段所在的中文名 ，目前检索大区所在列",
   "name_row":"姓名 字段所在的中文名 ，目前检索所在列",
   "read_cost_row":"罚款金额 字段所在的中文名 ，目前检索所在列",
   "read_cost_month_row":"扣除月份 字段所在的中文名 ，目前检索所在列"
}
```

#### bound 
奖金配置文件。值为单个对象。
匹配用户时，按规则一匹配。
```json lines
{
   "file_name": "要读取的文件名，带扩展名",
   "read_sheet_name": "要读取的sheet名", 
   "start_column":"读取表头的位置列开始，会读取范围内的第二行和第三行表头作为生成文件的表头  e.g. Q",
   "end_column":"读取表头的位置列开始，会读取范围内的第二行和第三行表头作为生成文件的表头  e.g. AH",
   "bound_title": "输出文件奖金部分单行文字，无需序号会自动生成 e.g. 3月份奖金："
}
``` 
注意   
1. start_column 和  end_column 之间的数据是按文本进行copy到生成表的 ，一般这里会把原始文件的涉及到时间的，设置字段为文本类型，再重新输入  
2. 由于无法copy 合并单元格的原样式， 这里会对同一行，后面值为空的 进行合并 ，例如 第二行的 J列有值， K，L 没值，M列有值， 此时会将第二列的J，K，L合并 保留J值
3. 行与行之间不会自动合并  


## exe使用需要关注
1. source 文件夹内的文件
2. config.json 

## 常见问题
文件报错无法读取，一般是因为源文件在不同系统之间兼容不一致，此时一般的处理方式是本地新建一个文件，将原内容copy到新文件中保存，注意移动sheet或创建副本没有用，需要复制粘贴，过程中不要因为筛选或隐藏列这种，少复制单元格