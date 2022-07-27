使用方式:
1. 双击导表.bat会解析./Excel目录下的配置表,解析结果放到assets/Script/config/下
2. <python convert.py 配置表目录>: 解析指定目录的配置表,解析结果也放到assets/Script/config/下

配置表格式要求
1. 命名: 以 XXConfig_任意中文描述.xlsx  命名, XXConfig就是解析后对应typescript文件名
2. 脚本只会解析配置表的第一个sheet,所以的后面的sheet格式不限制可作为备份或临时计算
3. sheet前3行是配置描述,第1行是中文描述方便阅读; 第2行是程序用的字段名,应该为英文,不用的字段留空; 
   第3行是该列数据的类型,暂支持number数字,boolean布尔,string字符串,number[]数字数组,string[]字符串数组
4. sheet第一列会作为程序取值的key,建议设置为数字且唯一,提前规划好各配置表的ID范围避免冲突
5. 每个单元格都不能为空(字符串除外),没有该属性的也需要填一个默认值,number类型为-1,数组类型为[]
