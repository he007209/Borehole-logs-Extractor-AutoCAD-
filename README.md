# Borehole-logs-Extractor-AutoCAD-
<pre>
这是一个CAD钻孔柱状图的信息提取工具，目的是批量从格式相同的柱状图中提取以下内容：  
  ①钻孔信息（钻孔编号、开孔日期、孔口标高等）
  ②分层信息（层底深度、剖面层号、时代成因等）
  ③土工样和标贯信息

读写cad用到的库：https://github.com/reclosedev/pyautocad

准备改下力学数据的识别方法...

20241121
  柱状图预处理建议：
  ①删除花纹，如果花纹是块对象，而你又需要执行全选-分解，则应先批量清理花纹块对象，否则花纹会被分解成很多短线，增加后续处理时间：选中其中一个花纹块，右键，选择类似对象，delete键删除。
  ②确保最多只有一个框框住主体表格，如果有多的，可通过快速选择，按面积大小来筛选出全部多余的框，删掉。
  下图中：
  红色：主体表格的框
  黄色：主体表格外第一个框（可接受，不用处理）
  蓝色：主体表格外第二个框（删掉）
  ![Image text](https://raw.githubusercontent.com/he007209/Borehole-logs-Extractor-AutoCAD-/blob/main/outline.png)


<pre>
