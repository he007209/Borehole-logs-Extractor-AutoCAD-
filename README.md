# Borehole-logs-Extractor-AutoCAD-
<pre>
这是一个CAD钻孔柱状图的信息提取工具，目的是批量从格式相同的柱状图中提取以下内容：  
  ①钻孔信息（钻孔编号、开孔日期、孔口标高等）
  ②分层信息（层底深度、剖面层号、时代成因等）
  ③土工样和标贯信息

读写cad用到的库：https://github.com/reclosedev/pyautocad

准备改下力学数据的识别方法...
  
20250324
  1、终于完成三年前就想做的钻孔信息字段的局部匹配，现在填写钻孔信息目标字段时也可以填写部分字符串了：
      举个例子：假设在cad中的一个钻孔信息字段名叫【钻孔深度(m)】，在之前的版本，在软件里填写字段名时要填写完整的
               【钻孔深度(m)】这7个字符才能匹配上，现在可以像土工标贯和分层信息目标字段那样，填写连续、局部且唯
                一（不能出现在其他格内）的字符串就行了，比如可以填【钻孔深度】、【孔深度(m)】、【孔深度(】、【钻孔深】等。
  
  2、最近新发现的某些柱状图存在一些隐藏对象，会导致报错"list index out of range"。有时全选复制到新建文件可以把它们排除掉，
    现在在代码中会将这些对象跳过，就不用再新建文件复制了。
  
20250323
  今天在图书馆呆了半天，终于有点看懂这个三年前写的代码是什么意思了，预计很快可以把钻孔信息的字符串部分匹配修好。
  
20250321
  处理一个cad文件中包含数十数百个柱状图的情况时，建议先随便复制一个或几个图到新建的文件中，测试下识别的效果。
  有时源文件有些特殊限制或者存在隐藏的对象可能会导致报错，通过复制到新建的文件可能有用。
  
20250320
  出现index out of range错误，可能是源文件中有特殊未知对象，可以试试选中目标柱状图图形，复制到新建的空白dwg再执行识别
  
20241121
  柱状图预处理建议：
  1. 删除花纹，如果花纹是块对象，而你又需要执行全选-分解，则应先批量清理花纹块对象，否则花纹会被分解成很多短线，
    增加后续处理时间：选中其中一个花纹块，右键，选择类似对象，delete键删除。
  2. 确保【最多】只有一个框框住主体表格，如果有多的，可通过快速选择，按面积大小来筛选出全部多余的框，删掉。

  主体表格是什么？当然是我自定义的概念。
  下图中：
  红色：主体表格
  黄色：主体表格外第一个框（不用处理）
  蓝色：主体表格外第二个框（假如存在，要删掉它）
  【怎么快速删除？】
  ①选中这个多段线，右键属性，查看它的面积，假设查到的面积为65432，
  ②esc键取消选中，空白处右键，快速选择，多段线，面积，条件写＞60000，添加到选择集，就可以选中所有这个类型的框，
  ③然后delete键删掉它们
  <a href="https://sm.ms/image/nh6j3UDOapZ2rAS" target="_blank"><img src="https://s2.loli.net/2024/11/21/nh6j3UDOapZ2rAS.png" ></a>


<pre>
