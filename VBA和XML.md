### 1. 日期：2018年6月23日
### 2. 问题：用Excel生成xml文档
### 3. 原因：
xml文档的结构和架构都是相同的，其中所引用到的数据都存储在excel中。可以利用excel中的
数据批量生成文件报表。
### 4. 具体实现：
#### 1. xml DOM 基础。
xml是一种扩展标记语言，是一种纯文本文件。可以将数据以纯文本的方式进行传递。

xml文档结构：
是一种树形文档结构，其包含三个重要结构：节点，属性，文本。

基本构成：
- baseName:名称
- attribute：属性
  - length ：长度
- nodeName：名称
- text ：文本

几种常见关系：
- 父子关系
  - 父节点：parentNode
  - 子节点：childNode
- 子节点
  - firstChild
  - lastChild
- 同级关系
  - nextSibling
  - previousSibling

下面是一个xml文档案例：
```
<bookstore>
<book category="COOKING">
  <title lang="en">Everyday Italian</title>
  <author>Giada De Laurentiis</author>
  <year>2005</year>
  <price>30.00</price>
</book>
<book category="CHILDREN">
  <title lang="en">Harry Potter</title>
  <author>J K. Rowling</author>
  <year>2005</year>
  <price>29.99</price>
</book>
<book category="WEB">
  <title lang="en">Learning XML</title>
  <author>Erik T. Ray</author>
  <year>2003</year>
  <price>39.95</price>
</book>
</bookstore>

```
文档结构：

```

```

#### 2. VBA的xml DOM操作语法

##### 1. 读取xml文档
用excel加载xml文档并读取其中的特定信息，我们需要完成一下三件事情
1. 加载文档
2. 选择需要读取节点和其属性
3. 将值写入单元格中
第一步、加载文档，首先我们需要知道文档名称和文档的路径。

XMLFileName = "D:\Sample.xml"

oXMLFile.load (XMLFileName)

第二步、读取

1. 读取节点的方法有两种
- SelectNodes，读取一系列节点
  ```
  Set TitleNodes = oXMLFile.SelectNodes("/catalog/book/title/text())
  ```
- SelectSingleNode，读取特定的节点
  ```
  Set Nodes_Particular = oXMLFile.SelectSingleNode("/catalog/book[4]/title/text())
  ```
2. 读取属性
- Attribute.Length
- getAttribute("属性名")
  三步走：
    1. 选择节点
    2. 选择属性
    3. 读取属性值

  ```
  Set Nodes_Attribute = oXMLFile.SelectNodes(“/catalog/book”)
  Attributes = Nodes_Attribute(i).getAttribute(“id”)
  ```
  - 选取特定的顺序的节点
      oXMLFile.SelectSingleNode(“/catalog/book[4]/title/text()”)
  - 选取包含特定值的节点
      oXMLFile.SelectNodes(“/catalog/book/title[../genre = ‘Fantasy’]/text()”)
3. 案例文档：
  ```
  <bookstore>
  	<book category="cooking">
  		<title lang="en">甲</title>
  		<author>Giada De Laurentiis</author>
  		<year>2005</year>
  		<price>30.00</price>
  	</book>
  	<book category="children">
  		<title lang="en">乙</title>
  		<author>J K. Rowling</author>
  		<year>2005</year>
  		<price>29.99</price>
  	</book>
  	<book category="web">
  		<title lang="en">丙</title>
  		<author>James McGovern</author>
  		<author>Per Bothner</author>
  		<author>Kurt Cagle</author>
  		<author>James Linn</author>
  		<author>Vaidyanathan Nagarajan</author>
  		<year>2003</year>
  		<price>49.99</price>
  	</book>
  	<book category="web" cover="paperback">
  		<title lang="en">丁</title>
  		<author>Erik T. Ray</author>
  		<year>2003</year>
  		<price>39.95</price>
  	</book>
  </bookstore>
```
4. 案例代码：

  ```
  Sub xmlRead()
  Dim a As DOMDocument
  a1 = "H:\Git\VBA\new.xml"
  Set a = New DOMDocument
  a.Load a1
  Set aNodes = a.SelectNodes("/bookstore/book")
  For i = 0 To aNodes.Length - 1
      ' read index
      aCategory = aNodes(i).getAttribute("category")
      Set aTitle = a.SelectSingleNode("/bookstore/book[" & i & "]/title/text()")
      Set aAuthor = a.SelectNodes("/bookstore/book[" & i & "]/author/text()")
      Set aYear = a.SelectNodes("/bookstore/book[" & i & "]/year/text()")
      Set aPrice = a.SelectNodes("/bookstore/book[" & i & "]/price/text()")
      ' write value
      Cells(i + 2, 1) = aCategory
      Cells(i + 2, 2) = aTitle.NodeValue
      Cells(i + 2, 3) = aAuthor(0).Data
      Cells(i + 2, 4) = aYear(0).Data
      Cells(i + 2, 5) = aPrice(0).Data
  Next
  End Sub
```
结果：

5. 注意事项：
- selectNodes 和 SelectSingleNode 对比
6. 资料参考
<a href="https://excel-macro.tutorialhorizon.com/vba-excel-read-data-from-xml-file/">VBA-Excel: Read Data from XML File</a>

##### 2. 将数据输出到xml文档

将数据输出为xml文档有两种操作，第一种生成一本新的xml文档。第二种更新现有的xml文档并将其另存。
前一种主要是涉及VBA对文本文档的写操作，只需将表格中的值放到字符串中的特定位置就行，然后将其保存为xml文档。
需要注意的是当要进行属性值的书写时，VBA使用string输出英文的双引号符号时需要将其换为两个引号。

即：变量a= 78,要输出“78”
要输出字符串c ：“78”，应该写成：
```
c =""" & a & """
```

第二种涉及VBA对XMLDOM的操作,主要包含更改和新建两个操作，具体方法如下：

- 更改现有节点的名称，属性
- 新建节点，属性

1. 更改
  1. 节点
  第一步选择要更改的节点，和前面讲的读取xml文档中的语法一样：
    Set TitleNode = oXMLFile.SelectSingleNode(“/catalog/book[0]/title”)
  第二步，对其内容进行更改：
    TitleNode.Text = “I am the new Title Here”
  2. 属性
  选择特定节点的属性：
    Set oAttribute = oXMLFile.SelectSingleNode(“/catalog/book[0]/@id”)
  更改：
    oAttribute.Text = “1111111111111111111111111111”
2. 新建
  1. 节点
    通过设定父节点确定新节点的位置，然后设置新节点的名称和文本，并新建节点。
```    
    ‘ select a parent node
    Set ParentNode = oXMLFile.SelectSingleNode(“/catalog/book[1]”)
    ‘ add a new childNode
    Set childNode = oXMLFile.CreateElement(“NewNode”)
    childNode.Text = “I am The New NOde Here”
    ParentNode.AppendChild (childNode)
```

  2. 属性
    和新建节点一样，通过设置其父节点的位置确定属性的位置，然后设置其名称和值，并新建属性。
```
    ‘Add new Attribute to the Node
    Set ParentNode = oXMLFile.SelectSingleNode(“/catalog /book[2]/publish_date”)
    ‘ add its attribute
    Set newChildAttribute = oXMLFile.CreateAttribute(“Status”)
    newChildAttribute.Text = “Very Old Book”
    ParentNode.Attributes.SetNamedItem (newChildAttribute)
```

参考资料：<a href="https://excel-macro.tutorialhorizon.com/vba-excel-update-xml-file/">VBA-Excel: Update XML File</a>
