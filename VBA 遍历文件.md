## 一、主要功能

1. 读取文件夹下的所有文件。
2. 将文件的文件名存储至工作表中。

## 二、功能实现

1. 主要代码Syntax说明：

   <code>
    a = Dir("FilePath & FileName")
   </code>
 
   Dir () 可以用于遍历该路径下的特定文件，并返回文件名称。
 
   FilePath用于填写文件路径，FileName可以用于指定特定的文件，并且可以使用通配符。
 
   如：Dir（"H:\*.txt"）
  
   读取H盘下所有的txt文档。
 
<pre><code>
Sub text()
  Dim a As String
     i = 1
     a = Dir("H:\Git\html\*.html")
   ' 设置循环体
  Do While a <> ""
  ' 开始读取下一个文件
     a = Dir
     Cells(i, 1) = a
     i = i + 1
  Loop
End Sub
</code>
</pre>
## 三、参考资料
1. <a href="https://www.cnblogs.com/gilgamesh-hjb/p/7291821.html">Dir()主要用于获取（遍历）目录下的文件名</a>
2. <a href="http://club.excelhome.net/thread-371188-1-1.html">VBA函数精选之十二（Dir函数）<a>
