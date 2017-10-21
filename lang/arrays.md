---
permalink: /lang/arrays/
---

# VBScript 数组

## 声明数组

```vbs
Dim a(9)  ' Declare an array with 10 elements.
Dim a()   ' Declare a dynamic array.
Dim a     ' Declare a variable.
```

第一种方式，9 是数组上界，索引从 0 开始，因此数组长度是 10。这种方式不能改变数组长度。

声明多维数组

```vbs
Dim Table(4, 4)
```

第二种方式，声明一个动态数组，数组长度和维数不定。在使用动态数组之前先使用 ReDim 声明数组长度和维数。

```vbs
ReDim a(4)
ReDim Preserve a(2)
```

Preserve 只能改变最后一维的长度，不能改变维数。

## 操作

```vbs
Dim colors(2)
colors(0) = "red"
colors(1) = "green"
colors(2) = "blue"
```

或者

```vbs
Dim colors
colors = Array("red", "green", "blue")
```
