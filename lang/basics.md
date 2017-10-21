---
permalink: /lang/basics/
---

# VBScript

VBScript 语法为 VB 风格，跟 C 风格区别较大。了解 VB 语法

- [Visual Basic Guide](https://docs.microsoft.com/en-us/dotnet/visual-basic/)

VBScript 标识符不区分大小写。

## 注释

第一种方式，使用单引号

```vbs
' comments
message = "hello world" ' comments
```

第二种方法，使用 `REM` 语句。

```vbs
REM remarks
message = "hello world" : REM remarks
```

既然是语句，`REM` 与其它语句同行时要用 `:` 隔开。

## 语句

语句不需要结束符。

多个语句写在一行时，用 `:` 隔开。

过长的语句可以跨行，行末用 `_`。字符串不可以跨行。

## 代码规范

可以借鉴 [Program Structure and Code Conventions (Visual Basic)](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/program-structure-and-code-conventions)

关键字，过程名字使用 PascalCase 风格。

变量名字使用 camelCase 风格。
