---
permalink: /lang/strings/
---

# Strings

Strings 使用双引号。在 strings 内可以使用单引号，若使用双引号则双倍它。

```vbs
str = "hello world"
str = "'good'"
str = """No"""
```

## 操作

[VBScript - Working With Strings](https://support.smartbear.com/testcomplete/docs/scripting/working-with/strings/vbscript.html)

获取字符串长度，`Len(str)`。

不能以索引的方式获取单个字符，可以用提取子串函数 `Left`, `Mid`, `Right`。注意字符位置从 1 开始。

使用 `&` 拼接字符串。`+` 也可以，不过也可能是数字加号。

查找子串，`InStr` 返回第一个匹配的位置，从 1 开始，没找到则返回 0。`InStrRev` 是从右向左查找。

