---
permalink: /lang/procedures/
---

# VBScript 过程

VBScript procedure：Sub, Function, Property

## Sub

```vbs
Sub Sum(a, b)
    WScript.Echo a + b
End Sub

Sum 1, 1
Call Sum(2, 2)
```

procedure 在直接调用时不能将参数放在括号内。用 Call 调用时则相反，必须将参数放在括号内。

procedure 可以在定义之前调用。

Sub 没有返回值。

## Function

函数有返回值，所以可以用在表达式中。

```vbs
Function Sum(a, b)
    Sum = a + b ' 返回值
End Function

WScript.Echo Sum(1, 2)
```


