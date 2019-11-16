# Visual Basic

## vb集成开发环境

### 工具箱窗口

- Label
- CommandButtom
- TextBox
- ...

### "属性"窗口

- caption : 显示的名称
- name : 对象的名字
- text : TextBox的显示内容
- left,top : 在显示屏中的位置
- width,height : 窗体的大小

### 对象属性赋值

    对象名.属性名 = 新的属性值

    Form1.caption = "Hwllo Visual Basic"

### 对象方法的调用

    对象名.方法名 [参数]

    object.Move left,[top,[width,[height]]]

    Form1.Move 1000,1000,2000,2000

    Form1.Move 1000     'error(缺少参数)

    object.Print "Hello"

### 窗体对象的常用事件

    Private Sub Form_事件名()      '过程首部
    ...
    End Sub

### a simple program

    Private Sub Form_Click()
        Form.Print("鼠标单击")
    End Sub

### 字符串连接

    用&     '&前后都有空格，

### 定义快捷键

>当Caption后有&字符，&不会显示在按钮表面，而会把紧接在&后面的字符定义为
快捷键；
>
>访问快捷键:ALT+&后的字符

### 窗体移动

Form：
    name：formMove

CommandButton：
    name：cmdMoveLeft

    Private Sub cmdMoveLeft_Click()
        formMove.left = formMove.left - 100
    End Sub

### TextBox 的常用属性

- Text
- MaxLength
- PasswordChar
- Alignment

### TextBox 的change事件过程

    Private Sub textInput_change()
        textOutput.Text = textInput.Text
    End Sub

>textInput中的内容一改变textOutput中的内容随之改变

### label

    Private Sub cmdDisplay_Click()
        label.Caption = "name:" & textName.Text & "age:" textAge.Text 
    End Sub

### Visual Basic 语法规则

1. 源程序中不区分大小写，但对象名，变量名，常量名和过程名只有一种大小写方式。

2. 续行：backspace+_

    lbDisplay.Caption = "name:" & textName.Text &  _
    "age:" textAge.Text

3. 一行可以写多句代码(不推荐)，用:分隔

    a = 4 : b = 5 :c = 6

4. 注释: 使用'

### Project

    工程文件.vbp + 窗体文件.frm

---

## 数据类型

### 数值型

类型名称|字节数|范围
:-:|:-:|:-:
Byte|1|0~255
Integer|2|-32768~32767
Long|4|
Single|4|
Double|8|

### String 型

### Boolean 型

    true/false
    2个字节

### Date 型

### 整形常量

- 十进制

>常量后加&会变为长整型，数值不变但内存不同

    10000， 10&， -200&

- 八进制

>以&O开头，如果表示长整型数，末尾加&

    &O11: 9
    &O176340: -800  Integer型，16位二进制数，最高位1，表示负值
    &O176340&: 64736 Long型，32位二进制

- 十六进制

> 以&H开头，如果表示长整型数，末尾加&
>
> 把一个表示负数的十六进制数转换为十进制数，需要用到补码

    &HFF: 255
    &HFFFF: -1
        原码：1111111111111111
        反码：1000000000000000(符号位不变，其余取反)->0
        补码：1000000000000001->-1
    &HFFFF&: 65535

### 浮点型常量

    3.14    24.    -.45    -0.05

    mEn = m*10^n
        -.5E-2 = -0.5*10^-2
> m不能省略，n必须为整数常量

### 字符串常量

> 用" "括起来

### 变量

> 字母数字下划线，字母打头

### 定义变量

- 过程级变量(局部变量)

    Dim|Static  变量名  [As 数据类型]

> Dim定义的变量，执行完毕就会消失，释放内存
>
> Static定义的变量，每次执行完毕后，变量的值被保存

- 模块级变量

> 定义语句必须放在模块开始的通用声明段中(所有过程前面)

    Private|Dim 变量名  [As 数据类型]

> 该模块所有过程都可以访问

- 程序级变量(全局变量)

### 变量默认值

    数值型： 0
    逻辑型： False
    日期： #0:00:00#
    字符串： " "

### 强制变量定义

> 加上 Option Explicit

### 变量的赋值

    Dim i As Integer，j As Integer
    i = 10 ：j = 10

### 使用过程级变量

    Private Sub command1_Click()
        Dim i As Integer        'i的默认值为0
        i = i + 1
        command1.Caption = i
    End Sub
> 结果一直为1

    Private Sub command2_Click()
        Static i As Integer     'i的默认值为0
        i = i + 1
        command.Caption = i
    End Sub
> 每点击一次，i的值加一

### 定长/变长字符串

变长：Public| Private| Dim| Static  变量名  As String

    Dim Str As String
    Str = "hello";

定长：Public| Private| Dim| Static  变量名  As String  *  字符串长度

    Dim Str2 As Integer * 4
    Str2 = "how are you";

> 超过长度的会自动截掉

### 对象型变量

Public| Private| Dim| Static  变量名  As Object

> 假设有一个cmdOK的按钮，设置其Caption属性

    Dim objButton As Object     '定义对象型变量
    *Set objButton = cmdOk       '把按钮对象赋给对象型变量
    objButton.Capiton = "Ok"

### 变体型数据

Public| Private| Dim| Static  变量名

    Dim var As Variant
    var = "haha"
    var = 15
    Set var = cmdOk
> 变体赋值为对象型时，必须用set

### 默认值

    Date：#0:00:00#
    逻辑：False
    对象：Nothing
    变体：Empty

### 类型转换

> 隐式转换

    浮点型->整型(>0.5,加一)(=0.5,向偶数靠拢)
    i = 4.56    '整数i为5
    i = 4.5     '整数i为4

    逻辑->数值
    False->0    *True->-1

> 显式转换

函数 | 转换为 | 函数| 转换为
:-:|:-:|:-:|:-:|
CBool()|Boolean |CDate()|Date|
CInt()|int|CStr()|String
...

> 不能进行转换

    1. 包含非数值字符的字符串向数值类型转换
    2. 非"true" or "False"的字符串向逻辑类型转换
    3. 超出表示范围

### 符号常量

    Private Const PI As Single = 3.14
> 赋值时右边不能出现变量

### tips

> int1为整型

    int1 = "2" + 3      '值为5
    int1 = "2" + "3"    '值为23

> True 转换为整形：-1

---

## 运算符&表达式

### 算数运算符

> 两种不同类型数值运算，结果的类型与精度高的保持一致
>
> \ 和 Mod 要求两个运算量是整数
>
> 求余结果的正负号与第一个运算符相同
>> a = -15 Mod 30
>>
>> a的结果为-15

### 比较运算符

- 不等于：<>
- =:赋值&判断是否相等

### 字符串运算

字符串连接

    "30" + "15"         '结果3015
    "30" & 15           '结果3015
    30 & 15             '结果3015

    "30" + 15           '结果45
    30 + "15"           '结果45

> &：将两边的数据都先转换为字符型，然后连接
>
> +：如果两边都为字符型，则连接为字符串。只要有一边不为字符，就先将字符转换为整数

字符串比较

> 一个字符一个字符接着比

字符串匹配运算符：like

### 逻辑运算

a|b|a And b|a Or b|Not a|a Eqv b|a Imp b|a Xor b|
:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:
True|True|True|True|False|True|True|False
True|False|False|True|False|False|False|True
False|True|False|True|True|False|True|True
False|False|False|False|True|True|True|False

> Eqv(等价):相同为True
>
>Imp(蕴含)：True&False->False,others->True
>
> Xor(异或)：不相等为True

    Private Sub command_Click()
        Dim a As Integer, b As Integer, c As Integer
        a = -3: b =-2: c = -1
        Print a<b And b<c           'True
        Print a <b <c               'False
    End Sub

### 按位逻辑运算

> 将十进制数按补码的形式，转换为二进制数

    正数：
    原码 = 补码

    Not 10
    原码：11110101(00001010->取反得到)(符号位为1，负数，求补码)
    反码：10001010
    补码：10001011(反码+1)->-11

### 表达式

> 优先级相同，从左往右

#### 优先级

算数运算符|比较运算符|逻辑运算符
:-:|:-:|:-:
^|=|Not
-|<>|And
*、/|<|Or
\\(整除)|>|Xor
Mod|<=
+-|Like
&|Is
>左上角优先级最高，右下角优先级最低

---

## 控制结构

### If语句

#### 单行形式if...then

    If i Mod 2 = 0 Then Print "偶数"

#### 块形式 if...then...endif

    If i Mod 2 = 0 Then
        Print "偶数"
    End If

#### 块形式 if...then...else...endif

    If i Mod 2 = 0 Then
        Print "偶数"
    Else
        Print "奇数"
    End If

#### 三个数求max

>在文本框text1，text2,text3中输入三个整数，单击按钮cmdMax，将最大值显示在textMax中

    Private Sub cmdMax_Click()
        Dim x As Integer, y As Integer, z As Integer
        x = text1.Text
        y = text2.Text
        z = text3.Text
        if x>y Then
            if x>z Then
                textMax.Text = x
            esle
                textMax.Text = z
            End if
        else
            if y>z Then
                textMax.Text = y
            else
                textMax.Text = z
            End if
        End if
    End Sub

#### if...then..elseif...endif

    Private Sub cmdRank_Click()
        Dim intMark As Integer
        intMark = CInt(textInput.Text)
        if intMark>90 Then
            textOutput.Text = "优秀"
        ElseIf intMark>80 Then
            textOutput.Text = "良好"
        ElseIf intMark >60 Then
            textOutput.Text = "及格"
        Else
            textOutput.Text = "不及格"
        End If
    End Sub

#### Select Case语句

    Private Sub cmd_Click()
        Dim int1 As Integer, int2 As integer
        int1 = text1.Text : int2 = text2.Text
        Select Case int1+int2
        Case 0
            text3.Text = "两数之和为0"
        Case 1 to 10
            text3.Text = "两数之和在1-10之间"
        Case Else
            text3.Text = "两数之和在>10"
        End Select
    End Sub

### Do...Loop语句

#### DoWhile...Loop

> 求1+2+3+...+100

    Private Sub Form_Click()
        Dim i As Integer, s As Integer
        i=0:s=0
        Do While i<100
            i = i+1
            s = s+i
        Loop
        Me.Print s
    End Sub

#### Do...Loop While

    Private Sub command1_Click()
        Dim i As Integer, s As Integer, n As Integer
        n  = text1.Text
        i=0:s=0
        Do 
            i = i+1
            s = s+i
        Loop While i<n
        text2.Text = s
    End Sub

#### Do...Loop

    Private Sub command1_Click()
        Dim i As Integer, s As Long, n As Integer
        n  = text1.Text
        s = 1
        Do 
            i = i+1
            s = s*i
            if i=n Then Exit Do
        Loop 
        text2.Text = s
    End Sub

### For...Next

    Private Sub command1_Click()
        Dim i As Integer, s As Long, n As Integer
        n  = text1.Text
        s = 1
        For i=1 To n
            s = s * i
        Next
        text2.Text = s
    End Sub

### With语句

> 使用with前

    cmdFirst.Height = cmdFirst.Height + 1000
    cmdFirst.Caption = "Hello"
    cmdFirst.Move 0,0

> 使用with后

    With cmdFirst
        .Height = .Height + 1000
        .Caption = "Hello"
        .Move 0,0
    End With

### 控制结构的应用

> 验证质数

    Private Sub cmdPrime_Click()
        Dim i As Integer, n As Integer
        n = textInput.Text
        if n<1
            textOuput.Text = "请输入自然数"
        Elseif n=1
            textOuput.Text = "1不是质数"
        Elseif n=2
            textOuput.Text = "2是质数"
        Else
            For i = 2 To sqr(n)
                if n Mod i = 0 Then
                    Exit For
                End if
            Next
            if i<n Then
                textOutput.Text = n & "不是质数"
            Else
                textOutput.Text = n & "是质数"
            End If
        End If
    End Sub

> 斐波那契数列(递推法)

    Private Sub cmdFib_Click()
        Dim n As Integer
        n = textInput.Text
        if n=1 or n=2 Then
            textOutput.Text = 1
        Else 
            Dim f1 As Integer, f2 As Ingteger,f3 As Integer
            Dim i As Integer
            f1 = 1: f2 = 1
            for i = 3 To n
                f3 = f1 +f2
                f1 = f2
                f2 = f3
            Next
            textOutput.Text = f3
        End If
    End Sub

> π/4 = 1 - 1/3 + 1/5 +...+(-1)^n+1*1/(2n+1),计算π

    Private Sub cmdPI_Click()
        Dim s As Single         '和
        Dim m As Single         '通项
        Dim flag As Integer
        Dim i As Integer        '第i项
        flag = 1 : i = 1
        Do
        m = 1/(2i-1)
        s = s + flag * m
        flag = -flag
        Loop While(m>=0.0001)
        Text1.Text = s*4
    End Sub

---

## 过程
