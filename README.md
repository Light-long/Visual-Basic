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

    Form1.Move 1000,1000,2000,2000

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

    .vbp + .frm

