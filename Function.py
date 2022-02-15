'和定义Sub子流程不同的是，定义函数可以返回值，可以直接在Excel使用
Function RandomLogic() As Boolean '返回值数据类型
    RandomLogic = Rnd() > 0.5 '内置函数
End Function


Function Add2Number(num1 As Double, num2 As Double) As Double '带参数的函数
    Add2Number = num1 + num2
End Function


Sub Main()'使用变量存储函数返回的值
    Dim result As Double
    result = Add2Number(12, 345) '调用函数
    If RandomLogic = True Then '注意这里相等不是用==
    	MsgBox result
    Else 
    	MsgBox "False"
    End If
End Sub

Sub Condition()
Dim Cities As String
Dim Judge As Integer
For i = 2 To 30
	Cities = Range("A"&i).Value
	Judge = Range("C"&i).Value
	Select Case Cities '对同一个值进行快速If Else筛选
	   Case Imp "省"
	   Judge = 1
	   Case  “A102”
	   Judge = 300
	   Case Else
	   Price=900
End Case