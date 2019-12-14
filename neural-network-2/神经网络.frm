VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   15060
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "导入神经网络"
      Height          =   495
      Left            =   9120
      TabIndex        =   19
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "导出神经网络"
      Height          =   495
      Left            =   9120
      TabIndex        =   18
      Top             =   4560
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawMode        =   2  'Blackness
      Height          =   2415
      Left            =   3960
      ScaleHeight     =   2415
      ScaleMode       =   0  'User
      ScaleWidth      =   6135
      TabIndex        =   10
      Top             =   480
      Width           =   6135
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "初始化/重置神经网络"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "开始/继续训练"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   6600
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   7080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始/继续识别"
      Height          =   495
      Left            =   11880
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   280
      Left            =   120
      ScaleHeight     =   280
      ScaleMode       =   0  'User
      ScaleWidth      =   280
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   280
   End
   Begin VB.Label Label11 
      Height          =   4095
      Left            =   11640
      TabIndex        =   17
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "识别结果："
      Height          =   495
      Left            =   11520
      TabIndex        =   16
      Top             =   5760
      Width           =   3735
   End
   Begin VB.Label Label10 
      Caption         =   "当前数据集详情"
      Height          =   615
      Left            =   960
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "loss"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "process"
      Height          =   255
      Left            =   10320
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   10440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Label Label6 
      Caption         =   "神经元个数："
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2805
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   2805
   End
   Begin VB.Label Label4 
      Caption         =   "图片长度："
      Height          =   855
      Left            =   960
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "图片宽度："
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "图片个数："
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "魔数："
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim trainimages() As Byte
Dim trainlabels() As Byte
Dim paint() As Integer
Dim amount As Long
Dim width1, height1 As Integer
Dim e As Double
Dim learning_rate As Double
Dim a() As Double
Dim w() As Double
Dim b() As Double
Dim z() As Double
Dim a_pd() As Double
Dim w_pd() As Double
Dim b_pd() As Double
Dim z_pd() As Double
Dim input_amount As Integer
Dim output_amount As Integer
Dim layer1_amount As Integer
Dim layer2_amount As Integer
Dim loss As Variant
Dim right As Integer
Dim process As Long
Dim training As Boolean
Dim ocring As Boolean
Dim result As Integer
Dim success As Long
Dim exist As Boolean
Dim mostpossible As Double

Private Sub Command1_Click()
Dim data As String
Dim i As Long
Open App.Path & "\data_w" For Output As #3
For i = 0 To UBound(w(), 1)
    data = data + CStr(w(i)) + vbCrLf
    Next i
Print #3, data
Close #3
data = ""
Open App.Path & "\data_b" For Output As #3
For i = 0 To UBound(b(), 1)
    data = data + CStr(b(i)) + vbCrLf
    Next i
Print #3, data
Close #3
MsgBox "神经网络导出成功！"
End Sub
Private Sub Command5_Click()
Dim i As Long
Dim data As String
Open App.Path & "\data_w" For Input As #3
Do While Not EOF(3)
    Line Input #3, data
    w(i) = Val(data)
    i = i + 1
    Loop
Close #3
i = 0
Open App.Path & "\data_b" For Input As #3
Do While Not EOF(3)
    Line Input #3, data
    b(i) = Val(data)
    i = i + 1
    Loop
Close #3
MsgBox "神经网络导入成功！"
End Sub
Private Sub Command2_Click()
If training = False Then
    If ocring Then
        ocring = False
    Else
        Command2.Caption = "暂停识别"
        ocring = True
        success = 0
        head
    End If
    Do While ocring And (process < amount)
        Forward_propagation
        most
        drawpicture
        graph
        Label7.Caption = "损失函数(loss)值:" + Str(Format(loss, "0.000000000000"))
        process = process + 1
        Label5.Caption = "当前进度：" + Str(process) + "/" + Str(amount)
        ProgressBar1.Value = process / amount * 100
        DoEvents
        Loop
    ocring = False
    Command2.Caption = "开始/继续识别"
Else
    MsgBox ("训练中无法进行识别")
End If
End Sub

Private Sub Command3_Click()

If ocring = False Then
    If training Then
        training = False
    Else
        Command3.Caption = "暂停训练"
        Dim i As Long
        training = True
        success = 0
        head
        For i = 0 To UBound(w(), 1)
            w(i) = Rnd() * 2 - 1
            Next i
    End If
    Do While training And (process < amount)
        Forward_propagation
        most
        Backward_propagation
        drawpicture
        graph
        Label7.Caption = "损失函数(loss)值:" + Str(Format(loss, "0.000000000000"))
        process = process + 1
        Label5.Caption = "当前进度：" + Str(process) + "/" + Str(amount)
        ProgressBar1.Value = process / amount * 100
        DoEvents
        Loop
    training = False
    Command3.Caption = "开始/继续训练"
Else
    MsgBox ("识别中无法进行训练")
End If
End Sub
Private Sub head()
process = 0
Close #1
Close #2
If training Then
Open App.Path & "\train-images.idx3-ubyte" For Binary As #1
Open App.Path & "\train-labels.idx1-ubyte" For Binary As #2
Else
Open App.Path & "\t10k-images.idx3-ubyte" For Binary As #1
Open App.Path & "\t10k-labels.idx1-ubyte" For Binary As #2
End If
ReDim trainimages(3)
Get #1, , trainimages()
Dim j, k As Long
Dim i As Variant
j = 3 '解析魔数
k = 0
For Each i In trainimages()
k = k + i * (256 ^ j)
j = j - 1
Next i
Label1.Caption = "魔数：" + Str(k)
Get #1, , trainimages() '解析图片个数
j = 3
k = 0
For Each i In trainimages()
k = k + i * (256 ^ j)
j = j - 1
Next i
amount = k
Label2.Caption = "图片个数：" + Str(amount)
Get #1, , trainimages() '解析图片宽度
j = 3
k = 0
For Each i In trainimages()
k = k + i * (256 ^ j)
j = j - 1
Next i
width1 = k
Label3.Caption = "图片宽度：" + Str(width1)
Get #1, , trainimages() '解析图片长度
j = 3
k = 0
For Each i In trainimages()
k = k + i * (256 ^ j)
j = j - 1
Next i
height1 = k
Label4.Caption = "图片长度：" + Str(height1)
ReDim trainlabels(3)
Get #2, , trainlabels()
j = 3
k = 0
For Each i In trainlabels()
k = k + i * (256 ^ j)
j = j - 1
Next i
Get #2, , trainlabels()
j = 3
k = 0
For Each i In trainlabels()
k = k + i * (256 ^ j)
j = j - 1
Next i
If exist = False Then
    Picture1.Width = Screen.TwipsPerPixelX * width1
    Picture1.Height = Screen.TwipsPerPixelY * height1
    input_amount = height1 * width1
    ReDim paint(height1 * width1 - 1)
    Label6.Caption = "神经元个数：" + Str(input_amount + layer1_amount + layer2_amount + output_amount)
    ReDim a(input_amount + output_amount + layer1_amount + layer2_amount - 1)
    ReDim w(input_amount * layer1_amount + layer1_amount * layer2_amount + layer2_amount * output_amount - 1)
    ReDim b(input_amount + output_amount + layer1_amount + layer2_amount - 1)
    ReDim z(input_amount + output_amount + layer1_amount + layer2_amount - 1)
    ReDim a_pd(input_amount + output_amount + layer1_amount + layer2_amount - 1)
    ReDim w_pd(input_amount * layer1_amount + layer1_amount * layer2_amount + layer2_amount * output_amount - 1)
    ReDim b_pd(input_amount + output_amount + layer1_amount + layer2_amount - 1)
    ReDim z_pd(input_amount + output_amount + layer1_amount + layer2_amount - 1)
End If
exist = True
End Sub



Private Sub Form_Load()
e = 2.71828
learning_rate = 0.01
Picture1.AutoRedraw = True
Picture1.ScaleMode = 3
Picture2.AutoRedraw = True
training = False
ocring = False
Line1.X1 = Picture2.Left - 8
Line1.X2 = Picture2.Left - 8
Line1.Y1 = Picture2.Top - 100
Line1.Y2 = Picture2.Top + Picture2.Height + 200
Line2.X1 = Picture2.Left - 200
Line2.X2 = Picture2.Left + Picture2.Width + 100
Line2.Y1 = Picture2.Top + Picture2.Height
Line2.Y2 = Picture2.Top + Picture2.Height
output_amount = 10
layer1_amount = 32
layer2_amount = 32
End Sub
Private Function drawpicture()
Dim xb, yb As Long
Dim i As Integer
i = 0
Dim X1, Y1 As Long
For Y1 = 0 To width1 - 1
    For X1 = 0 To height1 - 1
       Picture1.Circle (X1, Y1), 1, RGB(paint(i), paint(i), paint(i))        '绘制图像
       i = i + 1
       Next X1
    Next Y1
Image1.Picture = Picture1.Image
End Function
Private Function graph()
Dim xb, yb As Long
Dim X1, Y1 As Long
X1 = process / amount * Picture2.ScaleWidth
Y1 = -loss / 0.5 * Picture2.ScaleHeight + Picture2.ScaleHeight
Picture2.Circle (X1, Y1), 10, RGB(200, 200, 200)        '绘制loss-t图像
End Function
Private Function most()
Dim j As Integer
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    If a(j) > mostpossible Then
        mostpossible = a(j)
        result = j - (input_amount + layer1_amount + layer2_amount)
    End If
    Next j
Label17.Caption = "识别结果：" + Str(result)
If result = right Then
    success = success + 1
End If
If process = 0 Then
    Label17.Caption = Label17.Caption + " 成功率：null"
Else
    Label17.Caption = Label17.Caption + " 成功率：" + Str(CInt(success / process * 100)) + "%"
End If
End Function
Public Function sigmoid(x As Variant)
sigmoid = 1 / (1 + e ^ (-x))
End Function

Private Function Forward_propagation()      '正向传播
Dim sum As Double
Label11.Caption = ""
mostpossible = 0
Dim j, k As Long
Dim test As Double
Dim i As Variant
ReDim trainimages(0)
ReDim trainlabels(0)
For i = 0 To UBound(z(), 1)       '重置z()
    z(i) = 0
    Next i
loss = 0        '重置loss
Get #2, , trainlabels()     '获取目标数字
right = trainlabels(0)
For j = 0 To input_amount - 1
    Get #1, , trainimages()
    a(j) = sigmoid(trainimages(0))       '输入层赋值
    paint(j) = trainimages(0)
    Next j
For j = input_amount To input_amount + layer1_amount - 1
    For k = 0 To input_amount - 1
        z(j) = z(j) + a(k) * w((j - input_amount) * input_amount + k)       'layer1赋值
        Next k
        z(j) = z(j) + b(j)
        a(j) = sigmoid(z(j))
    Next j
For j = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
    For k = input_amount To input_amount + layer1_amount - 1
        z(j) = z(j) + a(k) * w(input_amount * layer1_amount + (j - input_amount - layer1_amount) * layer1_amount + k - input_amount)          'layer2赋值
        Next k
        z(j) = z(j) + b(j)
        a(j) = sigmoid(z(j))
    Next j
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    For k = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
        z(j) = z(j) + a(k) * w(input_amount * layer1_amount + layer1_amount * layer2_amount + (j - input_amount - layer1_amount - layer2_amount) * layer2_amount + k - input_amount - layer1_amount)    '输出层赋值
        Next k
        z(j) = z(j) + b(j)
        sum = sum + e ^ z(j)
    Next j
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    a(j) = e ^ z(j) / sum
    Next j
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1        '损失函数计算
    If j - (input_amount + layer1_amount + layer2_amount) = right Then
        loss = loss - Log(a(j))
    Else
        loss = loss
    End If
loss = loss / 2
    Label11.Caption = Label11.Caption + CStr(j - (input_amount + layer1_amount + layer2_amount)) + "的概率：" + CStr(Format(a(j), "0.000000")) + vbCrLf
        If a(j) > mostpossible Then
        mostpossible = a(j)
        result = j - (input_amount + layer1_amount + layer2_amount)
        End If
    Next j
End Function
Private Function Backward_propagation()     '反向传播
Dim j, k As Long
For j = 0 To UBound(a_pd(), 1)       '重置a_pd()
    a_pd(j) = 0
    Next j
For j = 0 To UBound(w_pd(), 1)       '重置w_pd()
    w_pd(j) = 0
    Next j
'输出层
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    If j - (input_amount + layer1_amount + layer2_amount) = right Then
        z_pd(j) = a(j) - 1      '损失函数对输出层神经元加权和的偏导数
    Else
        z_pd(j) = a(j)
    End If
    b_pd(j) = z_pd(j)
    Next j
For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    For k = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
        w_pd(input_amount * layer1_amount + layer1_amount * layer2_amount + (j - input_amount - layer1_amount - layer2_amount) * layer2_amount + k - input_amount - layer1_amount) = a(k) * z_pd(j)
        Next k
    Next j
'隐蔽层2
For k = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
    For j = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
        a_pd(k) = a_pd(k) + w(input_amount * layer1_amount + layer1_amount * layer2_amount + (j - input_amount - layer1_amount - layer2_amount) * layer2_amount + k - input_amount - layer1_amount) * z_pd(j)        '损失函数对隐蔽层2神经元数值的偏导数
        Next j
    Next k
For j = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
    z_pd(j) = a_pd(j) * (sigmoid(z(j)) * (1 - sigmoid(z(j))))
    b_pd(j) = z_pd(j)
    Next j
For j = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
    For k = input_amount To input_amount + layer1_amount - 1
        w_pd(input_amount * layer1_amount + (j - input_amount - layer1_amount) * layer1_amount + k - input_amount) = a(k) * z_pd(j)
        Next k
    Next j
'隐蔽层1
For k = input_amount To input_amount + layer1_amount - 1
    For j = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
        a_pd(k) = a_pd(k) + w(input_amount * layer1_amount + (j - input_amount - layer1_amount) * layer1_amount + k - input_amount) * z_pd(j)        '损失函数对隐蔽层1神经元数值的偏导数
        Next j
    Next k
For j = input_amount To input_amount + layer1_amount - 1
    z_pd(j) = a_pd(j) * (sigmoid(z(j)) * (1 - sigmoid(z(j))))
    b_pd(j) = z_pd(j)
    Next j
For j = input_amount To input_amount + layer1_amount - 1
    For k = 0 To input_amount - 1
        w_pd((j - input_amount) * input_amount + k) = a(k) * z_pd(j)
        Next k
    Next j
For j = 0 To UBound(w(), 1)
    w(j) = w(j) - w_pd(j) * learning_rate
     Next j
For j = 0 To UBound(b(), 1)
    b(j) = b(j) - b_pd(j) * learning_rate
    Next j
End Function

Private Sub Command4_Click()
Dim i As Long
For i = 0 To UBound(w(), 1)
    w(i) = Rnd() * 2 - 1
    Next i
End Sub


