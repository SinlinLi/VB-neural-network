VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11730
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "重画"
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "识别数字"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   1680
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   -360
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label11 
      Height          =   4095
      Left            =   8400
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label17 
      Caption         =   "识别结果："
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amount As Long
Dim width1, height1 As Integer
Dim e As Double
Dim a() As Double
Dim w() As Double
Dim b() As Double
Dim z() As Double
Dim input_amount As Integer
Dim output_amount As Integer
Dim layer1_amount As Integer
Dim layer2_amount As Integer
Dim loss As Variant
Dim right As Integer
Dim process As Long
Dim result As Integer
Dim mostpossible As Double

Private Sub Command1_Click()
Picture1.Picture = LoadPicture("")
End Sub

Private Sub Form_Load()
Dim data As String
Dim i As Long
Dim J As Variant
e = 2.71828
height1 = 28
width1 = 28
Picture1.AutoRedraw = True
Picture1.ScaleMode = 3
output_amount = 10
layer1_amount = 32
layer2_amount = 32
Picture1.Width = Screen.TwipsPerPixelX * width1 * 10
Picture1.Height = Screen.TwipsPerPixelY * height1 * 10
Picture1.ScaleWidth = width1
Picture1.ScaleHeight = height1
Image2.Picture = LoadPicture(App.Path & "\point.bmp")
input_amount = height1 * width1
ReDim Paint(height1 * width1 - 1)
ReDim a(input_amount + output_amount + layer1_amount + layer2_amount - 1)
ReDim w(input_amount * layer1_amount + layer1_amount * layer2_amount + layer2_amount * output_amount - 1)
ReDim b(input_amount + output_amount + layer1_amount + layer2_amount - 1)
ReDim z(input_amount + output_amount + layer1_amount + layer2_amount - 1)
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
End Sub
Private Function most()
Dim J As Integer
For J = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    If a(J) > mostpossible Then
        mostpossible = a(J)
        result = J - (input_amount + layer1_amount + layer2_amount)
    End If
    Next J
Label17.Caption = "识别结果：" + Str(result)
End Function
Public Function sigmoid(X As Variant)
sigmoid = 1 / (1 + e ^ (-X))
End Function

Private Function Forward_propagation()      '正向传播
Dim sum As Double
Label11.Caption = ""
mostpossible = 0
Dim J, k, l As Long
Dim test As Double
Dim i As Variant
For i = 0 To UBound(z(), 1)       '重置z()
    z(i) = 0
    Next i
For i = 1 To height1
    For J = 1 To width1
        If Picture1.Point(J, i) = 0 Then
            a(l) = sigmoid(Picture1.Point(J, i))        '输入层赋值
        Else
            a(l) = sigmoid(255)
        End If
        l = l + 1
    Next J
Next i
For J = input_amount To input_amount + layer1_amount - 1
    For k = 0 To input_amount - 1
        z(J) = z(J) + a(k) * w((J - input_amount) * input_amount + k)       'layer1赋值
        Next k
        z(J) = z(J) + b(J)
        a(J) = sigmoid(z(J))
    Next J
For J = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
    For k = input_amount To input_amount + layer1_amount - 1
        z(J) = z(J) + a(k) * w(input_amount * layer1_amount + (J - input_amount - layer1_amount) * layer1_amount + k - input_amount)          'layer2赋值
        Next k
        z(J) = z(J) + b(J)
        a(J) = sigmoid(z(J))
    Next J
For J = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    For k = input_amount + layer1_amount To input_amount + layer1_amount + layer2_amount - 1
        z(J) = z(J) + a(k) * w(input_amount * layer1_amount + layer1_amount * layer2_amount + (J - input_amount - layer1_amount - layer2_amount) * layer2_amount + k - input_amount - layer1_amount)    '输出层赋值
        Next k
        z(J) = z(J) + b(J)
        sum = sum + e ^ z(J)
    Next J
For J = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    a(J) = e ^ z(J) / sum
    Next J
For J = input_amount + layer1_amount + layer2_amount To input_amount + layer1_amount + layer2_amount + output_amount - 1
    Label11.Caption = Label11.Caption + CStr(J - (input_amount + layer1_amount + layer2_amount)) + "的概率：" + CStr(Format(a(J), "0.000000")) + vbCrLf
    If a(J) > mostpossible Then
    mostpossible = a(J)
    result = J - (input_amount + layer1_amount + layer2_amount)
    End If
    Next J
End Function
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i, J As Long
If Button = 1 Then
    Picture1.PaintPicture Image2.Picture, CInt(X), CInt(Y), 1, 1
    End If
End Sub
Private Sub Command6_Click()
Forward_propagation
most
End Sub
