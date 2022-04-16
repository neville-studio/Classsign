VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "班级签到系统"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   6570
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   6840
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   6840
   End
   Begin VB.CommandButton Command11 
      Caption         =   "登出"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "修改密码"
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2280
      TabIndex        =   24
      Text            =   "再输一遍新密码"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2280
      TabIndex        =   23
      Text            =   "输入新密码"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2280
      TabIndex        =   22
      Text            =   "输入原密码"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "请假更改"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   2280
      TabIndex        =   20
      Text            =   "输入请假人员"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      Caption         =   "删除值日生"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "添加值日生"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "更改签到截止时间"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "值日生签到"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "管理员操作"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2280
      TabIndex        =   13
      Text            =   "输入管理员密码操作"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "帮助"
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "输出记录"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "签到"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox List4 
      Height          =   780
      Left            =   4320
      TabIndex        =   5
      Top             =   8880
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   8160
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   0
      TabIndex        =   3
      Top             =   8160
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7560
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) XuPeng Studio 2019～2020. 感谢施振展同志提出的建设性意见。 "
      Height          =   180
      Left            =   0
      TabIndex        =   16
      Top             =   9840
      Width           =   6480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.03"
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "值日生签到记录"
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   8640
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签到记录"
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "未签名单，选择后单击“签到”"
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "值日生签到"
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   7920
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim names(1 To 100) As String
Dim qingjias(1 To 100) As Boolean
Dim qiandaos(1 To 100) As Boolean
Dim settime As Date
Dim total As Integer
Dim key(1 To 3) As Long
Private Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Dim passwordtime As Integer
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Sub GenKey()
Dim d As Long, phi As Long, e As Long
Dim m As Long, x As Long, q As Long
Dim p As Long
Randomize
On Error GoTo top
top:
p = Rnd * 1000 \ 1
If IsPrime(p) = False Then GoTo top
Sel_q:
q = Rnd * 1000 \ 1
If IsPrime(q) = False Then GoTo Sel_q
n = p * q \ 1
phi = (p - 1) * (q - 1) \ 1
d = Rnd * n \ 1
If d = 0 Or n = 0 Or d = 1 Then GoTo top
e = Euler(phi, d)
If e = 0 Or e = 1 Then GoTo top

x = Mult(255, e, n)
If Not Mult(x, d, n) = 255 Then
    DoEvents
    GoTo top
ElseIf Mult(x, d, n) = 255 Then
    key(1) = e
    key(2) = d
    key(3) = n
End If
End Sub

Private Function Euler(ByVal a As Long, ByVal b As Long) As Long
On Error GoTo error2
r1 = a: r = b
p1 = 0: p = 1
q1 = 2: q = 0
n = -1
Do Until r = 0
    r2 = r1: r1 = r
    p2 = p1: p1 = p
    q2 = q1: q1 = q
    n = n + 1
    r = r2 Mod r1
    c = r2 \ r1
    p = (c * p1) + p2
    q = (c * q1) + q2
Loop
s = (b * p1) - (a * q1)
If s > 0 Then
    x = p1
Else
    x = (0 - p1) + a
End If
Euler = x
Exit Function

error2:
Euler = 0
End Function

Private Function Mult(ByVal x As Long, ByVal p As Long, ByVal m As Long) As Long
y = 1
On Error GoTo error1
Do While p > 0
    Do While (p / 2) = (p \ 2)
        x = (x * x) Mod m
        p = p / 2
    Loop
    y = (x * y) Mod m
    p = p - 1
Loop
Mult = y
Exit Function

error1:
y = 0
End Function

Private Function IsPrime(lngNumber As Long) As Boolean
Dim lngCount As Long
Dim lngSqr As Long
Dim x As Long

    lngSqr = Sqr(lngNumber) ' get the int square root

    If lngNumber < 2 Then
IsPrime = False
Exit Function
End If

lngCount = 2
IsPrime = True

If lngNumber Mod lngCount = 0& Then
IsPrime = False
Exit Function
End If

lngCount = 3

For x& = lngCount To lngSqr Step 2
If lngNumber Mod x& = 0 Then
IsPrime = False
Exit Function
End If
Next
End Function

Private Function Base64_Encode(DecryptedText As String) As String
Dim c1, c2, c3 As Integer
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String
For n = 1 To Len(DecryptedText) Step 3
c1 = Asc(Mid$(DecryptedText, n, 1))
c2 = Asc(Mid$(DecryptedText, n + 1, 1) + Chr$(0))
c3 = Asc(Mid$(DecryptedText, n + 2, 1) + Chr$(0))
w1 = Int(c1 / 4)
w2 = (c1 And 3) * 16 + Int(c2 / 16)
If Len(DecryptedText) >= n + 1 Then w3 = (c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
      If Len(DecryptedText) >= n + 2 Then w4 = c3 And 63 Else w4 = -1

      retry = retry + mimeencode(w1) + mimeencode(w2) + mimeencode(w3) + mimeencode(w4)
   Next
   Base64_Encode = retry
End Function

Private Function Base64_Decode(a As String) As String
Dim w1 As Integer
Dim w2 As Integer
Dim w3 As Integer
Dim w4 As Integer
Dim n As Integer
Dim retry As String

   For n = 1 To Len(a) Step 4
      w1 = mimedecode(Mid$(a, n, 1))
      w2 = mimedecode(Mid$(a, n + 1, 1))
      w3 = mimedecode(Mid$(a, n + 2, 1))
      w4 = mimedecode(Mid$(a, n + 3, 1))
      If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
      If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
      If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
   Next
   Base64_Decode = retry
End Function

Private Function mimeencode(w As Integer) As String
   If w >= 0 Then mimeencode = Mid$(base64, w + 1, 1) Else mimeencode = ""
End Function

Private Function mimedecode(a As String) As Integer
   If Len(a) = 0 Then mimedecode = -1: Exit Function
   mimedecode = InStr(base64, a) - 1
End Function

Public Function Encode(ByVal Inp As String, ByVal e As Long, ByVal n As Long) As String
Dim s As String
s = ""
m = Inp
If m = "" Then Exit Function
s = Mult(CLng(Asc(Mid(m, 1, 1))), e, n)
For i = 2 To Len(m)
    s = s & "+" & Mult(CLng(Asc(Mid(m, i, 1))), e, n)
Next i
Encode = Base64_Encode(s)
End Function

Public Function Decode(ByVal Inp As String, ByVal d As Long, ByVal n As Long) As String
St = ""
ind = Base64_Decode(Inp)
For i = 1 To Len(ind)
    nxt = InStr(i, ind, "+")
    If Not nxt = 0 Then
        tok = Val(Mid(ind, i, nxt))
    Else
        tok = Val(Mid(ind, i))
    End If
    St = St + Chr(Mult(CLng(tok), d, n))
    If Not nxt = 0 Then
        i = nxt
    Else
        i = Len(ind)
    End If
Next i
Decode = St
End Function

Private Sub Combo1_Change()
    For i = 1 To Combo1.ListCount
        If Combo1.List(i - 1) = Combo1.Text Then
        Command1.Enabled = True: Command7.Enabled = True: Exit Sub
        End If
    Next i
    Command1.Enabled = False: Command7.Enabled = False
End Sub

Private Sub Combo1_Click()
    Command1.Enabled = True
    Command7.Enabled = True
End Sub

Private Sub Command1_Click()
    For i = 1 To List1.ListCount
        If List1.List(i - 1) = Combo1.Text Then Exit For
    Next i
    If i > List1.ListCount Then Exit Sub
    List3.AddItem List1.List(i - 1) + "  " + CStr(DateTime.Time)
    For j = 1 To n
    If names(j) = Combo1.Text Then
    qiandaos(j) = True
    End If
    Next j
    List1.RemoveItem (i - 1)
    Combo1.RemoveItem (i - 1)
    Combo1.Text = ""
    Command1.Enabled = False
    Command7.Enabled = False
    Command2.Enabled = True
End Sub


Private Sub Command10_Click()
 Const ForReading = 1, ForWriting = 2, ForAppending = 3
   Dim fs, f
   On Error GoTo err1
    If Text3.PasswordChar = "*" And Text4.PasswordChar = "*" And Text5.PasswordChar = "*" Then
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\pw")
    Set textf = f.openastextstream(1, -2)
    If Text3.Text = Decode(textf.readall, key(2), key(3)) Then
        textf.Close
        If Text4.Text = Text5.Text Then
            Set f = fs.createtextfile(App.Path + "\pw")
            f.writeline (Encode(Text4.Text, key(1), key(3)))
            MsgBox ("修改成功！")
            Text3.Text = "": Text4.Text = "": Text5.Text = ""
            Call text3_Lostfocus
            Call text4_Lostfocus
            Call text5_Lostfocus
        Else
            MsgBox ("两次新密码输入不一致！")
            Text4.Text = "": Text5.Text = ""
            Call text4_Lostfocus
            Call text5_Lostfocus

        End If
        Else
        textf.Close
        Text3.Text = ""
            Call text3_Lostfocus
            passwordtime = passwordtime - 1
            If passwordtime > 20 Then
                MsgBox ("原密码错误，还可以输入" + CStr(passwordtime Mod 5) + "次")
            Else
            MsgBox ("原密码错误次数过多，强制登出！")
            Call Command11_Click
        End If
    End If
    End If
err1:
    If Err.Number <> 0 Then
        MsgBox ("修改失败，当前存在着许多问题，导致密码修改遭到阻止。错误号：" + CStr(Err.Number))
    End If
        
End Sub

Private Sub Command11_Click()
    Command6.Visible = False
    Command7.Visible = False
    Command8.Visible = False
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text1.Visible = True
    Command4.Visible = True
End Sub

Private Sub Command2_Click()
    Dim fs, f, ts, s, ss
    Dim savetime As String
    savetime = CStr(DateTime.Year(Date)) + "-" + CStr(DateTime.Month(Date)) + "-" + CStr(DateTime.Day(Date)) + "-" + CStr(DateTime.Hour(Time)) + "-" + CStr(DateTime.Minute(Time)) + "-" + CStr(DateTime.Second(Time))
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path + "\签到记录" + savetime + ".txt")
    f.writeline ("记录时间：" + CStr(Date) + "  " + CStr(Time))
    f.writeline ("签到记录：")
    For i = 1 To List3.ListCount
    f.writeline (List3.List(i - 1))
    Next i
    If List4.ListCount > 0 Then
    f.writeline ("值日生：")
    For i = 1 To List4.ListCount
    f.writeline (List4.List(i - 1))
    Next i
    End If
    If List1.ListCount > 0 Then
    f.writeline ("未签名单：")
    For i = 1 To List1.ListCount
    f.writeline (List1.List(i - 1) + " 未签")
    Next i
    End If
    If List2.ListCount > 0 Then
    f.writeline ("值日生未签名单：")
    For i = 1 To List2.ListCount
    f.writeline (List2.List(i - 1) + " 未签")
    Next i
    End If
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Shell ("explorer.exe " + App.Path + "\help.html")
End Sub

Private Sub Command4_Click()
    If Text1.PasswordChar = "*" Then
    passwordtime = passwordtime - 1
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\pw")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    If Decode(ts.readall, key(2), key(3)) = Text1.Text Then
    Command6.Visible = True
    Command7.Visible = True
    Command8.Visible = True
    Command9.Visible = True
    Command10.Visible = True
    Command11.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Command4.Visible = False
    Text1.Visible = False
    passwordtime = 25
    ElseIf passwordtime = 25 Or (passwordtime Mod 5 <> 0 And passwordtime > 0) Then
        MsgBox ("密码错误，还可以输入" + CStr(passwordtime Mod 5) + "次")
    ElseIf passwordtime > 0 Then
        MsgBox ("密码错误，请在" + CStr(5 - passwordtime \ 5) + "分钟后重试")
        Timer2.Enabled = True
        Text1.Enabled = False
        Command4.Enabled = False
        Text1.Text = ""
        Call Text1_Lostfocus
    Else
        MsgBox ("还不赶紧好好学习！")
        Command4.Enabled = False
        Text1.Enabled = False
        Call Text1_Lostfocus
        Text1.Text = "您已无法输入密码。"
    End If
    End If
    Text1.PasswordChar = ""
    If Not Timer2.Enabled And passwordtime > 0 Then Text1.Text = "输入管理员密码操作"
End Sub

Private Sub Command5_Click()
If List2.ListIndex >= 0 Then
    If TimeSerial(Hour(settime), Minute(settime) - 10, Second(Time)) < DateTime.Time Then
        List4.AddItem "--------------"
        List4.AddItem "以下迟到"
    End If
    List4.AddItem List2.List(List2.ListIndex) + "  " + CStr(DateTime.Time)
    List2.RemoveItem (List2.ListIndex)
    Command5.Enabled = False
    Command8.Enabled = False
    Command2.Enabled = True
End If
End Sub

Private Sub Command6_Click()
    On Error GoTo rego
    settime = TimeValue(InputBox("输入签到截止时间："))
    For i = 1 To List3.ListCount
        If List3.List(i - 1) = "------------" Then
            List3.RemoveItem (i - 1)
            List3.RemoveItem (i - 1)
            Timer1.Enabled = True
            Exit For
        End If
    Next i
    For i = 1 To List4.ListCount
        If List4.List(i - 1) = "--------------" Then
            List4.RemoveItem (i - 1)
            List4.RemoveItem (i - 1)
        End If
    Next i
rego:
    If Err.Number = 13 Then
    MsgBox ("输入的签到截止时间非法！")
    End If
End Sub

Private Sub Command7_Click()
    For i = 1 To List1.ListCount
        If List1.List(i - 1) = Combo1.Text Then Exit For
    Next i
    If i > List1.ListCount Then Exit Sub
    Combo1.Text = ""
    Command1.Enabled = False
    Command7.Enabled = False
    List2.AddItem List1.List(i - 1)
    Combo1.RemoveItem i - 1
    List1.RemoveItem i - 1
    Command2.Enabled = True
End Sub

Private Sub Command8_Click()
    Combo1.AddItem List2.List(List2.ListIndex)
    List1.AddItem List2.List(List2.ListIndex)
    List2.RemoveItem List2.ListIndex
    Command8.Enabled = False
    Command5.Enabled = False
    Command2.Enabled = True
End Sub

Private Sub Command9_Click()
    For i = 1 To total
        If names(i) = Text2.Text Then qingjias(i) = Not qingjias(i): Exit For
    Next i
    If i <= total And Not qingjias(i) And Not qiandaos(i) Then
        List1.AddItem names(i)
        Combo1.AddItem names(i)
    End If
    If qingjias(i) And Not qiandaos(i) Then
    For j = 1 To List1.ListCount
    If List1.List(j - 1) = Text2.Text Then List1.RemoveItem (j - 1): Combo1.RemoveItem (j - 1): Combo1.Text = ""
    Next j
    End If
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\mingdan.txt")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    
    Do While Not ts.atendofstream
        total = total + 1
        s = ts.readline
        List1.AddItem s
        Combo1.AddItem s
    Loop
    For i = 1 To List1.ListCount
    names(i) = List1.List(i)
    Next i
    ts.Close
    settime = TimeValue("23:59:59")
     key(1) = 33965
     key(2) = 32717
     key(3) = 41831
     passwordtime = 25
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Timer1.Enabled And Not Command11.Visible And List1.ListCount + List2.ListCount > 0 Then
        Cancel = 1
        MsgBox ("无法关闭，原因是：还有人没签到，或签到时间没截止，或没以管理员身份登录。")
    Else
        If Command2.Enabled Then Call Command2_Click
        Cancel = 0
    End If
    
End Sub

Private Sub List1_Click()
    Combo1.Text = CStr(List1.List(List1.ListIndex))
    Command1.Enabled = True
    Command7.Enabled = True
End Sub

Private Sub List2_Click()
    Command5.Enabled = True
    Command8.Enabled = True
End Sub

Private Sub Text1_gotfocus()
    If Text1.Text = "输入管理员密码操作" Then Text1.PasswordChar = "*":     Text1.Text = ""
End Sub

Private Sub Text1_Lostfocus()
    If Text1.Text = "" Then Text1.PasswordChar = ""
    If Text1.Text = "" And ((passwordtime > 0 And passwordtime Mod 5 <> 0) Or passwordtime = 25) Then Text1.Text = "输入管理员密码操作"
End Sub
Private Sub Text2_gotfocus()
    If Text2.Text = "输入请假人员" Then Text2.Text = ""
End Sub
Private Sub text2_lostfocus()
    If Text2.Text = "" Then Text2.Text = "输入请假人员"
End Sub

Private Sub Text3_gotfocus()
    If Text3.PasswordChar = "" Then
    Text3.Text = ""
    Text3.PasswordChar = "*"
    End If
End Sub
Private Sub text3_Lostfocus()
    If Text3.Text = "" Then Text3.PasswordChar = "": Text3.Text = "输入原密码"
End Sub
Private Sub text4_gotfocus()
    If Text4.PasswordChar = "" Then
    Text4.Text = ""
    Text4.PasswordChar = "*"
    End If
End Sub
Private Sub text4_Lostfocus()
    If Text4.Text = "" Then Text4.PasswordChar = "": Text4.Text = "输入新密码"
End Sub
Private Sub Text5_gotfocus()
    If Text5.PasswordChar = "" Then
    Text5.Text = ""
    Text5.PasswordChar = "*"
    End If
End Sub
Private Sub text5_Lostfocus()
    If Text5.Text = "" Then Text5.PasswordChar = "": Text5.Text = "再输一遍新密码"
End Sub
Private Sub Timer1_Timer()
If settime < DateTime.Time Then
List3.AddItem "------------"
List3.AddItem "以下迟到："
Timer1.Enabled = False
Text1.Text = ""
Call Text1_gotfocus
Text1.Enabled = True
Command4.Enabled = True
passwordtime = 25
End If
End Sub

Private Sub Timer2_Timer()
Static timeout As Integer
timeout = timeout + 1
Text1.Enabled = False
Command4.Enabled = False
Text1.Text = CStr((5 - passwordtime \ 5) * 60 - timeout) + "秒后解锁。"
If timeout = (5 - passwordtime \ 5) * 60 Then
    Command4.Enabled = True
    timeout = 0
    Text1.Enabled = True
    Text1.Text = "输入管理员密码操作"
    Timer2.Enabled = False
End If
End Sub
