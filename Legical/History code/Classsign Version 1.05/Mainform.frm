VERSION 5.00
Begin VB.Form Mainform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "班级签到系统"
   ClientHeight    =   10125
   ClientLeft      =   4200
   ClientTop       =   885
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   6570
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer zttimer 
      Interval        =   1000
      Left            =   3840
      Top             =   6840
   End
   Begin VB.Timer zrshx 
      Interval        =   100
      Left            =   3360
      Top             =   6840
   End
   Begin VB.Timer passwordTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   6840
   End
   Begin VB.Timer lineTimer 
      Interval        =   100
      Left            =   2280
      Top             =   6840
   End
   Begin VB.CommandButton zrsCommand 
      Caption         =   "值日生签到"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton AdminCommand 
      Caption         =   "管理员操作"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox PasswordText 
      Height          =   270
      Left            =   2280
      TabIndex        =   13
      Text            =   "输入管理员密码操作"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton helpCommand 
      Caption         =   "帮助"
      Height          =   300
      Left            =   3480
      TabIndex        =   12
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton outputCommand 
      Caption         =   "输出记录"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton qiandaoCommand 
      Caption         =   "签到"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox NameCombo 
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
   Begin VB.ListBox records2List 
      Height          =   780
      Left            =   4320
      TabIndex        =   5
      Top             =   8880
      Width           =   2175
   End
   Begin VB.ListBox records1List 
      Height          =   8160
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox zrsList 
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
   Begin VB.ListBox qiandaoList 
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
   Begin VB.Label zttext 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   2280
      TabIndex        =   17
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Label CopyrightLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) XuPeng Studio 2019～2021. 感谢施振展同志提出的建设性意见。 "
      Height          =   180
      Left            =   0
      TabIndex        =   16
      Top             =   9840
      Width           =   6480
   End
   Begin VB.Label VersionLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.05"
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label records2Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "值日生签到记录"
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   8640
      Width           =   1260
   End
   Begin VB.Label records1Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签到记录"
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label qdLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "未签名单，选择后单击“签到”"
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label zrsqdLabel 
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
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tqtime As Integer
Dim names(1 To 100) As String
Public settime As Date
Public savepath As String
Public saveonexit As Integer
Dim total As Integer
Dim key(1 To 3) As Long
Private Const base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Public passwordtime As Integer
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

Private Sub NameCombo_Change()
    For i = 1 To NameCombo.ListCount
        If NameCombo.List(i - 1) = NameCombo.Text Then
            qiandaoCommand.Enabled = True: Exit Sub
        End If
    Next i
    qiandaoCommand.Enabled = False
End Sub

Private Sub NameCombo_Click()
    qiandaoCommand.Enabled = True
End Sub

Private Sub qiandaocommand_Click()
    For i = 1 To qiandaoList.ListCount
        If qiandaoList.List(i - 1) = NameCombo.Text Then Exit For
    Next i
    If i > qiandaoList.ListCount Then Exit Sub
    records1List.AddItem qiandaoList.List(i - 1) + "  " + CStr(DateTime.Time)
    qiandaoList.RemoveItem (i - 1)
    NameCombo.RemoveItem (i - 1)
    NameCombo.Text = ""
    qiandaoCommand.Enabled = False
    outputCommand.Enabled = True
End Sub
Private Sub outputcommand_Click()
    On Error GoTo errortest
    Dim fs, f, ts, s, ss
    Dim savetime As String
    savetime = CStr(DateTime.Year(Date)) + "-" + CStr(DateTime.Month(Date)) + "-" + CStr(DateTime.Day(Date)) + "-" + CStr(DateTime.Hour(Time)) + "-" + CStr(DateTime.Minute(Time)) + "-" + CStr(DateTime.Second(Time))
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(savepath + "\签到记录" + savetime + ".txt")
    f.writeline ("记录时间：" + CStr(Date) + "  " + CStr(Time))
    f.writeline ("签到记录：")
    For i = 1 To records1List.ListCount
    f.writeline (records1List.List(i - 1))
    Next i
    f.writeline ("--END OF RECORD--")
    If records2List.ListCount > 0 Then
    f.writeline ("值日生：")
    For i = 1 To records2List.ListCount
    f.writeline (records2List.List(i - 1))
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If qiandaoList.ListCount > 0 Then
    f.writeline ("未签名单：")
    For i = 1 To qiandaoList.ListCount
    f.writeline (qiandaoList.List(i - 1) + " 未签")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If zrslist.ListCount > 0 Then
    f.writeline ("值日生未签名单：")
    For i = 1 To zrslist.ListCount
    f.writeline (zrslist.List(i - 1) + " 未签")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If AdminForm.qingjialist.ListCount > 0 Then
    f.writeline ("请假名单：")
    For i = 1 To AdminForm.qingjialist.ListCount
    f.writeline (AdminForm.qingjialist.List(i - 1) + " 请假")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If fs.fileexists(App.Path + "\autosave.classsign") Then
    Set f = fs.getfile(App.Path + "\autosave.classsign")
    f.Delete
    End If
    outputCommand.Enabled = False
errortest:
    If Err.Number <> 0 Then MsgBox ("保存失败。请检查输出路径是否错误。错误原因：" & Chr(10) & Err.Number & Err.Description)
End Sub

Private Sub helpCommand_Click()
Shell ("explorer.exe " + App.Path + "\help.html")
End Sub

Private Sub adminCommand_Click()
    If PasswordText.PasswordChar = "*" Then
        passwordtime = passwordtime - 1
        Const ForReading = 1, ForWriting = 2, ForAppending = 3
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        Dim fs, f, ts, s, ss
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.getfile(App.Path + "\pw")
        Set ts = f.openastextstream(ForReading, TristateUseDefault)
        If Decode(ts.readall, key(2), key(3)) = PasswordText.Text Then
            AdminCommand.Visible = False
            PasswordText.Visible = False
            passwordtime = 25
            AdminForm.Show
        ElseIf passwordtime = 25 Or (passwordtime Mod 5 <> 0 And passwordtime > 0) Then
            MsgBox ("密码错误，还可以输入" + CStr(passwordtime Mod 5) + "次")
        ElseIf passwordtime > 0 Then
            MsgBox ("密码错误，请在" + CStr(5 - passwordtime \ 5) + "分钟后重试")
            passwordTimer.Enabled = True
            PasswordText.Enabled = False
            AdminCommand.Enabled = False
            PasswordText.Text = ""
            Call PasswordText_Lostfocus
        Else
            MsgBox ("还不赶紧好好学习！")
            AdminCommand.Enabled = False
            PasswordText.Enabled = False
            Call PasswordText_Lostfocus
            PasswordText.Text = "您已无法输入密码。"
        End If
    End If
    PasswordText.PasswordChar = ""
    If Not passwordTimer.Enabled And passwordtime > 0 Then PasswordText.Text = "输入管理员密码操作"
End Sub


Private Sub zrsCommand_Click()
If zrslist.ListIndex >= 0 Then

    records2List.AddItem zrslist.List(zrslist.ListIndex) + "  " + CStr(DateTime.Time)
    zrslist.RemoveItem (zrslist.ListIndex)
    zrsCommand.Enabled = False
    outputCommand.Enabled = True
End If
End Sub



Private Sub Form_Initialize()
    InitCommonControls
End Sub

Public Sub Form_Load()
On Error GoTo errorhandler1
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\mingdan.txt")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    total = 0
    Do While Not ts.atendofstream
        total = total + 1
        s = ts.readline
        qiandaoList.AddItem s
        NameCombo.AddItem s
        If total >= 100 Then MsgBox ("人数过多，后面就不能添加了。"): Exit Do
    Loop
    For i = 1 To qiandaoList.ListCount
        names(i) = qiandaoList.List(i)
    Next i
    ts.Close
    VersionLabel.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
    key(1) = 33965
    key(2) = 32717
    key(3) = 41831
    passwordtime = 25
    If fs.fileexists(App.Path + "\settingsaved.classsign") Then
    Set f = fs.getfile(App.Path + "\settingsaved.classsign")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    savepath = Decode(ts.readline, 32717, 41831)
    saveonexit = Val(Decode(ts.readline, 32717, 41831))
    settime = TimeValue(Decode(ts.readline, 32717, 41831))
    tqtime = Val(Decode(ts.readline, 32717, 41831))
    Else
    savepath = App.Path
    saveonexit = 1
    settime = TimeValue("23:59:59")
    tqtime = Val(Decode(ts.readline, 32717, 41831))
    End If
    If fs.fileexists(App.Path + "\autosave.classsign") Then
        sure = MsgBox("发现未完成的签到，是否恢复？", vbYesNo)
        If sure = vbYes Then
            Call recoveqiandao
        Else
            Set f = fs.getfile(App.Path + "\autosave.classsign")
            f.Delete
        End If
    End If
errorhandler1:
     If Err.Number <> 0 Then
     savepath = App.Path
    saveonexit = 1
        settime = TimeValue("23:59:59")
        tqtime = 10
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lineTimer.Enabled And AdminCommand.Visible And qiandaoList.ListCount + zrslist.ListCount > 0 Then
        Cancel = 1
        MsgBox ("无法关闭，原因是：还有人没签到，或签到时间没截止，或没以管理员身份登录。")
    Else
        If outputCommand.Enabled And saveonexit = 1 Then Call outputcommand_Click
        Cancel = 0
    End If
    If Cancel = 0 Then End
End Sub

Private Sub qiandaoList_Click()
Dim x As Integer
x = qiandaoList.ListIndex
If qiandaoList.ListIndex >= 0 Then
    NameCombo.Text = CStr(qiandaoList.List(qiandaoList.ListIndex))
    qiandaoCommand.Enabled = True
    End If
    If zrslist.ListIndex >= 0 Then
    zrslist.Selected(zrslist.ListIndex) = False
    zrsCommand.Enabled = False
    If x >= 0 Then qiandaoList.Selected(x) = True
    End If
End Sub

Private Sub zrshx_Timer()
If TimeSerial(Hour(settime), Minute(settime) - tqtime, Second(settime)) < DateTime.Time Then
    records2List.AddItem "--------------"
    records2List.AddItem "以下迟到："
    zrshx.Enabled = False
End If
End Sub

Private Sub zrslist_Click()
Dim x As Integer
    x = zrslist.ListIndex
    If zrslist.ListIndex >= 0 Then zrsCommand.Enabled = True
If qiandaoList.ListIndex >= 0 Then
    qiandaoCommand.Enabled = False
    NameCombo.Text = ""
    qiandaoList.Selected(qiandaoList.ListIndex) = False
    If x >= 0 Then zrslist.Selected(x) = True
    End If
End Sub

Private Sub PasswordText_gotfocus()
    If PasswordText.Text = "输入管理员密码操作" Then PasswordText.PasswordChar = "*":     PasswordText.Text = ""
End Sub

Private Sub PasswordText_Lostfocus()
    If PasswordText.Text = "" Then PasswordText.PasswordChar = ""
    If PasswordText.Text = "" And ((passwordtime > 0 And passwordtime Mod 5 <> 0) Or passwordtime = 25) Then PasswordText.Text = "输入管理员密码操作"
End Sub
Private Sub lineTimer_Timer()
    If settime < DateTime.Time Then
        records1List.AddItem "------------"
        records1List.AddItem "以下迟到："
        lineTimer.Enabled = False
        If passwordtime = 0 Then PasswordText.Text = ""
        Call PasswordText_Lostfocus
        PasswordText.Enabled = True
        AdminCommand.Enabled = True
        passwordtime = 25
    End If
End Sub

Private Sub passwordTimer_Timer()
Static timeout As Integer
timeout = timeout + 1
PasswordText.Enabled = False
AdminCommand.Enabled = False
PasswordText.Text = CStr((5 - passwordtime \ 5) * 60 - timeout) + "秒后解锁。"
If timeout = (5 - passwordtime \ 5) * 60 Then
    AdminCommand.Enabled = True
    timeout = 0
    PasswordText.Enabled = True
    PasswordText.Text = "输入管理员密码操作"
    passwordTimer.Enabled = False
End If
End Sub

Private Sub zttimer_Timer()
    Static autosave As Integer
    Dim ctd As Long
    Dim outputtext As String
    ctd = DateTime.DateDiff("s", DateTime.Time, settime)
    autosave = autosave + 1
    Debug.Print autosave
    If autosave = 60 Then
        autosave = 0
        If outputCommand.Enabled Then
            zttext.Caption = "正在保存自动恢复文件......"
            Call autosavefile
            zttext.Caption = "已保存自动恢复文件。"
            Exit Sub
        End If
    End If
    If Me.Enabled And qiandaoList.ListCount + zrslist.ListCount > 0 Then
        outputtext = "签到正在进行，"
    ElseIf Not Me.Enabled Then
        outputtext = "签到已暂停，管理员进行某些操作后继续"
        GoTo opt
    Else
        outputtext = "当前已完成签到，请单击“输出记录”按钮以输出记录"
        GoTo opt
    End If
    ctd = DateDiff("s", DateTime.Time, settime)
    If ctd > 999 Then
        outputtext = outputtext & CStr(ctd \ 60) & "分" & CStr(ctd Mod 60) & "秒后签到将成为迟到。"
    ElseIf ctd >= 0 Then
        outputtext = outputtext & CStr(ctd) & "秒后签到将成为迟到。"
    Else
        outputtext = outputtext & "现在签到将成为迟到。"
    End If
opt:
    zttext.Caption = outputtext
End Sub
Private Sub autosavefile()
    On Error GoTo errortest
    Dim fs, f, ts, s, ss
    Dim savetime As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path + "\autosave.classsign")
    f.writeline Encode(CStr(settime), key(1), key(3))
    f.writeline Encode(CStr(tqtime), key(1), key(3))
    For i = 1 To totals
        f.writeline Encode(names(i), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To records1List.ListCount
        f.writeline Encode(records1List.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To records2List.ListCount
        f.writeline Encode(records2List.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To qiandaoList.ListCount
        f.writeline Encode(qiandaoList.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To zrslist.ListCount
    f.writeline Encode(zrslist.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To AdminForm.qingjialist.ListCount
    f.writeline Encode(AdminForm.qingjialist.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
errortest:
    If Err.Number <> 0 Then MsgBox ("保存失败。请检查输出路径是否错误。错误原因：" & Chr(10) & Err.Number & Err.Description)

End Sub
Private Sub recoveqiandao()
    Dim fs, f, ts
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\autosave.classsign")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
     settime = TimeValue(Decode(ts.readline, key(2), key(3)))
     tqtime = Val(Decode(ts.readline, key(2), key(3)))
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        names(i) = x
        totals = totals + 1
    Loop
    NameCombo.Clear
    records1List.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        records1List.AddItem x
        NameCombo.AddItem x
    Loop
    records2List.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        records2List.AddItem x
    Loop
    qiandaoList.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        qiandaoList.AddItem x
    Loop
    zrslist.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        zrslist.AddItem x
    Loop
    AdminForm.qingjialist.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        AdminForm.qingjialist.AddItem x
    Loop
    outputCommand.Enabled = True
    ts.Close
    f.Delete
End Sub
