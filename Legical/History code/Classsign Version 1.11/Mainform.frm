VERSION 5.00
Begin VB.Form Mainform 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�༶ǩ����"
   ClientHeight    =   10125
   ClientLeft      =   4200
   ClientTop       =   885
   ClientWidth     =   6570
   Icon            =   "Mainform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   10125
   ScaleWidth      =   6570
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer undolimit 
      Interval        =   1000
      Left            =   2280
      Top             =   7320
   End
   Begin VB.CommandButton UndoCommand 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   8640
      Width           =   1935
   End
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
      Caption         =   "ֵ����ǩ��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton AdminCommand 
      Caption         =   "����Ա����"
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
      Text            =   "�������Ա�������"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton helpCommand 
      Caption         =   "����"
      Height          =   300
      Left            =   3360
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton outputCommand 
      Caption         =   "�����¼"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton qiandaoCommand 
      Caption         =   "ǩ��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ComboBox NameCombo 
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.ListBox recordsList 
      ForeColor       =   &H00000000&
      Height          =   780
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   8760
      Width           =   2175
   End
   Begin VB.ListBox recordsList 
      ForeColor       =   &H00000000&
      Height          =   8160
      Index           =   0
      Left            =   4320
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox zrsList 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��ֵ���������ڴ˷������롣"
      Height          =   540
      Left            =   2160
      TabIndex        =   19
      Top             =   840
      Width           =   2160
   End
   Begin VB.Image TMImage 
      Height          =   735
      Left            =   2280
      Picture         =   "Mainform.frx":12FA
      Stretch         =   -1  'True
      ToolTipText     =   "Copyright(C)2019��2021 XuPeng Studio,All rights reserved."
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label zttext 
      BackStyle       =   0  'Transparent
      Height          =   780
      Left            =   2280
      TabIndex        =   17
      Top             =   4200
      Width           =   2010
   End
   Begin VB.Label CopyrightLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) XuPeng Studio 2019��2021. ��лʩ��չͬ־����Ľ���������� "
      Height          =   180
      Left            =   0
      TabIndex        =   16
      Top             =   9840
      Width           =   6480
   End
   Begin VB.Label VersionLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.10"
      Height          =   180
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label records2Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ֵ����ǩ����¼"
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   8520
      Width           =   1260
   End
   Begin VB.Label records1Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ǩ����¼"
      Height          =   180
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label qdLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "δǩ������ѡ��󵥻���ǩ����"
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2520
   End
   Begin VB.Label zrsqdLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ֵ����ǩ��"
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
Public tqtime As Integer, undor As Integer
Dim timelimit As Integer
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
            qiandaoCommand.Enabled = True
            qiandaoList.Selected(i - 1) = True
            Exit Sub
        End If
    Next i
    qiandaoCommand.Enabled = False
End Sub

Private Sub NameCombo_Click()
    qiandaoCommand.Enabled = True
    qiandaoList.Selected(NameCombo.ListIndex) = True
    Exit Sub
End Sub

Private Sub NameCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call qiandaocommand_Click
End Sub

Private Sub PasswordText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call adminCommand_Click
End Sub

Private Sub qiandaocommand_Click()
    For i = 1 To qiandaoList.ListCount
        If qiandaoList.List(i - 1) = NameCombo.Text Then Exit For
    Next i
    If i > qiandaoList.ListCount Then Exit Sub
    recordsList(0).AddItem qiandaoList.List(i - 1) + "  " + CStr(DateTime.Time)
    qiandaoList.RemoveItem (i - 1)
    NameCombo.RemoveItem (i - 1)
    NameCombo.Text = ""
    qiandaoCommand.Enabled = False
    outputCommand.Enabled = True
    UndoCommand.Enabled = True
    undolimit.Enabled = True
    UndoCommand.Caption = "����(10s)"
    undor = 1
    timelimit = 10
End Sub
Private Sub outputcommand_Click()
    On Error GoTo errortest
    Dim fs, f, ts, s, ss
    Dim savetime As String
    savetime = CStr(DateTime.Year(Date)) + "-" + CStr(DateTime.Month(Date)) + "-" + CStr(DateTime.Day(Date)) + "-" + CStr(DateTime.Hour(Time)) + "-" + CStr(DateTime.Minute(Time)) + "-" + CStr(DateTime.Second(Time))
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(savepath + "\ǩ����¼" + savetime + ".txt")
    f.writeline ("��¼ʱ�䣺" + CStr(Date) + "  " + CStr(Time))
    f.writeline ("ǩ����¼��")
    For i = 1 To recordsList(0).ListCount
    f.writeline (recordsList(0).List(i - 1))
    Next i
    f.writeline ("--END OF RECORD--")
    If recordsList(1).ListCount > 0 Then
    f.writeline ("ֵ������")
    For i = 1 To recordsList(1).ListCount
    f.writeline (recordsList(1).List(i - 1))
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If qiandaoList.ListCount > 0 Then
    f.writeline ("δǩ������")
    For i = 1 To qiandaoList.ListCount
    f.writeline (qiandaoList.List(i - 1) + " δǩ")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If zrsList.ListCount > 0 Then
    f.writeline ("ֵ����δǩ������")
    For i = 1 To zrsList.ListCount
    f.writeline (zrsList.List(i - 1) + " δǩ")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If AdminForm.qingjialist.ListCount > 0 Then
    f.writeline ("���������")
    For i = 1 To AdminForm.qingjialist.ListCount
    f.writeline (AdminForm.qingjialist.List(i - 1) + " ���")
    Next i
    f.writeline ("--END OF RECORD--")
    End If
    If fs.fileexists(App.Path + "\autosave.classsign") Then
    Set f = fs.getfile(App.Path + "\autosave.classsign")
    f.Delete
    End If
    outputCommand.Enabled = False
errortest:
    If Err.Number <> 0 Then MsgBox ("����ʧ�ܡ��������·���Ƿ���󡣴���ԭ��" & Chr(10) & Err.Number & Err.Description)
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
            MsgBox ("������󣬻���������" + CStr(passwordtime Mod 5) + "��")
        ElseIf passwordtime > 0 Then
            MsgBox ("�����������" + CStr(25 - passwordtime) + "���Ӻ�����")
            passwordTimer.Enabled = True
            PasswordText.Enabled = False
            AdminCommand.Enabled = False
            PasswordText.Text = ""
            Call PasswordText_Lostfocus
        Else
            MsgBox ("�����Ͻ��ú�ѧϰ��")
            AdminCommand.Enabled = False
            PasswordText.Enabled = False
            Call PasswordText_Lostfocus
            PasswordText.Text = "�����޷��������롣"
        End If
    End If
    PasswordText.PasswordChar = ""
    If Not passwordTimer.Enabled And passwordtime > 0 Then
        PasswordText.Text = "�������Ա�������"
        If PasswordText.Visible Then PasswordText.SetFocus: Call PasswordText_gotfocus
    End If
End Sub


Private Sub recordsList_Click(Index As Integer)
    If recordsList(Index).ListIndex >= 0 Then recordsList(Index).Selected(recordsList(Index).ListIndex) = False
End Sub

Private Sub UndoCommand_Click()
    Call undoqd(undor - 1, recordsList(undor - 1).ListCount - 1)
    timelimit = 0
    UndoCommand.Enabled = False
    Call undolimit_Timer
End Sub

Private Sub undolimit_Timer()
    timelimit = timelimit - 1
    UndoCommand.Caption = "����(" + CStr(timelimit) + "s)"
    If timelimit <= 0 Then
        timelimit = 10
        UndoCommand.Caption = "����"
        undolimit.Enabled = False
        UndoCommand.Enabled = False
    End If
End Sub

Private Sub zrsCommand_Click()
If zrsList.ListIndex >= 0 Then

    recordsList(1).AddItem zrsList.List(zrsList.ListIndex) + "  " + CStr(DateTime.Time)
    zrsList.RemoveItem (zrsList.ListIndex)
    zrsCommand.Enabled = False
    outputCommand.Enabled = True
    UndoCommand.Enabled = True
    undolimit.Enabled = True
    undor = 2
    timelimit = 10
    UndoCommand.Caption = "����(10s)"
End If
End Sub

Public Sub undoqd(listnum As Integer, listind As Integer)
    On Error GoTo undoerr
    For i = 1 To total
        If names(i) = Mid(recordsList(listnum).List(listind), 1, Len(names(i))) Then
            recordsList(listnum).RemoveItem (listind)
            Select Case listnum
                Case 0
                    qiandaoList.AddItem (names(i)): NameCombo.AddItem (names(i))
                Case 1
                    zrsList.AddItem (names(i))
            End Select
            Exit For
        End If
    Next i
undoerr:
    If Err.Number <> 0 Then MsgBox ("����δ֪ԭ�򣬳�������ˣ���ϵ�����д�߸��Ĵ˴��롣")
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
        If total >= 100 Then MsgBox ("�������࣬����Ͳ�������ˡ�"): Exit Do
    Loop
    For i = 1 To qiandaoList.ListCount
        names(i) = qiandaoList.List(i - 1)
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
        sure = MsgBox("����δ��ɵ�ǩ�����Ƿ�ָ���", vbYesNo)
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
    If lineTimer.Enabled And AdminCommand.Visible And qiandaoList.ListCount + zrsList.ListCount > 0 Then
        Cancel = 1
        MsgBox ("�޷��رգ�ԭ���ǣ�������ûǩ������ǩ��ʱ��û��ֹ����û�Թ���Ա��ݵ�¼��")
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
    If zrsList.ListIndex >= 0 Then
    zrsList.Selected(zrsList.ListIndex) = False
    zrsCommand.Enabled = False
    If x >= 0 Then qiandaoList.Selected(x) = True
    End If
End Sub

Private Sub zrshx_Timer()
If TimeSerial(Hour(settime), Minute(settime) - tqtime, Second(settime)) < DateTime.Time Then
    recordsList(1).AddItem "--------------"
    recordsList(1).AddItem "���³ٵ���"
    zrshx.Enabled = False
End If
End Sub

Private Sub zrslist_Click()
Dim x As Integer
    x = zrsList.ListIndex
    If zrsList.ListIndex >= 0 Then zrsCommand.Enabled = True
If qiandaoList.ListIndex >= 0 Then
    qiandaoCommand.Enabled = False
    NameCombo.Text = ""
    qiandaoList.Selected(qiandaoList.ListIndex) = False
    If x >= 0 Then zrsList.Selected(x) = True
    End If
End Sub

Private Sub PasswordText_gotfocus()
    If PasswordText.Text = "�������Ա�������" Then PasswordText.PasswordChar = "*":     PasswordText.Text = ""
End Sub

Private Sub PasswordText_Lostfocus()
    If PasswordText.Text = "" Then PasswordText.PasswordChar = ""
    If PasswordText.Text = "" And ((passwordtime > 0 And passwordtime Mod 5 <> 0) Or passwordtime = 25) Then PasswordText.Text = "�������Ա�������"
End Sub
Private Sub lineTimer_Timer()
    If settime < DateTime.Time Then
        recordsList(0).AddItem "------------"
        recordsList(0).AddItem "���³ٵ���"
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
PasswordText.Text = CStr((25 - passwordtime) * 60 - timeout) + "��������"
If timeout = (25 - passwordtime) * 60 Then
    AdminCommand.Enabled = True
    timeout = 0
    PasswordText.Enabled = True
    PasswordText.Text = "�������Ա�������"
    passwordTimer.Enabled = False
End If
End Sub

Private Sub zttimer_Timer()
    Static autosave As Integer
    Dim ctd As Long
    Dim outputtext As String
    ctd = DateTime.DateDiff("s", DateTime.Time, settime)
    autosave = autosave + 1
    If autosave = 60 Then
        autosave = 0
        If outputCommand.Enabled Then
            zttext.Caption = "���ڱ����Զ��ָ��ļ�......"
            Call autosavefile
            zttext.Caption = "�ѱ����Զ��ָ��ļ���"
            Exit Sub
        End If
    End If
    If Me.Enabled And qiandaoList.ListCount + zrsList.ListCount > 0 Then
        outputtext = "ǩ�����ڽ��У�"
    ElseIf Not Me.Enabled Then
        outputtext = "ǩ������ͣ������Ա����ĳЩ���������"
        GoTo opt
    Else
        outputtext = "��ǰ�����ǩ�����뵥���������¼����ť�������¼"
        GoTo opt
    End If
    ctd = DateDiff("s", DateTime.Time, settime)
    If ctd > 999 Then
        outputtext = outputtext & CStr(ctd \ 60) & "��" & CStr(ctd Mod 60) & "���ǩ������Ϊ�ٵ���"
    ElseIf ctd >= 0 Then
        outputtext = outputtext & CStr(ctd) & "���ǩ������Ϊ�ٵ���"
    Else
        outputtext = outputtext & "����ǩ������Ϊ�ٵ���"
        GoTo opt
    End If
    ctd = DateTime.DateDiff("s", DateTime.Time, settime) - tqtime * 60
    If ctd > 999 Then
        outputtext = outputtext & "ֵ������" & CStr(ctd \ 60) & "��" & CStr(ctd Mod 60) & "���ǩ������Ϊ�ٵ���"
    ElseIf ctd >= 0 Then
        outputtext = outputtext & "ֵ������" & CStr(ctd) & "���ǩ������Ϊ�ٵ���"
    Else
        outputtext = outputtext & "ֵ��������ǩ������Ϊ�ٵ���"
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
    For i = 1 To recordsList(0).ListCount
        f.writeline Encode(recordsList(0).List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To recordsList(1).ListCount
        f.writeline Encode(recordsList(1).List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To qiandaoList.ListCount
        f.writeline Encode(qiandaoList.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To zrsList.ListCount
    f.writeline Encode(zrsList.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
    For i = 1 To AdminForm.qingjialist.ListCount
    f.writeline Encode(AdminForm.qingjialist.List(i - 1), key(1), key(3))
    Next i
    f.writeline Encode("--END OF RECORD--", key(1), key(3))
errortest:
    If Err.Number <> 0 Then MsgBox ("����ʧ�ܡ��������·���Ƿ���󡣴���ԭ��" & Chr(10) & Err.Number & Err.Description)

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
    recordsList(0).Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        recordsList(0).AddItem x
    Loop
    recordsList(1).Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        recordsList(1).AddItem x
    Loop
    qiandaoList.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        qiandaoList.AddItem x
        NameCombo.AddItem x
    Loop
    zrsList.Clear
    Do While True
        x = Decode(ts.readline, key(2), key(3))
        If x = "--END OF RECORD--" Then Exit Do
        zrsList.AddItem x
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
