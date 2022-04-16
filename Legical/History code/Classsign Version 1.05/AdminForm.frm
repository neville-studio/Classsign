VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AdminForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "班级签到系统 - 你好，管理员！"
   ClientHeight    =   3795
   ClientLeft      =   6435
   ClientTop       =   2505
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox zrstqtext 
      Height          =   270
      Left            =   2040
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton reset 
      Caption         =   "重新签到"
      Height          =   375
      Left            =   2040
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton endqd 
      Caption         =   "结束签到"
      Height          =   375
      Left            =   2040
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton mmggcommand 
      Caption         =   "密码修改"
      Height          =   375
      Left            =   2040
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox newmmagaintext 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   39
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox newmmtext 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox ylmmtext 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   35
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox nowtimetext 
      Height          =   270
      Left            =   2040
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox timerequire 
      Height          =   270
      Left            =   2040
      TabIndex        =   30
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton outputqr 
      Caption         =   "应用设置"
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox saveonexit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "退出后保存"
      Height          =   255
      Left            =   2040
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton getposition 
      Caption         =   "..."
      Height          =   255
      Left            =   4560
      TabIndex        =   26
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox outputpositiontext 
      Height          =   270
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton updatewq 
      Caption         =   "应用未签名单"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox zrslist 
      Enabled         =   0   'False
      Height          =   2940
      Left            =   5040
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox qingjialist 
      Enabled         =   0   'False
      Height          =   2580
      Left            =   5040
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton unzrsclick 
      Caption         =   "<<非值日生"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton zrsclick 
      Caption         =   "值日生>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton unqingjiaclick 
      Caption         =   "<<不请假"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton qingjiaClick 
      Caption         =   "请假>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton updateclick 
      Caption         =   "更新签到名单"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Filerequire 
      Left            =   3600
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox weiqianList 
      Enabled         =   0   'False
      Height          =   6180
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton recovemingdan 
      Caption         =   "恢复"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Clearmingdan 
      Caption         =   "删除名单"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton delmingdan 
      Caption         =   "删除名字"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton newname 
      Caption         =   "新名字"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton outputmingdan 
      Caption         =   "导出名单"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Importmingdan 
      Caption         =   "导入名单"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Savemingdan 
      Caption         =   "保存名单"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton mingdanCommand 
      Caption         =   "更改"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox MingdanText 
      Height          =   270
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox MingdanList 
      Height          =   3480
      Left            =   2040
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox MenuList 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      ItemData        =   "AdminForm.frx":0000
      Left            =   120
      List            =   "AdminForm.frx":0016
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label zrstqtip 
      BackStyle       =   0  'Transparent
      Caption         =   "值日生提前以下时间（分钟）签到："
      Height          =   420
      Left            =   2040
      TabIndex        =   43
      Top             =   1440
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label newmmagaintip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "再输一遍新密码"
      Height          =   180
      Left            =   2040
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label newmm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入新密码"
      Height          =   180
      Left            =   2040
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label ylmmtip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入原密码"
      Height          =   180
      Left            =   2040
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label mmxgtip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码修改"
      Height          =   180
      Left            =   2040
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Timenowtip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设定签到截止时间："
      Height          =   180
      Left            =   2040
      TabIndex        =   31
      Top             =   750
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label qdydtip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "签到时间预订"
      Height          =   180
      Left            =   2040
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label recordtip1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输出位置"
      Height          =   180
      Left            =   2040
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label tiptsqk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "值日生"
      Height          =   180
      Left            =   5040
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label tiptsqk2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请假名单"
      Height          =   180
      Left            =   5160
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label mingdanname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   180
      Left            =   3600
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label TipLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "你好，管理员！请选择一项进行操作。"
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3060
   End
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Dim passwordtime As Integer



Private Sub Clearmingdan_Click()
    delrequire = MsgBox("是否清空名单？", vbYesNo, "班级签到系统")
    If delrequire = vbYes Then MingdanList.Clear
End Sub

Private Sub Command1_Click()

End Sub

Private Sub delmingdan_Click()
    MingdanList.RemoveItem (MingdanList.ListIndex)
    newname.Enabled = MingdanList.ListCount < 100
    delmingdan.Enabled = False
End Sub

Private Sub endqd_Click()
    Unload Mainform
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
        MingdanList.AddItem ts.readline
    Loop
    For i = 1 To Mainform.qiandaoList.ListCount
        weiqianList.AddItem Mainform.qiandaoList.List(i - 1)
    Next i
    For i = 1 To Mainform.zrslist.ListCount
        zrslist.AddItem Mainform.zrslist.List(i - 1)
    Next i
    ts.Close
    If MingdanList.ListCount >= 100 Then newname.Enabled = False
    delmingdan.Enabled = False
    If fs.fileexists(App.Path + "\settingsaved.classsign") Then
    Set f = fs.getfile(App.Path + "\settingsaved.classsign")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    outputpositiontext.Text = Mainform.Decode(ts.readline, 32717, 41831)
    saveonexit.Value = Val(Mainform.Decode(ts.readline, 32717, 41831))
    timerequire.Text = Mainform.Decode(ts.readline, 32717, 41831)
    nowtimetext.Text = CStr(Mainform.settime)
    zrstqtext.Text = CStr(Mainform.tqtime)
    Else
        outputpositiontext.Text = App.Path
    saveonexit.Value = 1
    timerequire.Text = "23:59:59"
    nowtimetext.Text = CStr(Mainform.settime)
    End If
    passwordtime = 5
    ts.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Mainform.AdminCommand.Visible = True
    Mainform.PasswordText.Visible = True
    Me.Hide
    If updatewq.Enabled Then
        Call updatewq_Click
    End If
    Cancel = 1
End Sub

Private Sub getposition_Click()
    x = ShowFolderDialog
    If x <> "" Then
    outputpositiontext = x
    End If
End Sub

Private Sub Importmingdan_Click()
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    On Error GoTo errHandler
    Filerequire.CancelError = True
    Filerequire.Flags = cdIOFNHideReadOnly
    Filerequire.Filter = "所有文件(*.*)|*.*|TXT文本文档(*.txt)|*.txt"
    Filerequire.FilterIndex = 2
    Filerequire.ShowOpen
    MingdanList.Clear
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(Filerequire.FileName)
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    Do While Not ts.atendofstream
        MingdanList.AddItem ts.readline
    Loop
    If MingdanList.ListCount >= 100 Then newname.Enabled = False
    delmingdan.Enabled = False
errHandler:
    
End Sub

Private Sub MenuList_Click()
    TipLabel.Visible = False
    Select Case MenuList.ListIndex
        Case 0
            Me.Width = 5850
            Me.Height = 4155
            MingdanList.Visible = True
            MingdanText.Visible = True
            mingdanCommand.Visible = True
            Savemingdan.Visible = True
            Importmingdan.Visible = True
            outputmingdan.Visible = True
            newname.Visible = True
            delmingdan.Visible = True
            Clearmingdan.Visible = True
            recovemingdan.Visible = True
            weiqianList.Visible = False
            mingdanname.Visible = True
            updateclick.Visible = False
            qingjiaClick.Visible = False
            unqingjiaclick.Visible = False
            zrsclick.Visible = False
            unzrsclick.Visible = False
            qingjialist.Visible = False
            zrslist.Visible = False
            updatewq.Visible = False
            tiptsqk.Visible = False
            tiptsqk2.Visible = False
            recordtip1.Visible = False
            getposition.Visible = False
            saveonexit.Visible = False
            outputqr.Visible = False
            Timenowtip.Visible = False
            nowtimetext.Visible = False
            outputpositiontext.Visible = False
            timerequire.Visible = False
                       mmxgtip.Visible = False
           ylmmtip.Visible = False
           ylmmtext.Visible = False
           newmm.Visible = False
           newmmtext.Visible = False
           newmmagaintip.Visible = False
           newmmagaintext.Visible = False
           mmggcommand.Visible = False
           endqd.Visible = False
           reset.Visible = False
                      zrstqtext.Visible = False
           zrstqtip.Visible = False
        Case 1
            Me.Width = 7110
            Me.Height = 6915
            weiqianList.Visible = True
            mingdanname.Visible = False
            MingdanList.Visible = False
            MingdanText.Visible = False
            mingdanCommand.Visible = False
            Savemingdan.Visible = False
            Importmingdan.Visible = False
            outputmingdan.Visible = False
            newname.Visible = False
            delmingdan.Visible = False
            Clearmingdan.Visible = False
            recovemingdan.Visible = False
            updateclick.Visible = True
            qingjiaClick.Visible = True
            unqingjiaclick.Visible = True
            zrsclick.Visible = True
            unzrsclick.Visible = True
            qingjialist.Visible = True
            zrslist.Visible = True
            updatewq.Visible = True
            tiptsqk.Visible = True
            tiptsqk2.Visible = True
            recordtip1.Visible = False
            getposition.Visible = False
            saveonexit.Visible = False
            outputqr.Visible = False
            outputpositiontext.Visible = False
                        timerequire.Visible = False
            qdydtip.Visible = False
            Timenowtip.Visible = False
            nowtimetext.Visible = False
                       mmxgtip.Visible = False
           ylmmtip.Visible = False
           ylmmtext.Visible = False
           newmm.Visible = False
           newmmtext.Visible = False
           newmmagaintip.Visible = False
           newmmagaintext.Visible = False
           mmggcommand.Visible = False
           endqd.Visible = False
           reset.Visible = False
           zrstqtext.Visible = False
           zrstqtip.Visible = False
        Case 2
            Me.Width = 5205
            Me.Height = 4185
                 weiqianList.Visible = False
            mingdanname.Visible = False
            MingdanList.Visible = False
            MingdanText.Visible = False
            mingdanCommand.Visible = False
            Savemingdan.Visible = False
            Importmingdan.Visible = False
            outputmingdan.Visible = False
            newname.Visible = False
            delmingdan.Visible = False
            Clearmingdan.Visible = False
            recovemingdan.Visible = False
            updateclick.Visible = False
            qingjiaClick.Visible = False
            unqingjiaclick.Visible = False
            zrsclick.Visible = False
            unzrsclick.Visible = False
            qingjialist.Visible = False
            zrslist.Visible = False
            updatewq.Visible = False
            tiptsqk.Visible = False
            tiptsqk2.Visible = False
            recordtip1.Visible = True
            getposition.Visible = True
            saveonexit.Visible = True
            outputqr.Visible = True
            outputpositiontext.Visible = True
                        timerequire.Visible = False
            qdydtip.Visible = False
                        Timenowtip.Visible = False
            nowtimetext.Visible = False
                       mmxgtip.Visible = False
           ylmmtip.Visible = False
           ylmmtext.Visible = False
           newmm.Visible = False
           newmmtext.Visible = False
           newmmagaintip.Visible = False
           newmmagaintext.Visible = False
           mmggcommand.Visible = False
           endqd.Visible = False
           zrstqtip.Visible = False
                      zrstqtext.Visible = False
           reset.Visible = False
        Case 3
            Me.Width = 3915
            Me.Height = 4185
            weiqianList.Visible = False
            mingdanname.Visible = False
            MingdanList.Visible = False
            MingdanText.Visible = False
            mingdanCommand.Visible = False
            Savemingdan.Visible = False
            Importmingdan.Visible = False
            outputmingdan.Visible = False
            newname.Visible = False
            delmingdan.Visible = False
            Clearmingdan.Visible = False
            recovemingdan.Visible = False
            updateclick.Visible = False
            qingjiaClick.Visible = False
            unqingjiaclick.Visible = False
            zrsclick.Visible = False
            unzrsclick.Visible = False
            qingjialist.Visible = False
            zrslist.Visible = False
            updatewq.Visible = False
            tiptsqk.Visible = False
            tiptsqk2.Visible = False
            recordtip1.Visible = False
            getposition.Visible = False
            saveonexit.Visible = False
            outputpositiontext.Visible = False
            timerequire.Visible = True
            qdydtip.Visible = True
            outputqr.Visible = True
            Timenowtip.Visible = True
            nowtimetext.Visible = True
           mmxgtip.Visible = False
           ylmmtip.Visible = False
           ylmmtext.Visible = False
           newmm.Visible = False
           newmmtext.Visible = False
           newmmagaintip.Visible = False
           newmmagaintext.Visible = False
           mmggcommand.Visible = False
           endqd.Visible = False
           
           zrstqtip.Visible = True
           reset.Visible = False
                      zrstqtext.Visible = True
        Case 4
                    Me.Width = 3795
            Me.Height = 4350
        weiqianList.Visible = False
            mingdanname.Visible = False
            MingdanList.Visible = False
            MingdanText.Visible = False
            mingdanCommand.Visible = False
            Savemingdan.Visible = False
            Importmingdan.Visible = False
            outputmingdan.Visible = False
            newname.Visible = False
            delmingdan.Visible = False
            Clearmingdan.Visible = False
            recovemingdan.Visible = False
            updateclick.Visible = False
            qingjiaClick.Visible = False
            unqingjiaclick.Visible = False
            zrsclick.Visible = False
            unzrsclick.Visible = False
            qingjialist.Visible = False
            zrslist.Visible = False
            updatewq.Visible = False
            tiptsqk.Visible = False
            tiptsqk2.Visible = False
            recordtip1.Visible = False
            getposition.Visible = False
            saveonexit.Visible = False
            outputpositiontext.Visible = False
            timerequire.Visible = False
            qdydtip.Visible = False
            outputqr.Visible = False
            Timenowtip.Visible = False
            nowtimetext.Visible = False
           mmxgtip.Visible = True
           ylmmtip.Visible = True
           ylmmtext.Visible = True
           newmm.Visible = True
           newmmtext.Visible = True
           newmmagaintip.Visible = True
           newmmagaintext.Visible = True
           mmggcommand.Visible = True
           endqd.Visible = True
           reset.Visible = True
           zrstqtip.Visible = False
           zrstqtext.Visible = False
        Case 5
            MsgBox (App.ProductName & Chr(10) & "Version " & App.Major & "." & App.Minor & App.Revision & Chr(10) & App.LegalCopyright & "感谢施振展同志提出的建设性意见")
            MenuList.Selected(5) = False
    End Select
End Sub

Private Sub mingdanCommand_Click()
    MingdanList.List(MingdanList.ListIndex) = MingdanText.Text
End Sub

Private Sub MingdanList_Click()
    MingdanText.Text = MingdanList.List(MingdanList.ListIndex)
    delmingdan.Enabled = True
End Sub

Private Sub mmggcommand_Click()
    passwordtime = passwordtime - 1
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\pw")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    If Mainform.Decode(ts.readall, 32717, 41831) = ylmmtext.Text Then
        passwordtime = 5
        ts.Close
        If newmmtext.Text = newmmagaintext.Text Then
            Set f = fs.createtextfile(App.Path + "\pw")
            f.writeline (Mainform.Encode(newmmtext.Text, 33965, 41831))
            MsgBox ("修改成功")
        Else
            MsgBox ("两次密码输入不一致！")
        End If
    ElseIf passwordtime > 0 Then
        MsgBox ("密码错误，还可以输入" + CStr(passwordtime Mod 5) + "次")
    Else
        MsgBox ("密码错误，强制登出！")
        Call Form_Unload(0)
    End If
End Sub

Private Sub newname_Click()
    MingdanList.AddItem MingdanText.Text
    newname.Enabled = MingdanList.ListCount < 100
    delmingdan.Enabled = MingdanList.ListCount > 0
End Sub

Private Sub outputmingdan_Click()
    Filerequire.CancelError = True
    Filerequire.Flags = cdIOFNHideReadOnly
    Filerequire.Filter = "所有文件(*.*)|*.*|TXT文本文档(*.txt)|*.txt"
    Filerequire.FilterIndex = 2
    Filerequire.ShowSave
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(Filerequire.FileName, True)
    
    For i = 1 To MingdanList.ListCount
    f.writeline (MingdanList.List(i - 1))
    Next i
    f.Close
errHandler:
End Sub

Private Sub outputqr_Click()
    On Error GoTo errorHandler
    Dim fs, f, ts, s, ss
    Dim aaa As Date
    aaa = DateTime.TimeValue(timerequire.Text)
    bbb = DateTime.TimeValue(nowtimetext.Text)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path + "\settingsaved.classsign", True)
    f.writeline Mainform.Encode(outputpositiontext.Text, 33965, 41831)
    f.writeline Mainform.Encode(CStr(saveonexit.Value), 33965, 41831)
    f.writeline Mainform.Encode(CStr(timerequire.Text), 33965, 41831)
    f.writeline Mainform.Encode(CStr(zrstqtext.Text), 33965, 41831)
    Mainform.settime = TimeValue(nowtimetext.Text)
    Mainform.tqtime = Val(zrstqtext.Text)
    Mainform.lineTimer.Enabled = True
    Mainform.zrshx.Enabled = True
    For i = 1 To Mainform.records1List.ListCount
        If Mainform.records1List.List(i - 1) = "------------" Then
            Mainform.records1List.RemoveItem (i - 1)
            Mainform.records1List.RemoveItem (i - 1)
            Exit For
        End If
    Next i
    For i = 1 To Mainform.records2List.ListCount
        If Mainform.records2List.List(i - 1) = "--------------" Then
            Mainform.records2List.RemoveItem (i - 1)
            Mainform.records2List.RemoveItem (i - 1)
            Exit For
        End If
    Next i
    Mainform.savepath = outputpositiontext.Text
    Mainform.saveonexit = saveonexit.Value
errorHandler:
    If Err.Number <> 0 Then
    MsgBox "出错了，请检查输入的时间是否符合标准的时间格式。" + Chr(10) + CStr(Err.Number) + " " + Err.Description
    End If
End Sub

Private Sub qingjiaClick_Click()
    qingjialist.AddItem weiqianList.List(weiqianList.ListIndex)
    weiqianList.RemoveItem (weiqianList.ListIndex)
    qingjiaClick.Enabled = False
    zrsclick.Enabled = False
    Mainform.Enabled = False
    updateclick.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub qingjialist_Click()
    If qingjialist.ListIndex >= 0 Then
    unqingjiaclick.Enabled = True
    End If
End Sub

Private Sub recovemingdan_Click()
    MingdanList.Clear
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.getfile(App.Path + "\mingdan.txt")
    Set ts = f.openastextstream(ForReading, TristateUseDefault)
    Do While Not ts.atendofstream
        MingdanList.AddItem ts.readline
    Loop
    If MingdanList.ListCount >= 100 Then newname.Enabled = False
    delmingdan.Enabled = False
        
    ts.Close
End Sub

Private Sub reset_Click()
    Mainform.qiandaoList.Clear
    Mainform.records1List.Clear
    Mainform.zrslist.Clear
    Mainform.records2List.Clear
    Call Mainform.Form_Load
    Unload Me
End Sub

Private Sub Savemingdan_Click()
    Dim fs, f, ts, s, ss
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path + "\mingdan.txt")
    
    For i = 1 To MingdanList.ListCount
    f.writeline (MingdanList.List(i - 1))
    Next i
    MsgBox ("保存成功，签到名单将在下次应用")
End Sub

Private Sub unqingjiaclick_Click()
    weiqianList.AddItem qingjialist.List(qingjialist.ListIndex)
    qingjialist.RemoveItem (qingjialist.ListIndex)
    unqingjiaclick.Enabled = False
    Mainform.Enabled = False
    updateclick.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub unzrsclick_Click()
    Mainform.Enabled = False
    weiqianList.AddItem zrslist.List(zrslist.ListIndex)
    zrslist.RemoveItem zrslist.ListIndex
    unzrsclick.Enabled = False
    updateclick.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub updateclick_Click()
    weiqianList.Clear
    zrslist.Clear
    For i = 1 To Mainform.qiandaoList.ListCount
        weiqianList.AddItem Mainform.qiandaoList.List(i - 1)
    Next i
    For i = 1 To Mainform.zrslist.ListCount
        zrslist.AddItem Mainform.zrslist.List(i - 1)
    Next i
    updatewq.Enabled = True
    Mainform.Enabled = False
    zrslist.Enabled = True
    weiqianList.Enabled = True
    updateclick.Enabled = False
    qingjialist.Enabled = True
End Sub
Private Sub updatewq_Click()
    Mainform.qiandaoList.Clear
    Mainform.zrslist.Clear
    Mainform.NameCombo.Clear
    For i = 1 To weiqianList.ListCount
        Mainform.qiandaoList.AddItem weiqianList.List(i - 1)
        Mainform.NameCombo.AddItem weiqianList.List(i - 1)
    Next i
    For i = 1 To zrslist.ListCount
        Mainform.zrslist.AddItem zrslist.List(i - 1)
    Next i
    Mainform.Enabled = True
    updateclick.Enabled = True
    updatewq.Enabled = False
    zrslist.Enabled = False
    weiqianList.Enabled = False
    qingjialist.Enabled = False
    qingjiaClick.Enabled = False
    unqingjiaclick.Enabled = False
    unzrsclick.Enabled = False
    If weiqianList.ListIndex >= 0 Then weiqianList.Selected(weiqianList.ListIndex) = False
    If qingjialist.ListIndex >= 0 Then qingjialist.Selected(qingjialist.ListIndex) = False
    If zrslist.ListIndex >= 0 Then zrslist.Selected(zrslist.ListIndex) = False
End Sub

Private Sub weiqianList_Click()
    If weiqianList.ListIndex >= 0 Then
    qingjiaClick.Enabled = True
    zrsclick.Enabled = True
    End If
End Sub

Private Sub zrsclick_Click()
    zrslist.AddItem weiqianList.List(weiqianList.ListIndex)
    weiqianList.RemoveItem (weiqianList.ListIndex)
    zrsclick.Enabled = False
    qingjiaClick.Enabled = False
    updateclick.Enabled = False
    Mainform.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub zrslist_Click()
If zrslist.ListIndex >= 0 Then
    unzrsclick.Enabled = True
End If
End Sub

Public Function ShowFolderDialog() As String

'/最简单的显示文件夹选择对话框方法
Dim spShell, spFolder, spFolderItem, spPath As String
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
Set spShell = CreateObject("Shell.Application")
Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "选择目录:", NO_OPTIONS, "")
If spFolder Is Nothing Then
   ShowFolderDialog = ""
Else
    Set spFolderItem = spFolder.Self
    spPath = spFolderItem.Path
    spPath = Replace(spPath, "\", "\")
   ShowFolderDialog = spPath
End If
End Function

Private Sub zrstqtext_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then KeyAscii = 0
    If Len(zrstqtext.Text) >= 2 Then KeyAscii = 0
End Sub
