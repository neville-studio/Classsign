VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AdminForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�༶ǩ���� - ��ã�����Ա��"
   ClientHeight    =   3795
   ClientLeft      =   6435
   ClientTop       =   2505
   ClientWidth     =   5040
   Icon            =   "AdminForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox zrstqtext 
      Height          =   270
      Left            =   2040
      TabIndex        =   44
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton reset 
      Caption         =   "����ǩ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton endqd 
      Caption         =   "����ǩ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   41
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton mmggcommand 
      Caption         =   "�����޸�"
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
      Caption         =   "Ӧ������"
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox saveonexit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�˳��󱣴�"
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
      Caption         =   "Ӧ��δǩ����"
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
      Caption         =   "<<��ֵ����"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton zrsclick 
      Caption         =   "ֵ����>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton unqingjiaclick 
      Caption         =   "<<�����"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton qingjiaClick 
      Caption         =   "���>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton updateclick 
      Caption         =   "����ǩ������"
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
      Caption         =   "�ָ�"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Clearmingdan 
      Caption         =   "ɾ������"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton delmingdan 
      Caption         =   "ɾ������"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton newname 
      Caption         =   "������"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton outputmingdan 
      Caption         =   "��������"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Importmingdan 
      Caption         =   "��������"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Savemingdan 
      Caption         =   "��������"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton mingdanCommand 
      Caption         =   "����"
      Enabled         =   0   'False
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
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      ItemData        =   "AdminForm.frx":12FA
      Left            =   120
      List            =   "AdminForm.frx":1310
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image PrImage 
      Height          =   720
      Left            =   4080
      Picture         =   "AdminForm.frx":1357
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image TMImage 
      Height          =   735
      Left            =   2040
      Picture         =   "AdminForm.frx":2651
      Stretch         =   -1  'True
      ToolTipText     =   "Copyright(C)2019��2021 XuPeng Studio,All rights reserved."
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label CRlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   1215
      Left            =   2040
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label zrstqtip 
      BackStyle       =   0  'Transparent
      Caption         =   "ֵ������ǰ����ʱ�䣨���ӣ�ǩ����"
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
      Caption         =   "����һ��������"
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
      Caption         =   "����������"
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
      Caption         =   "����ԭ����"
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
      Caption         =   "�����޸�"
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
      Caption         =   "�趨ǩ����ֹʱ�䣺"
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
      Caption         =   "ǩ��ʱ��Ԥ��"
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
      Caption         =   "���λ��"
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
      Caption         =   "ֵ����"
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
      Caption         =   "�������"
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
      Caption         =   "����"
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
      Caption         =   "��ã�����Ա����ѡ��һ����в�����"
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
    delrequire = MsgBox("�Ƿ����������", vbYesNo, "�༶ǩ��ϵͳ")
    If delrequire = vbYes Then
        If MingdanList.ListIndex >= 0 Then
            newname.Enabled = MingdanList.ListCount < 100
            delmingdan.Enabled = False
            mingdanCommand.Enabled = False
            MingdanText.Text = ""
        End If
        MingdanList.Clear
    End If
End Sub
Private Sub delmingdan_Click()
    MingdanList.RemoveItem (MingdanList.ListIndex)
    newname.Enabled = MingdanList.ListCount < 100
    delmingdan.Enabled = False
    mingdanCommand.Enabled = False
    MingdanText.Text = ""
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
    For i = 1 To Mainform.zrsList.ListCount
        zrsList.AddItem Mainform.zrsList.List(i - 1)
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
    X = ShowFolderDialog
    If X <> "" Then
    outputpositiontext = X
    End If
End Sub

Private Sub Importmingdan_Click()
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    On Error GoTo errHandler
    Filerequire.CancelError = True
    Filerequire.Flags = cdIOFNHideReadOnly
    Filerequire.Filter = "�����ļ�(*.*)|*.*|TXT�ı��ĵ�(*.txt)|*.txt"
    Filerequire.FilterIndex = 2
    Filerequire.ShowOpen
    If MingdanList.ListIndex > -1 Then
        newname.Enabled = MingdanList.ListCount < 100
        delmingdan.Enabled = False
        mingdanCommand.Enabled = False
        MingdanText.Text = ""
    End If
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
     MingdanList.Visible = MenuList.ListIndex = 0
            MingdanText.Visible = MenuList.ListIndex = 0
            mingdanCommand.Visible = MenuList.ListIndex = 0
            Savemingdan.Visible = MenuList.ListIndex = 0
            Importmingdan.Visible = MenuList.ListIndex = 0
            outputmingdan.Visible = MenuList.ListIndex = 0
            newname.Visible = MenuList.ListIndex = 0
            delmingdan.Visible = MenuList.ListIndex = 0
            Clearmingdan.Visible = MenuList.ListIndex = 0
            recovemingdan.Visible = MenuList.ListIndex = 0
            mingdanname.Visible = MenuList.ListIndex = 0
            weiqianList.Visible = MenuList.ListIndex = 1
            updateclick.Visible = MenuList.ListIndex = 1
            qingjiaClick.Visible = MenuList.ListIndex = 1
            unqingjiaclick.Visible = MenuList.ListIndex = 1
            zrsclick.Visible = MenuList.ListIndex = 1
            unzrsclick.Visible = MenuList.ListIndex = 1
            qingjialist.Visible = MenuList.ListIndex = 1
            zrsList.Visible = MenuList.ListIndex = 1
            updatewq.Visible = MenuList.ListIndex = 1
            tiptsqk.Visible = MenuList.ListIndex = 1
            tiptsqk2.Visible = MenuList.ListIndex = 1
            recordtip1.Visible = MenuList.ListIndex = 2
            getposition.Visible = MenuList.ListIndex = 2
            saveonexit.Visible = MenuList.ListIndex = 2
            outputqr.Visible = MenuList.ListIndex = 2 Or MenuList.ListIndex = 3
            outputpositiontext.Visible = MenuList.ListIndex = 2
            timerequire.Visible = MenuList.ListIndex = 3
            qdydtip.Visible = MenuList.ListIndex = 3
            Timenowtip.Visible = MenuList.ListIndex = 3
            nowtimetext.Visible = MenuList.ListIndex = 3
            zrstqtip.Visible = MenuList.ListIndex = 3
            zrstqtext.Visible = MenuList.ListIndex = 3
           mmxgtip.Visible = MenuList.ListIndex = 4
           ylmmtip.Visible = MenuList.ListIndex = 4
           ylmmtext.Visible = MenuList.ListIndex = 4
           newmm.Visible = MenuList.ListIndex = 4
           newmmtext.Visible = MenuList.ListIndex = 4
           newmmagaintip.Visible = MenuList.ListIndex = 4
           newmmagaintext.Visible = MenuList.ListIndex = 4
           mmggcommand.Visible = MenuList.ListIndex = 4
           endqd.Visible = MenuList.ListIndex = 4
           reset.Visible = MenuList.ListIndex = 4
                     CRlabel.Visible = MenuList.ListIndex = 5
            TMImage.Visible = MenuList.ListIndex = 5
            PrImage.Visible = MenuList.ListIndex = 5
    Select Case MenuList.ListIndex
        Case 0
            Me.Width = 5850
            Me.Height = 4155
        Case 1
            Me.Width = 7110
            Me.Height = 6915
            
        Case 2
            Me.Width = 5205
            Me.Height = 4185

        Case 3
            Me.Width = 3915
            Me.Height = 4185

        Case 4
                    Me.Width = 3795
            Me.Height = 4350
        
        Case 5
            CRlabel.Caption = App.ProductName & Chr(10) & "Version " & App.Major & "." & App.Minor & App.Revision & Chr(10) & App.LegalCopyright & "��лʩ��չͬ־����Ľ��������"
  
            Me.Width = 5130
            Me.Height = 4275
    End Select
End Sub

Private Sub mingdanCommand_Click()
    MingdanList.List(MingdanList.ListIndex) = MingdanText.Text
    MingdanList.Selected(MingdanList.ListIndex) = False
    mingdanCommand.Enabled = False
    delmingdan.Enabled = False
End Sub

Private Sub MingdanList_Click()
    MingdanText.Text = MingdanList.List(MingdanList.ListIndex)
    delmingdan.Enabled = True
    mingdanCommand.Enabled = True
End Sub

Private Sub mmggcommand_Click()
Unlock #1
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
            MsgBox ("�޸ĳɹ�")
            newmmagaintext.Text = ""
            newmmtext.Text = ""
            ylmmtext.Text = ""
        Else
            MsgBox ("�����������벻һ�£�")
        End If
    ElseIf passwordtime > 0 Then
        MsgBox ("������󣬻���������" + CStr(passwordtime Mod 5) + "��")
    Else
        MsgBox ("�������ǿ�Ƶǳ���")
        Call Form_Unload(0)
    End If
    Lock #1
End Sub

Private Sub newname_Click()
    If MingdanText.Text = "" Then Exit Sub
    MingdanList.AddItem MingdanText.Text
    newname.Enabled = MingdanList.ListCount < 100
    delmingdan.Enabled = MingdanList.ListIndex > 0
    MingdanText.Text = ""
End Sub

Private Sub outputmingdan_Click()
    Filerequire.CancelError = True
    Filerequire.Flags = cdIOFNHideReadOnly
    Filerequire.Filter = "�����ļ�(*.*)|*.*|TXT�ı��ĵ�(*.txt)|*.txt"
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
    Dim aaa, bbb As Date
    Dim ccc
    aaa = DateTime.TimeValue(timerequire.Text)
    bbb = DateTime.TimeValue(nowtimetext.Text)
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.createtextfile(App.Path + "\settingsaved.classsign", True)
    f.writeline Mainform.Encode(outputpositiontext.Text, 33965, 41831)
    f.writeline Mainform.Encode(CStr(saveonexit.Value), 33965, 41831)
    f.writeline Mainform.Encode(CStr(timerequire.Text), 33965, 41831)
    f.writeline Mainform.Encode(CStr(zrstqtext.Text), 33965, 41831)
    Mainform.tqtime = Val(zrstqtext.Text)
    If TimeValue(nowtimetext.Text) - Val(zrstqtext.Text) / 60 / 24 < DateTime.Time Then
    ccc = MsgBox("��ǰʱ���ѳ���ֵ���������������ٵ��Ļ���ʱ���ˣ�����������»��ߣ������ǰ����ǩ�����ٵ�����Ա���ڻ��߷�Χ�ڡ��Ƿ�����������»��ߣ�", vbYesNo, "�༶ǩ��ϵͳ - Ӧ������")
    If ccc = vbYes Then
    Mainform.lineTimer.Enabled = True
    Mainform.zrshx.Enabled = True
    Mainform.settime = TimeValue(nowtimetext.Text)
    For i = 1 To Mainform.recordsList(0).ListCount
        If Mainform.recordsList(0).List(i - 1) = "------------" Then
            Mainform.recordsList(0).RemoveItem (i - 1)
            Mainform.recordsList(0).RemoveItem (i - 1)
            Exit For
        End If
    Next i
    For i = 1 To Mainform.recordsList(1).ListCount
        If Mainform.recordsList(1).List(i - 1) = "--------------" Then
            Mainform.recordsList(1).RemoveItem (i - 1)
            Mainform.recordsList(1).RemoveItem (i - 1)
            Exit For
        End If
    Next i
    Mainform.outputCommand.Enabled = True
    Else
    Mainform.lineTimer.Enabled = False
    Mainform.zrshx.Enabled = False
    End If
    Else
    Mainform.settime = TimeValue(nowtimetext.Text)
    For i = 1 To Mainform.recordsList(0).ListCount
        If Mainform.recordsList(0).List(i - 1) = "------------" Then
            Mainform.recordsList(0).RemoveItem (i - 1)
            Mainform.recordsList(0).RemoveItem (i - 1)
            Exit For
        End If
    Next i
    For i = 1 To Mainform.recordsList(1).ListCount
        If Mainform.recordsList(1).List(i - 1) = "--------------" Then
            Mainform.recordsList(1).RemoveItem (i - 1)
            Mainform.recordsList(1).RemoveItem (i - 1)
            Exit For
        End If
    Next i
    Mainform.lineTimer.Enabled = True
    Mainform.zrshx.Enabled = True
    Mainform.outputCommand.Enabled = True
    End If
    Mainform.savepath = outputpositiontext.Text
    Mainform.saveonexit = saveonexit.Value
errorHandler:
    If Err.Number <> 0 Then
    MsgBox "�����ˣ����������ʱ���Ƿ���ϱ�׼��ʱ���ʽ��" + Chr(10) + CStr(Err.Number) + " " + Err.Description
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
    If MingdanList.ListIndex > -1 Then
        newname.Enabled = MingdanList.ListCount < 100
        delmingdan.Enabled = False
        mingdanCommand.Enabled = False
        MingdanText.Text = ""
    End If
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
    If MingdanList.ListCount >= 100 Then newname.Enabled = False Else newname.Enabled = True
    delmingdan.Enabled = False
    mingdanCommand.Enabled = False
    ts.Close
End Sub

Private Sub reset_Click()
    Mainform.qiandaoList.Clear
    Mainform.recordsList(0).Clear
    Mainform.zrsList.Clear
    Mainform.NameCombo.Clear
    Mainform.recordsList(1).Clear
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
    If MingdanList.ListIndex > -1 Then
        MingdanList.Selected(MingdanList.ListIndex) = False
        newname.Enabled = MingdanList.ListCount < 100
        delmingdan.Enabled = False
        mingdanCommand.Enabled = False
        MingdanText.Text = ""
    End If
    MsgBox ("����ɹ���ǩ�����������´�Ӧ��")
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
    weiqianList.AddItem zrsList.List(zrsList.ListIndex)
    zrsList.RemoveItem zrsList.ListIndex
    unzrsclick.Enabled = False
    updateclick.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub updateclick_Click()
    weiqianList.Clear
    zrsList.Clear
    For i = 1 To Mainform.qiandaoList.ListCount
        weiqianList.AddItem Mainform.qiandaoList.List(i - 1)
    Next i
    For i = 1 To Mainform.zrsList.ListCount
        zrsList.AddItem Mainform.zrsList.List(i - 1)
    Next i
    updatewq.Enabled = True
    Mainform.Enabled = False
    zrsList.Enabled = True
    weiqianList.Enabled = True
    updateclick.Enabled = False
    qingjialist.Enabled = True
End Sub
Private Sub updatewq_Click()
    Mainform.qiandaoList.Clear
    Mainform.zrsList.Clear
    Mainform.NameCombo.Clear
    For i = 1 To weiqianList.ListCount
        Mainform.qiandaoList.AddItem weiqianList.List(i - 1)
        Mainform.NameCombo.AddItem weiqianList.List(i - 1)
    Next i
    For i = 1 To zrsList.ListCount
        Mainform.zrsList.AddItem zrsList.List(i - 1)
    Next i
    Mainform.Enabled = True
    updateclick.Enabled = True
    updatewq.Enabled = False
    zrsList.Enabled = False
    weiqianList.Enabled = False
    qingjialist.Enabled = False
    qingjiaClick.Enabled = False
    unqingjiaclick.Enabled = False
    unzrsclick.Enabled = False
    zrsclick.Enabled = False
    If weiqianList.ListIndex >= 0 Then weiqianList.Selected(weiqianList.ListIndex) = False
    If qingjialist.ListIndex >= 0 Then qingjialist.Selected(qingjialist.ListIndex) = False
    If zrsList.ListIndex >= 0 Then zrsList.Selected(zrsList.ListIndex) = False
End Sub

Private Sub weiqianList_Click()
    If weiqianList.ListIndex >= 0 Then
    qingjiaClick.Enabled = True
    zrsclick.Enabled = True
    End If
End Sub

Private Sub zrsclick_Click()
    zrsList.AddItem weiqianList.List(weiqianList.ListIndex)
    weiqianList.RemoveItem (weiqianList.ListIndex)
    zrsclick.Enabled = False
    qingjiaClick.Enabled = False
    updateclick.Enabled = False
    Mainform.Enabled = False
    updatewq.Enabled = True
End Sub

Private Sub zrslist_Click()
If zrsList.ListIndex >= 0 Then
    unzrsclick.Enabled = True
End If
End Sub

Public Function ShowFolderDialog() As String
'/��򵥵���ʾ�ļ���ѡ��Ի��򷽷�
Dim spShell, spFolder, spFolderItem, spPath As String
Const WINDOW_HANDLE = 0
Const NO_OPTIONS = 0
Set spShell = CreateObject("Shell.Application")
Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "ѡ��Ŀ¼:", NO_OPTIONS, "")
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
