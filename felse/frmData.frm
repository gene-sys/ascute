VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmData 
   Caption         =   "��������� ������"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form3"
   ScaleHeight     =   7050
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc AdoOpros 
      Height          =   330
      Left            =   0
      Top             =   6360
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmData.frx":0000
      Height          =   5895
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   4
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   820
      ButtonWidth     =   609
      ButtonHeight    =   767
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmData.frx":0017
         Left            =   1725
         List            =   "frmData.frx":0021
         TabIndex        =   13
         Top             =   -15
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   3525
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   6615
         TabIndex        =   11
         Top             =   -15
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   4905
         TabIndex        =   10
         Top             =   -15
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   8040
         TabIndex        =   8
         Top             =   -15
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39843
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   6315
         TabIndex        =   7
         Top             =   -15
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   39843
      End
      Begin VB.CommandButton cmdUnload 
         Height          =   345
         Left            =   0
         Picture         =   "frmData.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "����� (F2)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   330
         Picture         =   "frmData.frx":037A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "�������� ������ (F5)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   1380
         Picture         =   "frmData.frx":2074
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "�������� �����"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnLoad 
         Height          =   345
         Left            =   1020
         Picture         =   "frmData.frx":293E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "��������� ������ �� �����"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnSave 
         Height          =   345
         Left            =   660
         Picture         =   "frmData.frx":2C48
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "��������� ������ � ����"
         Top             =   0
         Width           =   345
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmData.frx":2F4E
         Height          =   315
         Left            =   8385
         TabIndex        =   1
         Top             =   -15
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "naim_uzl"
         BoundColumn     =   "kod_uzl"
         Text            =   ""
         Object.DataMember      =   "cmdUU"
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6675
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoX As ADODB.Connection
'
Private Sub btnLoad_Click()
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String
Dim pos As Long
On Error Resume Next
CDArh.CancelError = True
CDArh.ShowOpen ' ������� ����
filenum = FreeFile: filenum = filenum - 1
Close #filenum
If CDArh.FileName <> "" Then
    If OpenCSV(CDArh.FileName) Then MsgBox "���� ��������" _
    Else MsgBox "���� �� ��������"
End If
End Sub

Private Sub btnSave_Click()
Dim NFSO As New FileSystemObject
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String, i As Integer
Dim pos As Long
On Error GoTo errbtnSave_Click
' ������� ����
filenum = FreeFile ' ���������� c��������� �����
Open "tempbase.csv" For Input As #filenum ' ������� ����
' ���������� ������ � ������ ������
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
' �������������� ������ ���������
Line Input #filenum, strX '
pos = InStr(1, strX, "�������1", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "W3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������2", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "m3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������3", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "To;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������4", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t1;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������5", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t2;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������6", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������7", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "P1;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������8", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "P2;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "�������9", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "T���;" & Mid(strX, pos + 9)
strY = strY & strX & vbCrLf
' ��������� ������ �������� ��� ����
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' ������� ����
' ��������� ������ �� ����
CDArh.CancelError = True
CDArh.FileName = "���� ���������"
CDArh.ShowSave ' ������� ����
ChDir App.Path
filenum = FreeFile ' ������� ���� ��� ������
Open Mid(CDArh.FileName, 1, Len(CDArh.FileName) - 3) & "csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' ��������� ����
Close #filenum
' �������� �������� ������
TipOt = ""
Call AppResult(Mid(CDArh.FileName, 1, Len(CDArh.FileName) - 3) & "csv")
'Exit Sub
TipOt = ""
errbtnSave_Click:
ChDir App.Path
End Sub
' ������� ����������� �������� ������
Function AppResult(NofFile As String)
Dim FileNumber As Long
Dim sHead As String
On Error Resume Next
If Len(TipOt) = 0 Then TipOt = RemakeHead ' �������������� ��������� � ������ �������
sHead = "�����"
' �������� ��� ����� ������
Select Case TipOt
Case "�Journal", "�Journal" ' ���� ������� �������� �������
    'PrintVis
    If Not frmStart.mnuKT.Checked Then
        If DataEnvironment1.rsCommand8.State <> adStateOpen Then _
            DataEnvironment1.rsCommand8.Open ' ������� ������ � tembase.csv ��� ������
        ' ��������� �����
        With DataEnvironment1.rsCommand8
            .Requery
            sHead = sHead & ";" & Format(.Fields(0), "0.00") ' ���.�����. W1
            sHead = sHead & ";" & Format(.Fields(3), "0.00") ' ���.�����. W2
            sHead = sHead & ";" & Format(.Fields(1), "0.00") ' ����� m1
            sHead = sHead & ";" & Format(.Fields(4), "0.00") ' ����� m2
            sHead = sHead & ";" & str(.Fields(9)) ' ����� �����.
            sHead = sHead & ";" & Format(.Fields(6), "0.00") ' ���.�����. W3
            sHead = sHead & ";" & Format(.Fields(7), "0.00") ' ����� m3
            sHead = sHead & ";" & str(.Fields(8))  ' ����� �����.
            sHead = sHead & ";" & Format(.Fields(2), "0.00") '������. t1
            sHead = sHead & ";" & Format(.Fields(5), "0.00") '����-�� t2
            sHead = sHead & ";;" & Format(.Fields(10), "0.00") '����. �� ����. d1
            sHead = sHead & ";" & Format(.Fields(11), "0.00") '����.�� �����. d2
        End With
    Else
        If DataEnvironment1.rsCommand18.State <> adStateOpen Then _
            DataEnvironment1.rsCommand18.Open ' ������� ������ � tembase.csv ��� ������
        ' ��������� �����
        With DataEnvironment1.rsCommand18
            .Requery
            sHead = sHead & ";" & Format(.Fields(0), "0.00") ' ���.�����. W1
            sHead = sHead & ";" & Format(.Fields(1), "0.00") ' ���.�����. W2
            sHead = sHead & ";" & Format(.Fields(2), "0.00") ' ����� m1
            sHead = sHead & ";" & Format(.Fields(3), "0.00") ' ����� m2
            sHead = sHead & ";" & str(.Fields(4))  ' ����� �����.
            sHead = sHead & ";" & Format(.Fields(7), "0.00") ' ���.�����. W3
            sHead = sHead & ";" & Format(.Fields(8), "0.00") ' ����� m3
            sHead = sHead & ";" & str(.Fields(9)) ' ����� ������
            sHead = sHead & ";" & Format(.Fields(10), "0.00") ' ������. t1
            sHead = sHead & ";" & Format(.Fields(11), "0.00") ' ����-�� t2
            sHead = sHead & ";;" & Format(.Fields(5), "0.00") ' ����. �� ����. d1
            sHead = sHead & ";" & Format(.Fields(6), "0.00") ' ����.�� �����. d2
            sHead = sHead & ";;" & Format(.Fields(12), "0.00") ' ���-�� ������� ����.
            sHead = sHead & ";" & Format(.Fields(13), "0.00") ' ������-��� ����������
        End With
    End If
Case "�HV", "�HV"
'    �������� ���� ��� �������� ������
Case "�PAR"
    ' ��� �� ������� �� �������������� ��������
    'sHead = sHead & ";"
    If DataEnvironment1.rsCommand9.State <> adStateOpen Then _
        DataEnvironment1.rsCommand9.Open ' ������� ������ � tembase.csv ��� ������
    With DataEnvironment1.rsCommand9
        .Requery
        sHead = sHead & ";" & Format(.Fields(0), "0.00") ' ������. t1
        sHead = sHead & ";" & Format(.Fields(1), "0.00") ' ������ d1
        sHead = sHead & ";" & Format(.Fields(2), "0.00") '����� V1
        sHead = sHead & ";" & Format(.Fields(3), "0.00") ' ���.�����. W1
        sHead = sHead & ";" & Format(.Fields(4), "0.0") ' Tr1
        sHead = sHead & ";" & Format(.Fields(6), "0.00") ' ����-�� t2
        sHead = sHead & ";" & Format(.Fields(7), "0.00") ' ������. d2
        sHead = sHead & ";" & Format(.Fields(8), "0.00") ' ����� V2
        sHead = sHead & ";" & Format(.Fields(9), "0.00") ' ���.�����. W2
        sHead = sHead & ";" & Format(.Fields(10), "0.0") ' Tr2
        'sHead = sHead & ";" & Format(.Fields(5), "0.0") ' To1
        'sHead = sHead & ";" & Format(.Fields(11), "0.0") ' To2
    End With
Case "�PAR"
    ' ��� �� ������� �� �������������� ��������
    'sHead = sHead & ";"
    If DataEnvironment1.rsCommand10.State <> adStateOpen Then _
        DataEnvironment1.rsCommand10.Open ' ������� ������ � tembase.csv ��� ������
    With DataEnvironment1.rsCommand10
        .Requery
        sHead = sHead & ";" & Format(.Fields(0), "0.00") ' ������. t1
        sHead = sHead & ";" & Format(.Fields(1), "0.00") ' ������ d1
        sHead = sHead & ";" & Format(.Fields(2), "0.00") '����� V1
        sHead = sHead & ";" & Format(.Fields(3), "0.00") ' ���.�����. W1
        sHead = sHead & ";" & Format(.Fields(4), "0.0") ' Tr1
        sHead = sHead & ";" & Format(.Fields(6), "0.00") ' ����-�� t2
        sHead = sHead & ";" & Format(.Fields(7), "0.00") ' ������. d2
        sHead = sHead & ";" & Format(.Fields(8), "0.00") ' ����� V2
        sHead = sHead & ";" & Format(.Fields(9), "0.00") ' ���.�����. W2
        sHead = sHead & ";" & Format(.Fields(10), "0.0") ' Tr2
        'sHead = sHead & ";" & Format(.Fields(5), "0.0") ' To1
        'sHead = sHead & ";" & Format(.Fields(11), "0.0") ' To2
    End With
End Select
' ������� ����
FileNumber = FreeFile
Open NofFile For Append As #FileNumber
Print #FileNumber, sHead ' ����� ���������
Close #FileNumber
End Function

Private Sub Check1_Click()
On Error GoTo Check1_err
' ����� ������ ��������� ������
If Check1.Value = 1 Then Check1.Caption = "����� '����'" _
Else Check1.Caption = "����� '����'"
Exit Sub
Check1_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub


Private Sub Combo1_LostFocus()
    ' ��������� ��� ������
    Call WriteParameters("KindArchive", Combo1.Text)
End Sub

Private Sub Command2_Click()
Dim strD As String, n As String, x As Long
Dim TipUU As Long, dop As String
On Error Resume Next
' ������������ � ��������� ������
Me.StatusBar1.SimpleText = "����������� ������"
If Combo1.Text = "�������" Then
    n = "h": dop = " HOUR(vremy)"
Else
    n = "s": dop = " ''"
End If
' ���������� �������
If Check1.Value = 1 Then strD = "w3l,v3l" Else strD = "w3z,v3z"
x = Me.DataCombo1.BoundText   ' �������� ��� ����
'
'����� ���� ���� - �����, ���, ���.����, �����
With DataEnvironment1.rscmdUU
    .AbsolutePosition = Me.DataCombo1.SelectedItem
    TipUU = .Fields(3)
End With
' ������������ ������
Select Case TipUU
Case 1 ' ��������
    Me.AdoOpros.RecordSource = "SELECT  data," & dop & " as Vremja,w1,v1,w2,v2," & _
                        "t1,p1,t2,p2,vrem_n," & strD & " FROM teplo_" & n & "r " & _
                        "WHERE kod_uzl = " & x & " AND data BETWEEN " & _
                        "date_format('" & MySQLDate(Me.Text4.Text) & "','%Y-%m-%d') AND " & _
                        "date_format('" & MySQLDate(Me.Text5.Text) & "','%Y-%m-%d');"
Case 2 ' ���
    Me.AdoOpros.RecordSource = "SELECT data," & dop & " as Vremja,w1,m1,t1,p1," & _
                        "vrem_n1,m2,t2,p2,vrem_n2 FROM par_" & n & _
                        " WHERE kod_uzl = " & x & " AND data BETWEEN " & _
                        "date_format('" & MySQLDate(Me.Text4.Text) & "','%Y-%m-%d') AND " & _
                        "date_format('" & MySQLDate(Me.Text5.Text) & "','%Y-%m-%d');"
Case 3 ' ���.����
    Me.AdoOpros.RecordSource = "SELECT data," & dop & " as Vremja,v,vrem_n FROM voda_" & n & _
                        " WHERE kod_uzl = " & x & " AND data BETWEEN " & _
                        "date_format('" & MySQLDate(Me.Text4.Text) & "','%Y-%m-%d') AND " & _
                        "date_format('" & MySQLDate(Me.Text5.Text) & "','%Y-%m-%d');"
Case 4 ' �����
    Me.AdoOpros.RecordSource = "SELECT data," & dop & " as Vremja,h_min1,h_max1,v1," & _
                        "vrem_ot1,h_min2,h_max2,v2,vrem_ot2 FROM stok_" & n & _
                        " WHERE kod_uzl = " & x & " AND data BETWEEN " & _
                        "date_format('" & MySQLDate(Me.Text4.Text) & "','%Y-%m-%d') AND " & _
                        "date_format('" & MySQLDate(Me.Text5.Text) & "','%Y-%m-%d');"
End Select
Me.AdoOpros.Refresh ' �������� ���������
End Sub

Private Sub Command3_Click()
Dim i As Long
On Error GoTo Print_err
If Len(TipOt) = 0 Then TipOt = RemakeHead ' �������������� ��������� � ������ �������
Select Case TipOt
Case "�Journal", "�Journal" ' ���� ������� �������� �������
    If Dialog.Check3.Value And Left(TipOt, 1) = "�" Then
        'Call PrintTSRV(Me.Check1.Value) '������ ����
    Else
        ' ��������� ���� �� �������� ���-��
        If Not frmStart.mnuKT.Checked Then
            If DataEnvironment1.rsCommand3.State <> adStateOpen Then _
                DataEnvironment1.rsCommand3.Open ' ����������� ������
            DataEnvironment1.rsCommand3.Requery ' �������� ������
            ' �������� ������
            With DataEnvironment1.rsCommand8
                If .State <> adStateOpen Then .Open ' ����������� ������
                .Requery ' �������� ������
                ' ������������
                For i = 0 To 11
                    DataReport1.Sections(5).Controls("Label" & _
                        Trim(str(26 + i))).Caption = Format(.Fields(i), "0.00")
                Next
            End With
            ' ���������� ������ ��������� ��������� ����� �� ������
            DataReport1.Orientation = rptOrientLandscape
            ' ���������� ������� ������
            With DataEnvironment1.rsCommand11
                If .State <> adStateOpen Then .Open ' ����������� ������
                DataReport1.BottomMargin = .Fields("Niz") * Stwips
                DataReport1.TopMargin = .Fields("Verh") * Stwips
                DataReport1.LeftMargin = .Fields("Levo") * Stwips
                DataReport1.RightMargin = .Fields("Pravo") * Stwips
                DataReport1.Font.size = .Fields("Shrift")
            End With
            DataReport1.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
            DataReport1.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
            DataReport1.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
            DataReport1.Show ' ������������ � ������
        Else
            If DataEnvironment1.rsCommand17.State <> adStateOpen Then _
            DataEnvironment1.rsCommand17.Open ' ������� ������ � tembase.csv ��� ������
            If DataEnvironment1.rsCommand18.State <> adStateOpen Then _
                DataEnvironment1.rsCommand18.Open ' ������� ������ � tembase.csv ��� ������
            DataEnvironment1.rsCommand17.Requery ' �������� ������
            ' ���������� ������ ���������� ����� �� ������
            DataReport6.Orientation = rptOrientLandscape
            ' ��������� �����
            With DataEnvironment1.rsCommand18
                .Requery
                DataReport6.Sections(5).Controls(1).Caption = Format(.Fields(0), "0.00") ' ���.�����. W1
                DataReport6.Sections(5).Controls(3).Caption = Format(.Fields(2), "0.00") ' ����� V1
                DataReport6.Sections(5).Controls(4).Caption = Format(.Fields(10), "0.00") ' ������. t1
                DataReport6.Sections(5).Controls(5).Caption = Format(.Fields(1), "0.00")  ' ���.�����. W2
                DataReport6.Sections(5).Controls(6).Caption = Format(.Fields(3), "0.00") ' ����� V2
                DataReport6.Sections(5).Controls(7).Caption = Format(.Fields(11), "0.00") ' ����-�� t2
                DataReport6.Sections(5).Controls(8).Caption = Format(.Fields(7), "0.00") ' ���.�����. W3
                DataReport6.Sections(5).Controls(9).Caption = Format(.Fields(8), "0.00") ' ����� V3
                DataReport6.Sections(5).Controls(10).Caption = Format(.Fields(9), "0.00") ' ����� �����.
                DataReport6.Sections(5).Controls(11).Caption = Format(.Fields(4), "0.00")  ' ����� �����.
                DataReport6.Sections(5).Controls(12).Caption = Format(.Fields(5), "0.00") ' ����. �� ����. d1
                DataReport6.Sections(5).Controls(13).Caption = Format(.Fields(6), "0.00") ' ����.�� �����. d2
                DataReport6.Sections(5).Controls("Label40").Caption = _
                                                        Format(.Fields(12), "0.00") ' ���-�� ������� ����.
                DataReport6.Sections(5).Controls("Label41").Caption = _
                                                        Format(.Fields(13), "0.00") ' ������-��� ����������
            End With
            ' ���������� ������� ������
            With DataEnvironment1.rsCommand11
                If .State <> adStateOpen Then .Open ' ����������� ������
                DataReport6.BottomMargin = .Fields("Niz") * Stwips
                DataReport6.TopMargin = .Fields("Verh") * Stwips
                DataReport6.LeftMargin = .Fields("Levo") * Stwips
                DataReport6.RightMargin = .Fields("Pravo") * Stwips
                DataReport6.Font.size = .Fields("Shrift")
            End With
            ' ���������� "�����" ������
            DataReport6.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
            DataReport6.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
            DataReport6.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
            DataReport6.Show ' ������������ � ������
        End If
    End If
Case "�HV" ' ���� ������� ���.���� - ��������
    If DataEnvironment1.rsCommand4.State <> adStateOpen Then _
                                        DataEnvironment1.rsCommand4.Open ' ����������� ������
    DataEnvironment1.rsCommand4.Requery ' �������� ������
    ' ���������� ������� ������
    With DataEnvironment1.rsCommand13
        If .State <> adStateOpen Then .Open ' ����������� ������
        DataReport2.BottomMargin = .Fields("Niz") * Stwips
        DataReport2.TopMargin = .Fields("Verh") * Stwips
        DataReport2.LeftMargin = .Fields("Levo") * Stwips
        DataReport2.RightMargin = .Fields("Pravo") * Stwips
        DataReport2.Font.size = .Fields("Shrift")
    End With
    DataReport2.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport2.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport2.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport2.Show ' ������������ � ������
Case "�HV" ' ���� ������� ���.���� - �������
    If DataEnvironment1.rsCommand5.State <> adStateOpen Then _
        DataEnvironment1.rsCommand5.Open ' ����������� ������
    DataEnvironment1.rsCommand5.Requery ' �������� ������
    ' ���������� ������� ������
    With DataEnvironment1.rsCommand13
        If .State <> adStateOpen Then .Open ' ����������� ������
        DataReport4.BottomMargin = .Fields("Niz") * Stwips
        DataReport4.TopMargin = .Fields("Verh") * Stwips
        DataReport4.LeftMargin = .Fields("Levo") * Stwips
        DataReport4.RightMargin = .Fields("Pravo") * Stwips
        DataReport4.Font.size = .Fields("Shrift")
    End With
    DataReport4.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport4.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport4.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport4.Show ' ������������ � ������
Case "�PAR" ' ���� ������� ���� ���.
    If DataEnvironment1.rsCommand6.State <> adStateOpen Then _
        DataEnvironment1.rsCommand6.Open ' ����������� ������
    DataEnvironment1.rsCommand6.Requery ' �������� ������
    ' �������� ������
    With DataEnvironment1.rsCommand9
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery ' �������� ������
        ' ������������
        For i = 0 To 11
            DataReport3.Sections(5).Controls("Label" & _
                Trim(str(25 + i))).Caption = Format(.Fields(i), "0.00")
        Next
    End With
    ' ���������� ������ ��������� ��������� ����� �� ������
    DataReport3.Orientation = rptOrientLandscape
    ' ���������� ������� ������
    With DataEnvironment1.rsCommand12
        If .State <> adStateOpen Then .Open ' ����������� ������
        DataReport3.BottomMargin = .Fields("Niz") * Stwips
        DataReport3.TopMargin = .Fields("Verh") * Stwips
        DataReport3.LeftMargin = .Fields("Levo") * Stwips
        DataReport3.RightMargin = .Fields("Pravo") * Stwips
        DataReport3.Font.size = .Fields("Shrift")
    End With
    DataReport3.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport3.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport3.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport3.Show ' ������������ � ������
Case "�PAR" ' ���� ������� ���� ���.
    If DataEnvironment1.rsCommand7.State <> adStateOpen Then _
                                    DataEnvironment1.rsCommand7.Open ' ����������� ������
    DataEnvironment1.rsCommand7.Requery ' �������� ������
    ' �������� ������
    With DataEnvironment1.rsCommand10
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery ' �������� ������
        ' ������������
        For i = 0 To 11
            DataReport5.Sections(5).Controls("Label" & _
                Trim(str(25 + i))).Caption = Format(.Fields(i), "0.00")
        Next
    End With
    ' ���������� ������ ��������� ��������� ����� �� ������
    DataReport5.Orientation = rptOrientLandscape
    ' ���������� ������� ������
    With DataEnvironment1.rsCommand12
        If .State <> adStateOpen Then .Open ' ����������� ������
        DataReport5.BottomMargin = .Fields("Niz") * Stwips
        DataReport5.TopMargin = .Fields("Verh") * Stwips
        DataReport5.LeftMargin = .Fields("Levo") * Stwips
        DataReport5.RightMargin = .Fields("Pravo") * Stwips
        DataReport5.Font.size = .Fields("Shrift")
    End With
    DataReport5.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport5.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport5.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport5.Show ' ������������ � ������
End Select
' ������������ ��������� �����
'Call RecovHead
TipOt = ""
Exit Sub
Print_err:
 'MsgBox Err.Number & "->" & Err.Description
 Resume Next
End Sub

Private Sub DTPicker1_CloseUp()
Me.Text4.Text = Me.DTPicker1.Value
End Sub


Private Sub DTPicker2_CloseUp()
Me.Text5.Text = Me.DTPicker2.Value
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
    Call cmdUnload_Click
'Case vbKeyF3
'    Call Command4_Click
'Case vbKeyF4
'    Call Command1_Click
Case vbKeyF5
    Call Command2_Click
End Select
End Sub

Private Sub Form_load()
' ������������ C����� �����
' �������� ������ � ������� �����./�����.
    On Error GoTo Fload_err
    '��������� ������� �������
    With DataEnvironment1.rsCommand2 ' ����� ���������
        If .State <> adStateOpen Then .Open
        .Requery
        .MoveFirst
        Do While Not .EOF
            Select Case .Fields("NameSet") ' ����� �������� ��������� ���������
            Case "Mode" ' ������� ����� ��������� ������
                If .Fields("Set") = "True" Then
                    Check1.Value = 1
                    Check1.Caption = "����� '����'"
                Else
                    Check1.Value = 0
                    Check1.Caption = "����� '����'"
                End If
            Case "KindArchive" ' ������� ��� ������
                Combo1.Text = .Fields("Set")
            End Select
            .MoveNext ' �� ���� ���������� ���������
        Loop
    End With
    ' ������� ������ ��������� ������: �� ������ ������
    Text4.Text = CDate("01/" & Month(Date) & "/" & Year(Date))
    Text5.Text = Date ' �� ������� ����
    Me.DTPicker1.Value = Date
    Me.DTPicker2.Value = Date
    ' ����������� � �� ������
    Me.AdoOpros.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                    & "SERVER=192.168.100.23;" _
                    & " DATABASE=askute;" _
                    & "OPTION=0; PORT=3306"
    Me.AdoOpros.Password = frmLogin.strPSW ' ������
    Me.AdoOpros.UserName = frmLogin.strUser ' ��� �������
    ' ������������ �������
    Set adoX = Me.AdoOpros.Recordset.ActiveConnection
    Exit Sub
Fload_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
    '
End Sub

Private Sub Form_Resize()
Dim twips As Long
twips = 567
Me.DataGrid1.Height = Me.Height - 3 * twips
Me.DataGrid1.Width = Me.Width - 0.3 * twips
Me.AdoOpros.Top = Me.Height - 2 * twips
End Sub
