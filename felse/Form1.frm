VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A237EE18-33EE-468A-B4D8-07559BD2E396}#5.0#0"; "ProgBar.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Form1 
   Caption         =   "��������� ������"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11700
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   3  'Windows Default
   Begin ProgressBar.PrBar PrBar1 
      Height          =   345
      Left            =   6720
      TabIndex        =   0
      Top             =   7080
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   609
      BorderStyle     =   1
      ProgFont        =   "Times New Roman"
      ProgrColor      =   16711680
      BarColor        =   13619151
      CustomCaption   =   ""
      LabelType       =   2
      Value           =   0
      MaxValue        =   50
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   820
      ButtonWidth     =   609
      ButtonHeight    =   767
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form1.frx":08CA
         Height          =   315
         Left            =   9345
         TabIndex        =   16
         Top             =   -15
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "�������"
         Text            =   ""
         Object.DataMember      =   "Command1"
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   705
         Picture         =   "Form1.frx":08EA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "��������� ����� (F4)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnSave 
         Height          =   345
         Left            =   1380
         Picture         =   "Form1.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "��������� ������ � ����"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnLoad 
         Height          =   345
         Left            =   1740
         Picture         =   "Form1.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "��������� ������ �� �����"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   2100
         Picture         =   "Form1.frx":1D3C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "�������� �����"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   1050
         Picture         =   "Form1.frx":2606
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "�������� ������ (F5)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdUnload 
         Height          =   345
         Left            =   0
         Picture         =   "Form1.frx":4300
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "����� (F2)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Height          =   345
         Left            =   360
         Picture         =   "Form1.frx":4642
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "����������� (F3)"
         Top             =   0
         Width           =   345
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   7275
         TabIndex        =   7
         Top             =   -15
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   39843
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   9000
         TabIndex        =   6
         Top             =   -15
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16318465
         CurrentDate     =   39843
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
         Left            =   5865
         TabIndex        =   5
         Top             =   -15
         Width           =   1455
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
         Left            =   7575
         TabIndex        =   4
         Top             =   -15
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   4485
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form1.frx":494C
         Left            =   2685
         List            =   "Form1.frx":4956
         TabIndex        =   2
         Top             =   -15
         Width           =   1695
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   8280
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDArh 
      Left            =   8640
      Top             =   448
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "�������� �����"
      Filter          =   "������ � ������������  (*.csv)|*.csv"
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7920
      Top             =   480
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   7035
      Width           =   11700
      _ExtentX        =   20638
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
   Begin MSDataGridLib.DataGrid DBGrid1 
      Align           =   1  'Align Top
      Bindings        =   "Form1.frx":496D
      Height          =   5535
      Left            =   0
      TabIndex        =   18
      Top             =   465
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   9763
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      DataMember      =   "Command1"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "�������"
         Caption         =   "�������"
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
         DataField       =   "���������"
         Caption         =   "���������"
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
      BeginProperty Column02 
         DataField       =   "�����"
         Caption         =   "�����"
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
      BeginProperty Column03 
         DataField       =   "Users"
         Caption         =   "������"
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
      BeginProperty Column04 
         DataField       =   "���������"
         Caption         =   "�����"
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
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
            ColumnWidth     =   174,009
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   263,017
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   63,987
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   109,984
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   96,983
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":498C
      Height          =   6465
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   11404
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Dim Wrkspce As DAO.Workspace
'Dim mcnnUnits As DAO.Database ' ��������������� �����
'Dim mrstMain As DAO.Recordset  ' ����������� � ��������� �������
'Dim mrstUnits As DAO.Recordset ' ����������� � ������ ����
Dim blTP As Boolean
'
Dim sz As Long ' ������ ����� �����������
Dim size As Integer ' ���������� ������ ����� �� 8 ��
Dim fsize As Long ' ������ ������������� �����
Dim nameNode As String ' ��� ���� ��� ��������
Dim strTSRV1 As String ' ��� ���������� 1�� ������ ������ ����
Dim strTSRV2 As String ' ��� ���������� 2�� ������ ������ ����
Dim TipOt As String
Private Const chunk = 8000
Private Const Stwips = 537

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
'

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

'
'
Private Sub cmdUnload_Click()
'If MsgBox("������������� �����?", vbQuestion + vbYesNo) = vbYes Then
    'Unload frmLogin
    ' ��������� �����
    Unload Me
'End If
End Sub
'
'
Private Sub Combo1_LostFocus()
    ' ��������� ��� ������
    Call WriteParameters("KindArchive", Combo1.Text)
End Sub

Private Sub Command1_Click()
Dim strD As String, data8 As String
If MsgBox("�������� ������ �����?", vbYesNo) = vbYes Then
    Me.DBGrid1.Visible = True
    Me.DataGrid1.Visible = False
    ' �������� ������ ����� �����
    Me.StatusBar1.SimpleText = "��������� �������� ����� �����"
    strD = "node(" & Mid(frmStart.Caption, _
            InStr(1, frmStart.Caption, "=", 1) + 1) & ")" ' ��������� ������
    ' ��������� ������
    ws.SendData strD
Else
    Me.DBGrid1.Visible = True
    Me.DataGrid1.Visible = False
End If
End Sub


Private Sub Command2_Click()
Dim strD As String, n As Long
On Error Resume Next
Me.DBGrid1.Visible = False
Me.DataGrid1.Visible = True
' ������������ � ��������� ������
Me.StatusBar1.SimpleText = "���� ��������� ������"
If Combo1.Text = "�������" Then n = 2 Else n = 3
TipOt = ""
' ������������ ������
nameNode = Me.DataCombo1.Text  'Text6.Text
' ��������� ������ �� ���������� ���� �����
strD = "/get[" & Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1) & "][" & nameNode & _
        "][" & Text4.Text & "][" & Text5.Text & "][" & _
             Combo1.Text & "][" & Trim(str(Check1.Value)) & "]"
ws.SendData strD ' ��������� ������
End Sub

Private Sub Command4_Click()
'������� �� ������ ������, ���� ���������� ��� �������
Me.StatusBar1.SimpleText = ""
ws.Close
'����������� � �������� �� ����� 1001
'�������������� ����� ������ �������
ws.Connect Dialog.Text2.Text, Dialog.Text1.Text  ' 1001
'������ ���, ����� ������������ �� ���� ������ ��� ������
'������ ���������� ������ ��� ��� ���� �������� ������
Me.Command4.Enabled = False
End Sub


Private Sub DBGrid1_DblClick()
' ���������� �������� ��� �����
If Len(Trim(DBGrid1.Columns(4))) = 0 Then
    DBGrid1.Columns(4) = "��" ' ������� ���� ����� ����
ElseIf StrComp(DBGrid1.Columns(4), "��") = 0 Then
    DBGrid1.Columns(4) = "��" ' ������� ���� ���� ���.����
ElseIf StrComp(DBGrid1.Columns(4), "��") = 0 Then
    DBGrid1.Columns(4) = "" ' ������� ���� ����� �����
End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
' ��������� ������� �������� ����� �������
If DBGrid1.Col = 0 Then
    Me.StatusBar1.SimpleText = "��� ������"
    KeyAscii = vbKeyCancel
End If
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
Case vbKeyF3
    Call Command4_Click
Case vbKeyF4
    Call Command1_Click
Case vbKeyF5
    Call Command2_Click
End Select
End Sub


'Private Sub Form_Resize()
'Me.DataGrid1.Width = Me.ScaleWidth
'Me.DataGrid1.Height = Me.ScaleHeight - Me.StatusBar1.Height
'Me.PrBar1.Top = Me.DataGrid1.Height
'End Sub

'
Private Sub ws_Close()
'�������� � ����������� ����������, �������������� ������
Me.Command4.Enabled = True
End Sub

Private Sub ws_Connect()
' �������� � ���������� ����������
Me.StatusBar1.SimpleText = "���������� ������ �������."
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim i As Long, a As Double, b As Double
Dim data As String
Dim data4 As String
Dim data2 As String
Dim dataX As String
Dim data8 As String
Dim n As Long, n1 As Long
On Error GoTo wsDA_err
n = 0: n1 = 0
ws.GetData data, vbString ' �������� ���, ��� ������ �� �������
data2 = Left(data, 4) ' �������� ������� �� �������
Select Case data2 ' ���������������� ��
    Case "rqst"  ' ������ �� �������� ����� � ������� �� �������
        ' ������������ ��� ����� � ������� �������� ������
        dataX = Right(data, Len(data) - (4))
        fsize = CLng(dataX) ' ��������� �����:
        If fsize > 0 Then ' ���� ���� �� ����
            data4 = "Tempbase.csv" ' ��� ����� � ������� �������� ������
            PrBar1.CustomCaption = "���� ��������..."
            PrBar1.Value = 1
            PrBar1.MaxValue = (fsize \ chunk + 1)
            ' �������� ���� �� ���������� �������
            Open data4 For Output As #1
            Print #1, ""
            Close #1
            ' ������� ���� ��� ����� �������
            Open data4 For Binary As #1
            ws.SendData "okay"      ' ��������� ������ �� ��������� �����
        Else ' ���� ���� ����, �� ������ ��������������:
            MsgBox "������ �����������, ��������� �������� �������"
            Me.StatusBar1.SimpleText = ""
        End If
    Case "/reg"
        ' ����� ������� ���������������
        ws.SendData "NICK " & Trim$(Mid(frmStart.Caption, _
            InStr(1, frmStart.Caption, "=", 1) + 1)): Exit Sub
    Case "/bad"
        MsgBox "������ ������� �����������, ��������� ��� ���"
        Me.StatusBar1.SimpleText = ""
    Case "serv" ' �������������� ������
        Me.StatusBar1.SimpleText = Right(data, Len(data) - (4)) ' ����� ����������� � ��������
    Case "node"
        ' ��������� ������ ����� ����
        dataX = Right(data, Len(data) - (4))
        fsize = Len(dataX) ' ��������� �����:
        If fsize > 0 Then ' ���� ���� �� ����
            n = InStr(1, dataX, "@")
            Do While n > 0
                data8 = Mid(dataX, 1, n - 1) ' �������� ��������� ����������
                dataX = Mid(dataX, n + 1)
                With DataEnvironment1.rsCommand1
                    .MoveFirst
                    ' ���� ����� ���� ��� ����, �� ������ ������ ���������
                    .Find "������� = '" & Left(data8, InStr(1, data8, ";", 1) - 1) & "'"
                    If .EOF Then ' ���� ���, �� ...
                        .AddNew ' ��������� ����� ���� � ���������
                        .Fields("�������") = Left(data8, InStr(1, data8, ";", 1) - 1)
                        .Fields("���������") = Mid(data8, InStr(1, data8, ";", 1) + 1)
                        .Update
                    Else
                        .Fields("���������") = Mid(data8, InStr(1, data8, ";", 1) + 1)
                    End If
                End With
                n = InStr(1, dataX, "@")
            Loop
        Else ' ���� ���� ����, �� ������ ��������������:
            MsgBox "������ �� ����� �� ��������"
            Me.StatusBar1.SimpleText = ""
        End If
        '
    Case Else
        size = size + 1 '  ������� ���������� ������
        sz = size * 8 'chunk ' ��������� ������ �����
        'PrBar1.Value = PrBar1.Value + PrBar1.MaxValue / (size) ' * 100)
        PrBar1.Value = PrBar1.Value + size * 10
        PrBar1.CustomCaption = "�������� " & sz & "Kb"
        Put #1, , data ' ���������� ���������� ������ � ����
        i = Seek(1)
        If i >= fsize Then
'            'Mid(data, InStr(1, data, "EnDf"), 4) = "   "
            Close #1 ' ������� ���� � ����������� �������
            ' �������� ����� ��� ����� �������������� ������
            If frmStart.mnuKT.Checked Then
                a = ReadNParam("KOEFA"): b = ReadNParam("KOEFB")
                Open "Tempbase.csv" For Input As #1
                Line Input #1, dataX ' ����������
                Line Input #1, data ' ����������
                dataX = dataX & vbCrLf & data
                Line Input #1, data ' ������������ ����� ���������
                dataX = dataX & vbCrLf & data & "t2r;dt;" & vbCrLf
                Do While Not EOF(1)   ' ���� �� ����� �����
                    Line Input #1, data
                    data2 = data ' ��������� ��� ����������
                    ' �������� ��������
                    For n = 1 To 9
                        data = Mid(data, InStr(1, data, ";", 1) + 1)
                    Next
                    data = Mid(data, 1, Len(data) - 1)
                    ' ����������� �� �������
                    data4 = Mid(data, 1, InStr(1, data, ";", 1) - 1)
                    ' ����������� �� �������
                    data8 = Mid(data, InStr(1, data, ";", 1) + 1)
                    ' ��������� ���� �� �������� ���-��
                    data = Val(data4) * a + b
                    data4 = Val(Trim(data8)) - data
                    dataX = dataX & data2 & Trim(RemakeS(str(data), True)) & _
                                    ";" & Trim(RemakeS(str(data4), True)) & ";" & vbCrLf
                    'dataX = dataX & data2 & Trim(str(data)) & _
                                    ";" & Trim(str(data4)) & ";" & vbCrLf
                Loop
                Close #1
                dataX = Mid(dataX, 1, Len(dataX) - 2)
                ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
                'dataX = RemakeS(dataX)
                TipOt = ""
                ' ���������� ���-� ��������������
                Kill "tempbase.csv"
                Open "tempbase.csv" For Append As #1
                Print #1, dataX ' ������ ���������
                Close #1
            End If
            Me.StatusBar1.SimpleText = "������ �������� � �������: " & sz & "Kb"
            size = 0: sz = 0
            ' ���������� ���-�
            ViewRez
'       Else
        End If
End Select
Exit Sub
wsDA_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub
' ���������� ���-� ���������� �������
Function ViewRez()
Dim fldXs As Column
On Error Resume Next
' ���������� ����� �������
If Len(TipOt) = 0 Then TipOt = RemakeHead ' �������������� ��������� � ������ �������
If TipOt = "�Journal" Or TipOt = "�Journal" Then ' ���� ������� �������� �������
    ' ��������� ���� �� �������� ���-��
    If Not frmStart.mnuKT.Checked Then
        With DataEnvironment1.rscmdTeploRez
            If .State <> adStateOpen Then .Open ' ����������� ������
            .Requery  ' �������� ������
        End With
        ' ��������� �����������
        Me.DataGrid1.DataMember = "cmdTeploRez"
        Me.DataGrid1.Refresh
    Else
        With DataEnvironment1.rscmdTeploRezT
            If .State <> adStateOpen Then .Open ' ����������� ������`
            .Requery  ' �������� ������
        End With
        ' ��������� �����������
        Me.DataGrid1.DataMember = "cmdTeploRezT"
        Me.DataGrid1.Refresh
    End If
ElseIf TipOt = "�HV" Then
    With DataEnvironment1.rsCommand5
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery  ' �������� ������
    End With
    ' ��������� �����������
    Me.DataGrid1.DataMember = "Command5"
    Me.DataGrid1.Refresh
ElseIf TipOt = "�HV" Then
    With DataEnvironment1.rsCommand4
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery  ' �������� ������
    End With
    ' ��������� �����������
    Me.DataGrid1.DataMember = "Command4"
    Me.DataGrid1.Refresh
ElseIf TipOt = "�PAR" Then
    With DataEnvironment1.rscmdParhRez
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery  ' �������� ������
    End With
    ' ��������� �����������
    Me.DataGrid1.DataMember = "cmdParhRez"
    Me.DataGrid1.Refresh
ElseIf TipOt = "�PAR" Then
    With DataEnvironment1.rscmdParsRez
        If .State <> adStateOpen Then .Open ' ����������� ������
        .Requery  ' �������� ������
    End With
    ' ��������� �����������
    Me.DataGrid1.DataMember = "cmdParsRez"
    Me.DataGrid1.Refresh
End If
' ��������� �������
For Each fldXs In Me.DataGrid1.Columns
    If fldXs.Caption = "����" Or _
        fldXs.Caption = "datetime" Then fldXs.Width = 104 _
        Else fldXs.Width = 47
Next
' �������� ������������ �������
If DataEnvironment1.rscmdTeploRez.State = adStateOpen Then _
        DataEnvironment1.rscmdTeploRez.Close ' ������� ������
If DataEnvironment1.rscmdTeploRezT.State = adStateOpen Then _
        DataEnvironment1.rscmdTeploRezT.Close ' ������� ������
If DataEnvironment1.rsCommand4.State = adStateOpen Then _
        DataEnvironment1.rsCommand4.Close ' ������� ������
If DataEnvironment1.rsCommand5.State = adStateOpen Then _
        DataEnvironment1.rsCommand5.Close ' ������� ������
If DataEnvironment1.rscmdParsRez.State = adStateOpen Then _
        DataEnvironment1.rscmdParsRez.Close ' ������� ������
If DataEnvironment1.rscmdParhRez.State = adStateOpen Then _
        DataEnvironment1.rscmdParhRez.Close ' ������� ������
' ������� ���� � �������������� ���������
Call RecovHead
TipOt = ""
End Function
'
Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                    ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
                    CancelDisplay As Boolean)
'��� ������ �������� � ���, ������� ����������, �������������� ������
Me.StatusBar1.SimpleText = "������ Winsock #" & Number & "-" & Description
ws.Close ' ������� ����������
Me.Command4.Enabled = True
End Sub
' ����� ������ ���� �� ������
Sub PrintTSRV(ZL As Integer)
Dim notopen As String, i As Integer
Dim pos1 As Long, pos2 As Long, dRepDate As Date
Dim t1 As Double, t2 As Double, tf As Double
' ZL=1 - ����, ZL=0 - ����
On Error GoTo excl
Call SetLocaleInfo(LOCALE_SDECIMAL, ".")
If ZL = 1 Then
    otchetTSRVs.Sections(1).Controls(10).Caption = "�-0"
    otchetTSRVs.Sections(1).Controls(14).Caption = "dG=m1+m2"
    otchetTSRVs.Sections(1).Controls(15).Caption = "W3=W1 + W2"
    'otchetTSRVs.Sections(1).Controls(16).Caption = "dt=t1+t2"
End If
otchetTSRVs.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
otchetTSRVs.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
otchetTSRVs.Sections(1).Controls("lblDogovor").Caption = Dialog.txtDogov.Text
If DataEnvironment1.rsCommand19.State <> adStateOpen Then _
DataEnvironment1.rsCommand19.Open ' ������� ������ � tembase.csv ��� ������
DataEnvironment1.rsCommand19.Requery ' �������� ������
'dRepDate = Date
dRepDate = DataEnvironment1.rsCommand19.Fields(0)
otchetTSRVs.Sections(1).Controls(4).Caption = mon(Month(dRepDate)) & " " & Year(dRepDate) & " �."
notopen = DBGrid1.Columns(2).Value
otchetTSRVs.Sections(1).Controls(8).Caption = notopen
notopen = DBGrid1.Columns(3).Value
otchetTSRVs.Sections(1).Controls(9).Caption = notopen
If DataEnvironment1.rsCommand20.State <> adStateOpen Then _
    DataEnvironment1.rsCommand20.Open ' ������� ������ � tembase.csv ��� ������
' ���������� ������ ���������� ����� �� ������
'otchetTSRVs.Orientation = rptOrientLandscape
' ��������� �����
With DataEnvironment1.rsCommand20
    .Requery
    For i = 0 To 15
        otchetTSRVs.Sections(5).Controls("L" & Trim(str(i))).Caption = .Fields(i)
    Next
End With
' ���������� ��������� ��������� 1 ���.
pos1 = InStr(1, strTSRV1, "���������=", vbTextCompare) + 10
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblDV1").Caption = Mid(strTSRV1, pos1, pos2 - pos1 - 2)
pos1 = InStr(1, strTSRV1, "m1=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblm1_1").Caption = _
            Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)), "0.000")
pos1 = InStr(1, strTSRV1, "m2=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblm2_1").Caption = _
            Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)), "0.000")
pos1 = InStr(1, strTSRV1, "W1=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblW1_1").Caption = _
            Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)) / 4187, "0.000")
pos1 = InStr(1, strTSRV1, "W2=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblW2_1").Caption = _
            Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)) / 4187, "0.000")
pos1 = InStr(1, strTSRV1, "T��=", vbTextCompare)
If pos1 = 0 Then pos1 = InStr(1, strTSRV1, "T��4=", vbTextCompare) + 5 _
Else pos1 = pos1 + 4
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
t1 = Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)) / 60, "0.00")
otchetTSRVs.Sections(5).Controls("lblV1").Caption = t1
' ���������� ��������� ��������� 2 ���.
pos1 = InStr(1, strTSRV2, "���������=", vbTextCompare) + 10
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblDV2").Caption = Mid(strTSRV2, pos1, pos2 - pos1 - 2)
pos1 = InStr(1, strTSRV2, "m1=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblm1_2").Caption = _
            Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)), "0.000")
pos1 = InStr(1, strTSRV2, "m2=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblm2_2").Caption = _
            Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)), "0.000")
pos1 = InStr(1, strTSRV2, "W1=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblW1_2").Caption = _
            Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)) / 4187, "0.000")
pos1 = InStr(1, strTSRV2, "W2=", vbTextCompare) + 3
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
otchetTSRVs.Sections(5).Controls("lblW2_2").Caption = _
            Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)) / 4187, "0.000")
pos1 = InStr(1, strTSRV2, "T��=", vbTextCompare)
If pos1 = 0 Then pos1 = InStr(1, strTSRV2, "T��4=", vbTextCompare) + 5 _
Else pos1 = pos1 + 4
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
t2 = Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)) / 60, "0.00")
otchetTSRVs.Sections(5).Controls("lblV2").Caption = t2
t1 = Format(t2 - t1, "0.00")
otchetTSRVs.Sections(5).Controls("lblRV").Caption = t1
' ���������� ������� ������
With DataEnvironment1.rsCommand11
    If .State <> adStateOpen Then .Open ' ����������� ������
    otchetTSRVs.BottomMargin = .Fields("Niz") * Stwips
    otchetTSRVs.TopMargin = .Fields("Verh") * Stwips
    otchetTSRVs.LeftMargin = .Fields("Levo") * Stwips
    otchetTSRVs.RightMargin = .Fields("Pravo") * Stwips
    otchetTSRVs.Font.size = .Fields("Shrift")
End With
otchetTSRVs.Show ' ������������ � ������
Call SetLocaleInfo(LOCALE_SDECIMAL, ",")
Exit Sub
excl:
 If Err.Number = 13 Then
    notopen = ""
    Resume Next
 End If
End Sub
'
Sub RecovHead()
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String, x As Long
Dim priznak As String
On Error GoTo errRecHead
' ������� ����
filenum = FreeFile ' ���������� c��������� �����
Open "tempbase.csv" For Input As #filenum ' ������� ����
' ��������� ������ �������� ��� ����
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' ������� ����
filenum = FreeFile ' ������� ���� ��� ������
strY = strTSRV1 & vbCrLf & strTSRV2 & vbCrLf & strY
Open "tempbase.csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' ��������� ����
Close #filenum
Exit Sub
errRecHead:
Resume Next
End Sub
'
'
Private Sub Command3_Click()
Dim i As Long
On Error GoTo Print_err
If Len(TipOt) = 0 Then TipOt = RemakeHead ' �������������� ��������� � ������ �������
Select Case TipOt
Case "�Journal", "�Journal" ' ���� ������� �������� �������
    If Dialog.Check3.Value And Left(TipOt, 1) = "�" Then
        Call PrintTSRV(Me.Check1.Value) '������ ����
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
Call RecovHead
TipOt = ""
Exit Sub
Print_err:
 'MsgBox Err.Number & "->" & Err.Description
 Resume Next
End Sub
'������� ��������������� ���������� ����������� �����
Function RemakeHead() As String
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String, x As Long
Dim priznak As String
On Error GoTo errRHead
' ������� ����
filenum = FreeFile ' ���������� c��������� �����
Open "tempbase.csv" For Input As #filenum ' ������� ����
' ���������� ������ � ������ ������
Line Input #filenum, strTSRV1 '
Line Input #filenum, strTSRV2 '
' ������� ��� ������
x = InStr(1, strTSRV1, "�������", 1)
If x = 0 Then
    x = InStr(1, strTSRV1, "��������", 1)
    priznak = "�"
Else
    priznak = "�"
End If
priznak = priznak & Mid(strTSRV1, 1, x - 1)
' �������������� ������ ���������
Line Input #filenum, strX '
'strY = "���������;W1;W2;m1;m2;T��;P1;P2;W3;V3;����;t1;t2" & vbCrLf
If InStr(1, strX, ";T��4", 1) > 0 Then
    strY = Mid(strX, 1, InStr(1, strX, ";T��4", 1)) & "T��" & _
                Mid(strX, InStr(1, strX, ";T��4", 1) + 5) & vbCrLf
ElseIf InStr(1, strX, ";��", 1) > 0 Then
    strY = Mid(strX, 1, InStr(1, strX, ";��", 1)) & "T��" & _
                Mid(strX, InStr(1, strX, ";��", 1) + 3) & vbCrLf
Else
    strY = strX & vbCrLf
End If
' ��������� ������ �������� ��� ����
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   ' ���������� �� ������� ����������
   'strX = Milliard(strX)
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' ������� ����
filenum = FreeFile ' ������� ���� ��� ������
Open "tempbase.csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' ��������� ����
Close #filenum
RemakeHead = priznak
Exit Function
errRHead:
Resume Next
End Function
' ������������� ��������� �������� � ����������
Function Milliard(strMilli As String) As String
Dim stra(6) As String, tr As Boolean
Dim pos As Long, pos1 As Long, i As Long
Dim str1 As String
pos = InStr(1, strMilli, ";", 1) ' ���� ������ ���������
tr = False
Do While pos
    i = i + 1
    If i >= 2 Then
        stra(i - 2) = Mid(strMilli, pos1 + 1, pos - pos1 - 1) '
        If i = 8 Then Exit Do
    End If
    pos1 = pos
    pos = InStr(pos1 + 1, strMilli, ";", 1) ' ������ ���������
Loop
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(0))) = -1 Then
    If CDbl(stra(0)) * -1 > 1 Then stra(0) = Format(1000000000 / 4187 + stra(0), "0.000"): tr = True
End If
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(1))) = -1 Then
    If CDbl(stra(5)) > 1000000000 / 4187 Then
        stra(5) = Format(CDbl(stra(5)) - 1000000000 / 4187, "0.000"): tr = True
    End If
    If CDbl(stra(1)) * -1 > 1 Then stra(1) = Format(1000000000 / 4187 + stra(1), "0.000"): tr = True
End If
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(2))) = -1 Then
    If CDbl(stra(2)) * -1 > 1 Then stra(2) = Format(1000000 + stra(2), "0.000"): tr = True
End If
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(3))) = -1 Then
    If CDbl(stra(6)) > 1000000 Then
        stra(6) = Format(CDbl(stra(6)) - 1000000, "0.000"): tr = True
    End If
    If CDbl(stra(3)) * -1 > 1 Then stra(3) = Format(1000000 + stra(3), "0.000"): tr = True
End If
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(5))) = -1 Then
    If CDbl(stra(5)) * -1 > 1 Then stra(5) = Format(1000000000 / 4187 + stra(5), "0.000"): tr = True
End If
If CDbl(stra(4)) = 0 And Sgn(CDbl(stra(6))) = -1 Then
    If CDbl(stra(6)) * -1 > 1 Then stra(6) = Format(1000000 + stra(6), "0.000"): tr = True
End If
'
If tr = True Then
    str1 = NewPos(strMilli, 2): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(0) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 3): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(1) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 4): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(2) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 5): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(3) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 6): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(4) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 7): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(5) & Mid(strMilli, pos)
    str1 = NewPos(strMilli, 8): pos1 = Mid(str1, InStr(1, str1, "/", 1) + 1)
    pos = Mid(str1, 1, InStr(1, str1, "/", 1) - 1)
    strMilli = Mid(strMilli, 1, pos1) & stra(6) & Mid(strMilli, pos)
End If
Milliard = strMilli
End Function
' ��������� ����� �������
Function NewPos(strQ As String, num As Long) As String
Dim pos As Long, pos1 As Long, i As Long
i = 0
pos = InStr(1, strQ, ";", 1) ' ���� ������ ���������
Do While pos
    i = i + 1
    If i = num Then Exit Do
    pos1 = pos
    pos = InStr(pos1 + 1, strQ, ";", 1) ' ������ ���������
Loop
NewPos = CStr(pos) & "/" & CStr(pos1)
End Function
'
Private Sub DbGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If Not IsNull(LastRow) Then
On Error GoTo exit_DBG
    'Text6.Text = DBGrid1.Text ' �������� ������������ ���� �����
'End If
exit_DBG:
End Sub
'
Private Sub Form_Unload(Cancel As Integer) '��� ������ �� ���������:
    DataEnvironment1.rsCommand2.Close ' ������� ������� ��������
    ' =1 ��� ������ ������ ������������,
    ' =2 - ������ ��������������, =0 - �� ��������.
    Form1.WindowState = 0
    Cancel = False
End Sub
'
Private Sub Form_load()
    Dim Htenie
    Dim z As String, L As Long
    Dim strDBPath As String
    Dim rstTxt As Recordset
    Dim tmpPort As String, tmpServ As String
    On Error GoTo Fload_err
    'BDCmpct
    TipOt = ""
    strDBPath = App.Path & "\settings.mdb"
    '��������� ������� �������
    With DataEnvironment1.rsCommand2 ' ����� ���������
        If .State <> adStateOpen Then .Open
        .Requery
        .MoveFirst
        Do While Not .EOF
            Select Case .Fields("NameSet") ' ����� �������� ��������� ���������
            Case "PathAdmin"
                tmpServ = .Fields("Set") ' ����� �������� ������� ����������
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
            Case "Port"
                ' ����� ����� �����
                tmpPort = .Fields("Set")
            End Select
            .MoveNext ' �� ���� ���������� ���������
        Loop
    End With
    'Form1.Text2 = Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1)
    Dialog.Text1.Text = tmpPort ' ����������� ���������
    Dialog.Text2.Text = tmpServ 'Text1.Text ' ����������� ����.�������
    ' ������� ������ ��������� ������:
    ' �� ������ ������
    Text4.Text = CDate("01/" & Month(Date) & "/" & Year(Date))
    Text5.Text = Date ' �� ������� ����
    blTP = False
    'Me.Text6.Text = Me.DBGrid1.Text ' ������� ��� �������� ����
    Call Command4_Click ' ���������� ����� �����
    Me.DTPicker1.Value = Date
    Me.DTPicker2.Value = Date
    Me.DataCombo1.Text = Me.DBGrid1.Text
    Exit Sub
Fload_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
    '
End Sub
'
Private Sub Text1_Change()
On Error GoTo TC1_err
' ���������� � ������  ����������
'Call WriteParameters("PathAdmin", Text1.Text)
Exit Sub
TC1_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub
'
'������ ������ � ���� events.log � ������� ���������� ����
Public Sub writeLog(Text As String)
    Dim logFile As String
    Dim FileNr As Integer
    '���������� ��� ������� ����
    '�������� � �����-��������� ���� � ������� ����
    logFile = CurDir & "\events.log"
    '������� ����-��������
    FileNr = FreeFile:    Open logFile For Append As FileNr
    '�������� ��������� �������
    Print #FileNr, Format(Now, "dd.mm.yy hh:nn:ss ") & " : "; Text
    '������� ����-��������
    Close FileNr
End Sub
'�������� ������ � ������� "������"
Public Function protocol() As String
    protocol = Err.Number & " # " & _
    Left(Err.Description, 200) & " # " & _
    Err.LastDllError & " # " & _
    Err.Source
End Function
'
Private Sub Timer1_Timer()
If blTP Then PrBar1.Value = PrBar1.Value + 1
End Sub
'
' ������� �������� ������� ����������� Microsoft office
Public Function GetRegString(hKey As Long, _
    strSubKey As String, strValueName As String) As String
Dim strSetting As String
Dim lngDataLen As Long
Dim lngRes As Long
If RegOpenKey(hKey, strSubKey, _
    lngRes) = ERROR_SUCCESS Then
    strSetting = Space(255)
    lngDataLen = Len(strSetting)
    If RegQueryValueEx(lngRes, _
        strValueName, ByVal 0, _
        REG_EXPAND_SZ, ByVal strSetting, _
        lngDataLen) = ERROR_SUCCESS Then
        If lngDataLen > 1 Then
            GetRegString = Left(strSetting, lngDataLen - 1)
        End If
    End If
    If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
        MsgBox "RegCloseKey Failed: " & _
        strSubKey, vbCritical
    End If
End If
End Function

Function FileExists(sFileName$) As Boolean
On Error Resume Next
FileExists = IIf(Dir(Trim(sFileName)) <> "", _
True, False)
End Function

Public Function IsAppPresent(strSubKey$, _
strValueName$) As Boolean
IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, _
strSubKey, strValueName)))
End Function
' ��������������� ������� �������� � ����������� MS Office
Private Sub WhereMSOffice()
MsgBox "Access " & _
IsAppPresent("Access.Database\CurVer", "")
MsgBox "Excel " & _
IsAppPresent("Excel.Sheet\CurVer", "")
MsgBox "PowerPoint " & _
IsAppPresent("PowerPoint.Slide\CurVer", "")
MsgBox "Word " & _
IsAppPresent("Word.Document\CurVer", "")
End Sub
'
' ������ ��
'Sub BDCmpct()
'On Error GoTo ercmpct
'DBEngine.CompactDatabase "settings.mdb", "setting.mdb", dbLangCyrillic & ";pwd=MTWTFSS", , ";pwd=MTWTFSS" '
'Kill "settings.mdb"
'Name "setting.mdb" As "settings.mdb"
'Exit Sub
'ercmpct:
' writeLog ("compact DB: " & protocol())
'End Sub
