VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A237EE18-33EE-468A-B4D8-07559BD2E396}#5.0#0"; "ProgBar.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Form1 
   Caption         =   "Получение данных"
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
         ListField       =   "ИмяУзла"
         Text            =   ""
         Object.DataMember      =   "Command1"
      End
      Begin VB.CommandButton Command1 
         Height          =   345
         Left            =   705
         Picture         =   "Form1.frx":08EA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Состояние узлов (F4)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnSave 
         Height          =   345
         Left            =   1380
         Picture         =   "Form1.frx":172C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Сохранить данные в файл"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton btnLoad 
         Height          =   345
         Left            =   1740
         Picture         =   "Form1.frx":1A32
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Загрузить данные из файла"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Height          =   345
         Left            =   2100
         Picture         =   "Form1.frx":1D3C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Показать отчет"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command2 
         Height          =   345
         Left            =   1050
         Picture         =   "Form1.frx":2606
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Получить данные (F5)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdUnload 
         Height          =   345
         Left            =   0
         Picture         =   "Form1.frx":4300
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Выход (F2)"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Height          =   345
         Left            =   360
         Picture         =   "Form1.frx":4642
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Соединиться (F3)"
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
      DialogTitle     =   "Выберите архив"
      Filter          =   "Формат с разделителем  (*.csv)|*.csv"
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
         DataField       =   "ИмяУзла"
         Caption         =   "ИмяУзла"
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
         DataField       =   "Результат"
         Caption         =   "Результат"
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
         DataField       =   "Метка"
         Caption         =   "Метка"
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
         Caption         =   "Прибор"
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
         DataField       =   "ТипАрхива"
         Caption         =   "Номер"
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
'Dim mcnnUnits As DAO.Database ' подконнектиться базам
'Dim mrstMain As DAO.Recordset  ' подключение к служебной таблице
'Dim mrstUnits As DAO.Recordset ' подключение к архиву узла
Dim blTP As Boolean
'
Dim sz As Long ' размер файла вычисленный
Dim size As Integer ' количество блоков файла по 8 Кб
Dim fsize As Long ' размер передаваемого файла
Dim nameNode As String ' имя узла для передачи
Dim strTSRV1 As String ' для заполнения 1ой строки отчета ТСРВ
Dim strTSRV2 As String ' для заполнения 2ой строки отчета ТСРВ
Dim TipOt As String
Private Const chunk = 8000
Private Const Stwips = 537

Private Sub btnLoad_Click()
Dim filenum ' имя файла для сохранения данных
Dim strX As String, strY As String
Dim pos As Long
On Error Resume Next
CDArh.CancelError = True
CDArh.ShowOpen ' открыть файл
filenum = FreeFile: filenum = filenum - 1
Close #filenum
If CDArh.FileName <> "" Then
    If OpenCSV(CDArh.FileName) Then MsgBox "Файл загружен" _
    Else MsgBox "Файл не загружен"
End If
End Sub

Private Sub btnSave_Click()
Dim NFSO As New FileSystemObject
Dim filenum ' имя файла для сохранения данных
Dim strX As String, strY As String, i As Integer
Dim pos As Long
On Error GoTo errbtnSave_Click
' открыть файл
filenum = FreeFile ' дескриптор cвободного файла
Open "tempbase.csv" For Input As #filenum ' открыть файл
' пропустить первую и вторую строку
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
' корректировать строку заголовка
Line Input #filenum, strX '
pos = InStr(1, strX, "Формула1", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "W3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула2", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "m3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула3", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "To;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула4", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t1;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула5", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t2;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула6", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "t3;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула7", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "P1;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула8", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "P2;" & Mid(strX, pos + 9)
pos = InStr(1, strX, "Формула9", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "TНСч;" & Mid(strX, pos + 9)
strY = strY & strX & vbCrLf
' остальные строки записать как есть
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' закрыть файл
' сохранить данные на диск
CDArh.CancelError = True
CDArh.FileName = "Куда сохранить"
CDArh.ShowSave ' открыть файл
ChDir App.Path
filenum = FreeFile ' открыть файл для записи
Open Mid(CDArh.FileName, 1, Len(CDArh.FileName) - 3) & "csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' заполнить файл
Close #filenum
' добавить итоговую строку
TipOt = ""
Call AppResult(Mid(CDArh.FileName, 1, Len(CDArh.FileName) - 3) & "csv")
'Exit Sub
TipOt = ""
errbtnSave_Click:
ChDir App.Path
End Sub
' функция добавляющая итоговую строку
Function AppResult(NofFile As String)
Dim FileNumber As Long
Dim sHead As String
On Error Resume Next
If Len(TipOt) = 0 Then TipOt = RemakeHead ' преобразование заголовка к новому формату
sHead = "ИТОГО"
' получить код файла архива
Select Case TipOt
Case "сJournal", "чJournal" ' учет расхода тепловой энергии
    'PrintVis
    If Not frmStart.mnuKT.Checked Then
        If DataEnvironment1.rsCommand8.State <> adStateOpen Then _
            DataEnvironment1.rsCommand8.Open ' открыть запрос к tembase.csv для работы
        ' формируем Итого
        With DataEnvironment1.rsCommand8
            .Requery
            sHead = sHead & ";" & Format(.Fields(0), "0.00") ' теп.энерг. W1
            sHead = sHead & ";" & Format(.Fields(3), "0.00") ' теп.энерг. W2
            sHead = sHead & ";" & Format(.Fields(1), "0.00") ' масса m1
            sHead = sHead & ";" & Format(.Fields(4), "0.00") ' масса m2
            sHead = sHead & ";" & str(.Fields(9)) ' время остан.
            sHead = sHead & ";" & Format(.Fields(6), "0.00") ' теп.энерг. W3
            sHead = sHead & ";" & Format(.Fields(7), "0.00") ' масса m3
            sHead = sHead & ";" & str(.Fields(8))  ' время нараб.
            sHead = sHead & ";" & Format(.Fields(2), "0.00") 'темпер. t1
            sHead = sHead & ";" & Format(.Fields(5), "0.00") 'темп-ра t2
            sHead = sHead & ";;" & Format(.Fields(10), "0.00") 'давл. по прям. d1
            sHead = sHead & ";" & Format(.Fields(11), "0.00") 'давл.по обрат. d2
        End With
    Else
        If DataEnvironment1.rsCommand18.State <> adStateOpen Then _
            DataEnvironment1.rsCommand18.Open ' открыть запрос к tembase.csv для работы
        ' формируем Итого
        With DataEnvironment1.rsCommand18
            .Requery
            sHead = sHead & ";" & Format(.Fields(0), "0.00") ' теп.энерг. W1
            sHead = sHead & ";" & Format(.Fields(1), "0.00") ' теп.энерг. W2
            sHead = sHead & ";" & Format(.Fields(2), "0.00") ' масса m1
            sHead = sHead & ";" & Format(.Fields(3), "0.00") ' масса m2
            sHead = sHead & ";" & str(.Fields(4))  ' время остан.
            sHead = sHead & ";" & Format(.Fields(7), "0.00") ' теп.энерг. W3
            sHead = sHead & ";" & Format(.Fields(8), "0.00") ' масса m3
            sHead = sHead & ";" & str(.Fields(9)) ' время работы
            sHead = sHead & ";" & Format(.Fields(10), "0.00") ' темпер. t1
            sHead = sHead & ";" & Format(.Fields(11), "0.00") ' темп-ра t2
            sHead = sHead & ";;" & Format(.Fields(5), "0.00") ' давл. по прям. d1
            sHead = sHead & ";" & Format(.Fields(6), "0.00") ' давл.по обрат. d2
            sHead = sHead & ";;" & Format(.Fields(12), "0.00") ' тем-ра обратки расч.
            sHead = sHead & ";" & Format(.Fields(13), "0.00") ' темпер-ное отклонение
        End With
    End If
Case "сHV", "чHV"
'    холодная вода без итоговой строки
Case "сPAR"
    ' пар не зависит от температурного контроля
    'sHead = sHead & ";"
    If DataEnvironment1.rsCommand9.State <> adStateOpen Then _
        DataEnvironment1.rsCommand9.Open ' открыть запрос к tembase.csv для работы
    With DataEnvironment1.rsCommand9
        .Requery
        sHead = sHead & ";" & Format(.Fields(0), "0.00") ' темпер. t1
        sHead = sHead & ";" & Format(.Fields(1), "0.00") ' давлен d1
        sHead = sHead & ";" & Format(.Fields(2), "0.00") 'масса V1
        sHead = sHead & ";" & Format(.Fields(3), "0.00") ' теп.энерг. W1
        sHead = sHead & ";" & Format(.Fields(4), "0.0") ' Tr1
        sHead = sHead & ";" & Format(.Fields(6), "0.00") ' темп-ра t2
        sHead = sHead & ";" & Format(.Fields(7), "0.00") ' давлен. d2
        sHead = sHead & ";" & Format(.Fields(8), "0.00") ' масса V2
        sHead = sHead & ";" & Format(.Fields(9), "0.00") ' теп.энерг. W2
        sHead = sHead & ";" & Format(.Fields(10), "0.0") ' Tr2
        'sHead = sHead & ";" & Format(.Fields(5), "0.0") ' To1
        'sHead = sHead & ";" & Format(.Fields(11), "0.0") ' To2
    End With
Case "чPAR"
    ' пар не зависит от температурного контроля
    'sHead = sHead & ";"
    If DataEnvironment1.rsCommand10.State <> adStateOpen Then _
        DataEnvironment1.rsCommand10.Open ' открыть запрос к tembase.csv для работы
    With DataEnvironment1.rsCommand10
        .Requery
        sHead = sHead & ";" & Format(.Fields(0), "0.00") ' темпер. t1
        sHead = sHead & ";" & Format(.Fields(1), "0.00") ' давлен d1
        sHead = sHead & ";" & Format(.Fields(2), "0.00") 'масса V1
        sHead = sHead & ";" & Format(.Fields(3), "0.00") ' теп.энерг. W1
        sHead = sHead & ";" & Format(.Fields(4), "0.0") ' Tr1
        sHead = sHead & ";" & Format(.Fields(6), "0.00") ' темп-ра t2
        sHead = sHead & ";" & Format(.Fields(7), "0.00") ' давлен. d2
        sHead = sHead & ";" & Format(.Fields(8), "0.00") ' масса V2
        sHead = sHead & ";" & Format(.Fields(9), "0.00") ' теп.энерг. W2
        sHead = sHead & ";" & Format(.Fields(10), "0.0") ' Tr2
        'sHead = sHead & ";" & Format(.Fields(5), "0.0") ' To1
        'sHead = sHead & ";" & Format(.Fields(11), "0.0") ' To2
    End With
End Select
' открыть файл
FileNumber = FreeFile
Open NofFile For Append As #FileNumber
Print #FileNumber, sHead ' брать построчно
Close #FileNumber
End Function
'

Private Sub Check1_Click()
On Error GoTo Check1_err
' смена режима получения архива
If Check1.Value = 1 Then Check1.Caption = "Режим 'Лето'" _
Else Check1.Caption = "Режим 'Зима'"
Exit Sub
Check1_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub

'
'
Private Sub cmdUnload_Click()
'If MsgBox("Действительно выйти?", vbQuestion + vbYesNo) = vbYes Then
    'Unload frmLogin
    ' выгрузить форму
    Unload Me
'End If
End Sub
'
'
Private Sub Combo1_LostFocus()
    ' сохранить тип архива
    Call WriteParameters("KindArchive", Combo1.Text)
End Sub

Private Sub Command1_Click()
Dim strD As String, data8 As String
If MsgBox("Обновить список узлов?", vbYesNo) = vbYes Then
    Me.DBGrid1.Visible = True
    Me.DataGrid1.Visible = False
    ' получить список узлов учета
    Me.StatusBar1.SimpleText = "Получение описания узлов учета"
    strD = "node(" & Mid(frmStart.Caption, _
            InStr(1, frmStart.Caption, "=", 1) + 1) & ")" ' запросить список
    ' отправить запрос
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
' предупредить о получении данных
Me.StatusBar1.SimpleText = "Идет получение данных"
If Combo1.Text = "Часовой" Then n = 2 Else n = 3
TipOt = ""
' сформировать запрос
nameNode = Me.DataCombo1.Text  'Text6.Text
' запросить данные по выбранному узлу учета
strD = "/get[" & Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1) & "][" & nameNode & _
        "][" & Text4.Text & "][" & Text5.Text & "][" & _
             Combo1.Text & "][" & Trim(str(Check1.Value)) & "]"
ws.SendData strD ' отправить запрос
End Sub

Private Sub Command4_Click()
'закроем на всякий случай, если соединение уже открыто
Me.StatusBar1.SimpleText = ""
ws.Close
'соединяемся с сервером по порту 1001
'предварительно убрав лишние пробелы
ws.Connect Dialog.Text2.Text, Dialog.Text1.Text  ' 1001
'делаем так, чтобы пользователь не смог второй раз нажать
'кнопку соединения потому что при этом выскочит ошибка
Me.Command4.Enabled = False
End Sub


Private Sub DBGrid1_DblClick()
' установить признаки для узлов
If Len(Trim(DBGrid1.Columns(4))) = 0 Then
    DBGrid1.Columns(4) = "пр" ' признак узла учета пара
ElseIf StrComp(DBGrid1.Columns(4), "пр") = 0 Then
    DBGrid1.Columns(4) = "хв" ' признак узла учет хол.воды
ElseIf StrComp(DBGrid1.Columns(4), "хв") = 0 Then
    DBGrid1.Columns(4) = "" ' признак узла учета тепла
End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
' запретить править названия узлов вручную
If DBGrid1.Col = 0 Then
    Me.StatusBar1.SimpleText = "Так нельзя"
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
'сообщить о завершенном соединении, разблокировать кнопку
Me.Command4.Enabled = True
End Sub

Private Sub ws_Connect()
' сообщить о выполнении соединения
Me.StatusBar1.SimpleText = "Соединение прошло успешно."
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
ws.GetData data, vbString ' получить все, что пришло от сервера
data2 = Left(data, 4) ' выделить команду от сервера
Select Case data2 ' проанализировать ее
    Case "rqst"  ' запрос на передачу файла с данными от сервера
        ' сформировать имя файла в который получать данные
        dataX = Right(data, Len(data) - (4))
        fsize = CLng(dataX) ' получение файла:
        If fsize > 0 Then ' если файл не пуст
            data4 = "Tempbase.csv" ' имя файла в который получить данные
            PrBar1.CustomCaption = "Идет загрузка..."
            PrBar1.Value = 1
            PrBar1.MaxValue = (fsize \ chunk + 1)
            ' очистить файл от предыдущих записей
            Open data4 For Output As #1
            Print #1, ""
            Close #1
            ' открыть файл для новых записей
            Open data4 For Binary As #1
            ws.SendData "okay"      ' отправить запрос на получение файла
        Else ' если файл пуст, то выдать предупреждение:
            MsgBox "Запрос некорректен, проверьте значения запроса"
            Me.StatusBar1.SimpleText = ""
        End If
    Case "/reg"
        ' клиен успешно зарегистрирован
        ws.SendData "NICK " & Trim$(Mid(frmStart.Caption, _
            InStr(1, frmStart.Caption, "=", 1) + 1)): Exit Sub
    Case "/bad"
        MsgBox "Запрос передан некорректно, повторите еще раз"
        Me.StatusBar1.SimpleText = ""
    Case "serv" ' дополнительный сервис
        Me.StatusBar1.SimpleText = Right(data, Len(data) - (4)) ' обмен сообщениями с сервером
    Case "node"
        ' получение данных через файл
        dataX = Right(data, Len(data) - (4))
        fsize = Len(dataX) ' получение файла:
        If fsize > 0 Then ' если файл не пуст
            n = InStr(1, dataX, "@")
            Do While n > 0
                data8 = Mid(dataX, 1, n - 1) ' отделяем служебную информацию
                dataX = Mid(dataX, n + 1)
                With DataEnvironment1.rsCommand1
                    .MoveFirst
                    ' если такой узел уже есть, то ставим только результат
                    .Find "ИмяУзла = '" & Left(data8, InStr(1, data8, ";", 1) - 1) & "'"
                    If .EOF Then ' если нет, то ...
                        .AddNew ' добавляем новый узел и результат
                        .Fields("ИмяУзла") = Left(data8, InStr(1, data8, ";", 1) - 1)
                        .Fields("Результат") = Mid(data8, InStr(1, data8, ";", 1) + 1)
                        .Update
                    Else
                        .Fields("Результат") = Mid(data8, InStr(1, data8, ";", 1) + 1)
                    End If
                End With
                n = InStr(1, dataX, "@")
            Loop
        Else ' если файл пуст, то выдать предупреждение:
            MsgBox "Данные по узлам не получены"
            Me.StatusBar1.SimpleText = ""
        End If
        '
    Case Else
        size = size + 1 '  считать количество блоков
        sz = size * 8 'chunk ' вычислять размер файла
        'PrBar1.Value = PrBar1.Value + PrBar1.MaxValue / (size) ' * 100)
        PrBar1.Value = PrBar1.Value + size * 10
        PrBar1.CustomCaption = "Получено " & sz & "Kb"
        Put #1, , data ' записывать полученные данные в файл
        i = Seek(1)
        If i >= fsize Then
'            'Mid(data, InStr(1, data, "EnDf"), 4) = "   "
            Close #1 ' закрыть файл с полученными данными
            ' добавить часть для учета температурного режима
            If frmStart.mnuKT.Checked Then
                a = ReadNParam("KOEFA"): b = ReadNParam("KOEFB")
                Open "Tempbase.csv" For Input As #1
                Line Input #1, dataX ' пропустить
                Line Input #1, data ' пропустить
                dataX = dataX & vbCrLf & data
                Line Input #1, data ' сформировать новый заголовок
                dataX = dataX & vbCrLf & data & "t2r;dt;" & vbCrLf
                Do While Not EOF(1)   ' пока не конец файла
                    Line Input #1, data
                    data2 = data ' сохраняем для добавления
                    ' отрезаем ненужное
                    For n = 1 To 9
                        data = Mid(data, InStr(1, data, ";", 1) + 1)
                    Next
                    data = Mid(data, 1, Len(data) - 1)
                    ' температура по прямому
                    data4 = Mid(data, 1, InStr(1, data, ";", 1) - 1)
                    ' температура по обратке
                    data8 = Mid(data, InStr(1, data, ";", 1) + 1)
                    ' проверить есть ли контроль тем-ры
                    data = Val(data4) * a + b
                    data4 = Val(Trim(data8)) - data
                    dataX = dataX & data2 & Trim(RemakeS(str(data), True)) & _
                                    ";" & Trim(RemakeS(str(data4), True)) & ";" & vbCrLf
                    'dataX = dataX & data2 & Trim(str(data)) & _
                                    ";" & Trim(str(data4)) & ";" & vbCrLf
                Loop
                Close #1
                dataX = Mid(dataX, 1, Len(dataX) - 2)
                ' предварительная подготовка строки к записи (меняем [,] на [.])
                'dataX = RemakeS(dataX)
                TipOt = ""
                ' записываем рез-т преобразования
                Kill "tempbase.csv"
                Open "tempbase.csv" For Append As #1
                Print #1, dataX ' писать построчно
                Close #1
            End If
            Me.StatusBar1.SimpleText = "Данные получены в размере: " & sz & "Kb"
            size = 0: sz = 0
            ' отобразить рез-т
            ViewRez
'       Else
        End If
End Select
Exit Sub
wsDA_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub
' отображает рез-т выполнения запроса
Function ViewRez()
Dim fldXs As Column
On Error Resume Next
' обработать шапку запроса
If Len(TipOt) = 0 Then TipOt = RemakeHead ' преобразование заголовка к новому формату
If TipOt = "сJournal" Or TipOt = "чJournal" Then ' учет расхода тепловой энергии
    ' проверить есть ли контроль тем-ры
    If Not frmStart.mnuKT.Checked Then
        With DataEnvironment1.rscmdTeploRez
            If .State <> adStateOpen Then .Open ' переоткрыть запрос
            .Requery  ' обновить данные
        End With
        ' выполнить отображение
        Me.DataGrid1.DataMember = "cmdTeploRez"
        Me.DataGrid1.Refresh
    Else
        With DataEnvironment1.rscmdTeploRezT
            If .State <> adStateOpen Then .Open ' переоткрыть запрос`
            .Requery  ' обновить данные
        End With
        ' выполнить отображение
        Me.DataGrid1.DataMember = "cmdTeploRezT"
        Me.DataGrid1.Refresh
    End If
ElseIf TipOt = "чHV" Then
    With DataEnvironment1.rsCommand5
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery  ' обновить данные
    End With
    ' выполнить отображение
    Me.DataGrid1.DataMember = "Command5"
    Me.DataGrid1.Refresh
ElseIf TipOt = "сHV" Then
    With DataEnvironment1.rsCommand4
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery  ' обновить данные
    End With
    ' выполнить отображение
    Me.DataGrid1.DataMember = "Command4"
    Me.DataGrid1.Refresh
ElseIf TipOt = "чPAR" Then
    With DataEnvironment1.rscmdParhRez
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery  ' обновить данные
    End With
    ' выполнить отображение
    Me.DataGrid1.DataMember = "cmdParhRez"
    Me.DataGrid1.Refresh
ElseIf TipOt = "сPAR" Then
    With DataEnvironment1.rscmdParsRez
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery  ' обновить данные
    End With
    ' выполнить отображение
    Me.DataGrid1.DataMember = "cmdParsRez"
    Me.DataGrid1.Refresh
End If
' подгоняем размеры
For Each fldXs In Me.DataGrid1.Columns
    If fldXs.Caption = "Дата" Or _
        fldXs.Caption = "datetime" Then fldXs.Width = 104 _
        Else fldXs.Width = 47
Next
' закрыаем используемые запросы
If DataEnvironment1.rscmdTeploRez.State = adStateOpen Then _
        DataEnvironment1.rscmdTeploRez.Close ' закрыть запрос
If DataEnvironment1.rscmdTeploRezT.State = adStateOpen Then _
        DataEnvironment1.rscmdTeploRezT.Close ' закрыть запрос
If DataEnvironment1.rsCommand4.State = adStateOpen Then _
        DataEnvironment1.rsCommand4.Close ' закрыть запрос
If DataEnvironment1.rsCommand5.State = adStateOpen Then _
        DataEnvironment1.rsCommand5.Close ' закрыть запрос
If DataEnvironment1.rscmdParsRez.State = adStateOpen Then _
        DataEnvironment1.rscmdParsRez.Close ' закрыть запрос
If DataEnvironment1.rscmdParhRez.State = adStateOpen Then _
        DataEnvironment1.rscmdParhRez.Close ' закрыть запрос
' вернуть файл в первоначальное состояние
Call RecovHead
TipOt = ""
End Function
'
Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                    ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
                    CancelDisplay As Boolean)
'при ошибке сообщить о ней, закрыть соединение, разблокировать кнопку
Me.StatusBar1.SimpleText = "Ошибка Winsock #" & Number & "-" & Description
ws.Close ' закрыть соединение
Me.Command4.Enabled = True
End Sub
' вывод отчета тсрв на печать
Sub PrintTSRV(ZL As Integer)
Dim notopen As String, i As Integer
Dim pos1 As Long, pos2 As Long, dRepDate As Date
Dim t1 As Double, t2 As Double, tf As Double
' ZL=1 - лето, ZL=0 - зима
On Error GoTo excl
Call SetLocaleInfo(LOCALE_SDECIMAL, ".")
If ZL = 1 Then
    otchetTSRVs.Sections(1).Controls(10).Caption = "В-0"
    otchetTSRVs.Sections(1).Controls(14).Caption = "dG=m1+m2"
    otchetTSRVs.Sections(1).Controls(15).Caption = "W3=W1 + W2"
    'otchetTSRVs.Sections(1).Controls(16).Caption = "dt=t1+t2"
End If
otchetTSRVs.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
otchetTSRVs.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
otchetTSRVs.Sections(1).Controls("lblDogovor").Caption = Dialog.txtDogov.Text
If DataEnvironment1.rsCommand19.State <> adStateOpen Then _
DataEnvironment1.rsCommand19.Open ' открыть запрос к tembase.csv для работы
DataEnvironment1.rsCommand19.Requery ' обновить данные
'dRepDate = Date
dRepDate = DataEnvironment1.rsCommand19.Fields(0)
otchetTSRVs.Sections(1).Controls(4).Caption = mon(Month(dRepDate)) & " " & Year(dRepDate) & " г."
notopen = DBGrid1.Columns(2).Value
otchetTSRVs.Sections(1).Controls(8).Caption = notopen
notopen = DBGrid1.Columns(3).Value
otchetTSRVs.Sections(1).Controls(9).Caption = notopen
If DataEnvironment1.rsCommand20.State <> adStateOpen Then _
    DataEnvironment1.rsCommand20.Open ' открыть запрос к tembase.csv для работы
' установить формат выводимого листа на печать
'otchetTSRVs.Orientation = rptOrientLandscape
' формируем Итого
With DataEnvironment1.rsCommand20
    .Requery
    For i = 0 To 15
        otchetTSRVs.Sections(5).Controls("L" & Trim(str(i))).Caption = .Fields(i)
    Next
End With
' информация показания счетчиков 1 стр.
pos1 = InStr(1, strTSRV1, "ДатаВремя=", vbTextCompare) + 10
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
pos1 = InStr(1, strTSRV1, "TНС=", vbTextCompare)
If pos1 = 0 Then pos1 = InStr(1, strTSRV1, "Tвр4=", vbTextCompare) + 5 _
Else pos1 = pos1 + 4
pos2 = InStr(pos1, strTSRV1, ";", vbTextCompare) - 1
t1 = Format(CDbl(Mid(strTSRV1, pos1, pos2 - pos1 + 1)) / 60, "0.00")
otchetTSRVs.Sections(5).Controls("lblV1").Caption = t1
' информация показания счетчиков 2 стр.
pos1 = InStr(1, strTSRV2, "ДатаВремя=", vbTextCompare) + 10
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
pos1 = InStr(1, strTSRV2, "TНС=", vbTextCompare)
If pos1 = 0 Then pos1 = InStr(1, strTSRV2, "Tвр4=", vbTextCompare) + 5 _
Else pos1 = pos1 + 4
pos2 = InStr(pos1, strTSRV2, ";", vbTextCompare) - 1
t2 = Format(CDbl(Mid(strTSRV2, pos1, pos2 - pos1 + 1)) / 60, "0.00")
otchetTSRVs.Sections(5).Controls("lblV2").Caption = t2
t1 = Format(t2 - t1, "0.00")
otchetTSRVs.Sections(5).Controls("lblRV").Caption = t1
' установить границы печати
With DataEnvironment1.rsCommand11
    If .State <> adStateOpen Then .Open ' переоткрыть запрос
    otchetTSRVs.BottomMargin = .Fields("Niz") * Stwips
    otchetTSRVs.TopMargin = .Fields("Verh") * Stwips
    otchetTSRVs.LeftMargin = .Fields("Levo") * Stwips
    otchetTSRVs.RightMargin = .Fields("Pravo") * Stwips
    otchetTSRVs.Font.size = .Fields("Shrift")
End With
otchetTSRVs.Show ' предпросмотр и печать
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
Dim filenum ' имя файла для сохранения данных
Dim strX As String, strY As String, x As Long
Dim priznak As String
On Error GoTo errRecHead
' открыть файл
filenum = FreeFile ' дескриптор cвободного файла
Open "tempbase.csv" For Input As #filenum ' открыть файл
' остальные строки записать как есть
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' закрыть файл
filenum = FreeFile ' открыть файл для записи
strY = strTSRV1 & vbCrLf & strTSRV2 & vbCrLf & strY
Open "tempbase.csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' заполнить файл
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
If Len(TipOt) = 0 Then TipOt = RemakeHead ' преобразование заголовка к новому формату
Select Case TipOt
Case "сJournal", "чJournal" ' учет расхода тепловой энергии
    If Dialog.Check3.Value And Left(TipOt, 1) = "с" Then
        Call PrintTSRV(Me.Check1.Value) 'печать ТСРВ
    Else
        ' проверить есть ли контроль тем-ры
        If Not frmStart.mnuKT.Checked Then
            If DataEnvironment1.rsCommand3.State <> adStateOpen Then _
                DataEnvironment1.rsCommand3.Open ' переоткрыть запрос
            DataEnvironment1.rsCommand3.Requery ' обновить данные
            ' итоговая строка
            With DataEnvironment1.rsCommand8
                If .State <> adStateOpen Then .Open ' переоткрыть запрос
                .Requery ' обновить данные
                ' распределить
                For i = 0 To 11
                    DataReport1.Sections(5).Controls("Label" & _
                        Trim(str(26 + i))).Caption = Format(.Fields(i), "0.00")
                Next
            End With
            ' установить формат альбомный выводмого листа на печать
            DataReport1.Orientation = rptOrientLandscape
            ' установить границы печати
            With DataEnvironment1.rsCommand11
                If .State <> adStateOpen Then .Open ' переоткрыть запрос
                DataReport1.BottomMargin = .Fields("Niz") * Stwips
                DataReport1.TopMargin = .Fields("Verh") * Stwips
                DataReport1.LeftMargin = .Fields("Levo") * Stwips
                DataReport1.RightMargin = .Fields("Pravo") * Stwips
                DataReport1.Font.size = .Fields("Shrift")
            End With
            DataReport1.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
            DataReport1.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
            DataReport1.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
            DataReport1.Show ' предпросмотр и печать
        Else
            If DataEnvironment1.rsCommand17.State <> adStateOpen Then _
            DataEnvironment1.rsCommand17.Open ' открыть запрос к tembase.csv для работы
            If DataEnvironment1.rsCommand18.State <> adStateOpen Then _
                DataEnvironment1.rsCommand18.Open ' открыть запрос к tembase.csv для работы
            DataEnvironment1.rsCommand17.Requery ' обновить данные
            ' установить формат выводимого листа на печать
            DataReport6.Orientation = rptOrientLandscape
            ' формируем Итого
            With DataEnvironment1.rsCommand18
                .Requery
                DataReport6.Sections(5).Controls(1).Caption = Format(.Fields(0), "0.00") ' теп.энерг. W1
                DataReport6.Sections(5).Controls(3).Caption = Format(.Fields(2), "0.00") ' масса V1
                DataReport6.Sections(5).Controls(4).Caption = Format(.Fields(10), "0.00") ' темпер. t1
                DataReport6.Sections(5).Controls(5).Caption = Format(.Fields(1), "0.00")  ' теп.энерг. W2
                DataReport6.Sections(5).Controls(6).Caption = Format(.Fields(3), "0.00") ' масса V2
                DataReport6.Sections(5).Controls(7).Caption = Format(.Fields(11), "0.00") ' темп-ра t2
                DataReport6.Sections(5).Controls(8).Caption = Format(.Fields(7), "0.00") ' теп.энерг. W3
                DataReport6.Sections(5).Controls(9).Caption = Format(.Fields(8), "0.00") ' масса V3
                DataReport6.Sections(5).Controls(10).Caption = Format(.Fields(9), "0.00") ' время остан.
                DataReport6.Sections(5).Controls(11).Caption = Format(.Fields(4), "0.00")  ' время остан.
                DataReport6.Sections(5).Controls(12).Caption = Format(.Fields(5), "0.00") ' давл. по прям. d1
                DataReport6.Sections(5).Controls(13).Caption = Format(.Fields(6), "0.00") ' давл.по обрат. d2
                DataReport6.Sections(5).Controls("Label40").Caption = _
                                                        Format(.Fields(12), "0.00") ' тем-ра обратки расч.
                DataReport6.Sections(5).Controls("Label41").Caption = _
                                                        Format(.Fields(13), "0.00") ' темпер-ное отклонение
            End With
            ' установить границы печати
            With DataEnvironment1.rsCommand11
                If .State <> adStateOpen Then .Open ' переоткрыть запрос
                DataReport6.BottomMargin = .Fields("Niz") * Stwips
                DataReport6.TopMargin = .Fields("Verh") * Stwips
                DataReport6.LeftMargin = .Fields("Levo") * Stwips
                DataReport6.RightMargin = .Fields("Pravo") * Stwips
                DataReport6.Font.size = .Fields("Shrift")
            End With
            ' заполнение "шапки" отчета
            DataReport6.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
            DataReport6.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
            DataReport6.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
            DataReport6.Show ' предпросмотр и печать
        End If
    End If
Case "сHV" ' учет расхода хол.воды - суточный
    If DataEnvironment1.rsCommand4.State <> adStateOpen Then _
                                        DataEnvironment1.rsCommand4.Open ' переоткрыть запрос
    DataEnvironment1.rsCommand4.Requery ' обновить данные
    ' установить границы печати
    With DataEnvironment1.rsCommand13
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        DataReport2.BottomMargin = .Fields("Niz") * Stwips
        DataReport2.TopMargin = .Fields("Verh") * Stwips
        DataReport2.LeftMargin = .Fields("Levo") * Stwips
        DataReport2.RightMargin = .Fields("Pravo") * Stwips
        DataReport2.Font.size = .Fields("Shrift")
    End With
    DataReport2.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport2.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport2.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport2.Show ' предпросмотр и печать
Case "чHV" ' учет расхода хол.воды - часовой
    If DataEnvironment1.rsCommand5.State <> adStateOpen Then _
        DataEnvironment1.rsCommand5.Open ' переоткрыть запрос
    DataEnvironment1.rsCommand5.Requery ' обновить данные
    ' установить границы печати
    With DataEnvironment1.rsCommand13
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        DataReport4.BottomMargin = .Fields("Niz") * Stwips
        DataReport4.TopMargin = .Fields("Verh") * Stwips
        DataReport4.LeftMargin = .Fields("Levo") * Stwips
        DataReport4.RightMargin = .Fields("Pravo") * Stwips
        DataReport4.Font.size = .Fields("Shrift")
    End With
    DataReport4.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport4.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport4.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport4.Show ' предпросмотр и печать
Case "сPAR" ' учет расхода пара сут.
    If DataEnvironment1.rsCommand6.State <> adStateOpen Then _
        DataEnvironment1.rsCommand6.Open ' переоткрыть запрос
    DataEnvironment1.rsCommand6.Requery ' обновить данные
    ' итоговая строка
    With DataEnvironment1.rsCommand9
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery ' обновить данные
        ' распределить
        For i = 0 To 11
            DataReport3.Sections(5).Controls("Label" & _
                Trim(str(25 + i))).Caption = Format(.Fields(i), "0.00")
        Next
    End With
    ' установить формат альбомный выводмого листа на печать
    DataReport3.Orientation = rptOrientLandscape
    ' установить границы печати
    With DataEnvironment1.rsCommand12
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        DataReport3.BottomMargin = .Fields("Niz") * Stwips
        DataReport3.TopMargin = .Fields("Verh") * Stwips
        DataReport3.LeftMargin = .Fields("Levo") * Stwips
        DataReport3.RightMargin = .Fields("Pravo") * Stwips
        DataReport3.Font.size = .Fields("Shrift")
    End With
    DataReport3.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport3.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport3.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport3.Show ' предпросмотр и печать
Case "чPAR" ' учет расхода пара час.
    If DataEnvironment1.rsCommand7.State <> adStateOpen Then _
                                    DataEnvironment1.rsCommand7.Open ' переоткрыть запрос
    DataEnvironment1.rsCommand7.Requery ' обновить данные
    ' итоговая строка
    With DataEnvironment1.rsCommand10
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        .Requery ' обновить данные
        ' распределить
        For i = 0 To 11
            DataReport5.Sections(5).Controls("Label" & _
                Trim(str(25 + i))).Caption = Format(.Fields(i), "0.00")
        Next
    End With
    ' установить формат альбомный выводмого листа на печать
    DataReport5.Orientation = rptOrientLandscape
    ' установить границы печати
    With DataEnvironment1.rsCommand12
        If .State <> adStateOpen Then .Open ' переоткрыть запрос
        DataReport5.BottomMargin = .Fields("Niz") * Stwips
        DataReport5.TopMargin = .Fields("Verh") * Stwips
        DataReport5.LeftMargin = .Fields("Levo") * Stwips
        DataReport5.RightMargin = .Fields("Pravo") * Stwips
        DataReport5.Font.size = .Fields("Shrift")
    End With
    DataReport5.Sections(1).Controls("lblPotreb").Caption = Dialog.txtPotreb.Text
    DataReport5.Sections(1).Controls("lblAdres").Caption = Dialog.txtAdres.Text
    DataReport5.Sections(1).Controls("lblDogov").Caption = Dialog.txtDogov.Text
    DataReport5.Show ' предпросмотр и печать
End Select
' восстановить заголовок файла
Call RecovHead
TipOt = ""
Exit Sub
Print_err:
 'MsgBox Err.Number & "->" & Err.Description
 Resume Next
End Sub
'функция предварительной подготовки полученного файла
Function RemakeHead() As String
Dim filenum ' имя файла для сохранения данных
Dim strX As String, strY As String, x As Long
Dim priznak As String
On Error GoTo errRHead
' открыть файл
filenum = FreeFile ' дескриптор cвободного файла
Open "tempbase.csv" For Input As #filenum ' открыть файл
' пропустить первую и вторую строку
Line Input #filenum, strTSRV1 '
Line Input #filenum, strTSRV2 '
' выбрать тип отчета
x = InStr(1, strTSRV1, "Часовой", 1)
If x = 0 Then
    x = InStr(1, strTSRV1, "Суточный", 1)
    priznak = "с"
Else
    priznak = "ч"
End If
priznak = priznak & Mid(strTSRV1, 1, x - 1)
' корректировать строку заголовка
Line Input #filenum, strX '
'strY = "ДатаВремя;W1;W2;m1;m2;TНС;P1;P2;W3;V3;Траб;t1;t2" & vbCrLf
If InStr(1, strX, ";Tвр4", 1) > 0 Then
    strY = Mid(strX, 1, InStr(1, strX, ";Tвр4", 1)) & "TНС" & _
                Mid(strX, InStr(1, strX, ";Tвр4", 1) + 5) & vbCrLf
ElseIf InStr(1, strX, ";НС", 1) > 0 Then
    strY = Mid(strX, 1, InStr(1, strX, ";НС", 1)) & "TНС" & _
                Mid(strX, InStr(1, strX, ";НС", 1) + 3) & vbCrLf
Else
    strY = strX & vbCrLf
End If
' остальные строки записать как есть
Do While Not EOF(filenum)
   Line Input #filenum, strX '
   ' обработать на предмет превышения
   'strX = Milliard(strX)
   strY = strY & strX & vbCrLf
Loop
Close #filenum ' закрыть файл
filenum = FreeFile ' открыть файл для записи
Open "tempbase.csv" For Output As #filenum
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' заполнить файл
Close #filenum
RemakeHead = priznak
Exit Function
errRHead:
Resume Next
End Function
' преобразовать минусовые значения в правильные
Function Milliard(strMilli As String) As String
Dim stra(6) As String, tr As Boolean
Dim pos As Long, pos1 As Long, i As Long
Dim str1 As String
pos = InStr(1, strMilli, ";", 1) ' ищем первое вхождение
tr = False
Do While pos
    i = i + 1
    If i >= 2 Then
        stra(i - 2) = Mid(strMilli, pos1 + 1, pos - pos1 - 1) '
        If i = 8 Then Exit Do
    End If
    pos1 = pos
    pos = InStr(pos1 + 1, strMilli, ";", 1) ' искать следующий
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
' получение новых позиций
Function NewPos(strQ As String, num As Long) As String
Dim pos As Long, pos1 As Long, i As Long
i = 0
pos = InStr(1, strQ, ";", 1) ' ищем первое вхождение
Do While pos
    i = i + 1
    If i = num Then Exit Do
    pos1 = pos
    pos = InStr(pos1 + 1, strQ, ";", 1) ' искать следующий
Loop
NewPos = CStr(pos) & "/" & CStr(pos1)
End Function
'
Private Sub DbGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If Not IsNull(LastRow) Then
On Error GoTo exit_DBG
    'Text6.Text = DBGrid1.Text ' выделить наименование узла учета
'End If
exit_DBG:
End Sub
'
Private Sub Form_Unload(Cancel As Integer) 'При выходе из программы:
    DataEnvironment1.rsCommand2.Close ' закрыть таблицу настроек
    ' =1 при выходе эффект сворачивания,
    ' =2 - эффект разворачивания, =0 - то мерцания.
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
    'Открываем рабочую область
    With DataEnvironment1.rsCommand2 ' берем настройки
        If .State <> adStateOpen Then .Open
        .Requery
        .MoveFirst
        Do While Not .EOF
            Select Case .Fields("NameSet") ' взять название параметра настройки
            Case "PathAdmin"
                tmpServ = .Fields("Set") ' взять название сервера соединения
            Case "Mode" ' выбрать режим получения данных
                If .Fields("Set") = "True" Then
                    Check1.Value = 1
                    Check1.Caption = "Режим 'Лето'"
                Else
                    Check1.Value = 0
                    Check1.Caption = "Режим 'Зима'"
                End If
            Case "KindArchive" ' выбрать тип архива
                Combo1.Text = .Fields("Set")
            Case "Port"
                ' взять номер порта
                tmpPort = .Fields("Set")
            End Select
            .MoveNext ' по всем параметрам настройки
        Loop
    End With
    'Form1.Text2 = Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1)
    Dialog.Text1.Text = tmpPort ' подготовить настройки
    Dialog.Text2.Text = tmpServ 'Text1.Text ' подготовить назв.сервера
    ' выбрать период получения архива:
    ' от начала месяца
    Text4.Text = CDate("01/" & Month(Date) & "/" & Year(Date))
    Text5.Text = Date ' до текущей даты
    blTP = False
    'Me.Text6.Text = Me.DBGrid1.Text ' выбрать имя текущего узла
    Call Command4_Click ' установить связь сразу
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
' Записываем в раздел  переменную
'Call WriteParameters("PathAdmin", Text1.Text)
Exit Sub
TC1_err:
    'MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub
'
'запись текста в файл events.log в текущей директории базы
Public Sub writeLog(Text As String)
    Dim logFile As String
    Dim FileNr As Integer
    'определить имя текущей базы
    'добавить к файлу-протоколу путь к текущей базы
    logFile = CurDir & "\events.log"
    'открыть файл-протокол
    FileNr = FreeFile:    Open logFile For Append As FileNr
    'записать вызванное событие
    Print #FileNr, Format(Now, "dd.mm.yy hh:nn:ss ") & " : "; Text
    'закрыть файл-протокол
    Close FileNr
End Sub
'протокол ошибок в таблицу "Ошибки"
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
' функции проверки наличия компонентов Microsoft office
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
' непосредственно выводит сведения о компонентах MS Office
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
' сжатие БД
'Sub BDCmpct()
'On Error GoTo ercmpct
'DBEngine.CompactDatabase "settings.mdb", "setting.mdb", dbLangCyrillic & ";pwd=MTWTFSS", , ";pwd=MTWTFSS" '
'Kill "settings.mdb"
'Name "setting.mdb" As "settings.mdb"
'Exit Sub
'ercmpct:
' writeLog ("compact DB: " & protocol())
'End Sub
