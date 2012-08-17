VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Графопостроитель"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form3"
   ScaleHeight     =   7215
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5565
      Left            =   75
      OleObjectBlob   =   "frmGraph.frx":0000
      TabIndex        =   1
      Top             =   90
      Width           =   9105
   End
   Begin VB.Frame Frame1 
      Height          =   1425
      Left            =   45
      TabIndex        =   0
      Top             =   5745
      Width           =   9195
      Begin VB.CheckBox chkArh 
         Caption         =   "добавить архив"
         Height          =   240
         Left            =   180
         TabIndex        =   22
         Top             =   1035
         Width           =   1485
      End
      Begin VB.CheckBox chkTek 
         Caption         =   "начать заново"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   780
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox Check4 
         Caption         =   "по часам"
         Height          =   225
         Left            =   7575
         TabIndex        =   9
         Top             =   210
         Width           =   1035
      End
      Begin VB.Frame Frame2 
         Height          =   1425
         Left            =   2580
         TabIndex        =   7
         Top             =   0
         Width           =   4755
         Begin VB.CommandButton btnPrint 
            Caption         =   "Печать"
            Height          =   315
            Left            =   3390
            TabIndex        =   25
            Top             =   900
            Width           =   945
         End
         Begin VB.CheckBox chkGis 
            Caption         =   "Гистограмма"
            Height          =   195
            Left            =   255
            TabIndex        =   24
            Top             =   1110
            Width           =   1605
         End
         Begin VB.CheckBox chkLin 
            Caption         =   "Линейный график"
            Height          =   195
            Left            =   255
            TabIndex        =   23
            Top             =   825
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1650
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "23"
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   825
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   825
            TabIndex        =   14
            Top             =   150
            Width           =   1185
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2685
            TabIndex        =   13
            Top             =   150
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmGraph.frx":24D6
            Left            =   2460
            List            =   "frmGraph.frx":2501
            TabIndex        =   12
            Text            =   "Параметр"
            Top             =   480
            Width           =   1785
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Добавить"
            Height          =   315
            Left            =   2430
            TabIndex        =   8
            Top             =   900
            Width           =   975
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   3885
            TabIndex        =   10
            Top             =   150
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   39861
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   1980
            TabIndex        =   11
            Top             =   150
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            _Version        =   393216
            Format          =   16515073
            CurrentDate     =   39861
         End
         Begin VB.Label Label4 
            Caption         =   "до:"
            Height          =   255
            Left            =   1380
            TabIndex        =   20
            Top             =   495
            Width           =   210
         End
         Begin VB.Label Label3 
            Caption         =   "время от:"
            Height          =   285
            Left            =   45
            TabIndex        =   18
            Top             =   495
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "период с:"
            Height          =   255
            Left            =   60
            TabIndex        =   16
            Top             =   195
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "по:"
            Height          =   210
            Left            =   2400
            TabIndex        =   15
            Top             =   195
            Width           =   255
         End
      End
      Begin VB.CheckBox Check3 
         Caption         =   "по годам"
         Height          =   225
         Left            =   7575
         TabIndex        =   6
         Top             =   1065
         Width           =   1185
      End
      Begin VB.CheckBox Check2 
         Caption         =   "по месяцам"
         Height          =   255
         Left            =   7575
         TabIndex        =   5
         Top             =   750
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "по дням"
         Height          =   255
         Left            =   7575
         TabIndex        =   4
         Top             =   465
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Посмотреть"
         Height          =   315
         Left            =   1305
         TabIndex        =   3
         Top             =   345
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Подготовить"
         Height          =   315
         Left            =   165
         TabIndex        =   2
         Top             =   345
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LineNum As Long

Private Sub btnPrint_Click()
Dim Msg   ' Declare variable.
Dim lH As Long, lW As Long
On Error GoTo ErrorbtnPrint   ' Set up error handler.
If MsgBox("Действительно печатать?", vbQuestion + vbYesNo) = vbYes Then
    Me.Frame1.Visible = False
    Me.BackColor = RGB(255, 255, 255)
    lW = Me.Width
    lH = Me.Height
    Me.Width = 567 * 28
    Me.Height = 567 * 21
    Printer.Orientation = vbPRORLandscape
    PrintForm   ' печать
    Printer.Orientation = vbPRORPortrait
    Me.Width = lW
    Me.Height = lH
    Me.BackColor = &H8000000F
    Me.Frame1.Visible = True
End If
Exit Sub
ErrorbtnPrint:
    Msg = "Форма не может быть распечатана"
    MsgBox Msg   ' Display message.
    Resume Next
End Sub

'
Private Sub Check1_Click()
If Me.Check1.Value Then
    Me.Check2.Value = False
    Me.Check3.Value = False
    Me.Check4.Value = False
End If
End Sub

Private Sub Check2_Click()
If Me.Check2.Value Then
    Me.Check1.Value = False
    Me.Check3.Value = False
    Me.Check4.Value = False
End If
End Sub

Private Sub Check3_Click()
If Me.Check3.Value Then
    Me.Check2.Value = False
    Me.Check1.Value = False
    Me.Check4.Value = False
End If
End Sub

Private Sub Check4_Click()
If Me.Check4.Value Then
    Me.Check2.Value = False
    Me.Check3.Value = False
    Me.Check1.Value = False
End If
End Sub


Private Sub chkArh_Click()
If Me.chkArh.Value Then Me.chkTek.Value = False _
Else Me.chkTek.Value = 1
End Sub

Private Sub chkGis_Click()
If Me.chkGis.Value Then
    Me.chkLin.Value = False
    Me.MSChart1.chartType = VtChChartType2dBar
    Me.MSChart1.Stacking = True
Else
    Me.chkLin.Value = 1
    Me.MSChart1.chartType = VtChChartType2dLine
    Me.MSChart1.Stacking = False
End If
End Sub

Private Sub chkLin_Click()
If Me.chkLin.Value Then
    Me.chkGis.Value = False
    Me.MSChart1.chartType = VtChChartType2dLine
    Me.MSChart1.Stacking = False
Else
    Me.chkGis.Value = 1
    Me.MSChart1.chartType = VtChChartType2dBar
    Me.MSChart1.Stacking = True
End If
End Sub

Private Sub chkTek_Click()
If Me.chkTek.Value Then Me.chkArh.Value = False _
Else Me.chkArh.Value = 1
End Sub

Private Sub Command1_Click()
frmOpenFiles.Show
End Sub
'
Private Sub Command2_Click()
' посмотреть полученную таблицу
frmtblGraf.Show
End Sub

Private Sub Command3_Click()
Dim i As Long, ip As Long, predel As Long
Dim cnn As ADODB.Connection
Dim rst As ADODB.Recordset
On Error GoTo errNetGraf
Set cnn = New ADODB.Connection
' Контакт! - Есть контакт!
cnn.Open "DefaultDir=;Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
    "DriverId=27;Extensions=csv;FIL=text;FILEDSN=vzljot_csv.dsn;MaxBufferSize=2048;" & _
    "MaxScanRows=25;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"
' открыть запрос к tempgraf.csv для работы
Set rst = New ADODB.Recordset
If Check1.Value Then ' по дням
    rst.Open "SELECT format([ДатаВремя],'dd.mm.yyyy') as DataT,W1,W2,m1,m2,[TНС],[Формула1] as W3," & _
            "[Формула2] as m3,[Формула3],[Формула4] as t1,[Формула5] as t2, [Формула6]," & _
            "[Формула7] as P1,[Формула8] as P2,[Формула9],t2r,dt,[Формула7] - [Формула8] as Pr FROM `tempgraf.csv`" & _
            " WHERE format([ДатаВремя],'dd.mm.yyyy') Between " & SQLDate(Me.Text1.Text) & _
            " AND " & SQLDate(Me.Text2.Text), cnn, adOpenStatic, adLockReadOnly
ElseIf Check2.Value Then ' по месяцам
    rst.Open "SELECT Month(ДатаВремя) AS DataT, SUM(W1) AS sW1, SUM(W2) AS sW2," & _
            "SUM(m1) AS sm1, SUM(m2) AS sm2, SUM(Tнс) AS sTнс, SUM(Формула1) AS sW3," & _
            "SUM(Формула2) AS sm3, AVG(Формула4) AS at1,AVG(Формула5) AS at2, AVG(Формула7) AS aP1," & _
            "AVG(Формула8) AS aP2, SUM(t2r) AS str2, SUM(dt) AS sdt,SUM([Формула7] - [Формула8]) as Pr FROM `tempgraf.csv`" & _
            " WHERE format([ДатаВремя],'dd.mm.yyyy') Between " & SQLDate(Me.Text1.Text) & _
            " AND " & SQLDate(Me.Text2.Text) & " GROUP BY Month(ДатаВремя)", cnn, _
            adOpenStatic, adLockReadOnly
ElseIf Check3.Value Then ' по годам
    MsgBox "Функция в стадии разработки"
    cnn.Close
    Exit Sub
ElseIf Check4.Value Then ' по часам
    rst.Open "SELECT format([ДатаВремя],'dd hh:mm') as DataT,W1,W2,m1,m2,[TНС],[Формула1] as W3," & _
            "[Формула2] as m3,[Формула3],[Формула4] as t1,[Формула5] as t2, [Формула6]," & _
            "[Формула7] as P1,[Формула8] as P2,[Формула9],t2r,dt,[Формула7] - [Формула8] as Pr FROM `tempgraf.csv`" & _
            " WHERE format([ДатаВремя],'dd.mm.yyyy hh:mm') Between " & SQLDate(Me.Text1.Text, _
            TimeValue(Text3.Text & ":00")) & " AND " & SQLDate(Me.Text2.Text, TimeValue(Text4.Text & ":00")), _
            cnn, adOpenStatic, adLockReadOnly
Else
    MsgBox "Выберите период анализа"
    cnn.Close
    Exit Sub
End If
' выбираем анализируемый параметр
ip = Me.Combo1.ListIndex
ip = Switch(ip = 0, 1, ip = 1, 3, ip = 2, 12, ip = 3, 9, _
        ip = 4, 2, ip = 5, 4, ip = 6, 13, ip = 7, 10, ip = 8, 6, _
        ip = 9, 7, ip = 10, 15, ip = 11, 16, ip = 12, 17)
' выбрать данные
With rst
    '.Requery ' обновить данные
    .MoveLast: .MoveFirst: predel = .RecordCount
    ' добавить линию в график
    LineNum = LineNum + 1 'выбрать линию
    MSChart1.Column = LineNum
    MSChart1.ColumnLabel = .Fields(ip).Name
    MSChart1.RowCount = predel
    For i = 1 To predel
        ' определить значения по оси х
        MSChart1.Row = i
        MSChart1.RowLabel = .Fields(0)
        ' определить значение самого графика
        MSChart1.data = IIf(IsNull(.Fields(ip)), 0, .Fields(ip))
        .MoveNext
    Next
End With
rst.Close
Set rst = Nothing
cnn.Close
extGraf:
Exit Sub
errNetGraf:
MsgBox Err.Description
Resume extGraf
End Sub


Private Sub DTPicker1_CloseUp()
Me.Text1.Text = Me.DTPicker1.Value
End Sub


Private Sub DTPicker2_CloseUp()
Me.Text2.Text = Me.DTPicker2.Value
End Sub

Private Sub Form_Load()
Dim i As Long, j As Long
LineNum = 0
For j = 1 To MSChart1.ColumnCount
    MSChart1.Column = j
    For i = 1 To MSChart1.RowCount
        MSChart1.Row = i
        MSChart1.data = 0
    Next
Next
End Sub

Private Sub Form_Resize()
Me.Frame1.Top = Me.Height * 0.9 - Me.Frame1.Height ' сдвинуть управляющие команды
Me.MSChart1.Height = Me.Height * 0.9 - Me.Frame1.Height ' изменить высоту графика
Me.MSChart1.Width = Me.Width * 0.95 ' изменить ширину графика
End Sub

Private Sub MSChart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
MSChart1.Column = Series
MsgBox MSChart1.ColumnLabel
End Sub

