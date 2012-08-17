VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTG 
   Caption         =   "Температурный график"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form4"
   ScaleHeight     =   6825
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDel 
      Caption         =   "Удалить"
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox txtOT 
      Height          =   285
      Left            =   2640
      TabIndex        =   16
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtPT 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox txtNV 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Редактировать"
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Добавить"
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Загрузить"
      Height          =   315
      Left            =   3130
      TabIndex        =   11
      Top             =   6480
      Width           =   945
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Сохранить"
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   6480
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Расчитать"
      Height          =   315
      Left            =   5010
      TabIndex        =   8
      Top             =   6480
      Width           =   945
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5850
      TabIndex        =   4
      Top             =   630
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5850
      TabIndex        =   2
      Top             =   195
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTG.frx":0000
      Height          =   5805
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   10239
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.CommandButton Command1 
      Caption         =   "Выход"
      Height          =   315
      Left            =   5955
      TabIndex        =   0
      Top             =   6480
      Width           =   945
   End
   Begin MSComDlg.CommonDialog CDArh 
      Left            =   5160
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Температурный график"
      Filter          =   "Формат с разделителем  (*.csv)|*.csv"
   End
   Begin MSAdodcLib.Adodc adoTG 
      Height          =   330
      Left            =   120
      Top             =   5520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   582
      ConnectMode     =   2
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   2
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
      Caption         =   "Температурный график"
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
   Begin VB.Label Label8 
      Caption         =   "Тем-ра об.тр-да"
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Тем-ра пр.тр-да"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Тем-ра н.возд."
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Примечание: TN-тем-ра наружная; Т1-тем-ра по прямому тр-ду;       Т2-тем-ра по обратному тр-ду; Т2R-тем-ра расчетн."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   5130
      TabIndex        =   9
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "dT = T2 - T2R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5610
      TabIndex        =   7
      Top             =   1530
      Width           =   1305
   End
   Begin VB.Label Label3 
      Caption         =   "T2R = A*T1+B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5610
      TabIndex        =   6
      Top             =   1230
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "B ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5490
      TabIndex        =   5
      Top             =   660
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "А ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5475
      TabIndex        =   3
      Top             =   225
      Width           =   345
   End
End
Attribute VB_Name = "frmTG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCnct As New ADODB.Connection
'Dim adoRst As New ADODB.Recordset
'
Private Sub btnLoad_Click()
Dim NFSO As New FileSystemObject
Dim FileNumber As Long
Dim str1 As String
Dim x As Long
Dim i As Long
Dim str4 As String
On Error GoTo errLoad
CDArh.CancelError = True
CDArh.ShowOpen ' открыть файл
ChDir App.Path
' открыть файл
FileNumber = FreeFile
Open CDArh.FileName For Input As #FileNumber
If frmStart.mnu61.Checked = True Then strNT = "TG" _
Else strNT = "sp_tr"
adoCnct.Execute "DELETE FROM " & strNT & ";"
With Me.adoTG.Recordset
    .Requery
    Do While Not EOF(FileNumber)   ' пока не конец файла
        Line Input #FileNumber, str1 ' брать построчно
        ' скорректировать десятичный разделитель
        If KindOfDecSep = "," Then str1 = RemakeS(str1, True) _
        Else str1 = RemakeS(str1, False)
        str1 = str1 & ";" ' подготовить строку для анализа
        x = InStr(1, str1, ";"):    i = 0 ' взять первое вхождение
        .AddNew ' начать добавление значений
        Do While x
            ' вставлять значение в соответств.поле
            .Fields(i) = CDbl(Mid(str1, 1, x - 1))
            ' перейти к следующ.значению
            str1 = Mid(str1, x + 1):  i = i + 1
            x = InStr(1, str1, ";") ' продолжить анализ
        Loop
        .Update ' сохранить введенные значения
    Loop
End With
Close #FileNumber ' закрыть файл
Me.adoTG.Recordset.Requery
Me.adoTG.Refresh
Me.DataGrid1.Refresh
errLoad:
ChDir App.Path
'MsgBox Err.Description
End Sub
'
Private Sub btnSave_Click()
Dim FileNumber As Long
Dim rstTG As Recordset
On Error GoTo errSave
CDArh.CancelError = True
CDArh.FileName = "Куда сохранить"
CDArh.ShowSave ' открыть файл
'
FileNumber = FreeFile
Open CDArh.FileName For Output As #FileNumber
Set rstTG = Me.adoTG.Recordset
With rstTG
    .MoveFirst
    Do While Not .EOF
        Print #FileNumber, .Fields("TN") & ";";  ' темп.наружн.возд.
        Print #FileNumber, .Fields("T1") & ";"; ' темп.по прям.тр-ду
        Print #FileNumber, .Fields("T2") & ";";  ' темп.по обрат.тр-ду
        Print #FileNumber, .Fields("T2R")  ' темп.по обрат.тр-ду расчетн.
        .MoveNext
    Loop
End With
'rstTG.Close: Set rstTG = Nothing
Close #FileNumber
CDArh.FileName = ""
errSave:
ChDir App.Path
End Sub

Private Sub cmdAdd_Click()
Dim lTN As Double, lT1 As Double, lT2 As Double
Dim strNT As String
On Error GoTo errcmdAdd
If Len(txtNV.Text) > 0 And Len(txtPT.Text) > 0 And Len(txtOT.Text) > 0 Then
    If frmStart.mnu61.Checked = True Then strNT = "TG" Else strNT = "sp_tr"
    txtPT.Text = RemakeS(txtPT.Text, False)
    txtOT.Text = RemakeS(txtOT.Text, False)
    lTN = Val(txtNV.Text)
    lT1 = Val(txtPT.Text)
    lT2 = Val(txtOT.Text)
    adoCnct.Execute "INSERT INTO " & strNT & " (tn,t1,t2) VALUES (" & txtNV.Text & _
            "," & txtPT.Text & "," & txtOT.Text & ");"
End If
Me.adoTG.Recordset.Requery
Me.adoTG.Refresh
Me.DataGrid1.Refresh
Exit Sub
errcmdAdd:
'MsgBox "Невозможно добавить данные"
MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub cmdDel_Click()
Dim lTN As Double, lT1 As String, strNT As String
On Error GoTo errcmdDel
lTN = Me.DataGrid1.Columns(0)
lT1 = Me.DataGrid1.Columns(1)
lT1 = RemakeS(lT1, False)
If frmStart.mnu61.Checked = True Then strNT = "TG" _
Else strNT = "sp_tr"
adoCnct.Execute "DELETE FROM " & strNT & _
            " WHERE tn=" & lTN & " and t1=" & lT1 & ";"
Me.adoTG.Recordset.Requery
Me.adoTG.Refresh
Me.DataGrid1.Refresh
Exit Sub
errcmdDel:
MsgBox Err.Number & "-" & Err.Description
Resume Next
End Sub

Private Sub cmdEdit_Click()
Dim lTN As Double, lT1 As String, lT2 As String
Dim lP As Long, strNT As String
On Error GoTo errcmdEdit
If Len(txtNV.Text) > 0 And Len(txtPT.Text) > 0 And Len(txtOT.Text) > 0 Then
    lP = DataGrid1.Row ' сохраняем позицию
    lTN = Me.DataGrid1.Columns(0)
    lT1 = Me.DataGrid1.Columns(1)
    lT2 = Me.DataGrid1.Columns(2)
    If frmStart.mnu61.Checked = True Then strNT = "TG" Else strNT = "sp_tr"
    txtPT.Text = RemakeS(txtPT.Text, False)
    txtOT.Text = RemakeS(txtOT.Text, False)
    lT1 = RemakeS(lT1, False)
    lT2 = RemakeS(lT2, False)
        adoCnct.Execute "UPDATE " & strNT & " SET tn = " & txtNV.Text & _
                    ",t1 = " & txtPT.Text & "," & _
                    " t2 = " & txtOT.Text & _
                    " WHERE tn=" & lTN & " and t1=" & _
                    lT1 & ";"
End If
Me.adoTG.Recordset.Requery
Me.adoTG.Refresh
Me.DataGrid1.Refresh
DataGrid1.Row = lP
Exit Sub
errcmdEdit:
'MsgBox "Невозможно добавить данные"
MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim strNT As String
On Error Resume Next
' выполнить непосредственно расчет
If frmStart.mnu61.Checked = True Then strNT = "TG" Else strNT = "sp_tr"
Call koefInter(Me.adoTG.Recordset, strNT)
' сохранить расчитанные параметры
Call WriteParameters("KOEFA", Me.Text1.Text)
Call WriteParameters("KOEFB", Me.Text2.Text)
End Sub
'
' выполнение расчетов коэф. интерполяции
Function koefInter(rstX As Recordset, strE As String) As Boolean
On Error GoTo errRasKoef
If IsNull(isit) Then isit = False
' выполнить расчет методом простой линейной аппроксимации (ПЛА)
With rstX
    ' 1 вычислить кол-во значений
    .MoveLast
    .MoveFirst
    n = .RecordCount
    Do While Not .EOF
        ' 2 вычислить сумму произведений T1*T2
        s1 = s1 + .Fields(1) * .Fields(2)
        ' 3 вычислить сумму T1
        s2 = s2 + .Fields(1)
        ' 4 вычислить сумму T2
        s3 = s3 + .Fields(2)
        ' 6 вычислить сумму квадратов T1
        s4 = s4 + .Fields(1) * .Fields(1)
        .MoveNext
    Loop
    ' 8 вычислить коэфф. А
    Text1.Text = Format((n * s1 - s2 * s3) / (n * s4 - s2 * s2), "0.0000")
    ' 9 вычислить коэфф. B
    Text2.Text = Format((s3 * s4 - s2 * s1) / (n * s4 - s2 * s2), "0.0000")
    ' пересчитать расчетное значение
    .MoveFirst
    Do While Not .EOF
        adoCnct.Execute "UPDATE " & strE & " SET t2r = " _
                    & RemakeS(.Fields(1) * CDbl(Text1.Text) + CDbl(Text2.Text), False) & _
                                    " WHERE tn=" & .Fields(0) & " and t1=" & _
                                    RemakeS(.Fields(1), False) & ";"
        .MoveNext
    Loop
    .Requery
End With
Me.adoTG.Refresh
Me.DataGrid1.Refresh
Exit Function
errRasKoef:
MsgBox Err.Number & "-" & Err.Description  '"Невозможно выполнить расчет"
End Function

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
txtNV.Text = Me.DataGrid1.Columns(0)
txtPT.Text = Me.DataGrid1.Columns(1)
txtOT.Text = Me.DataGrid1.Columns(2)
End Sub

'
Private Sub Form_Load()
On Error Resume Next
If frmStart.mnu61.Checked = True Then
'    adoCnct.ConnectionString = "Provider=MSDASQL.1;FILEDSN=setting.dsn;DATABASE=setting.mdb;" & _
'    "UID=admin;PWD=MTWTFSS;"
    adoTG.ConnectionString = _
    "DBQ=settings.mdb;DefaultDir=.;Driver={Microsoft Access Driver (*.mdb)};" & _
    "DriverId=281;FIL=MS Access;FILEDSN=setting.dsn;MaxBufferSize=2048;MaxScanRows=8;" & _
    "PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;PWD=MTWTFSS;"
    adoTG.RecordSource = "TG"
    Me.adoTG.Refresh
ElseIf frmStart.mnu62.Checked = True Then
    ' сформировать и подключить данные по АСКУТЭ
'    Provider=MSDASQL.1;Persist Security Info=False;_
'    Extended Properties = "DATABASE=askute;DRIVER={MySQL ODBC 3.51 Driver}; _
'    OPTION=0;PORT=3306;SERVER=192.168.100.23"
    Me.adoTG.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                      & "SERVER=192.168.100.23;" _
                      & " DATABASE=askute;" _
                      & "OPTION=0; PORT=3306"
    Me.adoTG.Password = frmLogin.strPSW
    Me.adoTG.UserName = frmLogin.strUser
    Me.adoTG.RecordSource = "sp_tr" ' открываем таблицу температурного графика
    Me.adoTG.Refresh
End If
' загрузить расчитанные параметры аппроксимации
Text1.Text = ReadNParam("KOEFA")
Text2.Text = ReadNParam("KOEFB")
Set adoCnct = Me.adoTG.Recordset.ActiveConnection
End Sub
