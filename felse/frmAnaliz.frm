VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAnaliz 
   Caption         =   "Анализ состояния узлов"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   LinkTopic       =   "Form3"
   ScaleHeight     =   5175
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2580
      TabIndex        =   6
      Top             =   4575
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   315
      TabIndex        =   5
      Top             =   4575
      Width           =   1725
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAnaliz.frx":0000
      Height          =   4215
      Left            =   165
      TabIndex        =   4
      Top             =   45
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   7435
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Узел"
         Caption         =   "Узел"
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
         DataField       =   "ДатаВремя"
         Caption         =   "ДатаВремя"
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
         DataField       =   "Параметр"
         Caption         =   "Параметр"
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
            ColumnWidth     =   2280,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2190,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2759,811
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnTie 
      Caption         =   "Соединить"
      Height          =   330
      Left            =   4650
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   30
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnUpdt 
      Caption         =   "Обновить"
      Height          =   330
      Left            =   5865
      TabIndex        =   1
      Top             =   4560
      Width           =   1080
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Выход"
      Height          =   330
      Left            =   6945
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   210
      Top             =   3930
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
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
      Connect         =   $"frmAnaliz.frx":0015
      OLEDBString     =   $"frmAnaliz.frx":0147
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tempbase.csv"
      Caption         =   "Adodc1"
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
   Begin VB.Label Label4 
      Caption         =   "по"
      Height          =   240
      Left            =   2370
      TabIndex        =   8
      Top             =   4605
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "c"
      Height          =   180
      Left            =   165
      TabIndex        =   7
      Top             =   4605
      Width           =   135
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4920
      Width           =   8085
   End
End
Attribute VB_Name = "frmAnaliz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnTie_Click()
On Error GoTo errbtnTie
'сообщить о завершенном соединении,
Me.Label1.Caption = ""
'закроем на всякий случай, если соединение уже открыто
ws.Close
'соединяемся с сервером по порту
ws.Connect Dialog.Text2.Text, Dialog.Text1.Text ' 1001
'делаем так, чтобы пользователь не смог второй раз нажать кнопку соединения
Me.btnTie.Enabled = False
errbtnTie:
End Sub

Private Sub btnUpdt_Click()
' обновить данные
Dim strD As String, strDate As String, strDate1 As String
Dim n As Integer
On Error GoTo errbtnUpdt
    Me.Label1.Caption = "Исполнение запроса..."
        'Открываем рабочую область
    With DataEnvironment1.rsCommand2 ' берем настройки
        If .State <> adStateOpen Then .Open
        .MoveFirst
        Do While Not .EOF
            ' взять дату последнего анализа
            If .Fields("NameSet") = "Mode" Then ' выбрать режим получения данных
                If .Fields("Set") = "True" Then n = 1 Else n = 0
            End If
            .MoveNext ' по всем параметрам настройки
        Loop
    End With
    ' сформировать дату и время запроса анализа
'    strDate = Format(Date - 1, "dd/mm/yyyy") & " " & _
'            Mid(strDate, 1, InStr(1, strDate, ":", 1)) & "00"
'    strDate1 = Format(Now, "dd/mm/yyyy hh:mm")
'    ' сформировать запрос за сутки
'    If DateDiff("h", strDate, strDate1, vbMonday) > 24 Then
'        strDate = Format(Date - 1, "dd/mm/yyyy") & _
'                            " " & Format(Time(), "hh") & ":00"
'    End If
    strDate = Format(Me.Text1.Text, "dd.mm.yyyy hh:mm:ss")
    strDate1 = Format(Me.Text2.Text, "dd.mm.yyyy hh:mm:ss")
    ' сформировать запрос
    strD = "anlz[" & Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1) & _
            "][][" & strDate & "][" & strDate1 & "][" & "Часовой][" & Trim(str(n)) & "]"
    ' отправить запрос
    ws.SendData strD
'End If
errbtnUpdt:
End Sub

Private Sub Text1_Change()
Me.Text1.Text = Format(Me.Text1.Text, "dd.mm.yyyy hh:mm")
End Sub

Private Sub Text2_Change()
Me.Text2.Text = Format(Me.Text2.Text, "dd.mm.yyyy hh:mm")
End Sub

'
Private Sub ws_Close()
'разблокировать кнопку
Me.btnTie.Enabled = True
End Sub

Private Sub ws_Connect()
' сообщить о выполнении соединения
Me.Label1.Caption = "Соединение прошло успешно"
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim data As String, data2, data4, dataX, fsize, size, sz, i
'On Error Resume Next
On Error GoTo errDA 'Resume Next
ws.GetData data, vbString ' получить все, что пришло от сервера
data2 = Left(data, 4) ' выделить команду от сервера
Select Case data2 ' проанализировать ее
    Case "rqst"  ' запрос на передачу файла с данными от сервера
        ' сформировать имя файла в который получать данные
        dataX = Right(data, Len(data) - (4))
        fsize = CLng(dataX) ' получение файла:
        If fsize > 0 Then ' если файл не пуст
            Reset
            'data4 = App.Path & "\Tempbase.csv" ' имя файла в который получить данные
            data4 = "Tempbase.csv" ' имя файла в который получить данные
            ' очистить файл от предыдущих записей
            Open data4 For Output As #1
            Print #1, ""
            Close #1
            ' открыть файл для новых записей
            Open data4 For Binary As #1
            ws.SendData "okay"      ' отправить запрос на получение файла
        End If
    Case "/reg"
        ' клиен успешно зарегистрирован
        ws.SendData "NICK " & Trim$(Mid(frmStart.Caption, _
            InStr(1, frmStart.Caption, "=", 1) + 1)): Exit Sub
    Case Else
        size = size + 1 '  считать количество блоков
        sz = size * 8 'chunk ' вычислять размер файла
        Put #1, , data ' записывать полученные данные в файл
        i = Seek(1)
        If i >= fsize Then
            Close #1 ' закрыть файл с полученными данными
            Label1.Caption = "Данные получены в размере: " & sz & "Kb"
            'Label2.Caption = ""
            size = 0: sz = 0
            ' записать дату последнего анализа
            Call WriteParameters("DateAnlz", Format(Time, "hh:mm"))
            Me.Adodc1.Refresh
'            Me.DataGrid1.DataMember = "command15"
'            Me.DataGrid1.Refresh
        End If
End Select
Exit Sub
errDA:
    MsgBox Err.Number & "-" & Err.Description
    Resume Next
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                    ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
                    CancelDisplay As Boolean)
'при ошибке сообщить о ней, закрыть соединение, разблокировать кнопку
Me.Label1.Caption = "Ошибка соединения #" & Number & "-" & Description
ws.Close ' закрыть соединение
Me.btnTie.Enabled = True
End Sub
'
Private Sub Form_Load()
Dim tmpPort As String, tmpServ As String
On Error GoTo Fload_err
'Form1.Text2 = Mid(frmStart.Caption, InStr(1, frmStart.Caption, "=", 1) + 1)
'Открываем рабочую область
With DataEnvironment1.rsCommand2 ' берем настройки
    If .State <> adStateOpen Then .Open
    .MoveFirst
    Do While Not .EOF
        Select Case .Fields("NameSet") ' взять название параметра настройки
        Case "PathAdmin"
            tmpServ = .Fields("Set") ' взять название сервера соединения
        Case "Port"
            tmpPort = .Fields("Set") ' взять номер порта
        Case "DateAnlz"
            Me.Text1.Text = .Fields("Set")
        End Select
        .MoveNext ' по всем параметрам настройки
    Loop
End With
Dialog.Text1.Text = tmpPort ' подготовить настройки
Dialog.Text2.Text = tmpServ ' подготовить назв.сервера
Me.Text1.Text = Now - 1: Me.Text2.Text = Now
'Me.DataGrid1.DataMember = "Command15"
'Me.DataGrid1.Refresh
Call btnTie_Click ' установить связь сразу
Exit Sub
Fload_err:
    Me.Label1.Caption = Err.Number & "->" & Err.Description
End Sub

