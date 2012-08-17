VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStart 
   Caption         =   "Учет расхода энергии v2.00"
   ClientHeight    =   7875
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10080
   LinkTopic       =   "Form3"
   ScaleHeight     =   7875
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7620
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:09"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "27.11.2009"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman CYR"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu1 
      Caption         =   "Файл"
      Begin VB.Menu mnuSoed 
         Caption         =   "Соединение"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRaz 
         Caption         =   "Рассоединение"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu11 
         Caption         =   "Выход"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu2 
      Caption         =   "Отчеты"
      Begin VB.Menu mnuKT 
         Caption         =   "Контроль температуры"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnu21 
         Caption         =   "Получить"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu22 
         Caption         =   "Анализ"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu23 
         Caption         =   "Данные"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu24 
         Caption         =   "Анализ"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuA1 
      Caption         =   "Архивы"
      Begin VB.Menu mnuA11 
         Caption         =   "Графики"
      End
      Begin VB.Menu mnuA12 
         Caption         =   "Тренды и диаграммы"
      End
   End
   Begin VB.Menu mnu4 
      Caption         =   "Настройки"
      Begin VB.Menu mnu41 
         Caption         =   "Общие"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnu42 
         Caption         =   "Отчеты"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuTG 
         Caption         =   "Температурный график"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu43 
         Caption         =   "Справочник подразделений"
      End
      Begin VB.Menu mnu44 
         Caption         =   "Справочник УУ ТЭ"
      End
      Begin VB.Menu mnu45 
         Caption         =   "Справочник оборудования"
      End
      Begin VB.Menu mnu46 
         Caption         =   "Справочник местоположений"
      End
      Begin VB.Menu mnu47 
         Caption         =   "Справочник наименований УУ"
      End
   End
   Begin VB.Menu mnu6 
      Caption         =   "Профиль"
      Begin VB.Menu mnu61 
         Caption         =   "Взлет-СП"
      End
      Begin VB.Menu mnu62 
         Caption         =   "АСКУТЭ"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu5 
      Caption         =   "Администрирование"
      Begin VB.Menu mnu51 
         Caption         =   "Защита"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oConnectionState As New CWinInetConnection
'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Disconnect Net Drive"
            'ToDo: Add 'Disconnect Net Drive' button code.
            MsgBox "Add 'Disconnect Net Drive' button code."
        Case "Delete"
            'ToDo: Add 'Delete' button code.
            MsgBox "Add 'Delete' button code."
        Case "Open"
            'ToDo: Add 'Open' button code.
            MsgBox "Add 'Open' button code."
        Case "Properties"
            'ToDo: Add 'Properties' button code.
            MsgBox "Add 'Properties' button code."
        Case "Sort Ascending"
            'ToDo: Add 'Sort Ascending' button code.
            MsgBox "Add 'Sort Ascending' button code."
        Case "Sum"
            'ToDo: Add 'Sum' button code.
            MsgBox "Add 'Sum' button code."
        Case "Up One Level"
            'ToDo: Add 'Up One Level' button code.
            MsgBox "Add 'Up One Level' button code."
        Case "View Large Icons"
            'ToDo: Add 'View Large Icons' button code.
            MsgBox "Add 'View Large Icons' button code."
    End Select
End Sub


Private Sub Form_load()
' открыть setting.mdb и выяснить какой профиль сделать активным
' prflVzjot или prflAskute
If ReadNParam("PrflVzljot") = "1" Then Call mnu61_Click
If ReadNParam("PrflAskute") = "1" Then Call mnu62_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If MsgBox("Действительно выйти?", vbQuestion + vbYesNo) = vbYes Then
    oConnectionState.HangUp
    ' выполнить корректное закрытие всех открытых форм
    For Each Form In Forms
        Unload Form
    Next Form
    'Unload Me
'End If
End Sub

Private Sub mnu11_Click()
Unload Me ' выгрузить форму
End Sub

Private Sub mnu21_Click()
 Form1.Show 'vbModal, Me ' форма получения данных
End Sub

Private Sub mnu22_Click()
' сформировать анализ состояния узлов учета
frmAnaliz.Show
End Sub

Private Sub mnu23_Click()
frmData.Show
End Sub

'
Private Sub mnu41_Click()
Dialog.Show ' форма основных настроек
End Sub

Private Sub mnu42_Click()
frmOptions.Show ' открыть форму настроек отчетов
End Sub

Private Sub mnu43_Click()
frmAReport.Show
End Sub

Private Sub mnu44_Click()
frmUU.Show
End Sub

Private Sub mnu45_Click()
frmOb.Show
End Sub

Private Sub mnu46_Click()
frmMest.Show
End Sub

Private Sub mnu47_Click()
frmNUZ.Show
End Sub

Private Sub mnu51_Click()
 Form2.Show ' форма настроек получения данных
End Sub

Private Sub mnu61_Click()
mnu61.Checked = True
mnu62.Checked = False
' скрыть не нужное и открыть нужное
mnu23.Visible = False
mnu24.Visible = False
mnu43.Visible = False
mnu44.Visible = False
mnu45.Visible = False
mnu46.Visible = False
mnu47.Visible = False
mnu21.Visible = True
mnu22.Visible = True
mnuSoed.Visible = True
mnuRaz.Visible = True
' изменить в setting.mdb обозначение пункта
Call WriteParameters("PrflVzljot", "1")
Call WriteParameters("PrflAskute", "0")
End Sub

Private Sub mnu62_Click()
mnu61.Checked = False
mnu62.Checked = True
' скрыть не нужное и открыть нужное
mnu23.Visible = True
mnu24.Visible = True
mnu43.Visible = True
mnu44.Visible = True
mnu45.Visible = True
mnu46.Visible = True
mnu47.Visible = True
mnu21.Visible = False
mnu22.Visible = False
mnuSoed.Visible = False
mnuRaz.Visible = False
' изменить в setting.mdb обозначение пункта
Call WriteParameters("PrflVzljot", "0")
Call WriteParameters("PrflAskute", "1")
End Sub

Private Sub mnuA11_Click()
frmGraph.Show
End Sub

' открытие функции для организации контроля температуры
Private Sub mnuKT_Click()
mnuKT.Checked = Not mnuKT.Checked
End Sub

Private Sub mnuRaz_Click()
If MsgBox("Действительно разорвать связь?", _
                vbQuestion + vbYesNo) = vbYes Then _
                                oConnectionState.HangUp
End Sub

Private Sub mnuSoed_Click()
Dim tmp As String
With DataEnvironment1.rsCommand2 ' берем настройки
    If .State <> adStateOpen Then .Open
    .Requery:  .MoveFirst
    Do While Not .EOF
        ' взять название параметра настройки-взять наименование соединения
        If .Fields("NameSet") = "Connect" Then
            tmp = .Fields("Set"): Exit Do
        End If
        .MoveNext ' по всем параметрам настройки
    Loop
End With
' выполнить соединение с сервером
'oConnectionState.Dial Me.hWnd, tmp, DF_FORCE_ONLINE, False
oConnectionState.Dial 0, tmp, DF_FORCE_ONLINE, False
End Sub

Private Sub mnuSql_Click()
First.Show
End Sub

' форма контроля температурного графика
Private Sub mnuTG_Click()
    frmTG.Show
End Sub
