VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Построитель запросов"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form2"
   ScaleHeight     =   7125
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Список таблиц и шаблонов"
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   7455
      Begin VB.OptionButton WinterMode 
         Caption         =   "Шаблон зимнего режима"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   2295
      End
      Begin VB.OptionButton SummerMode 
         Caption         =   "Шаблон летнего режима"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   5280
         Pattern         =   "*.ptn"
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton btnTmpl 
         Height          =   375
         Left            =   4560
         Picture         =   "Form2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Сопоставить шаблон таблице"
         Top             =   960
         Width           =   375
      End
      Begin VB.ListBox lstTable 
         Height          =   2010
         ItemData        =   "Form2.frx":018A
         Left            =   120
         List            =   "Form2.frx":018C
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "Изменить"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      ToolTipText     =   "редактировать шаблон запроса"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Сохранить"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Сохранить шаблон запроса в списке"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton btnOut 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "отказаться от поля и формулы в запросе"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton btnFrm 
      Caption         =   "...="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Задать формулу в запросе"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton btnTo 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "сортировать по возрастанию"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Выход"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Работа с таблицей образцом"
      Height          =   3255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdRule 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         ToolTipText     =   "Установить правила для строк "
         Top             =   1680
         Width           =   375
      End
      Begin VB.ListBox lstPtn 
         Height          =   2400
         ItemData        =   "Form2.frx":018E
         Left            =   2640
         List            =   "Form2.frx":0190
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton btnCond 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         ToolTipText     =   "Задать условие"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox TxtFrm 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   4575
      End
      Begin VB.ListBox lstStruc 
         Height          =   2400
         ItemData        =   "Form2.frx":0192
         Left            =   120
         List            =   "Form2.frx":0194
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Управление"
      Height          =   735
      Left            =   1440
      TabIndex        =   7
      Top             =   6360
      Width           =   4215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Примечание:"
      Height          =   3255
      Left            =   4920
      TabIndex        =   12
      Top             =   120
      Width           =   2535
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         Caption         =   "Изменение в формуле и условии запоминаются по нажатии клавиши ENTER после ввода формулы (условия) "
         Height          =   1095
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "Имя файла шаблона по умолчанию - default.ptn"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Описание структуры таблицы вызывается двойным щелчком манипулятора ""мышь"" по наименованию соответствующей таблицы"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Объявляем API для чтения
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Объявляем API для записи
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'
Private FileName
Private addFormula As Boolean
Private addCondition As Boolean
Private addRule As Boolean
'
Private Sub btnCond_Click()
Me.TxtFrm.Text = "Условие: "
Me.TxtFrm.SetFocus
addCondition = True
End Sub

Private Sub btnEdit_Click()
' переименование и открытие файла *.ptn для редактирования
On Error GoTo err_Edit
lstPtn.Clear
FileName = File1.FileName
Open FileName For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
    Input #1, MyString
    If InStr(1, MyString, "[", vbTextCompare) > 0 And InStr(1, MyString, "]", vbTextCompare) > 0 Then
        MyString = Mid(MyString, 2, Len(MyString) - 2)
        lstPtn.AddItem MyString
    End If
Loop
Close #1
Exit Sub
err_Edit:
If Err.Number = 75 Then MsgBox "Нет выбранного файла"
End Sub

'
Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnFrm_Click()
Me.TxtFrm.Text = "Формула: "
Me.TxtFrm.SetFocus
addFormula = True
End Sub
' функция проверки последнего номера формулы или условия
Function ContFN(what As String) As String
Dim pos As Long, MyString As String
Dim stroka As String
stroka = "0]"
Open FileName For Input As #1
Do While Not EOF(1)   '
    Input #1, MyString
    pos = InStr(1, MyString, what, vbTextCompare)
    If pos > 0 Then stroka = Mid(MyString, pos + Len(what), Len(MyString))
Loop
Close #1
stroka = Trim(Left(stroka, Len(stroka) - 1))
pos = CLng(stroka)
pos = pos + 1
ContFN = Trim(str(pos))
End Function

Private Sub btnOut_Click()
Dim stroka As String, MyString As String, MyString1 As String
Dim pos As Long
On Error GoTo errOut
' удалить строку из листинга
stroka = lstPtn.List(lstPtn.ListIndex)
lstPtn.RemoveItem (lstPtn.ListIndex)
MyString1 = ""
' отредактировать файл
Open FileName For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
    Input #1, MyString
    MyString1 = MyString1 & MyString & vbCrLf
Loop
Close #1
pos = InStr(1, MyString1, stroka, vbTextCompare)
If pos > 0 Then
    MyString = Mid(MyString1, 1, pos - 2)
    pos = InStr(pos, MyString1, "[", vbTextCompare)
    If pos <> 0 Then
        MyString1 = Mid(MyString1, pos)
        MyString = MyString & MyString1
    End If
End If
Open FileName For Output As #1
Print #1, MyString
Close #1
Exit Sub
errOut:
MsgBox Err.Description
End Sub

Private Sub btnSave_Click()
' переименование файла *.ptn для редактирования
On Error GoTo err_save
Dim strName As String
strName = InputBox("Введите имя файла для переименования текущего.")
If Len(strName) > 0 Then
    FileCopy FileName, strName & ".ptn"
    FileName = strName & ".ptn"
    File1.Refresh
End If
Exit Sub
err_save:
If Err.Number = 70 Then
MsgBox "Такой файл уже существует"
End If
End Sub

Private Sub btnTmpl_Click()
' сопоставить шаблон соответствующей таблице
If WinterMode.Value = True Then
    WritePrivateProfileString lstTable.List(lstTable.ListIndex()), "ШаблонЗима", _
                                               File1.FileName, App.Path & "/pattern.ini"
ElseIf SummerMode.Value = True Then
    WritePrivateProfileString lstTable.List(lstTable.ListIndex()), "ШаблонЛето", _
                                              File1.FileName, App.Path & "/pattern.ini"
End If
End Sub

' назначить сортировку по возрастанию
Private Sub btnTo_Click()
Dim strZ As String
' организовать просмотр формул, условий, сортировок
Me.TxtFrm.Text = ""
strZ = String$(255, " ")
GetPrivateProfileString lstPtn.List(lstPtn.ListIndex), "ORDER BY", "", strZ, 255, App.Path & "/" & FileName
strZ = Trim(strZ)
If InStr(1, strZ, "1") = 1 Then
    WritePrivateProfileString lstPtn.List(lstPtn.ListIndex), "ORDER BY", _
                                        "0", App.Path & "/" & FileName
Else
    WritePrivateProfileString lstPtn.List(lstPtn.ListIndex), "ORDER BY", _
                                        "1", App.Path & "/" & FileName
    Me.TxtFrm.Text = "ORDER BY " & lstPtn.List(lstPtn.ListIndex)
End If
End Sub

Private Sub cmdRule_Click()
Me.TxtFrm.Text = "Правило: "
Me.TxtFrm.SetFocus
addRule = True
End Sub

Private Sub Form_Load()
' вывести список таблиц
' взять настройки из ini-файла
Dim pos As Long, MyString As String
On Error GoTo errNi
Open "node.ini" For Input As #1
Do While Not EOF(1)   '
    Input #1, MyString
    pos = InStr(1, MyString, "Часовой архив", vbTextCompare)
    If pos > 0 Then
        pos = InStr(1, MyString, "=", vbTextCompare)
        MyString = Mid(MyString, pos + 1, Len(MyString))
        MyString = Trim(MyString)
        lstTable.AddItem MyString
    End If
    pos = InStr(1, MyString, "Суточный архив", vbTextCompare)
    If pos > 0 Then
        pos = InStr(1, MyString, "=", vbTextCompare)
        MyString = Mid(MyString, pos + 1, Len(MyString))
        MyString = Trim(MyString)
        lstTable.AddItem MyString
    End If
Loop
Close #1
' вывести список шаблонов
File1.Path = App.Path
File1.Refresh
' вывести состояние шаблона - default
FileName = "default.ptn"
Open FileName For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
    Input #1, MyString
    If InStr(1, MyString, "[", vbTextCompare) > 0 And InStr(1, MyString, "]", vbTextCompare) > 0 Then
        MyString = Mid(MyString, 2, Len(MyString) - 2)
        lstPtn.AddItem MyString
    End If
Loop
Close #1
'
Exit Sub
errNi:
MsgBox Err.Description
End Sub

Private Sub lstPtn_Click()
Dim strZ As String
' организовать просмотр формул, условий, сортировок и правил
Me.TxtFrm.Text = ""
strZ = String$(255, " ")
GetPrivateProfileString lstPtn.List(lstPtn.ListIndex), "ORDER BY", "", strZ, 255, App.Path & "/" & FileName
strZ = Trim(strZ)
If InStr(1, strZ, "1") = 1 Then _
    Me.TxtFrm.Text = "ORDER BY " & lstPtn.List(lstPtn.ListIndex)
strZ = String$(255, " ")
GetPrivateProfileString lstPtn.List(lstPtn.ListIndex), "fx", "", strZ, 255, App.Path & "/" & FileName
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
strZ = Trim(strZ)
If Len(strZ) > 0 Then Me.TxtFrm.Text = "Формула: " & strZ
strZ = String$(255, " ")
GetPrivateProfileString lstPtn.List(lstPtn.ListIndex), "Cnd", "", strZ, 255, App.Path & "/" & FileName
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
strZ = Trim(strZ)
If Len(strZ) > 0 Then Me.TxtFrm.Text = "Условие: " & strZ
strZ = String$(255, " ")
GetPrivateProfileString lstPtn.List(lstPtn.ListIndex), "Rule", "", strZ, 255, App.Path & "/" & FileName
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
strZ = Trim(strZ)
If Len(strZ) > 0 Then Me.TxtFrm.Text = "Правило: " & strZ
End Sub

Private Sub lstStruc_DblClick()
lstPtn.AddItem lstStruc.List(lstStruc.ListIndex)
WritePrivateProfileString lstStruc.List(lstStruc.ListIndex), "", _
                                        "", App.Path & "/" & FileName
End Sub

'
Private Sub lstTable_DblClick()
Dim fld As Field
 On Error GoTo Kosjak
    Form1.Data1.DatabaseName = Form1.Text1.Text
    ' запрос служебной таблицы для работы пользователей
    Form1.Data1.RecordSource = lstTable.List(lstTable.ListIndex())
    Form1.Data1.Refresh
    lstStruc.Clear
    With Form1.Data1.Recordset
        .Requery
        For Each fld In .Fields
            lstStruc.AddItem fld.Name
        Next
    End With
    lstStruc.Refresh
    Form1.Data1.DatabaseName = ""
    Form1.Data1.RecordSource = ""
    Form1.Data1.Refresh
Exit Sub
Kosjak:
    MsgBox Err.Description
End Sub
'
Private Sub TxtFrm_KeyDown(KeyCode As Integer, Shift As Integer)
'mstrTable = "SELECT ДатаВремя, W1,W2,m1,m2, 0 as W3, 0 as m3,t1,t2, Tнс,0 as xНС," & _
'                                    "P1,P2 FROM " & nameOfNode & " WHERE ДатаВремя " & _
'                                    "BETWEEN #" & SQLData(CDate(firstP) - 1) & " 23:00# AND #" & _
'                                    SQLData(CDate(secondP)) & " 23:00# ORDER BY ДатаВремя;"
'mstrTable = "SELECT * FROM " & nameOfNode & " WHERE ДатаВремя BETWEEN #" & _
'            SQLData(CDate(firstP)) & " 00:00# AND #" & SQLData(CDate(secondP)) & _
'            IIf(StrComp(KindOfArh, "Часовой") = 0, " 23:59#", " 00:00#") & " ORDER BY ДатаВремя;"
'
'Print #filenum, "DateTimes"; ";"; "W1"; ";"; "M1"; ";"; "t1"; ";"; "W2"; ";"; "M2"; _
'    ";"; "t2"; ";"; "W3"; ";"; "M3"; ";"; "T"; ";"; "Tp"; ";"; "P1"; ";"; "P2"
'Print #filenum, Format(.Fields(0), "dd.mm.yy hh:mm"); ";"; dbl2W1; ";"; Format(dbl2M1, "#0.000"); ";"; _
'.Fields(7) / 100; ";"; dbl2W2; ";"; Format(dbl2M2, "#0.000"); ";"; .Fields(8) / 100; ";"; _
'    IIf(KindOfMode = "1", dbl2W1 + dbl2W2, dbl2W1 - dbl2W2); ";"; _
'    IIf(KindOfMode = "1", dbl2M1 + dbl2M2, dbl2M1 - dbl2M2); ";"; _
'    IIf(KindOfArh = "Часовой", 60 - dbl2Time, 1440 - dbl2Time); ";"; dbl2Time; ";"; _
'    Format(.Fields(11), "#0.000"); ";"; Format(.Fields(12), "#0.000")
Dim str As String, frm As String
If KeyCode = vbKeyReturn Then
    If addFormula Or addCondition Then
        If InStr(1, TxtFrm.Text, "Условие") > 0 Then
            ' проверить имя условия в файле *.ptn
            str = ContFN("Условие")
            lstPtn.AddItem "Условие" & str
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString "Условие" & str, "Cnd", frm, App.Path & "/" & FileName
        ElseIf InStr(1, TxtFrm.Text, "Формула") > 0 Then
            ' проверить имя формулы в файле *.ptn
            str = ContFN("Формула")
            lstPtn.AddItem "Формула" & str
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString "Формула" & str, "fx", frm, App.Path & "/" & FileName
        End If
        addFormula = False
        addCondition = False
    ElseIf addRule Then
            str = ContFN("Правило")
            lstPtn.AddItem "Правило" & str
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString "Правило" & str, "Rule", frm, App.Path & "/" & FileName
    Else
        If InStr(1, TxtFrm.Text, "Условие") > 0 Then
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString lstPtn.List(lstPtn.ListIndex), "Cnd", frm, App.Path & "/" & FileName
        ElseIf InStr(1, TxtFrm.Text, "Формула") > 0 Then
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString lstPtn.List(lstPtn.ListIndex), "fx", frm, App.Path & "/" & FileName
        ElseIf InStr(1, TxtFrm.Text, "Правило") > 0 Then
            frm = Trim(Mid(TxtFrm.Text, InStr(1, TxtFrm.Text, ":", vbTextCompare) + 1))
            WritePrivateProfileString lstPtn.List(lstPtn.ListIndex), "Rule", frm, App.Path & "/" & FileName
        End If
    End If
End If
End Sub
