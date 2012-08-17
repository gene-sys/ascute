VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Пользователи"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form3"
   ScaleHeight     =   5160
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Тип архива"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form3.frx":0000
         Left            =   120
         List            =   "Form3.frx":000D
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Поле сообщений:"
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   6375
      Begin VB.Label lblMsg 
         BackColor       =   &H80000018&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Выход"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton btnList 
      Height          =   375
      Left            =   3000
      Picture         =   "Form3.frx":002A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Составить список узлов"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton btnUser 
      Height          =   375
      Left            =   3000
      Picture         =   "Form3.frx":055C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Сопоставить пользователей узлу"
      Top             =   1440
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Пользователи и узлы"
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin VB.ListBox lstUsers 
         Height          =   2310
         ItemData        =   "Form3.frx":06E6
         Left            =   3600
         List            =   "Form3.frx":06E8
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Height          =   2400
         ItemData        =   "Form3.frx":06EA
         Left            =   120
         List            =   "Form3.frx":06EC
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "Form3"
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
Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnList_Click()
Dim strTable As String, strZ As String
Dim chas As String, sut As String
Dim I As Long
'
' взять настройки из ini-файла
strZ = String$(255, " ")
GetPrivateProfileString "StartNode", "query", "x", strZ, 255, App.Path & "/set.ini"
strZ = Trim(strZ)
'
' 1 открыть таблицу с описанием узлов в spdef.mdb
On Error GoTo btnList
If strZ <> "x" Then
    lblMsg.Caption = "Подождите до завершения процесса..."
    Form3.Refresh
    strTable = strZ
    Form1.Data2.DatabaseName = Form1.Text3.Text
    ' запрос служебной таблицы для работы пользователей
    Form1.Data2.RecordSource = strTable
    Form1.Data2.Refresh
    List1.Clear
    If Not (Form1.Data2.Recordset.BOF And Form1.Data2.Recordset.EOF) Then
        With Form1.Data2.Recordset
            .Requery
            .MoveFirst
            ' 2 сформировать файл описаний node.ini
            Do While Not .EOF
                chas = .Fields(2) & Chr(0): sut = .Fields(3) & Chr(0)
                ' записать данные в файл
                WritePrivateProfileString .Fields(0), "Часовой архив", chas, App.Path & "/node.ini"
                WritePrivateProfileString .Fields(0), "Суточный архив", sut, App.Path & "/node.ini"
                ' 3 отобразить результат на экране
                List1.AddItem .Fields(0)
                .MoveNext
            Loop
        End With
        lblMsg.Caption = ""
    End If
    Form1.Data2.DatabaseName = ""
    Form1.Data2.RecordSource = ""
    Form1.Data2.Refresh
End If
Exit Sub
btnList:
MsgBox Err.Description
End Sub
'
Private Sub btnUser_Click()
Dim I As Long, str As String
str = ""
   For I = 0 To lstUsers.ListCount - 1
      If lstUsers.Selected(I) Then
        str = str & "," & lstUsers.List(I)
      End If
   Next I
   str = Right(str, Len(str) - 1)
WritePrivateProfileString List1.List(List1.ListIndex), "Пользователь", str, App.Path & "/node.ini"
End Sub


Private Sub Combo1_Click()
Dim strZ As String
If Combo1.ListIndex = 0 Then strZ = "PAR"
If Combo1.ListIndex = 1 Then strZ = "HV"
If Combo1.ListIndex = 2 Then strZ = "Journal"
WritePrivateProfileString List1.List(List1.ListIndex), "Тип архива", strZ, App.Path & "/node.ini"
End Sub


Private Sub Form_Load()
Dim MyString As String, mloc
' вывести список узлов
' взять настройки из ini-файла
Open "node.ini" For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
    Input #1, MyString
    If InStr(1, MyString, "[", vbTextCompare) > 0 And InStr(1, MyString, "]", vbTextCompare) > 0 Then
        MyString = Mid(MyString, 2, Len(MyString) - 2)
        List1.AddItem MyString
    End If
Loop
Close #1
' вывести список пользователей
Open "users.lst" For Input As #1
Do While Not EOF(1)   ' Loop until end of file.
   Input #1, MyString
   lstUsers.AddItem MyString
Loop
Close #1
' вывести список шаблонов
End Sub

