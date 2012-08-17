VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2010
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1187.574
   ScaleMode       =   0  'User
   ScaleWidth      =   3718.226
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtUserName 
      Height          =   315
      ItemData        =   "frmLogin.frx":0000
      Left            =   1440
      List            =   "frmLogin.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Новый пароль"
      Height          =   390
      Left            =   120
      TabIndex        =   5
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Вход"
      Default         =   -1  'True
      Height          =   390
      Left            =   1560
      TabIndex        =   3
      Top             =   1500
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Выход"
      Height          =   390
      Left            =   2760
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2325
   End
   Begin VB.Label Label1 
      Caption         =   "Новый пароль:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Имя:"
      Height          =   270
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Пароль:"
      Height          =   270
      Index           =   1
      Left            =   585
      TabIndex        =   1
      Top             =   615
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public LoginSucceeded As Boolean
Public strPSW As String ' пароль доступа к АСКУТЭ
Public strUser As String ' имя доступа к АСКУТЭ

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    'LoginSucceeded = False
    'Me.Hide
    Unload Me
End Sub
Private Sub cmdOK_Click()
Dim sHead As String, strz As String
Dim Control As Control, pos As Long
DeCode
strz = String$(255, " ")
GetPrivateProfileString txtUserName.Text, "PWD", "x", strz, 255, App.Path & "/Users"
strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
If txtPassword.Text = strz Then
    strz = String$(255, " ")
    GetPrivateProfileString txtUserName.Text, "username", "x", strz, 255, App.Path & "/Users"
    strUser = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
    strz = String$(255, " ")
    GetPrivateProfileString txtUserName.Text, "password", "x", strz, 255, App.Path & "/Users"
    strPSW = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
    strz = String$(255, " ")
    GetPrivateProfileString txtUserName.Text, "SGN", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
    'If strz <> Chr(42) Then
        strz = String$(255, " ")
        GetPrivateProfileString txtUserName.Text, "false", "x", strz, 255, App.Path & "/Users"
        strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
        pos = InStr(1, strz, ",", 1): If pos = 0 Then pos = Len(strz) + 1
        Do
            For Each Control In frmStart.Controls
                If Control.Name = Left(strz, pos - 1) Then
                    Control.Visible = False
                Else
                    If TypeOf Control Is CommandButton Then Control.Visible = True
                End If
            Next Control
            strz = Mid(strz, pos + 1)
            pos = InStr(1, strz, ",", 1): If pos = 0 Then pos = Len(strz) + 1
        Loop Until Len(strz) = 0
    'End If
    'LoginSucceeded = True
    'Form1.Text2 = txtUserName.Text ' имя открывшего пользователя
    frmStart.Caption = frmStart.Caption & "=" & txtUserName.Text
    Me.Hide
    Code
    frmStart.Show
Else
    Code
    MsgBox "Неверный пароль!", , "Пароль"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub Form_Load()
Dim sHead As String
DeCode
    Open "Users" For Input As #1
    Line Input #1, sHead
    Do While Not EOF(1)   ' Loop until end of file.
        Input #1, sHead
        If InStr(1, sHead, "[", 1) > 0 And InStr(1, sHead, "]", 1) > 0 Then
            sHead = Mid(sHead, 2, Len(sHead) - 2)
            txtUserName.AddItem sHead
        End If
    Loop
    Close #1
Code
End Sub

Private Sub Text1_Change()
    Dim lFileLen As Long
    Dim sHead As String
    ' Проверяем наличие файла
    On Error Resume Next
    lFileLen = Len(Text1.Text)
    ' Проверяем на наличие ошибок в имени файла
    If Err <> 0 Or lFileLen = 0 Or Len(Text1.Text) = 0 Then
        Exit Sub
    End If
    ' Проверяем по строке [Secret] в начале файла,
    ' что он зашифрован нашим классом
    Open "Users" For Binary As #1
    sHead = Space(8)
    Get #1, , sHead
    Close #1
End Sub


Public Function Code()
    Refresh
    Encrypt
    Text1_Change
End Function

Public Function DeCode()
    Refresh
    Decrypt
    Text1_Change
End Function

Private Sub Command1_Click()
Dim strz As String, sHead As String
Dim strLong As String
If Not Me.Text1.Visible Then
    Me.Command1.Caption = "Принять"
    Me.Label1.Visible = True
    Me.Text1.Visible = True
    Me.Text1.SetFocus
Else
    DeCode
    strz = String$(255, " ")
    GetPrivateProfileString txtUserName.Text, "PWD", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' удалить признак конца строки
    Code
    'check for correct password
    If txtPassword.Text <> strz Then
        MsgBox "Текущие значения имени пользователя и пароля неверны!"
        Me.Command1.Caption = "Новый пароль"
        Me.Command1.SetFocus
        Me.Label1.Visible = False
        Me.Text1.Visible = False
    Else
        DeCode
        WritePrivateProfileString txtUserName.Text, "PWD", Text1.Text, App.Path & "/Users"  ' Записываем в раздел  переменную
        Code
        MsgBox "Запомните новый пароль, иначе вы не сможете воспользоваться программой!"
        Me.Command1.Caption = "Новый пароль"
        Me.Command1.SetFocus
        Me.Label1.Visible = False
        Me.Text1.Visible = False
    End If
End If
End Sub
' выбирать пользователей только из списка
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
KeyAscii = vbKeyClear
End Sub
