VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "������� ������"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6945
   LinkTopic       =   "Form2"
   ScaleHeight     =   3840
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "������ � ������"
      Height          =   1215
      Left            =   2760
      TabIndex        =   13
      Top             =   2040
      Width           =   4095
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtPsw 
         Height          =   405
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "������������"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "������"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�������"
      Height          =   375
      Left            =   3945
      TabIndex        =   9
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���������"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      ItemData        =   "Form2.frx":0000
      Left            =   120
      List            =   "Form2.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������������� ������� �������"
      Height          =   2055
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "�������������"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "������ ����"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "���"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "������� ������"
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lIndx As Long
'
Private Sub Command1_Click()
Dim sHead As String, i As Long
If MsgBox("��������� ?", vbQuestion + vbYesNo) = vbYes Then
    Decrypt ' ������������
    '' ���������� � ������  ����������
    WritePrivateProfileString Me.Text1.Text, "PWD", Text2.Text, App.Path & "/Users"
    If Check1.Value = 1 Then _
        WritePrivateProfileString Me.Text1.Text, _
            "SGN", "*", App.Path & "/Users"  ' ���������� � ������  ����������
    WritePrivateProfileString Me.Text1.Text, "false", _
        Text3.Text, App.Path & "/Users"  ' ���������� � ������  ����������
    WritePrivateProfileString Text1.Text, "username", _
        Me.txtUser.Text, App.Path & "/Users"  ' ���������� � ������  ����������
    WritePrivateProfileString Text1.Text, "password", _
        Me.txtPsw.Text, App.Path & "/Users"  ' ���������� � ������  ����������
    Encrypt ' �����������
    Me.List1.Clear ' �������� ������ ������ �������������
    OpneNew ' ����������� ������ ������������� ������
End If
End Sub

Private Sub Command2_Click()
If MsgBox("������������� �����?", vbQuestion + vbYesNo) _
                    = vbYes Then Unload Me ' ����� �� �������
End Sub
' ������� ������� ������� ������
Private Sub Command4_Click()
Dim strX As String, Mstr As String
Dim pos As Long
If MsgBox("�� ������������� ������ ������� ������� ������", _
                            vbQuestion + vbYesNo) = vbYes Then
    Decrypt ' ������������
    Open "Users" For Input As #1 ' ������� ��� ������
    Do While Not EOF(1)
        Line Input #1, strX ' ������� ������
        Mstr = Mstr & strX & vbCrLf ' ����������� ������
    Loop
    Close #1 ' ������� ����
    pos = InStr(1, Mstr, "[" & Text1.Text, 1) ' ����� ��������� ������
    strX = Left(Mstr, pos - 1)
    pos = InStr(pos + 1, Mstr, "[", 1) ' ����� ����� �� �������� �������
    ' ������������ ����� ������ ����� ��������
    If pos > 0 Then strX = strX & Mid(Mstr, pos) & vbCrLf
    Open "Users" For Output As #1 ' ������� ���� ��� ������
    Print #1, strX ' �������� ����� ������
    Close #1 ' ������� ����
    Encrypt ' �����������
    Me.List1.Clear ' �������� ������ ������ �������������
    OpneNew ' ����������� ������ ������������� ������
End If
End Sub

Private Sub Form_load()
OpneNew ' ������� ������ �������������
End Sub
' ��������� � ������� ������ �������������
Sub OpneNew()
Dim sHead As String
Decrypt ' ������������
    Open "Users" For Input As #1 ' ������� ���� ��� ������
    Line Input #1, sHead ' ������� 1�� ������
    Do While Not EOF(1)   ' ���� ���� �� ����� �����
        Input #1, sHead ' ������� ��������� ������
        ' �������� ��� �����������
        If InStr(1, sHead, "[", 1) > 0 And _
                InStr(1, sHead, "]", 1) > 0 Then
            sHead = Mid(sHead, 2, Len(sHead) - 2) ' �������� � ������
            List1.AddItem sHead ' �������� ������ � ������
        End If
    Loop ' ����� �����
    Close #1 ' ������� ����
Encrypt ' �����������
' ���������� ������� ������������ � ������
Me.Text1.Text = Me.List1.List(0)
End Sub
' ���������� ��������� ������������ ��� ��������������
Private Sub List1_Click()
Dim strz As String
lIndx = List1.ListIndex ' ������ ������������ � ������
Me.Text1.Text = Me.List1.List(lIndx) ' ���������� ��� ������������ � ������
' ����� ��������� �� ini-�����
Decrypt ' ������������
    strz = String$(255, " ") ' ����������� ������
    ' ����� ������ ������������
    GetPrivateProfileString Me.List1.List(lIndx), "PWD", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' ������� ������� ����� ������
    Me.Text2.Text = strz ' ���������� ������
    strz = String$(255, " ") ' ����������� ������
    ' �������� ������� ���� (��� ����������) ��������������
    GetPrivateProfileString Me.List1.List(lIndx), "SGN", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' ������� ������� ����� ������
    ' ���������� ������
    If strz = Chr(42) Then Me.Check1.Value = 1 Else Me.Check1.Value = 0
    strz = String$(255, " ") ' ����������� ������
    ' �������� ������ ����������� ������� ���� (���� ����)
    GetPrivateProfileString Me.List1.List(lIndx), "false", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' ������� ������� ����� ������
    Me.Text3.Text = strz ' ���������� ������
    strz = String$(255, " ") ' ����������� ������
    ' �������� ��� ������� � ������
    GetPrivateProfileString Me.List1.List(lIndx), "username", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' ������� ������� ����� ������
    Me.txtUser.Text = strz ' ���������� ������
    strz = String$(255, " ") ' ����������� ������
    ' �������� ������ ������� � ������
    GetPrivateProfileString Me.List1.List(lIndx), "password", "x", strz, 255, App.Path & "/Users"
    strz = Mid(strz, 1, InStr(1, strz, Chr(0)) - 1) ' ������� ������� ����� ������
    Me.txtPsw.Text = strz ' ���������� ������
Encrypt ' �����������
End Sub
