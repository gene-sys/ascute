VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Настройки печати"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   3975
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Выход"
      Height          =   375
      Left            =   4185
      TabIndex        =   2
      Top             =   3570
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Применить"
      Height          =   375
      Left            =   3105
      TabIndex        =   1
      Top             =   3570
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3315
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   5847
      _Version        =   327681
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Тепло"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Пар"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Вода"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Стоки"
      TabPicture(3)   =   "frmOptions.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Граница листа:"
         Height          =   1665
         Left            =   240
         TabIndex        =   42
         Top             =   510
         Width           =   4755
         Begin VB.TextBox txtSBottom 
            DataField       =   "Niz"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command14"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   1455
            TabIndex        =   46
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtSTop 
            DataField       =   "Verh"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command14"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   1455
            TabIndex        =   45
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtSLeft 
            DataField       =   "Levo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command14"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   3165
            TabIndex        =   44
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtSRight 
            DataField       =   "Pravo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command14"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   3165
            TabIndex        =   43
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Сверху"
            Height          =   375
            Left            =   675
            TabIndex        =   50
            Top             =   450
            Width           =   645
         End
         Begin VB.Label Label20 
            Caption         =   "Снизу"
            Height          =   345
            Left            =   675
            TabIndex        =   49
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label19 
            Caption         =   "Слева"
            Height          =   315
            Left            =   2535
            TabIndex        =   48
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label18 
            Caption         =   "Справа"
            Height          =   345
            Left            =   2505
            TabIndex        =   47
            Top             =   990
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Height          =   825
         Left            =   240
         TabIndex        =   39
         Top             =   2175
         Width           =   4755
         Begin VB.TextBox txtSSize 
            Alignment       =   1  'Right Justify
            DataField       =   "Shrift"
            DataMember      =   "Command14"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   2550
            TabIndex        =   40
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label17 
            Caption         =   "Размер шрифта:"
            Height          =   405
            Left            =   1080
            TabIndex        =   41
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   -74760
         TabIndex        =   36
         Top             =   2175
         Width           =   4755
         Begin VB.TextBox txtVSize 
            Alignment       =   1  'Right Justify
            DataField       =   "Shrift"
            DataMember      =   "Command13"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   2550
            TabIndex        =   37
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label16 
            Caption         =   "Размер шрифта:"
            Height          =   405
            Left            =   1080
            TabIndex        =   38
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Граница листа:"
         Height          =   1665
         Left            =   -74760
         TabIndex        =   27
         Top             =   510
         Width           =   4755
         Begin VB.TextBox txtVRight 
            DataField       =   "Pravo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command13"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   3165
            TabIndex        =   31
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtVLeft 
            DataField       =   "Levo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command13"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   3165
            TabIndex        =   30
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtVTop 
            DataField       =   "Verh"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command13"
            DataSource      =   "DataEnvironment1"
            Height          =   330
            Left            =   1455
            TabIndex        =   29
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtVBottom 
            DataField       =   "Niz"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command13"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   1455
            TabIndex        =   28
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label15 
            Caption         =   "Справа"
            Height          =   345
            Left            =   2505
            TabIndex        =   35
            Top             =   990
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "Слева"
            Height          =   315
            Left            =   2535
            TabIndex        =   34
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label13 
            Caption         =   "Снизу"
            Height          =   345
            Left            =   675
            TabIndex        =   33
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label12 
            Caption         =   "Сверху"
            Height          =   375
            Left            =   675
            TabIndex        =   32
            Top             =   450
            Width           =   645
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Граница листа:"
         Height          =   1665
         Left            =   -74760
         TabIndex        =   18
         Top             =   510
         Width           =   4755
         Begin VB.TextBox txtPBottom 
            DataField       =   "Niz"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command12"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   1455
            TabIndex        =   22
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtPTop 
            DataField       =   "Verh"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command12"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   1455
            TabIndex        =   21
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtPLeft 
            DataField       =   "Levo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command12"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   3165
            TabIndex        =   20
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtPRight 
            DataField       =   "Pravo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command12"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   3150
            TabIndex        =   19
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label11 
            Caption         =   "Сверху"
            Height          =   375
            Left            =   675
            TabIndex        =   26
            Top             =   450
            Width           =   645
         End
         Begin VB.Label Label10 
            Caption         =   "Снизу"
            Height          =   345
            Left            =   675
            TabIndex        =   25
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label9 
            Caption         =   "Слева"
            Height          =   315
            Left            =   2535
            TabIndex        =   24
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label8 
            Caption         =   "Справа"
            Height          =   345
            Left            =   2505
            TabIndex        =   23
            Top             =   990
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Height          =   825
         Left            =   -74760
         TabIndex        =   15
         Top             =   2175
         Width           =   4755
         Begin VB.TextBox txtPSize 
            Alignment       =   1  'Right Justify
            DataField       =   "Shrift"
            DataMember      =   "Command12"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   2580
            TabIndex        =   16
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label7 
            Caption         =   "Размер шрифта:"
            Height          =   405
            Left            =   1080
            TabIndex        =   17
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Height          =   825
         Left            =   -74760
         TabIndex        =   12
         Top             =   2175
         Width           =   4755
         Begin VB.TextBox txtSize 
            Alignment       =   1  'Right Justify
            DataField       =   "Shrift"
            DataMember      =   "Command11"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   2550
            TabIndex        =   13
            Top             =   255
            Width           =   1125
         End
         Begin VB.Label Label5 
            Caption         =   "Размер шрифта:"
            Height          =   405
            Left            =   1080
            TabIndex        =   14
            Top             =   315
            Width           =   1395
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Граница листа:"
         Height          =   1665
         Left            =   -74760
         TabIndex        =   3
         Top             =   510
         Width           =   4755
         Begin VB.TextBox txtRight 
            DataField       =   "Pravo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command11"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   3165
            TabIndex        =   10
            Top             =   960
            Width           =   915
         End
         Begin VB.TextBox txtLeft 
            DataField       =   "Levo"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command11"
            DataSource      =   "DataEnvironment1"
            Height          =   345
            Left            =   3165
            TabIndex        =   8
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtTop 
            DataField       =   "Verh"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command11"
            DataSource      =   "DataEnvironment1"
            Height          =   330
            Left            =   1425
            TabIndex        =   5
            Top             =   390
            Width           =   915
         End
         Begin VB.TextBox txtBottom 
            DataField       =   "Niz"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   0
            EndProperty
            DataMember      =   "Command11"
            DataSource      =   "DataEnvironment1"
            Height          =   375
            Left            =   1455
            TabIndex        =   4
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "Справа"
            Height          =   345
            Left            =   2505
            TabIndex        =   11
            Top             =   990
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Слева"
            Height          =   315
            Left            =   2535
            TabIndex        =   9
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label2 
            Caption         =   "Снизу"
            Height          =   345
            Left            =   675
            TabIndex        =   7
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "Сверху"
            Height          =   375
            Left            =   675
            TabIndex        =   6
            Top             =   450
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Private Sub Command1_Click()
'' сохранить
If SSTab1.Tab = 0 Then _
    DataEnvironment1.rsCommand11.Update
If SSTab1.Tab = 1 Then _
    DataEnvironment1.rsCommand12.Update
If SSTab1.Tab = 2 Then _
    DataEnvironment1.rsCommand13.Update
If SSTab1.Tab = 3 Then _
    DataEnvironment1.rsCommand14.Update
'Call WritePrintParam("тепло", "top", _
'        Mid(Me.txtTop.Text, 1, InStr(1, Me.txtTop.Text, " см") - 1))
'Call WritePrintParam("тепло", "bottom", _
'        Mid(Me.txtBottom.Text, 1, InStr(1, Me.txtBottom.Text, " см") - 1))
'Call WritePrintParam("тепло", "left", _
'        Mid(Me.txtLeft.Text, 1, InStr(1, Me.txtLeft.Text, " см") - 1))
'Call WritePrintParam("тепло", "right", _
'        Mid(Me.txtRight.Text, 1, InStr(1, Me.txtRight.Text, " см") - 1))
''Call WritePrintParam("тепло", "linesize", _
''        Mid(Me.txtLine.Text, 1, InStr(1, Me.txtLine.Text, " см") - 1))
'Call WritePrintParam("тепло", "fontsize", Me.txtSize)
'' пар
'Call WritePrintParam("пар", "top", _
'        Mid(Me.txtPTop.Text, 1, InStr(1, Me.txtPTop.Text, " см") - 1))
'Call WritePrintParam("пар", "bottom", _
'        Mid(Me.txtPBottom.Text, 1, InStr(1, Me.txtPBottom.Text, " см") - 1))
'Call WritePrintParam("пар", "left", _
'        Mid(Me.txtPLeft.Text, 1, InStr(1, Me.txtPLeft.Text, " см") - 1))
'Call WritePrintParam("пар", "right", _
'        Mid(Me.txtPRight.Text, 1, InStr(1, Me.txtPRight.Text, " см") - 1))
'Call WritePrintParam("пар", "fontsize", Me.txtPSize)
'' вода
'Call WritePrintParam("Вода", "top", _
'        Mid(Me.txtVTop.Text, 1, InStr(1, Me.txtVTop.Text, " см") - 1))
'Call WritePrintParam("Вода", "bottom", _
'        Mid(Me.txtVBottom.Text, 1, InStr(1, Me.txtVBottom.Text, " см") - 1))
'Call WritePrintParam("Вода", "left", _
'        Mid(Me.txtVLeft.Text, 1, InStr(1, Me.txtVLeft.Text, " см") - 1))
'Call WritePrintParam("Вода", "right", _
'        Mid(Me.txtVRight.Text, 1, InStr(1, Me.txtVRight.Text, " см") - 1))
'Call WritePrintParam("Вода", "fontsize", Me.txtVSize)
'' Стоки
'Call WritePrintParam("Стоки", "top", _
'        Mid(Me.txtSTop.Text, 1, InStr(1, Me.txtSTop.Text, " см") - 1))
'Call WritePrintParam("Стоки", "bottom", _
'        Mid(Me.txtSBottom.Text, 1, InStr(1, Me.txtSBottom.Text, " см") - 1))
'Call WritePrintParam("Стоки", "left", _
'        Mid(Me.txtSLeft.Text, 1, InStr(1, Me.txtSLeft.Text, " см") - 1))
'Call WritePrintParam("Стоки", "right", _
'        Mid(Me.txtSRight.Text, 1, InStr(1, Me.txtSRight.Text, " см") - 1))
'Call WritePrintParam("Стоки", "fontsize", Me.txtSSize)
End Sub

Private Sub Command2_Click()
' выйти без сохранения
Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 44 Then
'    If TypeOf Me.ActiveControl Is TextBox Then
'        ' "Должна быть использована точка"
'        KeyAscii = Asc(Chr(46))
'    End If
'ElseIf KeyAscii = 8 Or KeyAscii = 46 Then
'' пропускаем
'ElseIf (KeyAscii < 48) Or (KeyAscii > 57) Then
'    ' ругаемся на неправильный ввод
'    MsgBox "Вводите только числа"
'    KeyAscii = 0
'End If
End Sub

Private Sub Form_Load()
'' тепло
'Me.txtTop = RemakeS(ReadPrintParam("Тепло", "top")) & " см"
'Me.txtBottom = RemakeS(ReadPrintParam("Тепло", "bottom")) & " см"
'Me.txtLeft = RemakeS(ReadPrintParam("Тепло", "left")) & " см"
'Me.txtRight = RemakeS(ReadPrintParam("Тепло", "right")) & " см"
''Me.txtLine = RemakeS(ReadPrintParam("Тепло", "linesize")) & " см"
'Me.txtSize = ReadPrintParam("Тепло", "fontsize")
'' пар
'Me.txtPTop = RemakeS(ReadPrintParam("Пар", "top")) & " см"
'Me.txtPBottom = RemakeS(ReadPrintParam("Пар", "bottom")) & " см"
'Me.txtPLeft = RemakeS(ReadPrintParam("Пар", "left")) & " см"
'Me.txtPRight = RemakeS(ReadPrintParam("Пар", "right")) & " см"
'Me.txtPSize = ReadPrintParam("Пар", "fontsize")
'' вода
'Me.txtVTop = RemakeS(ReadPrintParam("Вода", "top")) & " см"
'Me.txtVBottom = RemakeS(ReadPrintParam("Вода", "bottom")) & " см"
'Me.txtVLeft = RemakeS(ReadPrintParam("Вода", "left")) & " см"
'Me.txtVRight = RemakeS(ReadPrintParam("Вода", "right")) & " см"
'Me.txtVSize = ReadPrintParam("Вода", "fontsize")
'' Стоки
'Me.txtSTop = RemakeS(ReadPrintParam("Стоки", "top")) & " см"
'Me.txtSBottom = RemakeS(ReadPrintParam("Стоки", "bottom")) & " см"
'Me.txtSLeft = RemakeS(ReadPrintParam("Стоки", "left")) & " см"
'Me.txtSRight = RemakeS(ReadPrintParam("Стоки", "right")) & " см"
'Me.txtSSize = ReadPrintParam("Стоки", "fontsize")
End Sub
'
' тепло
'Private Sub txtline_GotFocus()
'Me.txtLine.Text = Mid(Me.txtLine.Text, 1, InStr(1, Me.txtLine.Text, " см") - 1)
'End Sub
'Private Sub txtline_LostFocus()
'Me.txtLine.Text = Me.txtLine.Text & " см"
'End Sub
'Private Sub txttop_GotFocus()
'Me.txtTop.Text = Mid(Me.txtTop.Text, 1, InStr(1, Me.txtTop.Text, " см") - 1)
'End Sub
'
'Private Sub txttop_LostFocus()
'Me.txtTop.Text = Me.txtTop.Text & " см"
'End Sub
'
'Private Sub txtBottom_GotFocus()
'Me.txtBottom.Text = Mid(Me.txtBottom.Text, 1, InStr(1, Me.txtBottom.Text, " см") - 1)
'End Sub
'
'Private Sub txtBottom_LostFocus()
'Me.txtBottom.Text = Me.txtBottom.Text & " см"
'End Sub
'
'Private Sub txtleft_GotFocus()
'Me.txtLeft.Text = Mid(Me.txtLeft.Text, 1, InStr(1, Me.txtLeft.Text, " см") - 1)
'End Sub
'
'Private Sub txtleft_LostFocus()
'Me.txtLeft.Text = Me.txtLeft.Text & " см"
'End Sub
'
'Private Sub txtRight_GotFocus()
'Me.txtRight.Text = Mid(Me.txtRight.Text, 1, InStr(1, Me.txtRight.Text, " см") - 1)
'End Sub
'
'Private Sub txtRight_LostFocus()
'Me.txtRight.Text = Me.txtRight.Text & " см"
'End Sub
'' пар
'Private Sub txtPtop_GotFocus()
'Me.txtPTop.Text = Mid(Me.txtPTop.Text, 1, InStr(1, Me.txtPTop.Text, " см") - 1)
'End Sub
'
'Private Sub txtPtop_LostFocus()
'Me.txtPTop.Text = Me.txtPTop.Text & " см"
'End Sub
'
'Private Sub txtPBottom_GotFocus()
'Me.txtPBottom.Text = Mid(Me.txtPBottom.Text, 1, InStr(1, Me.txtPBottom.Text, " см") - 1)
'End Sub
'
'Private Sub txtPBottom_LostFocus()
'Me.txtPBottom.Text = Me.txtPBottom.Text & " см"
'End Sub
'
'Private Sub txtPleft_GotFocus()
'Me.txtPLeft.Text = Mid(Me.txtPLeft.Text, 1, InStr(1, Me.txtPLeft.Text, " см") - 1)
'End Sub
'
'Private Sub txtPleft_LostFocus()
'Me.txtPLeft.Text = Me.txtPLeft.Text & " см"
'End Sub
'
'Private Sub txtPRight_GotFocus()
'Me.txtPRight.Text = Mid(Me.txtPRight.Text, 1, InStr(1, Me.txtPRight.Text, " см") - 1)
'End Sub
'
'Private Sub txtPRight_LostFocus()
'Me.txtPRight.Text = Me.txtPRight.Text & " см"
'End Sub
'' вода
'Private Sub txtVtop_GotFocus()
'Me.txtVTop.Text = Mid(Me.txtVTop.Text, 1, InStr(1, Me.txtVTop.Text, " см") - 1)
'End Sub
'
'Private Sub txtVtop_LostFocus()
'Me.txtVTop.Text = Me.txtVTop.Text & " см"
'End Sub
'
'Private Sub txtVBottom_GotFocus()
'Me.txtVBottom.Text = Mid(Me.txtVBottom.Text, 1, InStr(1, Me.txtVBottom.Text, " см") - 1)
'End Sub
'
'Private Sub txtVBottom_LostFocus()
'Me.txtVBottom.Text = Me.txtVBottom.Text & " см"
'End Sub
'
'Private Sub txtVleft_GotFocus()
'Me.txtVLeft.Text = Mid(Me.txtVLeft.Text, 1, InStr(1, Me.txtVLeft.Text, " см") - 1)
'End Sub
'
'Private Sub txtVleft_LostFocus()
'Me.txtVLeft.Text = Me.txtVLeft.Text & " см"
'End Sub
'
'Private Sub txtVRight_GotFocus()
'Me.txtVRight.Text = Mid(Me.txtVRight.Text, 1, InStr(1, Me.txtVRight.Text, " см") - 1)
'End Sub
'
'Private Sub txtVRight_LostFocus()
'Me.txtVRight.Text = Me.txtVRight.Text & " см"
'End Sub
'' Стоки
'Private Sub txtStop_GotFocus()
'Me.txtSTop.Text = Mid(Me.txtSTop.Text, 1, InStr(1, Me.txtSTop.Text, " см") - 1)
'End Sub
'
'Private Sub txtStop_LostFocus()
'Me.txtSTop.Text = Me.txtSTop.Text & " см"
'End Sub
'
'Private Sub txtSBottom_GotFocus()
'Me.txtSBottom.Text = Mid(Me.txtSBottom.Text, 1, InStr(1, Me.txtSBottom.Text, " см") - 1)
'End Sub
'
'Private Sub txtSBottom_LostFocus()
'Me.txtSBottom.Text = Me.txtSBottom.Text & " см"
'End Sub
'
'Private Sub txtSleft_GotFocus()
'Me.txtSLeft.Text = Mid(Me.txtSLeft.Text, 1, InStr(1, Me.txtSLeft.Text, " см") - 1)
'End Sub
'
'Private Sub txtSleft_LostFocus()
'Me.txtSLeft.Text = Me.txtSLeft.Text & " см"
'End Sub
'
'Private Sub txtSRight_GotFocus()
'Me.txtSRight.Text = Mid(Me.txtSRight.Text, 1, InStr(1, Me.txtSRight.Text, " см") - 1)
'End Sub
'
'Private Sub txtSRight_LostFocus()
'Me.txtSRight.Text = Me.txtSRight.Text & " см"
'End Sub
'
