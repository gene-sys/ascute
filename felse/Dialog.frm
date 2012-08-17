VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Основные настройки"
   ClientHeight    =   4305
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   2220
      Left            =   45
      TabIndex        =   10
      Top             =   1290
      Width           =   5910
      Begin VB.TextBox txtDogov 
         Height          =   360
         Left            =   2370
         TabIndex        =   18
         Top             =   1770
         Width           =   3435
      End
      Begin VB.TextBox txtAdres 
         Height          =   360
         Left            =   2370
         TabIndex        =   17
         Top             =   1290
         Width           =   3435
      End
      Begin VB.TextBox txtPotreb 
         Height          =   375
         Left            =   2385
         TabIndex        =   16
         Top             =   780
         Width           =   3405
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   2385
         TabIndex        =   11
         Top             =   315
         Width           =   3405
      End
      Begin VB.Label Label7 
         Caption         =   "Договор:"
         Height          =   285
         Left            =   1515
         TabIndex        =   21
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Адрес:"
         Height          =   315
         Left            =   1710
         TabIndex        =   20
         Top             =   1350
         Width           =   600
      End
      Begin VB.Label Label5 
         Caption         =   "Потребитель:"
         Height          =   315
         Left            =   1185
         TabIndex        =   19
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Наименование соединения:"
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   345
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1395
      Left            =   60
      TabIndex        =   4
      Top             =   -90
      Width           =   4485
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1665
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Номер порта:"
         Height          =   225
         Left            =   585
         TabIndex        =   8
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Сервер соединения:"
         Height          =   225
         Left            =   60
         TabIndex        =   7
         Top             =   810
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Сохранить"
      Height          =   345
      Left            =   4665
      TabIndex        =   1
      Top             =   705
      Width           =   1230
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Выход"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   45
      TabIndex        =   2
      Top             =   3495
      Width           =   3645
      Begin VB.CheckBox Check1 
         Height          =   345
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Режим выполнения расчета:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   315
         Width           =   2205
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3660
      TabIndex        =   13
      Top             =   3495
      Width           =   2295
      Begin VB.CheckBox Check3 
         Caption         =   "Отчет ТСРВ"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   465
         Width           =   1545
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Обычный отчет"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private z As String

Private Sub Check1_Click()
On Error GoTo TC4_err
' смена режима получения архива
If Check1.Value Then
    Check1.Caption = "Режим 'Лето'": z = "True"
Else
    Check1.Caption = "Режим 'Зима'": z = "False"
End If
Exit Sub
TC4_err:
    MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub

Private Sub Check2_Click()
If Check2.Value Then Check3.Value = 0 Else _
    Check3.Value = 1
End Sub

Private Sub Check3_Click()
If Check3.Value Then Check2.Value = 0 Else _
    Check2.Value = 1
End Sub

Private Sub Command1_Click()
' сохранить новое значение порта
Call WriteParameters("Port", Text1.Text)
' сохранить новое значение сервера
If frmStart.mnu61.Checked Then
    Call WriteParameters("PathAdmin", Text2.Text)
Else
    Call WriteParameters("PathASKUTE", Text2.Text)
End If
' Записываем в раздел  переменную режима получения архива
Call WriteParameters("Mode", z)
' новое наименование соединения
Call WriteParameters("Connect", Text3.Text)
If Check3.Value Then
    Call WriteParameters("TipRep", "2")
Else
    Call WriteParameters("TipRep", "1")
End If
Call WriteParameters("Potreb", txtPotreb.Text)
Call WriteParameters("Adres", txtAdres.Text)
Call WriteParameters("Dogov", txtDogov.Text)
End Sub

Private Sub Form_Load()
With DataEnvironment1.rsCommand2 ' берем настройки
If .State <> adStateOpen Then .Open  ' переоткрыть запрос
.Requery
.MoveFirst
Do While Not .EOF
    Select Case .Fields("NameSet") ' взять название параметра настройки
    Case "PathAdmin"
        ' взять название сервера соединения
        If frmStart.mnu61.Checked Then Dialog.Text2.Text = .Fields("Set")
    Case "PathASKUTE"
        ' взять название сервера соединения
        If frmStart.mnu62.Checked Then Dialog.Text2.Text = .Fields("Set")
    Case "Port"
        ' взять номер порта
        Dialog.Text1.Text = .Fields("Set")
    Case "Mode" ' выбрать режим получения данных
        If .Fields("Set") = "True" Then
            Check1.Value = 1
            Check1.Caption = "Режим 'Лето'": z = "True"
        Else
            Check1.Value = 0
            Check1.Caption = "Режим 'Зима'": z = "False"
        End If
    Case "Connect" ' выбрать наименование соединения
        Dialog.Text3.Text = .Fields("Set")
    Case "Potreb" ' выбрать потребителя
        Dialog.txtPotreb.Text = "" & .Fields("Set")
    Case "Adres" ' выбрать адрес потреб.
        Dialog.txtAdres.Text = "" & .Fields("Set")
    Case "Dogov" ' выбрать договор
        Dialog.txtDogov.Text = "" & .Fields("Set")
    Case "TipRep" ' взять тип отчета
        If .Fields("Set") = 1 Then
            Check2.Value = 1:  Check3.Value = 0
        ElseIf .Fields("Set") = 2 Then
            Check3.Value = 1:  Check2.Value = 0
        End If
    End Select
    .MoveNext ' по всем параметрам настройки
Loop
End With
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

