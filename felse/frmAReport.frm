VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmAReport 
   Caption         =   "Справочник подразделений"
   ClientHeight    =   8235
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   7680
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   6720
      Width           =   7695
      Begin VB.CommandButton CancelButton 
         Caption         =   "Выход"
         Height          =   375
         Left            =   6015
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "Обновить"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "Применить"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "Удалить"
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox swtRed 
         Caption         =   "Редактировать"
         Height          =   315
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtKod 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtPrim 
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Наименование"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Примечание"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Код"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAReport.frx":0000
      Height          =   6675
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   11774
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   582
      ConnectMode     =   2
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
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
End
Attribute VB_Name = "frmAReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim adoX As ADODB.Connection
'
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_load()
On Error GoTo errLoad
Me.Adodc1.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                & "SERVER=192.168.100.23;" _
                & " DATABASE=askute;" _
                & "OPTION=0; PORT=3306"
Me.Adodc1.RecordSource = "sp_podraz"
Me.Adodc1.UserName = frmLogin.strUser
Me.Adodc1.Password = frmLogin.strPSW
Me.Adodc1.Refresh
Set adoX = Me.Adodc1.Recordset.ActiveConnection
Exit Sub
errLoad:
MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub Form_Resize()
Me.DataGrid1.Width = Me.Width - 120
Me.Frame2.Top = Me.Height - Me.Frame2.Height - 570
Me.DataGrid1.Height = Me.Height - Me.Frame2.Height - 585
End Sub

Private Sub swtRed_Click()
If swtRed.Value Then
    swtRed.Caption = "Добавить"
    Me.txtKod.Enabled = True
Else
    swtRed.Caption = "Редактировать"
    Me.txtKod.Enabled = False
End If
End Sub

Private Sub cmdDel_Click()
Dim lV As Long
On Error GoTo errcmdDel
lV = DataGrid1.Columns(0)
adoX.Execute "DELETE FROM " & Me.Adodc1.RecordSource & _
            " WHERE " & Me.Adodc1.Recordset.Fields(0).Name & "=" & lV & ";"
Me.Adodc1.Refresh
Exit Sub
errcmdDel:
MsgBox Err.Number & "-" & Err.Description
Resume Next
End Sub

Private Sub cmdRed_Click()
Dim lV As Long, lP As Long
On Error GoTo errcmdRed
With adoX
    lP = DataGrid1.Row ' сохраняем позицию
    If swtRed.Value = 0 Then
        lV = DataGrid1.Columns(0) ' сохраняем знач.поз.в таблице
        ' выполняем запрос на изменение данных
        If Len(Text1.Text) > 0 Then _
            .Execute "UPDATE " & Me.Adodc1.RecordSource & _
                    " SET " & Me.Adodc1.Recordset.Fields(1).Name & " = '" & Text1.Text & _
                    "' WHERE " & Me.Adodc1.Recordset.Fields(0).Name & "=" & lV & ";"
        ' выполняем запрос на изменение данных
        If Len(txtPrim.Text) > 0 Then _
            .Execute "UPDATE " & Me.Adodc1.RecordSource & _
                    " SET " & Me.Adodc1.Recordset.Fields(2).Name & " = '" & txtPrim.Text & _
                    "' WHERE " & Me.Adodc1.Recordset.Fields(0).Name & "=" & lV & ";"
    Else
        If Len(txtKod.Text) > 0 Then _
        .Execute "INSERT INTO " & Me.Adodc1.RecordSource & _
                 " (" & Me.Adodc1.Recordset.Fields(0).Name & "," & _
                 Me.Adodc1.Recordset.Fields(1).Name & "," & _
                 Me.Adodc1.Recordset.Fields(2).Name & ") VALUE (" & txtKod.Text & _
                 ",'" & Text1.Text & "','" & txtPrim.Text & "');"
    End If
End With
Me.Adodc1.Refresh
DataGrid1.Row = lP
Exit Sub
errcmdRed:
Resume Next
End Sub
'
Private Sub OKButton_Click()
Me.DataGrid1.Refresh
Me.Adodc1.Recordset.Requery
Me.Adodc1.Refresh
End Sub


