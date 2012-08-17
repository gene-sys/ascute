VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmUU 
   Caption         =   "Ведение узлов учета"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form3"
   ScaleHeight     =   7485
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   9135
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   6960
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   6000
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox swchRed 
         Caption         =   "Редактировать"
         Height          =   315
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "Применить"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Выход"
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Удалить"
         Height          =   375
         Left            =   6000
         TabIndex        =   1
         Top             =   960
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmUU.frx":0000
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "naim_mest"
         BoundColumn     =   ""
         Text            =   "DataCombo1"
         Object.DataMember      =   "cmdMEST"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmUU.frx":0035
         Height          =   315
         Left            =   4080
         TabIndex        =   21
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "naim_naz"
         Text            =   "DataCombo2"
         Object.DataMember      =   "cmdNuz"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmUU.frx":0054
         Height          =   315
         Left            =   4560
         TabIndex        =   22
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "naim_ob"
         Text            =   "DataCombo3"
         Object.DataMember      =   "cmdObor"
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmUU.frx":0073
         Height          =   315
         Left            =   6360
         TabIndex        =   23
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "naim_pod"
         Text            =   "DataCombo4"
         Object.DataMember      =   "cmdPodraz"
      End
      Begin VB.Label Label7 
         Caption         =   "Код подраз."
         Height          =   255
         Left            =   6960
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Призн.УУ"
         Height          =   255
         Left            =   6000
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Код оборуд."
         Height          =   255
         Left            =   5040
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Код назнач."
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Код местоп."
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Наименование УУ"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Код УУ"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmUU.frx":0092
      Height          =   5895
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10398
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
   Begin MSAdodcLib.Adodc adoUU 
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
Attribute VB_Name = "frmUU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoX As ADODB.Connection

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim lV As Long
On Error GoTo errDel
lV = DataGrid1.Columns(0) ' сохраняем знач.поз.в таблице
' удаляем
adoX.Execute "DELETE FROM sp_uzl WHERE kod_uzl=" & lV & ";"
Me.adoUU.Refresh ' обновляем отображение
Exit Sub
errDel:
MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub DataCombo1_Change()
' выбираем код выбранного местоположения
DataEnvironment1.rscmdMEST.AbsolutePosition = Me.DataCombo1.SelectedItem
Me.Text3.Text = DataEnvironment1.rscmdMEST.Fields(0)
End Sub

Private Sub DataCombo2_Change()
' выбираем код выбранного названия узла
DataEnvironment1.rscmdNuz.AbsolutePosition = Me.DataCombo2.SelectedItem
Me.Text4.Text = DataEnvironment1.rscmdNuz.Fields(0)
End Sub


Private Sub DataCombo3_Change()
' выбираем код выбранного оборудования
DataEnvironment1.rscmdObor.AbsolutePosition = Me.DataCombo3.SelectedItem
Me.Text5.Text = DataEnvironment1.rscmdObor.Fields(0)
End Sub

Private Sub DataCombo4_Change()
' выбираем код выбранного подразделения
DataEnvironment1.rscmdPodraz.AbsolutePosition = Me.DataCombo4.SelectedItem
Me.Text7.Text = DataEnvironment1.rscmdPodraz.Fields(0)
End Sub

Private Sub Form_load()
On Error GoTo errLoad
' подключение к БД АСКУТЭ
Me.adoUU.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                & "SERVER=192.168.100.23;" _
                & " DATABASE=askute;" _
                & "OPTION=0; PORT=3306" 'OPTION=16906
Me.adoUU.Password = frmLogin.strPSW ' пароль
Me.adoUU.UserName = frmLogin.strUser ' имя доступа
' формирование запроса
Me.adoUU.RecordSource = "SELECT  sp_uzl.kod_uzl, sp_podraz.naim_pod,sp_uzl.naim_uzl," & _
                        "sp_mest.naim_mest, sp_mest.tel,sp_ob.naim_ob," & _
                        "sp_ob.nom_ob,sp_nuz.naim_naz FROM sp_uzl " & _
                        "LEFT JOIN sp_ob ON sp_uzl.kod_ob=sp_ob.kod_ob " & _
                        "LEFT JOIN sp_mest ON sp_uzl.kod_mest=sp_mest.kod_mest " & _
                        "LEFT JOIN sp_nuz ON sp_uzl.kod_naz=sp_nuz.kod_naz " & _
                        "LEFT JOIN sp_podraz ON sp_uzl.kod_pod=sp_podraz.kod_pod;"
Me.adoUU.Refresh ' обновить результат
Set adoX = Me.adoUU.Recordset.ActiveConnection
Exit Sub
errLoad:
MsgBox Err.Number & "-" & Err.Description
End Sub

Private Sub Form_Resize()
' изменение размеров элементов при изменении размеров формы
Me.DataGrid1.Width = Me.Width - 120
Me.Frame3.Top = Me.Height - Me.Frame3.Height - 570
Me.DataGrid1.Height = Me.Height - Me.Frame3.Height - 585
End Sub
Private Sub swchRed_Click()
' переключить м/д редактированием и добавлением
If swchRed.Value Then
    swchRed.Caption = "Добавить"
    Me.Text1.Enabled = True
Else
    swchRed.Caption = "Редактировать"
    Me.Text1.Enabled = False
End If
End Sub
Private Sub cmdRed_Click()
Dim lV As Long, lP As Long
On Error GoTo errcmdRed
With adoX
    lP = DataGrid1.Row ' сохраняем позицию
    If swchRed.Value = 0 Then
        lV = DataGrid1.Columns(0) ' сохраняем знач.поз.в таблице
        ' выполняем запрос на изменение данных
        If Len(Text2.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET naim_uzl = '" & Text2.Text & "' WHERE kod_uzl=" & lV & ";"
        If Len(Text3.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET kod_mest = '" & Text3.Text & "' WHERE kod_uzl=" & lV & ";"
        If Len(Text4.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET kod_naz = '" & Text4.Text & "' WHERE kod_uzl=" & lV & ";"
        If Len(Text5.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET kod_ob = '" & Text5.Text & "' WHERE kod_uzl=" & lV & ";"
        If Len(Text6.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET pr_uzl = '" & Text6.Text & "' WHERE kod_uzl=" & lV & ";"
        If Len(Text7.Text) > 0 Then .Execute "UPDATE sp_uzl " & _
                " SET kod_pod = '" & Text7.Text & "' WHERE kod_uzl=" & lV & ";"
    Else
        If Len(Text1.Text) > 0 Then .Execute "INSERT INTO sp_uzl " & _
                 " (kod_uzl,naim_uzl,kod_mest,kod_naz,kod_ob,pr_uzl,kod_pod) VALUE (" & _
                 Text1.Text & ",'" & _
                 Text2.Text & "'," & _
                 nz(Text3.Text) & "," & _
                 nz(Text4.Text) & "," & _
                 nz(Text5.Text) & "," & _
                 nz(Text6.Text) & "," & _
                 nz(Text7.Text) & ");" '
    End If
End With
Me.adoUU.Refresh
DataGrid1.Row = lP
Exit Sub
errcmdRed:
MsgBox Err.Number & "-" & Err.Description
Resume Next
End Sub
' если числовое поле NULL то выдавать 0
Function nz(f As String) As Variant
If Len(f) = 0 Then nz = 0 Else nz = f
End Function

