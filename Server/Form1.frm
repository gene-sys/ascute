VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{0853907C-ABF5-470D-A3CB-AB5C07EDC088}#1.0#0"; "TRAYCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "�������������� ������ �����-�� v1.33"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4905
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnRef 
      Height          =   450
      Left            =   4305
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "�������� ����� ������"
      Top             =   0
      Width           =   495
   End
   Begin trayctl.Tray Tray1 
      Left            =   15
      Top             =   4965
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���������� ������"
      Height          =   375
      Left            =   1725
      TabIndex        =   10
      Top             =   2640
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������ ��������"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Text            =   "c:\��� ���������\spdef.mdb"
      Top             =   3765
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "c:\��� ���������\spbase.mdb"
      Top             =   3390
      Width           =   4455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "��� ���������"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "Form1.frx":0614
         Left            =   120
         List            =   "Form1.frx":061B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "������� Escape ��� ���������� ���������� �������"
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����, ������������ ��������"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   765
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   3840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "���� � �����:"
      Height          =   1845
      Left            =   90
      TabIndex        =   8
      Top             =   3180
      Width           =   4695
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   3915
         TabIndex        =   16
         Top             =   1440
         Width           =   705
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Text            =   "0.7"
         Top             =   1005
         Width           =   615
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "����� �����:"
         Height          =   255
         Left            =   2850
         TabIndex        =   17
         Top             =   1470
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "����� �������� ����� �������� ������ (���):"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1005
         Width           =   3735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "����������"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4695
      Begin VB.CommandButton btnPerf 
         Caption         =   "���������"
         Height          =   375
         Left            =   3315
         TabIndex        =   19
         ToolTipText     =   "��������� ���������� ������"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "���������:"
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   4695
      Begin VB.Label stat 
         Height          =   750
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   4455
      End
   End
   Begin VB.Label clients 
      Caption         =   "0"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "���������� ����������� ��������:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu mnu 
      Caption         =   "TMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "����������"
      End
      Begin VB.Menu mnu2 
         Caption         =   "�����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��� ���. ����� ��������� ���-�� �������� ����������
Public indmax As Integer
'��� ���������� ����� �������� ��� ��� �� ������������ �����-�� ������, ������� ������ ��� ����������
Public again As Boolean
' ���������� ���������� ��������� ����������� ������� � ��
Public nameOfNode As String ' ��� ����
Public nameOfTable As String ' ��� ������� ����
Public firstP As String ' ������ �������
Public secondP As String ' ����� �������
Public KindOfArh As String  ' ��� ������
Public TipArh As String  ' ��� ������ (����� ����, ���, ���. ����)
Public KindOfMode As String ' ����� (����/����)
Public mstrTable As String ' �������������� ������ �������
'
Private TipNode As String ' ��� ���� (����., ���.�., ���)
Private CountHour As Long ' ������� ����� � ������������� �������
Private Const chunk = 8000
Private lSchet As Long
Private isp As Boolean ' ���������� ����� �� �������

'
'Suspends operation of a thread for the specified time.
'dwMilliseconds  Long�The time to suspend the thread in milliseconds. The constant
'INFINITE to put a thread permanently to sleep.
'bAlertable  Long�SleepEx only. Set to True if an asynchronous I/O transfer has
'been initiated with a ReadFileEx or WriteFileEx function call, and you want the function
'to return so that the I/O completion routine specified by those functions may execute.
'Return Value:
'Long�SleepEx only. Zero if the timeout elapses, WAIT_IO_COMPLETION if the function returned
'because of the completion of an asynchronous I/O operation.
'Private Declare Function SleepEx& Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal 'bAlertable As Long)

Private Sub btnPerf_Click()
Call Performance(Me.Text1.Text)
End Sub

Private Sub btnRef_Click()
Call RemakeThisLink(Me.Text1.Text)
End Sub

'
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    ws(Me.Combo1.ListIndex).Close
    Call ws_Close(Me.Combo1.ListIndex)
End If
End Sub

Private Sub Command1_Click()
' ������� ����� ����������� �������� (� ������������ �������� ������)
Form2.Show
End Sub

Private Sub Command2_Click()
' ������������� ������������� ������ �� ����������� �����
' �������� ����� � �������������� �������� ��������
Form3.Show
End Sub

Private Sub Form_Load()
Dim strZ As String
On Error GoTo errLoad
lSchet = 0: isp = True
' ����� ��������� �� ini-�����
strZ = String$(255, " ")
GetPrivateProfileString "Setup", "Database", "x", strZ, 255, App.Path & "/set.ini"
Text1.Text = Trim(strZ)
strZ = String$(255, " ")
GetPrivateProfileString "Setup", "Table", "x", strZ, 255, App.Path & "/set.ini"
Text3.Text = Trim(strZ)
strZ = String$(255, " ")
'GetPrivateProfileString "Setup", "Time", "0.5", strz, 255, App.Path & "/set.ini"
'Text4.Text = Trim(strz)
strZ = String$(255, " ")
GetPrivateProfileString "Setup", "Port", "x", strZ, 255, App.Path & "/set.ini"
Text5.Text = Trim(strZ)
'����� ������� ����.
ws(0).LocalPort = 1001 'Text5.Text
ws(0).Listen
Combo1.ListIndex = 0
Tray1.AddToTray Me.Icon, Me.Caption, True
errLoad:
'
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then _
    Tray1.AddToTray Me.Icon, Me.Caption, True
End Sub

Private Sub mnu1_Click()
Me.WindowState = vbNormal: Me.Show
End Sub

Private Sub mnu2_Click()
Unload Me
End Sub

'
Private Sub Tray1_MouseUp(Button As trayctl.TrayMouseConstants)
If Button = RightButton Then PopupMenu mnu, , , , mnu1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Tray1.ToReturn = True
Tray1.DeleteFromTray
End Sub

'
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim i2 As Integer
'���� ������ ������� Enter
If KeyAscii = 13 Then
    i2 = 1 '��� ����
    '''''if instr(1,combo1.List(combo1.ListIndex)," (��������)") then
    If Combo1.ListIndex <> 0 And InStr(1, Combo1.List(Combo1.ListIndex), _
        " (��������)") = 0 Then ws(Combo1.ListIndex).SendData "serv������: " _
                                                        & Text2.Text: GoTo Label2
    '���������� � ����� ��� ����������� ������� (�������� �������)
    For i2 = ws.LBound + 1 To ws.ubound
        '���� ������ ������ (�.�. ������ ����������),
        '�� ���������� ��� � ���� ������
        If ws(i2).State = sckClosed Then GoTo NextFor
        '��������� ������ �������� �������
        ws(i2).SendData "serv������: " & Me.Text2
        '��� �����, ����� �������� �������� ������
        '������������ �������
NextFor:
    Next
Label2:
    '�������� � �������, ��� ���������
    Me.stat.Caption = "������ ����������"
    '�������� ���� �����
    Text2.Text = ""
    '�� ������ ������... (����� �� �������� ���� Enter ������ �� ���������,
    '���� ������� ��������, ��� ����� ����� (���� keypreview=true)
    KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()
    On Error GoTo TC4_err
    ' ���������� � ������  ����������
    WritePrivateProfileString "Setup", "Time", _
                    Text4.Text, App.Path & "/set.ini"
    'Call WriteParameters("PathDatabase", Text2.Text)
    Exit Sub
TC4_err:
    MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub

Private Sub Text5_Change()
On Error GoTo TC5_err
' ���������� � ������  ����������
WritePrivateProfileString "Setup", "Port", _
            Text5.Text, App.Path & "/set.ini"
'Call WriteParameters("PathDatabase", Text2.Text)
Exit Sub
TC5_err:
    MsgBox Err.Number & "->" & Err.Description
    Resume Next
End Sub

Private Sub ws_Close(Index As Integer)
On Error Resume Next
'������� ������, �� �������� ���������� ������
ws(Index).Close
'Combo1.List(Index) = Combo1.List(Index) & " (��������)"
'�������� � �������, ��� ���������
Me.stat.Caption = "���� �� �������� ���������� (" & Index & ")"
Combo1.RemoveItem Index
Combo1.ListIndex = 0
'���������������� ���-�� ������ ����������� � ������������ ���������.
'��� �� �����������, �� �� ������ ������, ����� ����������� ���
'����� �������, ����� �� �������� �����
If Me.clients.Caption = "0" Then indmax = 0
'��������, ��� �������� ����� ������ �� ������
Me.clients.Caption = CInt(Me.clients.Caption) - 1
Beep
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim I As Long
On Error GoTo errCR
If Index = 0 Then
    '��������������� ������� �����������
    indmax = indmax + 1
    '������������ ������������ ��� �������
    For I = ws.LBound To ws.ubound
        '���� ������ ��� ��������, � ����� ����� ������������, ��
        '��� ���� ��� ��������������� �����, � ������ � �� ���� ���������
        '����� ��-� ������� (�.�. �� ��� ����)
        If ws(I).State = sckClosed Then
            indmax = indmax - 1
            ws(I).LocalPort = 0 'Text5.Text '1001
            ws(I).Accept requestID
            Me.clients.Caption = CInt(Me.clients.Caption) + 1
            ws(I).SendData "/reg"
            Me.stat.Caption = "����������� ������ ����� �����������!"
            'Beep
            again = True
            Exit Sub
        End If
    Next
    '�� � ���� ��� ����� ������, ��
    '... �������� ����� ������� �������...
    Load ws(indmax)
    '�������������� ��� �� ���� ����������...
    ws(indmax).LocalPort = 0 'Text5.Text '1001
    ws(indmax).Accept requestID
    '� ������� � ����� �������
    Me.clients.Caption = CInt(Me.clients.Caption) + 1
    ws(indmax).SendData "/reg"
    Me.stat.Caption = "��������� ����� ������!"
    again = False
    'Beep
End If
errCR:
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim I As Integer, tmp, handlef
Dim data As String, lenFile As Long
Dim bname As Boolean
On Error GoTo errDA
'
For I = ws.LBound + 1 To ws.ubound
    ' ���� ������ ������, �� ��� ��������� � ���� �������� ������
    If ws(I).State = sckClosed Then GoTo NextFor
    '�������� ���������� ������ � ������� �� ����� (��� � ����� �-���)
    '����� �� ������� �����, ����� ������������ peekdata
    ws(I).GetData data, vbString
    If InStr(1, data, "NICK ") Then
        data = Replace(data, "NICK ", "")
'        For tmp = 0 To Me.Combo1.ListCount
'            If Replace(Me.Combo1.List(tmp), " (��������)", "") = data Then _
'                Me.Combo1.List(tmp) = Replace(Me.Combo1.List(tmp), " (��������)", ""): GoTo exitme
'        Next
        Combo1.AddItem data, I
exitme:
        Exit Sub
    End If
'
    bname = False: lenFile = 0
    Select Case Left(data, 4)
        Case "/get"
            writeLog (data) ' �������� � log-���� ��� �������� � ����� ������ ������
            bname = strAnalize(data) ' ���������������� �������� ����������� �������
            If bname Then
                ' �������� �������
                If Getnameoftable() Then
                    'lenFile = getRqst(Combo1.List(I) & "base.tmp")       '��������� ������ � ����
                    lenFile = DataPass(Combo1.List(I) & "base.tmp", 1, Empty)     '��������� ������ � ����
                    handlef = "rqst" & Trim(str(lenFile)) ' ������������ ��������� � ������������ �����
                    ws(I).SendData handlef '  �������� ���������
                Else
                    ws(I).SendData "/bad" ' ����� ������ ������
                End If
            Else
                ws(I).SendData "/bad" ' ����� ������ ������
            End If
        Case "okay"
            send Combo1.List(I) & "base.tmp", I ' �������� ���� �� ����������� ���������� ������� �� ������� /get
            'If IsNumeric(Val(Text4.Text)) Then SleepEx 1000 * Val(Text4.Text), 0
            Wait Val(Text4.Text)
            'ws(i).SendData "EnDf"
        Case "node"
            writeLog (data) ' �������� � log-���� ��� �������� � ����� ������ ������
            getNodeList I
        Case "anlz"
            ' ��������� ������ ��������� �����
            writeLog (data) ' �������� � log-���� ��� �������� � ����� ������ ������
            bname = strAnalize(data) ' ���������������� �������� ����������� �������
            lenFile = NodesStatus(Combo1.List(I) & "base.tmp") ' ��������� �����
            handlef = "rqst" & Trim(str(lenFile)) ' ������������ ��������� � ������������ �����
            ws(I).SendData handlef ' ��������� ������� ������������.����
    End Select
NextFor:
Next
'�������� � �������, ��� ���������
Me.stat.Caption = "�������� ������"
Exit Sub
errDA:
Resume Next
End Sub
' ������� ������������ ������ � ��������� �����
Function NodesStatus(UsrName As String) As Long
Dim filenum ' ��� ����� ��� ���������� ������ �� ���������� �������
Dim lenFile As Long
On Error Resume Next
' �������� ini-����� ����� ������������ � ������ �����
' ���������� ������� �� ������� ����
Data2.DatabaseName = Text3.Text
' ����� ������� ������ �����
Data2.RecordSource = "����"
Data2.Refresh
lenFile = 0
' ����������� ���� ��� ������
filenum = FreeFile
Open UsrName For Output As #filenum
'Print #filenum, "" ' �������� ����
Print #filenum, "����;���������;��������" '& vbCrLf
Close #filenum
With Data2.Recordset
    .MoveFirst
    Do While Not .EOF
        nameOfNode = .Fields("�������")
        If InStr(1, nameOfNode, "���", 1) > 0 Then
            TipNode = "��" ' ���
        ElseIf InStr(1, nameOfNode, "��� ����", 1) > 0 Then
            TipNode = "��" ' ��� ����
        Else
            TipNode = "" ' ���������
        End If
        If Getnameoftable() Then
            ' ����� ��������� �������� ���� � �������
            CountHour = CLng(Mid(firstP, InStr(1, firstP, ":", 1) - 2, 2))
            If CLng(Mid(firstP, InStr(1, firstP, ":", 1) + 1, 2)) > 0 Then _
                                                    CountHour = CountHour + 1
            ' ��������� ������ � ����
            lenFile = lenFile + getAnlz(UsrName) '=?
        End If
        .MoveNext
    Loop
End With
Data2.DatabaseName = "":  Data2.RecordSource = "":  Data2.Refresh
' ������������ ��������� �����
NodesStatus = lenFile
End Function
'
'
Private Function getAnlz(FN As String) As Long
Dim filenum ' ��� ����� ��� ���������� ������ �� ���������� �������
Dim I As Long ' �������� ���������� ������� �� ���������� �������
Dim J As Long ' ��������� ���������� ������ � �������
Dim lngRule As Long ' ���������� ������ � �������
Dim arrNameFields() As String ' ������ ���� ����� ��� ���. ����������� ��������  �������� ������
Dim arrFields() As Double ' ������ �������� ����� ��� ���������� ������ ��� ����
Dim arrOper() As String ' ������ �������� ��� ������ �������� ������
Dim strZ As String ' ����� �������� ��� �������
Dim strRule As String ' ����� �������� ������� ��� �������
Dim strFields As String ' ����� ���������� ����� ��������� �����
Dim strNameFields As String ' ����� �������� ����� ��������� �����
Dim pos As Long ' ������� ; ��� ��������� ���� �����
Dim pos1 As Long ' ������� [��������] ��� ��������� ���� ����� � �������
Dim strTemp As String ' ������ ���������� ��� �������� �������������� ����
Dim strForma As String ' ������ ��� ������ �� ������ ������� ������
Dim strFunc As String ' �������� ������ �������
Dim strFunc1 As String ' �������� ������������� ������ �������
Dim dblFormula1 As Double '��� ���������� ���������� �������
Dim dblFormula2 As Double '��� ���������� ���������� �������
'
    On Error GoTo exit_geta
    ' ������� ������
    ' ������� pattern.ini
    ' ����� ��� �������
    ' ������� �� pattern.ini �� ����� ������� � ���� ������� (����/����) ��������������� ������
    strZ = String$(255, " ")
    If KindOfMode = "1" Then ' ���� ����
        GetPrivateProfileString nameOfTable, "����������", "", strZ, 255, App.Path & "/pattern.ini"
    ElseIf KindOfMode = "0" Then ' ���� ����
        GetPrivateProfileString nameOfTable, "����������", "", strZ, 255, App.Path & "/pattern.ini"
    End If
    strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
    strZ = Trim(strZ) ' ������ ��� �������
    ' ����� �������� ������ �������� �������
    mstrTable = "SELECT * FROM " & nameOfTable
    '
    ' �������������� ������� �������
   mstrTable = mstrTable & " WHERE [���������] BETWEEN #" & _
                            SQLData(CDate(firstP), True) & "# AND #" & _
                        SQLData(CDate(secondP), True) & "# ORDER BY [���������];"
    ' ���������� � ����� ������
    Data1.DatabaseName = Text1.Text
    ' ������ ��������� ������� ��� ������ �������������
    Data1.RecordSource = mstrTable
    Data1.Refresh
    I = 0
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        Data1.Recordset.Requery
        ' ���������������� ����������
        Data1.Recordset.MoveFirst
        ' ���������� ���������� ������ � ����� ������� � ������� ��������������� ������
        lngRule = CountRule(strZ)
        If lngRule > 0 Then
            ReDim arrFields(lngRule): ReDim arrOper(lngRule): ReDim arrNameFields(lngRule)
        End If
       ' ������� ��������� �������
        strForma = "": filenum = FreeFile
        Open FN For Append As #filenum
        'Print #filenum, TipArh & KindOfArh
        ' ������ ����� ���������� � 0 �������������� ��������
        With Data1.Recordset
            ' ��������� �������� ����� ��������� � ��������
            If lngRule > 0 Then
                For J = 1 To lngRule
                    strRule = RulesAnalize(J, strZ)
                    arrOper(J) = Left$(strRule, 1) & "," & Rules2Make(J, strZ)
                    arrNameFields(J) = Trim(Mid(strRule, 2))
                    arrFields(J) = Format(NZ(.Fields(arrNameFields(J))), "#0.000")
                Next J
            End If
            For I = 0 To .RecordCount - 1
                ' �������� � ���� ����������� ��������, ��������� 1-�� ������
                ' � ����� ���������� ������������ �������, ����� � ������� �������
                If .AbsolutePosition = 0 Then
                    '����� ����� �� * .ptn
                    strNameFields = NameOfFields(strZ)
                '    strForma = strNameFields
                Else
                    ' ����� ����� ����������� ������� � ���������� ����� �����
                    strFields = strNameFields
                    pos = InStr(1, strFields, ";", 1)
                    Do While pos
                        ' ����� ��� ����
                        strTemp = Mid(strFields, 1, pos - 1)
                        strFields = Mid(strFields, pos + 1)
                        pos = InStr(1, strFields, ";", 1)
                        ' �������� ��� � ������� ����������� ������
                        strFunc = "": strFunc1 = ""
                        ' ����� �������� ���������� �������� �������
                        If InStr(1, strTemp, "�������", 1) > 0 Then
                            ' ������ ������ �������
                            strFunc = String$(255, " ")
                            GetPrivateProfileString strTemp, "fx", "", strFunc, 255, App.Path & "/" & strZ
                            strFunc = Left$(strFunc, InStr(1, strFunc, Chr(0), 1) - 1)
                            strFunc = Trim(strFunc)
                            ' ������ �������
                            If Left$(strFunc, 1) = "(" Then
                                pos1 = InStr(1, strFunc, ")", 1)
                                ' �������� � strFunc1 �� ��� �� ��������
                                strFunc1 = Mid(strFunc, pos1 + 1)
                                ' �������� � strFunc �� ��� ������ ������
                                strFunc = Mid(strFunc, 2, pos1 - 2)
                            End If
                            ' ���������� ����� �������� ���������
                            pos1 = InStr(1, strFunc, "*", 1)
                            If pos1 = 0 Then
                                pos1 = InStr(1, strFunc, "/", 1)
                                If pos1 = 0 Then
                                    pos1 = InStr(1, strFunc, "+", 1)
                                    If pos1 = 0 Then
                                        pos1 = InStr(1, strFunc, "-", 1)
                                    End If
                                End If
                            End If
                            ' ���������� ������ (����������� �������� ��� ��������: +, -, *, /) ���
                            ' ����� ����������� � �������
                            If IsNumeric(Trim(Left$(strFunc, pos1 - 1))) Then _
                                dblFormula1 = CDbl(Trim(Left$(strFunc, pos1 - 1))) _
                            Else _
                                dblFormula1 = GetNumber(Trim(Left$(strFunc, pos1 - 1)) _
                                                    , NZ(.Fields(Trim(Left$(strFunc, pos1 - 1)))) _
                                                                    , lngRule, arrOper, arrNameFields, arrFields)
                            If IsNumeric(Trim(Mid(strFunc, pos1 + 1))) Then _
                                                dblFormula2 = CDbl(Trim(Mid(strFunc, pos1 + 1))) _
                            Else dblFormula2 = GetNumber(Trim(Mid(strFunc, pos1 + 1)), _
                                                                    NZ(.Fields(Trim(Mid(strFunc, _
                                                        pos1 + 1)))), lngRule, arrOper, arrNameFields, arrFields)
                            ' ��������������� ���������� �������
                            dblFormula1 = doOperate(Mid(strFunc, pos1, 1), dblFormula1, dblFormula2)
                            If Len(strFunc1) > 0 Then dblFormula1 = doOperate(Left$(strFunc1, 1), _
                                                                dblFormula1, CDbl(Trim(Mid(strFunc1, 2))))
                            ' ��������������� ��� ������
                            strForma = strForma & Format(dblFormula1, "#0.000") & ";"
                        Else
                            ' ��������� ����� ��� ������: ����, �����, ����� � �.�.
                            If TypeName(.Fields(strTemp)) = "String" Then _
                                                    strForma = strForma & .Fields(strTemp) & ";"
                            If IsNull(.Fields(strTemp)) Then strForma = strForma & "0" & ";" '& "NULL" & ";"
                            If IsNumeric(.Fields(strTemp)) Then
                                ' ���������� ������ (����������� �������� ��� ��������: +, -, *, /)
                                dblFormula1 = GetNumber(strTemp, NZ(.Fields(strTemp)), _
                                                        lngRule, arrOper, arrNameFields, arrFields)
                                ' ������������ ������ ���������� �������
                                J = IndxOper(strTemp, lngRule, arrNameFields)
                                If J > 0 Then
                                    pos1 = InStr(1, arrOper(J), ",", 1) + 1
                                    If Len(Mid(arrOper(J), pos1, 1)) > 0 Then dblFormula1 = doOperate _
                                    (Mid(arrOper(J), pos1, 1), dblFormula1, CDbl(Trim(Mid(arrOper(J), pos1 + 1))))
                                End If
                                strForma = strForma & Format(dblFormula1, "#0.000") & ";"
                            End If
                            If IsDate(.Fields(strTemp)) Then _
                                strForma = strForma & Format(.Fields(strTemp), "dd.mm.yy hh:mm") & ";"
                        End If
                    Loop
                     ' ��������� �������� ����� ��������� � ��������
                    If lngRule > 0 Then
                        For J = 1 To lngRule
                            strRule = RulesAnalize(J, strZ)
                            arrOper(J) = Left$(strRule, 1) & "," & Rules2Make(J, strZ)
                            arrNameFields(J) = Trim(Mid(strRule, 2))
                            arrFields(J) = Format(NZ(.Fields(arrNameFields(J))), "#0.000")
                        Next J
                    End If
                    ' ��������� �������� � ���� ���������� ������ ������������� ��������
                    ' ���������� �� ���� ��� - ����������
                    If CountHour = 23 Then CountHour = 0 Else CountHour = CountHour + 1
                    strForma = Strokaanal(strForma)
                    If Len(strForma) > 0 Then
                        Print #filenum, strForma
                        Exit For
                    End If
                    strForma = ""
                End If
                .MoveNext
            Next I
        End With
        I = Seek(filenum)
        Close #filenum    ' ��������� ����.
    End If
    If I = 0 Then
        filenum = FreeFile
        Open FN For Append As #filenum
        Print #filenum, Mid(nameOfTable, 9) & ";;" & "��� ������"
        I = Seek(filenum)
        Close #filenum    ' ��������� ����.
    End If
    Data1.DatabaseName = "":  Data1.RecordSource = "":  Data1.Refresh
    getAnlz = I
    Exit Function
exit_geta:
    If Len(nameOfTable) > 0 Then writeLog (nameOfTable & ":" & protocol())
End Function
'
' ������ ������ ��� ��������
Function Strokaanal(strF As String) As String
Dim pos As Long, str As String, str1 As String
Dim I As Integer, strE As String
Dim priznak As Boolean
str = strF
I = 0
priznak = False
'str = Mid(str, 1, Len(str) - 1) ' ���������� ��������� ;
pos = InStr(1, str, ";", 1) ' ���� ������ ���������
Do While pos
    I = I + 1 ' ������� �������
    str1 = Mid(str, 1, pos - 1)  ' ���������
    str = Mid(str, pos + 1)
    Select Case I
    Case 1
        ' ��������� �� ������������������ �����
        ' ����� ������� �������� ���� � �������
        If CLng(Mid(str1, InStr(1, str1, ":", 1) - 2, 2)) <> CountHour Then
            strE = str1 & ";" & "��� ������"
            priznak = True
            Exit Do
        End If
        strE = str1 & ";"
    Case 2
        If TipNode = "��" Then
            If CDbl(str1) = 0 Then
                strE = strE & "V=0"
                priznak = True
                Exit Do
            End If
        Else
            If str1 = "0" Then
                strE = strE & "��� ��������"
                priznak = True
                Exit Do
            End If
            If KindOfMode = "0" Then ' ���� ����
                If CDbl(str1) <= 0 Then
                    strE = strE & "W1=" & str1
                    priznak = True
                    Exit Do
                End If
            End If
        End If
    Case 3
        If str1 = "0" Then
            strE = strE & "��� ��������"
            priznak = True
            Exit Do
        End If
        If KindOfMode = "0" Then ' ���� ����
            If CDbl(str1) <= 0 Then
                strE = strE & "W2=" & str1
                priznak = True
                Exit Do
            End If
        End If
    Case 4
        If TipNode = "��" Then
            If CDbl(str1) <> 0 Then
                strE = strE & "T�������=" & str1
                priznak = True
                Exit Do
            End If
        End If
        If TipNode = "��" Then
            If CDbl(str1) = 0 Then
                strE = strE & "�1=0"
                priznak = True
                Exit Do
            End If
        End If
        If str1 = "0" Then
            strE = strE & "��� ��������"
            priznak = True
            Exit Do
        End If
        If KindOfMode = "0" Then ' ���� ����
            If CDbl(str1) <= 0 Then
                strE = strE & "m1=" & str1
                priznak = True
                Exit Do
            End If
        End If
    Case 5
        If str1 = "0" Then
            strE = strE & "��� ��������"
            priznak = True
            Exit Do
        End If
        If KindOfMode = "0" Then ' ���� ����
            If CDbl(str1) <= 0 Then
                strE = strE & "m2=" & str1
                priznak = True
                Exit Do
            End If
        End If
    Case 7
        ' ���� ������������ �������
        If CDbl(str1) <= 0 Then
            strE = strE & "W����=" & str1
            priznak = True
            Exit Do
        End If
    Case 8 ' ���� ������������ �����
        If TipNode = "��" Then
            If CDbl(str1) = 0 Then
                strE = strE & "�2=0"
                priznak = True
                Exit Do
            End If
        End If
        If CDbl(str1) <= 0 Then
            strE = strE & "M����=" & str1
            priznak = True
            Exit Do
        End If
    Case 9 ' ���� ����� ���������<>60
         If CDbl(str1) < 60 Then
            strE = strE & "T���������=" & str1
            priznak = True
            Exit Do
        End If
    Case 10 ' ���� ����-�� �� 1�� ��-��=0
        If CDbl(str1) = 0 Then
            strE = strE & "t1=0"
            priznak = True
            Exit Do
        End If
    Case 11 ' ���� ����-�� �� 2�� ��-��=0
        If CDbl(str1) = 0 Then
            strE = strE & "t2=0"
            priznak = True
            Exit Do
        End If
    Case 13 ' ���� �������� �� 1�� ��-��<=0
        If CDbl(str1) = 0 Then
            strE = strE & "P1=0"
            priznak = True
            Exit Do
        End If
    Case 14 ' ���� �������� �� 2�� ��-��<=0
        If CDbl(str1) = 0 Then
            strE = strE & "P2=0"
            priznak = True
            Exit Do
        End If
    End Select
    pos = InStr(1, str, ";", 1) ' ������ ���������
Loop
If priznak Then
    Strokaanal = Mid(nameOfTable, 9) & ";" & strE
Else
    Strokaanal = ""
End If
End Function
'
'
Function Getnameoftable() As Boolean
Dim strZ As String
strZ = String$(255, " ")
GetPrivateProfileString nameOfNode, "��� ������", "", strZ, 255, App.Path & "/node.ini"
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
strZ = Trim(strZ) ' ������ ��� �������
TipArh = strZ
If Len(TipArh) = 0 Then TipArh = "Journal"
strZ = String$(255, " ")
GetPrivateProfileString nameOfNode, KindOfArh & " �����", "", strZ, 255, App.Path & "/node.ini"
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
strZ = Trim(strZ) ' ������ ��� �������
nameOfTable = strZ
If Len(nameOfTable) > 0 Then Getnameoftable = True Else Getnameoftable = False
End Function
'
' �������� ������ �����
Private Sub getNodeList(I As Integer)
' ��� ������� ���������� � ����� ������
Dim strZ As String, MyString As String, namestr As String
Dim pos As Long, ppos As Long, data As String
' ���������� � ����� ������
On Error GoTo erComC
data = ""
MyString = NodeList
' ��� ������������
' ������������� ���� ������ ����� � ������������� �� �������������
pos = InStr(1, MyString, ",", vbTextCompare): ppos = 1
Data2.DatabaseName = Text3.Text
' ����� ������� ������ �����
Data2.RecordSource = "����"
Data2.Refresh
Do
    ' ����� ������������ �� ����� ���� � ini-�����
    strZ = String$(255, " ")
    namestr = Mid(MyString, ppos, pos - ppos)
    GetPrivateProfileString namestr, "������������", "", strZ, 255, App.Path & "/node.ini"
    Data2.Recordset.MoveFirst
    Data2.Recordset.FindFirst "�������='" & namestr & "'"
    If Not Data2.Recordset.NoMatch Then _
        namestr = namestr & ";" & Data2.Recordset.Fields("���������������") _
    Else namestr = namestr & ";-"
    ' �������� ��������� ������
    If Len(Trim(strZ)) > 1 Then
        If InStr(1, strZ, Combo1.List(I), vbTextCompare) > 0 Then
'            ws(I).SendData "node" & namestr & "@"
'            Wait Val(Text4.Text)
            data = data & namestr & "@"
        End If
    Else
'        ws(I).SendData "node" & namestr & "@"
'        Wait Val(Text4.Text)
        data = data & namestr & "@"
    End If
    ppos = pos
    pos = InStr(ppos + 1, MyString, ",", vbTextCompare): ppos = ppos + 1
Loop While pos > 0
Data2.DatabaseName = "":  Data2.RecordSource = "":  Data2.Refresh
ws(I).SendData "node" & data
Wait Val(Text4.Text)
Exit Sub
erComC:
writeLog ("����:" & protocol())
End Sub
' ������ � ��������� ������ ������ �����
Function NodeList() As String
Dim strZ As String, MyString As String
Dim filenum As Long
On Error GoTo errNodeList
' �������� ������ ����� �� node.ini
filenum = FreeFile
Open "node.ini" For Input As #filenum
Do While Not EOF(filenum)   ' Loop until end of file.
    Input #filenum, MyString
    If InStr(1, MyString, "[", vbTextCompare) > 0 And InStr(1, MyString, "]", vbTextCompare) > 0 Then
        MyString = Mid(MyString, 2, Len(MyString) - 2)
        strZ = strZ & MyString & ","
    End If
Loop
Close #filenum
'strz = Left(strz, Len(strz) - 1) ' ������ ��������� ','
NodeList = Trim(strZ)
Exit Function
errNodeList:
writeLog ("������ �����:" & protocol())
End Function
'
Private Sub send(fname As String, I As Integer)
Dim data As String
Dim a As Long
Dim Data1 As String
Dim Data2 As String
Dim FileNumber
On Error Resume Next
Reset ' �������� ��� ����� �������� �����
FileNumber = FreeFile ' �������� ���������� �����
Open fname For Binary As #FileNumber ' ������� ���� �������� �����������

Do While Not EOF(FileNumber) ' ���� �� ����� �����
    data = Input(chunk, #FileNumber) ' �������� ���������� ������������������ ������� ������ �����
    ws(I).SendData data ' ���������
'    Wait Val(Text4.Text)
'    DoEvents
Loop

Close #FileNumber ' ������� ����
End Sub
'
' ������������ ��� �����
Function strAnalize(strData As String) As Boolean
Dim n As Long, I As Long
Dim arrOIndx(5) As Long, arrZIndx(5) As Long
On Error GoTo errSA
' ���� ��� 6 ����������� ������ �������
n = InStr(1, strData, "["): I = 0
Do While n > 0
    arrOIndx(I) = n ' ��������� ������� ��������� ���� ����������� ������
    I = I + 1 ' ������� ���������� ������
    n = InStr(n + 1, strData, "[") ' ���� ��������� ������
Loop
' ���� ��� ����������� ������� ������� ���� �����������
If I = 6 Then
    n = InStr(1, strData, "]"): I = 0
    Do While n > 0
        arrZIndx(I) = n ' ��������� ������� ��������� ���� ����������� ������
        I = I + 1 ' ������� ���������� ������
        n = InStr(n + 1, strData, "]") ' ���� ��������� ������
    Loop
    ' ���� ��� ����������� ������ ������� ��������� ���������� �������
    If I = 6 Then
        nameOfNode = Mid(strData, arrOIndx(1) + 1, arrZIndx(1) - arrOIndx(1) - 1)
        firstP = Mid(strData, arrOIndx(2) + 1, arrZIndx(2) - arrOIndx(2) - 1)
        secondP = Mid(strData, arrOIndx(3) + 1, arrZIndx(3) - arrOIndx(3) - 1)
        KindOfArh = Mid(strData, arrOIndx(4) + 1, arrZIndx(4) - arrOIndx(4) - 1)
        KindOfMode = Mid(strData, arrOIndx(5) + 1, arrZIndx(5) - arrOIndx(5) - 1)
    Else
        strAnalize = False ' ����� ������� � �������
        Exit Function
    End If
Else
    strAnalize = False ' ����� ������� � �������
    Exit Function
End If
strAnalize = True ' ����� ������� � �������
Exit Function
errSA:
strAnalize = False
End Function
'
' ������� ������������ ����� �������� ������
Function GetNumber(strTemp As String, dblField As Double, lngRule As Long, aOper() As String, _
                                                aNameFields() As String, aFields() As Double) As Double
Dim dblForma As Double
Dim J As Long
    '
    ' ���� ��� � ������ ������
    J = IndxOper(strTemp, lngRule, aNameFields)
    If J > 0 Then
        ' ������ ����������-���������� ��������
        If dblField > 0 Then
            If aFields(J) > 1000000 Then
                If dblField / aFields(J) < 0.001 Then dblField = dblField + 1000000000
            ElseIf aFields(J) > 0 Then
                If dblField / aFields(J) < 0.001 Then dblField = dblField + 1000000
            ElseIf aFields(J) = 0 Then dblField = 0
            End If
        End If
    End If
    ' ��������� ��������������� ��������
    If J > 0 Then dblForma = doOperate(Left$(aOper(J), 1), dblField, aFields(J)) Else dblForma = dblField
    GetNumber = dblForma
End Function
'
' ������� ��������� ������ � ������� �������� �� ����� ����
Function IndxOper(strTemp As String, lngRule As Long, aNameFields() As String) As Long
Dim L As Long
Dim Ok As Boolean
    Ok = False
    ' ���� ��� � ������ ������
    For L = 1 To lngRule
        If UCase(strTemp) = UCase(aNameFields(L)) Then
            Ok = True:  Exit For
        End If
    Next L
    If Ok Then IndxOper = L Else IndxOper = 0
End Function
'
' ������� ���������� ��������
Function doOperate(typOper As String, dbl1 As Double, dbl2 As Double) As Double
    Select Case typOper
    Case "*"
        doOperate = dbl1 * dbl2
    Case "/"
        doOperate = dbl1 / dbl2
    Case "+"
        doOperate = dbl1 + dbl2
    Case "-"
        doOperate = dbl1 - dbl2
    End Select
End Function
'
' �������� ����� �����
Function NameOfFields(strC As String) As String
Dim MyString As String
Dim nameField As String
Dim posX As Long, posY As Long, filenum As Long
    '
    On Error GoTo errNOF
    nameField = ""
    filenum = FreeFile
    Open strC For Input As #filenum
    Do While Not EOF(filenum)   '
        Input #filenum, MyString
        ' �������� ������-��������� (���������� [])
        posX = InStr(1, MyString, "[", 1):  posY = InStr(1, MyString, "]", 1)
        If posX > 0 And posY > 0 Then
            If InStr(1, MyString, "Rule", 1) = 0 And InStr(1, MyString, "�������", 1) = 0 And _
                                                        InStr(1, MyString, "�������", 1) = 0 Then _
                                nameField = nameField & Mid(MyString, posX + 1, posY - posX - 1) & ";"
        End If
    Loop
    Close #filenum
NameOfFields = nameField
Exit Function
errNOF:
NameOfFields = Null
End Function
'
Function SQLData(dData As Date, Optional time As Boolean) As String
    'SQLData = "#" & Year(dData) & "/" & Month(dData) & "/" & Day(dData) & "#"
' ������������ ������� ��� �������
If IsNull(time) Then time = False
If time Then
    SQLData = Year(dData) & "/" & Month(dData) & "/" & Day(dData) & _
        " " & Hour(dData) & ":" & Minute(dData)
Else
    SQLData = Year(dData) & "/" & Month(dData) & "/" & Day(dData)
End If
End Function
'
' ������� ���������� ������ � ����� �������
Function CountRule(strX As String) As Long
Dim MyString As String
Dim posX As Long, intY As Long, filenum As Long
    '
    intY = 0
    filenum = FreeFile
    Open strX For Input As #filenum
    Do While Not EOF(filenum)   '
        Input #filenum, MyString
        posX = InStr(1, MyString, "�������", 1)
        If posX > 0 Then intY = intY + 1
    Loop
    Close #filenum    ' ��������� ����.
    'If intY = 0 Then intY = 1
    CountRule = intY
End Function
'
' ��������� �������� ������� �� ��� ������
Function RulesAnalize(K As Long, strX As String) As String
Dim MyString As String, posX As Long, posY As Long
    ' �������� ������ ������ � ��������
    MyString = String$(255, " ")
    GetPrivateProfileString "�������" & Trim(str(K)), "Rule", "", MyString, 255, App.Path & "/" & strX
    MyString = Trim(MyString)
    ' ����������� ������ � ��������
    posX = InStr(1, MyString, "{N}", 1)
    If posX > 0 Then posY = InStr(1, MyString, "{N-1}", 1)
    If posY > 0 Then MyString = Mid(MyString, posX + 3, posY - posX - 3) ' �������� ������ � ��������
    RulesAnalize = Trim(MyString)
End Function
'
' ��������� 2-�� �������� ������� �� ��� ������ (���� ��� ����)
Function Rules2Make(K As Long, strX As String) As String
Dim MyString As String, posX As Long, posY As Long
    ' �������� ������ ������ � ��������
    MyString = String$(255, " ")
    GetPrivateProfileString "�������" & Trim(str(K)), "Rule", "", MyString, 255, App.Path & "/" & strX
    MyString = Trim(MyString)
    ' ����������� ������ � ��������
    posX = InStr(1, MyString, "(", 1)
    If posX > 0 Then posY = InStr(1, MyString, ")", 1)
    ' �������� ������ � ��������
    If posY > 0 Then MyString = Mid(MyString, posY + 1) Else MyString = ""
    Rules2Make = Trim(MyString)
End Function
'
' ����������� ����-������ �� ������� � ����������
' � strX ��� ����� �������
Sub TemplateAn(strX As String)
Dim MyString As String, TempStr As String
Dim nameField As String
Dim posX As Long, posY As Long, filenum As Long
    '
    On Error GoTo errTA
    filenum = FreeFile
    Open strX For Input As #filenum
    Do While Not EOF(filenum)   '
        Input #filenum, MyString
        ' �������� ������-��������� (���������� [])
        posX = InStr(1, MyString, "[", 1):  posY = InStr(1, MyString, "]", 1)
        If posX > 0 And posY > 0 Then nameField = Mid(MyString, posX + 1, posY - posX - 1)
        If InStr(1, MyString, "Cnd", 1) > 0 Then
            ' Cnd=��������� between DATA1(23:00) AND DATA2(23:00)
            posX = InStr(1, MyString, "DATA1(", 1)
            If posX > 0 Then
                ' ��������� ������� �� ���� �� ���������
                posY = InStr(1, MyString, "=", 1)
                
                mstrTable = mstrTable & " WHERE " & Trim(Mid(MyString, posY + 1, InStr(1, MyString, "between", 1) _
                - posY - 1)) & " BETWEEN #" & SQLData(CDate(firstP) - 1) & " " & Mid(MyString, posX + 6, _
                InStr(posX, MyString, ")", 1) - 6 - posX) & "#" '" 23:00#"
            End If
            posX = InStr(1, MyString, "DATA2(", 1)
            If posX > 0 Then
                ' ��������� ������� �� ���� �� ��������� (� �������� � �������)
                mstrTable = mstrTable & " AND #" & SQLData(CDate(secondP)) & " " & Mid(MyString, posX + 6, _
                InStr(posX, MyString, ")", 1) - 6 - posX) & "#" '" 23:00#"
            End If
        ElseIf InStr(1, MyString, "ORDER BY", 1) > 0 Then
            TempStr = " ORDER BY " & nameField & ";"
        End If
        ' ��������� ���������� �� ����
    Loop
    Close #filenum    ' ��������� ����.
    mstrTable = mstrTable & TempStr
errTA:
End Sub
'
' ������� ������ �������� NULL � ���� �� 0
Function NZ(X As Variant) As Variant
NZ = IIf(IsNull(X), 0, X)
End Function
'������ ������ � ���� events.log � ������� ���������� ����
Public Sub writeLog(Text As String)
    Dim logFile As String
    Dim FileNr As Integer
    '���������� ��� ������� ����
    '�������� � �����-��������� ���� � ������� ����
    logFile = CurDir & "\events.log"
    '������� ����-��������
    FileNr = FreeFile:    Open logFile For Append As FileNr
    '�������� ��������� �������
    Print #FileNr, Format(Now, "dd.mm.yy hh:nn:ss ") & " : "; Text
    '������� ����-��������
    Close FileNr
End Sub
'�������� ������ � ������� "������"
Public Function protocol() As String
    protocol = Err.Number & " # " & _
    Left(Err.Description, 200) & " # " & _
    Err.LastDllError & " # " & _
    Err.Source
    'MsgBox Err.Description
End Function
'
Private Sub Text1_Change()
    On Error GoTo TC2_err
    WritePrivateProfileString "Setup", "Database", Text1.Text, _
            App.Path & "/set.ini"  ' ���������� � ������  ����������
    Exit Sub
TC2_err:
        MsgBox Err.Number & "->" & Err.Description
        Resume Next
End Sub
'
Private Sub Text3_Change()
On Error GoTo TC3_err
WritePrivateProfileString "Setup", "Table", Text3.Text, _
            App.Path & "/set.ini"  ' ���������� � ������  ����������
Exit Sub
TC3_err:
    MsgBox Err.Number & "->" & Err.Description
    Resume Next '

End Sub
' �����. �������� ������ � ���������� � ������� ��������
' Ext=1 ��� �������� �������
' Ext=0 ��� ��������������� ���������� � MySQL
' ��� ����. ������������ � ���� ������:
' 1. ��������� ������ �� ������� - � ws.dataarrival
' 2. �������������� �������� � mysql - ��� �������
' strH = ��� ����������� ��������� �������� ("autosut") ��� ������� ("autochs")
Private Function DataPass(FN As String, Ext As Integer, strH As String) As Long
Dim strZ As String ' ����� �������� ��� �������
Dim filenum As Long, I As Long
'
On Error GoTo exit_DataPass
strZ = String$(255, " ")
If Ext = 1 Then
    ' ����������� ���� ��� ������
    filenum = FreeFile
    Open FN For Output As #filenum
    Print #filenum, "" ' �������� ����
    Close #filenum
    ' ������� �� pattern.ini �� ����� ������� � ���� ������� (����/����) ��������������� ������
    If KindOfMode = "1" Then ' ���� ����
        GetPrivateProfileString nameOfTable, "����������", "", strZ, 255, App.Path & "/pattern.ini"
    ElseIf KindOfMode = "0" Then ' ���� ����
        GetPrivateProfileString nameOfTable, "����������", "", strZ, 255, App.Path & "/pattern.ini"
    End If
    strH = FN
ElseIf Ext = 0 Then
    ' � nameOfNode - ��� ������� ����, � ���������� ��� ��� ���� �� sp_uzl
    ' ������� �� pattern.ini �� ����� ������� � ���� ������� (����/����) ��������������� ������
    GetPrivateProfileString nameOfNode, strH, "", strZ, 255, App.Path & "/node.ini"
End If
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
strZ = Trim(strZ) ' ������ ��� �������
If Len(strZ) > 0 Then
    ' ����� �������� ������ �������� �������
    mstrTable = "SELECT * FROM [" & nameOfTable & "]"
    ' ����������� ������
    Call TemplateAn(strZ)  ' � mstrTable - �������� ��������� (�������������� ������)
    DataPass = DataCulc(strZ, mstrTable, strH, Ext) ' ���������� ������ �� �������
Else
    DataPass = 0
End If
Exit Function
exit_DataPass:
DataPass = 0
End Function
'
' ������� ���������� ����������������� ������� �� �������
Function DataCulc(strN As String, mstrTable As String, strTA As String, Optional Ext As Integer) As Long
Dim I As Long ' �������� ���������� ������� �� ���������� �������
Dim J As Long ' ��������� ���������� ������ � �������
Dim lngRule As Long ' ���������� ������ � �������
Dim arrNameFields() As String ' ������ ���� ����� ��� ���. ����������� ��������  �������� ������
Dim arrFields() As Double ' ������ �������� ����� ��� ���������� ������ ��� ����
Dim arrOper() As String ' ������ �������� ��� ������ �������� ������
Dim strRule As String ' ����� �������� ������� ��� �������
Dim strFields As String ' ����� ���������� ����� ��������� �����
Dim strNameFields As String ' ����� �������� ����� ��������� �����
Dim pos As Long ' ������� ; ��� ��������� ���� �����
Dim pos1 As Long ' ������� [��������] ��� ��������� ���� ����� � �������
Dim strTemp As String ' ������ ���������� ��� �������� �������������� ����
Dim strForma As String ' ������ ��� ������ �� ������ ������� ������
Dim strFunc As String ' �������� ������ �������
Dim strFunc1 As String ' �������� ������������� ������ �������
Dim dblFormula1 As Double '��� ���������� ���������� �������
Dim dblFormula2 As Double '��� ���������� ���������� �������
Dim filenum As Long
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim intKU As Integer
Dim strZ As String
Dim a As Double, b As Double, sdata As String, sdata1 As String
    On Error GoTo exit_DataCulc
    If IsNull(Ext) Then Ext = 1
    ' ���������� � ����� ������
    Me.Data1.DatabaseName = Text1.Text
    Me.Data1.RecordSource = mstrTable
    Me.Data1.Refresh
    If Not (Me.Data1.Recordset.BOF And Me.Data1.Recordset.EOF) Then
        Me.Data1.Recordset.Requery
        ' ���������������� ����������
        Me.Data1.Recordset.MoveFirst
        ' ���������� ���������� ������ � ����� ������� � ������� ��������������� ������
        lngRule = CountRule(strN)
        If lngRule > 0 Then
            ReDim arrFields(lngRule): ReDim arrOper(lngRule): ReDim arrNameFields(lngRule)
        End If
        If Ext = 1 Then
            ' ������� ��������� �������
            strForma = TipArh & KindOfArh
            With Me.Data1.Recordset
                For I = 0 To .Fields.Count - 1
                    strForma = strForma & ";" & .Fields(I).Name & "=" & .Fields(I)
                Next
                strForma = strForma & vbCrLf & TipArh & KindOfArh: .MoveLast
                For I = 0 To .Fields.Count - 1
                    strForma = strForma & ";" & .Fields(I).Name & "=" & .Fields(I)
                Next
                .MoveFirst
            End With
            filenum = FreeFile
            Open strTA For Output As #filenum
            Print #filenum, strForma
        ElseIf Ext = 0 Then
            ' connectionString;PORT=3306
            Set cnn = New ADODB.Connection
            cnn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                                  & "SERVER=192.168.100.23;" _
                                  & " DATABASE=askute;" _
                                  & "UID=root;PWD=111; OPTION=16"
            cnn.ConnectionTimeout = 30
            cnn.Mode = adModeReadWrite
            cnn.Open
            ' �� �������� �������/�� ������� �������
            Set rst = New ADODB.Recordset
            rst.Open "teplo_" & IIf(Right(strTA, 3) = "sut", "s", "h") & "r", _
                                    cnn, adOpenKeyset, adLockOptimistic, adCmdTable
        End If
        I = 0: strForma = ""
        ' ������ ����� ���������� � 0 �������������� ��������
        With Me.Data1.Recordset
            ' ��������� �������� ����� ��������� � ��������
            If lngRule > 0 Then
                For J = 1 To lngRule
                    strRule = RulesAnalize(J, strN)
                    arrOper(J) = Left$(strRule, 1) & "," & Rules2Make(J, strN)
                    arrNameFields(J) = Trim(Mid(strRule, 2))
                    arrFields(J) = Format(NZ(.Fields(arrNameFields(J))), "#0.000")
                Next J
            End If
            For I = 0 To .RecordCount - 1
                ' �������� � ���� ����������� ��������, ��������� 1-�� ������
                ' � ����� ���������� ������������ �������, ����� � ������� �������
                If .AbsolutePosition = 0 Then
                    '����� ����� �� * .ptn
                    strNameFields = NameOfFields(strN)
                    strForma = strNameFields
                Else
                    ' ����� ����� ����������� ������� � ���������� ����� �����
                    strFields = strNameFields
                    pos = InStr(1, strFields, ";", 1)
                    Do While pos
                        ' ����� ��� ����
                        strTemp = Mid(strFields, 1, pos - 1)
                        strFields = Mid(strFields, pos + 1)
                        pos = InStr(1, strFields, ";", 1)
                        ' �������� ��� � ������� ����������� ������
                        strFunc = "": strFunc1 = ""
                        ' ����� �������� ���������� �������� �������
                        If InStr(1, strTemp, "�������", 1) > 0 Then
                            ' ������ ������ �������
                            strFunc = String$(255, " ")
                            GetPrivateProfileString strTemp, "fx", "", strFunc, 255, App.Path & "/" & strN
                            strFunc = Left$(strFunc, InStr(1, strFunc, Chr(0), 1) - 1)
                            strFunc = Trim(strFunc)
                            ' ������ �������
                            If Left$(strFunc, 1) = "(" Then
                                pos1 = InStr(1, strFunc, ")", 1)
                                ' �������� � strFunc1 �� ��� �� ��������
                                strFunc1 = Mid(strFunc, pos1 + 1)
                                ' �������� � strFunc �� ��� ������ ������
                                strFunc = Mid(strFunc, 2, pos1 - 2)
                            End If
                            ' ���������� ����� �������� ���������
                            pos1 = InStr(1, strFunc, "*", 1)
                            If pos1 = 0 Then
                                pos1 = InStr(1, strFunc, "/", 1)
                                If pos1 = 0 Then
                                    pos1 = InStr(1, strFunc, "+", 1)
                                    If pos1 = 0 Then
                                        pos1 = InStr(1, strFunc, "-", 1)
                                    End If
                                End If
                            End If
                            ' ���������� ������ (����������� �������� ��� ��������: +, -, *, /) ���
                            ' ����� ����������� � �������
                            If IsNumeric(Trim(Left$(strFunc, pos1 - 1))) Then _
                                dblFormula1 = CDbl(Trim(Left$(strFunc, pos1 - 1))) _
                            Else _
                                dblFormula1 = GetNumber(Trim(Left$(strFunc, pos1 - 1)) _
                                                    , NZ(.Fields(Trim(Left$(strFunc, pos1 - 1)))) _
                                                                    , lngRule, arrOper, arrNameFields, arrFields)
                            If IsNumeric(Trim(Mid(strFunc, pos1 + 1))) Then _
                                                dblFormula2 = CDbl(Trim(Mid(strFunc, pos1 + 1))) _
                            Else dblFormula2 = GetNumber(Trim(Mid(strFunc, pos1 + 1)), _
                                                                    NZ(.Fields(Trim(Mid(strFunc, _
                                                        pos1 + 1)))), lngRule, arrOper, arrNameFields, arrFields)
                            ' ��������������� ���������� �������
                            dblFormula1 = doOperate(Mid(strFunc, pos1, 1), dblFormula1, dblFormula2)
                            If Len(strFunc1) > 0 Then dblFormula1 = doOperate(Left$(strFunc1, 1), _
                                                                dblFormula1, CDbl(Trim(Mid(strFunc1, 2))))
                            ' ��������������� ��� ������
                            strForma = strForma & Format(dblFormula1, "#0.000") & ";"
                        Else
                            ' ��������� ����� ��� ������: ����, �����, ����� � �.�.
                            If TypeName(.Fields(strTemp)) = "String" Then _
                                                    strForma = strForma & .Fields(strTemp) & ";"
                            If IsNull(.Fields(strTemp)) Then strForma = strForma & "0" & ";" '& "NULL" & ";"
                            If IsNumeric(.Fields(strTemp)) Then
                                ' ���������� ������ (����������� �������� ��� ��������: +, -, *, /)
                                dblFormula1 = GetNumber(strTemp, NZ(.Fields(strTemp)), _
                                                        lngRule, arrOper, arrNameFields, arrFields)
                                ' ������������ ������ ���������� �������
                                J = IndxOper(strTemp, lngRule, arrNameFields)
                                If J > 0 Then
                                    pos1 = InStr(1, arrOper(J), ",", 1) + 1
                                    If Len(Mid(arrOper(J), pos1, 1)) > 0 Then dblFormula1 = doOperate _
                                    (Mid(arrOper(J), pos1, 1), dblFormula1, CDbl(Trim(Mid(arrOper(J), pos1 + 1))))
                                End If
                                strForma = strForma & Format(dblFormula1, "#0.000") & ";"
                            End If
                            If IsDate(.Fields(strTemp)) Then _
                                strForma = strForma & Format(.Fields(strTemp), "dd.mm.yy hh:mm") & ";"
                        End If
                    Loop
                     ' ��������� �������� ����� ��������� � ��������
                    If lngRule > 0 Then
                        For J = 1 To lngRule
                            strRule = RulesAnalize(J, strN)
                            arrOper(J) = Left$(strRule, 1) & "," & Rules2Make(J, strN)
                            arrNameFields(J) = Trim(Mid(strRule, 2))
                            arrFields(J) = Format(NZ(.Fields(arrNameFields(J))), "#0.000")
                        Next J
                    End If
                    '
                End If
                If Ext = 1 Then
                    Print #filenum, strForma
                ElseIf Ext = 0 Then ' ����� �������� ������ � teplo_hr, teplo_sr
                    If .AbsolutePosition > 0 Then
                        ' �������� ��� ���� �� ����� ���� � ������� � �������
                        strZ = String$(255, " ")
                        GetPrivateProfileString nameOfNode, "code", "", strZ, 255, App.Path & "/node.ini"
                        strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
                        intKU = CInt(Trim(strZ)) '  ����� ���� ����
                        With rst
                            .Requery
                            .AddNew
                            .Fields("kod_uzl") = intKU
                            ' ��������� ������ ������ ��� �������� � �����
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("data") = DateValue(Left(strForma, pos - 1))
                            ' ����� ��� �����.���.
                            If strTA = "autochs" Then .Fields("vremy") = TimeValue(Left(strForma, pos - 1))
                            strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("w1") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("w2") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("v1") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("v2") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("vrem_n") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("w3l") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("v3l") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("t1") = Left(strForma, pos - 1) ' ����������� �� �������
                            sdata = Left(strForma, pos - 1):    strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("t2") = Left(strForma, pos - 1) ' ����������� �� �������
                            sdata1 = Left(strForma, pos - 1):   strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("p1") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("p2") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("w3z") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            pos = InStr(1, strForma, ";", 1)
                            .Fields("v3z") = Left(strForma, pos - 1): strForma = Mid(strForma, pos + 1)
                            ' ��������� ���� �� �������� ���-��
                            ' ����������� ��� ������ �� ������� sp_tr
                            strZ = String$(255, " ")
                            GetPrivateProfileString "Culc", "koefA", "", strZ, 255, App.Path & "/set.ini"
                            strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
                            a = Val(strZ)
                            strZ = String$(255, " ")
                            GetPrivateProfileString "Culc", "koefB", "", strZ, 255, App.Path & "/set.ini"
                            strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1) ' ������� ������� ����� ������
                            b = Val(strZ)
                            ' ��������� ������ �� ������-���� �������
                            a = Val(sdata) * a + b
                            .Fields("t2r") = a
                            .Fields("dt") = Val(sdata1) - a
                            .Update
                        End With
                        rst.Close
                        cnn.Close
                        Set cnn = Nothing
                    End If
                End If
                strForma = ""
                .MoveNext
            Next I
        End With
        If Ext = 1 Then
            I = Seek(filenum)
            Close #filenum    ' ��������� ����.
        Else
            I = 1
        End If
    End If
    Data1.DatabaseName = "":  Data1.RecordSource = "":  Data1.Refresh
    DataCulc = I
    Exit Function
exit_DataCulc:
    If Len(nameOfTable) > 0 Then writeLog (nameOfTable & ":" & protocol())
    If Ext = 1 Then
        Close #filenum
    ElseIf Ext = 0 Then
        If rst.State = adStateOpen Then
            rst.CancelBatch
            rst.Close
        End If
        If cnn.State = adStateOpen Then cnn.Close
    End If
    DataCulc = False
End Function
' ��������� ����� �������
Function NewPos(strQ As String, num As Long) As String
Dim pos As Long, pos1 As Long, I As Long
I = 0
pos = InStr(1, strQ, ";", 1) ' ���� ������ ���������
Do While pos
    I = I + 1
    If I = num Then Exit Do
    pos1 = pos
    pos = InStr(pos1 + 1, strQ, ";", 1) ' ������ ���������
Loop
NewPos = CStr(pos) & "/" & CStr(pos1)
End Function

