VERSION 5.00
Begin VB.Form frmOpenFiles 
   Caption         =   "������� ���� ������"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   LinkTopic       =   "Form4"
   ScaleHeight     =   6765
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox swtchReg 
      Alignment       =   1  'Right Justify
      Caption         =   "����� ����(false)/����(true) ��� ������ ""����� ����"""
      Height          =   375
      Left            =   -30
      TabIndex        =   10
      Top             =   6255
      Value           =   1  'Checked
      Width           =   3045
   End
   Begin VB.TextBox txtPar 
      Height          =   345
      Left            =   2835
      TabIndex        =   9
      Text            =   "���961"
      Top             =   5910
      Width           =   2115
   End
   Begin VB.TextBox txtVoda 
      Height          =   390
      Left            =   2835
      TabIndex        =   8
      Text            =   "VZLJOT"
      Top             =   5475
      Width           =   2115
   End
   Begin VB.TextBox txtVzljot 
      Height          =   405
      Left            =   2835
      TabIndex        =   7
      Text            =   "��˨� ���"
      Top             =   5025
      Width           =   2115
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   2475
      Pattern         =   "*.txt;*.arh;*.ard;*.csv"
      TabIndex        =   3
      Top             =   345
      Width           =   2490
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   4410
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
      Height          =   315
      Left            =   3990
      TabIndex        =   0
      Top             =   4530
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "�������� ����� ��� ������ �� ����"
      Height          =   465
      Left            =   0
      TabIndex        =   13
      Top             =   5865
      Width           =   2835
   End
   Begin VB.Label Label5 
      Caption         =   "�������� ����� ��� ������ �� ����"
      Height          =   465
      Left            =   0
      TabIndex        =   12
      Top             =   5460
      Width           =   2835
   End
   Begin VB.Label Label4 
      Caption         =   "�������� ����� ��� ������ ""Visikal Pro"""
      Height          =   465
      Left            =   0
      TabIndex        =   11
      Top             =   4995
      Width           =   2835
   End
   Begin VB.Label Label3 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   75
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "����� �������:"
      Height          =   270
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "�����..."
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
      Left            =   2610
      TabIndex        =   4
      Top             =   4455
      Width           =   1140
   End
End
Attribute VB_Name = "frmOpenFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
Public CodeFiles As Long ' ��� ��������������� ����
'
Private SavePath As String ' ������� ���� ������� �����
Private NewPath As String ' ������� ���� ��������� �����
Private KoefA As Double
Private KoefB As Double
'
Const Stwips = 567
' ��������� ������ �������� ������ � ����� �����
Const Wn = 239000
Const Vn = 1000000
'
'
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Me.File1.Path = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Me.Dir1.Path = Me.Drive1.Drive
End Sub

Private Sub File1_DblClick()
Dim FileNumber As Integer
Dim ThisFileName As String
On Error GoTo FuckExit
Me.Label1.Visible = True
SavePath = CurDir
bTipArh = False
ThisFileName = Me.File1.FileName ' �������� ��� ������������ �����
ChDir Me.File1.Path ' ������� ��������� ����
NewPath = CurDir
CodeFiles = 0
CodeFiles = PreAnalizeFiles(ThisFileName) '(�������� ������ ��� ����� )
' ������� ��������� ������������� �����
If CodeFiles > 0 Then
    ' �������� ���� tempbase.csv �� ���������� �������
    ChDir SavePath ' ��������� � ������� �����
    FileNumber = FreeFile
    If frmGraph.chkTek.Value Then
        Open "tempgraf.csv" For Output As #FileNumber
        Close #FileNumber
    End If
    KoefA = ReadNParam("KOEFA"): KoefB = ReadNParam("KOEFB")
    ChDir NewPath ' ��������� � ����� �����
    ' ������������ tempbase.csv
    Call MakeRezFile(CodeFiles, ThisFileName)
Else
    ChDir SavePath ' ��������� � ������� �����
    MsgBox "���� �� �������������� �������!"
    Exit Sub
End If
' ���������� �������������� ������
ChDir SavePath ' ��������� � ������� �����
Me.Label1.Visible = False
Unload frmOpenFiles
Exit Sub
FuckExit:
    If Err.Number = cdlCancel Then Exit Sub Else Resume Next
End Sub
'
'
' ���� ������������ ������ ��� ������
Function MakeRezFile(CodeArh As Long, ThisFile As String) As Boolean
Select Case CodeArh
Case 11
    Call ReadHourVis(ThisFile) ' ������� ����� ����� ���
Case 12
    Call ReadDayVis(ThisFile) ' �������� ����� ����� ���
Case 21
    Call ReadHourHVod(ThisFile) ' ������� ����� ����
Case 22
    Call ReadDayHVod(ThisFile) ' �������� ����� ����
Case 31
    Call ReadHourPar(ThisFile) ' ������� ������ ����
Case 32
    Call ReadDayPar(ThisFile) ' �������� ������ ����
Case 41, 42
    Call RepTSRV(ThisFile, CodeArh) ' ���������� �����-����
Case 5
    Call OpenCurFile(ThisFile) ' ���������� ������� ����
End Select
End Function
'
' ������� ������� ��������� ������� ����
Function ReadDayPar(ThisFile As String) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String
Dim dblX As Double
On Error Resume Next
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead ' ����� ���������
            ' ������ ������ ������ = ����
            str1 = Mid(sHead, 1, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ������ ������ ������ = �����
            str1 = Mid(sHead, 11, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����������� �1
            str1 = Mid(sHead, 60, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� �1
            str1 = Mid(sHead, 70, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� �1
            str1 = Mid(sHead, 90, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� ������� W1
            str1 = Mid(sHead, 99, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����������� �2
            str1 = Mid(sHead, 161, 14)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� �2
            str1 = Mid(sHead, 175, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� �2
            str1 = Mid(sHead, 195, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� ������� W2
            str1 = Mid(sHead, 204, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� ��������� 1
            str1 = Mid(sHead, 117, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����� ������ 1
            dblX = Val(str1): dblX = 24 - dblX
            strRez = strRez & Trim(str(dblX)): strRez = strRez & ";"
            ' ����� ��������� 2
            str1 = Mid(sHead, 222, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����� ������ 2
            dblX = Val(str1): dblX = 24 - dblX
            strRez = strRez & Trim(str(dblX))
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez, 2)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' ������� ������� �������� ������� ����
Function ReadHourPar(ThisFile As String) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String
On Error Resume Next
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead ' ����� ���������
            ' ������ ������ ������ = ����
            str1 = Mid(sHead, 1, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ������ ������ ������ = �����
            str1 = Mid(sHead, 11, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����������� �1
            str1 = Mid(sHead, 24, 16)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� �1
            str1 = Mid(sHead, 40, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� �1
            str1 = Mid(sHead, 60, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� ������� W1
            str1 = Mid(sHead, 69, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����������� �2
            str1 = Mid(sHead, 101, 14)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� �2
            str1 = Mid(sHead, 115, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� �2
            str1 = Mid(sHead, 135, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' �������� ������� W2
            str1 = Mid(sHead, 144, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            strRez = strRez & "-": strRez = strRez & ";" ' ����� ��������� 1
            strRez = strRez & "-": strRez = strRez & ";" ' ����� ������ 1
            strRez = strRez & "-": strRez = strRez & ";" ' ����� ��������� 2
            strRez = strRez & "-" ' ����� ������ 2
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez, 2)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' ������� ��������� ������ ����
Function RepTSRV(ThisFile As String, TipArh As Long) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim dblT As Double, dblT1 As Double, i As Integer
Dim pos As Long, pos1 As Long
Dim H As Integer, strV(12) As String
Dim dblt2 As Double, delta As Double, a As Double, b As Double ' ��� �������� ����-��
'
On Error Resume Next
        a = KoefA: b = KoefB
        If TipArh = 41 Then H = 60 Else H = 1
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead ' ����� ���������
            ' ������ ������ ������ = ����-�����
            pos = InStr(1, sHead, ";", 1) ' ���� ������ ���������
            str1 = Mid(sHead, 1, pos - 1)
            strV(0) = Trim(str1)
            ' ����� �� �������
            pos = InStr(1, sHead, ";", 1) ' ���� ������ ���������
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' ������ ����������
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(1) = Format(str1, "0.00")
            dblW = CDbl(str1)
            ' ����� �� �������
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(2) = Format(str1, "0.00")
            dblX = CDbl(str1)
            ' ��������� ������������ �������' ����� ����/����
            If Me.swtchReg.Value Then dblW = dblW + dblX Else dblW = dblW - dblX
            ' ����� �� �������
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(3) = Format(str1, "0.00")
            dblm = CDbl(str1)
            ' ����� �� �������
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(4) = Format(str1, "0.00")
            dblX = CDbl(str1)
            ' ��������� ������������ �����' ����� ����/����
            If Me.swtchReg.Value Then dblm = dblm + dblX Else dblm = dblm - dblX
            ' ����� ��������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 18
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT = H * CDbl(str1)
            str1 = Mid(sHead, pos + 1)
            dblT1 = H * CDbl(str1)
            strV(5) = Format(dblT, "0.00") & "-" & Format(dblT1, "0.00")
            ' �������� �� �������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 12
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(6) = Format(str1, "0.00")
            ' �������� �� �������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 13
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(7) = Format(str1, "0.00")
            ' ������ ������������ ��������
            strV(8) = Format(dblW, "0.00")
            strV(9) = Format(dblm, "0.00")
            ' ����� ������ ������� ���������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 16
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT = H * CDbl(str1)
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT1 = H * CDbl(str1)
            strV(10) = Format(dblT, "0.00") & "-" & Format(dblT1, "0.00")
            ' ����������� �� �������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 7
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            ' ��������� ���� �� �������� ���-��
            strV(11) = Format(str1, "0.00")
            ' ����������� �� �������
            pos = InStr(1, sHead, ";", 1)  ' ������ ���������
            For i = 1 To 9
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' ������ ���������
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(12) = Format(str1, "0.00")
            ' ��������� ���� �� �������� ���-��
            dblt2 = CDbl(strV(11)) * a + b:  delta = CDbl(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & _
                    Format(delta, "0.00")
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
End Function
'
' ������� ������ ��������� ������ ����
Function ReadDayHVod(ThisFile As String) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String
Dim lX As Long, lM As Long
On Error Resume Next
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead ' ����� ���������
            ' ������ ������ ������ = ����-�����
            str1 = Mid(sHead, 1, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����� ������������
            str1 = Mid(sHead, 62, 24)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' ����� ������
            str1 = Mid(sHead, 86, 32)
            strRez = strRez & Trim(str1) & " ��": strRez = strRez & ";" ' ������������ �������� ������
            ' ����� ������
            lX = 24 - Val(Mid(str1, 1, InStr(1, str1, "��", vbTextCompare) - 1))
            lM = 60 - Val(Mid(str1, InStr(1, str1, "��", vbTextCompare) + 2))
            ' ������������ �������� ������
            strRez = strRez & Trim(str(lX)) & " � " & Trim(str(lM)) & " �"
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez, 1)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' ������� ������ �������� ������ ����
Function ReadHourHVod(ThisFile As String) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String, lX As Long
On Error Resume Next
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead ' ����� ���������
            ' ������ ������ ������ = ����-�����
            str1 = Mid(sHead, 1, 15)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' ����� ������������
            str1 = Mid(sHead, 64, 24)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' ������������ �������� ������
            ' ����� ������
            str1 = Mid(sHead, 88, 34)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' ������������ �������� ������
            ' ����� ������
            lX = 3600 - Val(str1)
            strRez = strRez & Trim(str(lX))  ' ������������ �������� ������
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez, 1)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' ������� ������ �������� ������ VISIKAL
Function ReadHourVis(ThisFile As String) As Boolean
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim strRez As String
Dim strV(12) As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim lngT As Long, bTipArh As Boolean
' ��� �������� ����-��
Dim dblt2 As Double, delta As Double, str3 As String, a As Double, b As Double
'
On Error GoTo errOPA
        a = KoefA: b = KoefB: bTipArh = False
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        If InStr(1, sHead, "�-0", 1) > 0 Then bTipArh = True ' �-0=����
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead1 ' ����� ���������
            ' ������ ������ ������ = ����-�����
            str1 = Mid(sHead1, 1, 15)
            If InStr(1, str1, ":") > 0 Then str1 = Left(str1, InStr(1, str1, ":") - 1) & _
                    " " & Mid(str1, InStr(1, str1, ":") + 1) & ":00" Else str1 = str1 & " 00:00"
            strV(0) = Trim(str1)
            ' ����� �� �������
            str1 = Mid(sHead1, 16, 16)
            str2 = Mid(sHead, 16, 16)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            strV(1) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 32, 16)
            str2 = Mid(sHead, 32, 16)
            dblW = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            ' ��������� ������������ ������� ' ����� ����/����
            If bTipArh Then dblW = dblW + dblX Else dblW = dblW - dblX
            strV(2) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 48, 13)
            str2 = Mid(sHead, 48, 13)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            strV(3) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 61, 13)
            str2 = Mid(sHead, 61, 13)
            dblm = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            ' ��������� ������������ ����� ' ����� ����/����
            If bTipArh Then dblm = dblm + dblX Else dblm = dblm - dblX
            strV(4) = Format(dblX, "0.00")
            ' ����� ��������
            str1 = Mid(sHead1, 152, 14)
            str2 = Mid(sHead, 152, 14)
            lngT = CLng(str1) - CLng(str2)
            strV(5) = Format(lngT, "0.00")
            ' �������� �� �������
            str1 = Mid(sHead1, 166, 13):  strV(6) = Format(Val(str1), "0.00")
            ' �������� �� �������
            str1 = Mid(sHead1, 179, 14):  strV(7) = Format(Val(str1), "0.00")
            strV(8) = Format(dblW, "0.00"):   strV(9) = Format(dblm, "0.00")
            ' ����� ������� �����
            strV(10) = Format((60 - lngT), "0.00")
            ' ����������� �� �������
            str1 = Mid(sHead1, 100, 9):   strV(11) = Format(Val(str1), "0.00")
            ' ����������� �� �������
            str1 = Mid(sHead1, 109, 9):   strV(12) = Format(Val(str1), "0.00")
            ' �������� ���-��
            dblt2 = Val(strV(11)) * a + b:   delta = Val(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & _
                    Format(delta, "0.00")
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
        Exit Function
errOPA:
    Resume Next
End Function
'
' ������� ������ ��������� ������ VISIKAL
Function ReadDayVis(ThisFile As String) As Boolean
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim strRez As String
Dim strV(12) As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim lngT As Long, bTipArh As Boolean
' ��� �������� ����-��
Dim dblt2 As Double, delta As Double, str3 As String, a As Double, b As Double
'
On Error GoTo errOPAd
        a = KoefA: b = KoefB: bTipArh = False
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        If InStr(1, sHead, "�-0", 1) > 0 Then bTipArh = True
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' ���� �� ����� �����
            Line Input #FileNumber, sHead1 ' ����� ���������
            ' ������ ������ ������ = ����-�����
            str1 = Mid(sHead1, 1, 12)
            If InStr(1, str1, ":") > 0 Then str1 = Left(str1, InStr(1, str1, ":") - 1) & _
                " " & Mid(str1, InStr(1, str1, ":") + 1) & ":00" Else str1 = str1 & " 00:00"
            strV(0) = Trim(str1)
            ' ����� �� �������
            str1 = Mid(sHead1, 13, 16)
            str2 = Mid(sHead, 13, 16)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            strV(1) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 29, 16)
            str2 = Mid(sHead, 29, 16)
            dblW = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            ' ��������� ������������ �������' ����� ����/����
            If bTipArh Then dblW = dblW + dblX Else dblW = dblW - dblX
            strV(2) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 45, 13)
            str2 = Mid(sHead, 45, 13)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            strV(3) = Format(dblX, "0.00")
            ' ����� �� �������
            str1 = Mid(sHead1, 58, 13)
            str2 = Mid(sHead, 58, 13)
            dblm = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            ' ��������� ������������ �����' ����� ����/����
            If bTipArh Then dblm = dblm + dblX Else dblm = dblm - dblX
            strV(4) = Format(dblX, "0.00")
            ' ����� ��������
            str1 = Mid(sHead1, 133, 14)
            str2 = Mid(sHead, 133, 14)
            lngT = CLng(str1) - CLng(str2)
            strV(5) = Format(lngT / 60, "0.00")
            ' �������� �� �������
            str1 = Mid(sHead1, 147, 13):     strV(6) = Format(Val(str1), "0.00")
            ' �������� �� �������
            str1 = Mid(sHead1, 160, 14):     strV(7) = Format(Val(str1), "0.00")
            strV(8) = Format(dblW, "0.00"):  strV(9) = Format(dblm, "0.00")
            '  ��������  �����
            str(10) = Format((1440 - lngT) / 60, "0.00")
            ' ����������� �� �������
            str1 = Mid(sHead1, 97, 9):  strV(11) = Format(Val(str1), "0.00")
            ' ����������� �� �������
            str1 = Mid(sHead1, 106, 9): strV(12) = Format(Val(str1), "0.00")
            ' ��������� ���� �� �������� ���-��
            dblt2 = Val(strV(11)) * a + b:   delta = Val(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & Format(delta, "0.00")
            ' ��������������� ���������� ������ � ������ (������ [,] �� [.])
            strRez = RemakeS(strRez)
           ' ������ � �������� ����
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
        Exit Function
errOPAd:
    Resume Next
End Function
'
' ������� ������������ ��������� �����
Function MakeTempbase(strExit As String, Optional xHead As Integer) As Integer
Dim FileNumber As Integer
On Error Resume Next
If IsNull(xHead) Then xHead = 0
    FileNumber = FreeFile
    ChDir SavePath ' ��������� � ������� �����
    Open "tempgraf.csv" For Append As #FileNumber
    If LOF(FileNumber) = 0 Then
        If xHead = 0 Then
                Print #FileNumber, "���������;W1;W2;m1;m2;T��;�������1;�������2;�������3;�������4;" & _
                                                    "�������5;�������6;�������7;�������8;�������9;t2r;dt"
        ElseIf xHead = 1 Then
            Print #FileNumber, "DateTime;W1;�������1;�������2"
        ElseIf xHead = 2 Then
            Print #FileNumber, "Date;Time;t1;P1;M1;W1;t2;P2;M2;W2;Tr1;To1;Tr2;To2"
       End If
    End If
    Print #FileNumber, strExit ' ������ ���������
    Close #FileNumber
    ChDir NewPath
    MakeTempbase = 1
End Function
'
' ������� ���������������� ������� �����
Function PreAnalizeFiles(strNameFile As String) As Long
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim iNum As Integer, i As Integer, ExitCode As String
Dim k As Integer
On Error Resume Next
    k = 0: ExitCode = 0
    FileNumber = FreeFile
    Open strNameFile For Input As #FileNumber
    For i = 1 To 6 ' ��������� ���� ��� ���������� ������ �����
        Line Input #FileNumber, sHead ' ����� ���������
        'sHead = ToOEM(sHead)
        sHead1 = ToAnsi(sHead)
        ' ������ �������� �����
        ' ���� �������� ������ �� VISIKAL PRO
        iNum = InStr(1, sHead, Me.txtVzljot, vbTextCompare)
        If iNum > 0 Then ExitCode = "1"
        ' ���� ������ �� ����
        iNum = InStr(1, sHead, Me.txtVoda, vbTextCompare)
        If iNum > 0 Then ExitCode = "2"
        ' ���� ������ �� ����
        iNum = InStr(1, sHead1, Me.txtPar, vbTextCompare)
        If iNum > 0 Then ExitCode = "3"
        ' ���� ������ �� ������� ���������
        iNum = InStr(1, sHead1, "Journal", vbTextCompare)
        If iNum > 0 Then
            PreAnalizeFiles = "5": Exit Function
        End If
        iNum = InStr(1, sHead1, "Par", vbTextCompare)
        If iNum > 0 Then
            PreAnalizeFiles = "6": Exit Function
        End If
        iNum = InStr(1, sHead1, "HV", vbTextCompare)
        If iNum > 0 Then
            PreAnalizeFiles = "7": Exit Function
        End If
        ' ���� ������ �� �� �����-����
        iNum = InStr(1, sHead, ";", vbTextCompare): k = 1
        iNum = InStr(iNum + 1, sHead, ";", vbTextCompare): k = 2
        iNum = InStr(iNum + 1, sHead, ";", vbTextCompare): k = 3
        If iNum > 0 And i = 1 And k = 3 Then
            Line Input #FileNumber, sHead ' ����� ���������
            iNum = InStr(1, sHead, "23:00", vbTextCompare)
            Line Input #FileNumber, sHead ' ����� ���������
            k = InStr(1, sHead, "23:00", vbTextCompare)
            Close #FileNumber
            If iNum > 0 And k > 0 Then ExitCode = "42" Else ExitCode = "41"
            PreAnalizeFiles = CLng(ExitCode)
            Exit Function
        End If
        iNum = InStr(1, sHead & sHead1, "���", vbTextCompare)
        If iNum > 0 Then ExitCode = ExitCode & "1"
        iNum = InStr(1, sHead & sHead1, "���", vbTextCompare)
        If iNum > 0 Then ExitCode = ExitCode & "2"
    Next
    Close #FileNumber
PreAnalizeFiles = CLng(ExitCode)
End Function
'
Private Sub Form_Load()
On Error Resume Next
Me.Dir1.Path = CurDir ' ������ � ������� �����
Me.Label1.Visible = False
End Sub
' �������� ������� ���������� ������ ��� ���������� �������
Function OpenCurFile(stFileName As String)
' �������� ������ ��� ������ � �������
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String
Dim priznak As String
On Error GoTo errPol
' ������� ����
If OpenCSV(stFileName) Then
    filenum = FreeFile ' ���������� c��������� �����
    Open "tempbase.csv" For Input As #filenum ' ������� ����
    ' ���������� ������ 3-� ������
    Line Input #filenum, strX '
    Line Input #filenum, strX '
    Line Input #filenum, strX '
    ' ��������� ������ �������� ��� ����
    Do While Not EOF(filenum)
        Line Input #filenum, strX '
        strX = RemakeS(strX) ' ������ [,] �� [.]
        strX = strX & ";;"
        Call MakeTempbase(strX)
    Loop
    Close #filenum ' ������� ����
Else
    MsgBox "���� �� ��������"
End If
Exit Function
errPol:
Resume Next
End Function
' ������� ��������� ������ �� DOS � WIN
Function ToAnsi(S As String) As String
    Dim ss As String
    ss = S: OemToChar S, ss: ToAnsi = ss
End Function
' ������� ��������� ������ �� WIN � DOS
Function ToOEM(S As String) As String
    Dim ss As String
    ss = S: CharToOem S, ss: ToOEM = ss
End Function

