VERSION 5.00
Begin VB.Form frmOpenFiles 
   Caption         =   "Открыть файл данных"
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
      Caption         =   "Режим зима(false)/лето(true) для данных ""Отчет ТСРВ"""
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
      Text            =   "СПТ961"
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
      Text            =   "ВЗЛЁТ ТСР"
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
      Caption         =   "Выход"
      Height          =   315
      Left            =   3990
      TabIndex        =   0
      Top             =   4530
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "Ключевое слово для данных по пару"
      Height          =   465
      Left            =   0
      TabIndex        =   13
      Top             =   5865
      Width           =   2835
   End
   Begin VB.Label Label5 
      Caption         =   "Ключевое слово для данных по воде"
      Height          =   465
      Left            =   0
      TabIndex        =   12
      Top             =   5460
      Width           =   2835
   End
   Begin VB.Label Label4 
      Caption         =   "Ключевое слово для данных ""Visikal Pro"""
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
      Caption         =   "Режим расчета:"
      Height          =   270
      Left            =   45
      TabIndex        =   5
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Ждите..."
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
Public CodeFiles As Long ' код характеризующий файл
'
Private SavePath As String ' хранить путь рабочей папки
Private NewPath As String ' хранить путь исходного файла
Private KoefA As Double
Private KoefB As Double
'
Const Stwips = 567
' константы сброса счетчика объема и массы тепла
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
ThisFileName = Me.File1.FileName ' получить имя открываемого файла
ChDir Me.File1.Path ' открыть выбранный путь
NewPath = CurDir
CodeFiles = 0
CodeFiles = PreAnalizeFiles(ThisFileName) '(передать только имя файла )
' функция обработки распознанного файла
If CodeFiles > 0 Then
    ' очистить файл tempbase.csv от предыдущих записей
    ChDir SavePath ' вернуться в текущую папку
    FileNumber = FreeFile
    If frmGraph.chkTek.Value Then
        Open "tempgraf.csv" For Output As #FileNumber
        Close #FileNumber
    End If
    KoefA = ReadNParam("KOEFA"): KoefB = ReadNParam("KOEFB")
    ChDir NewPath ' вернуться в новую папку
    ' сформировать tempbase.csv
    Call MakeRezFile(CodeFiles, ThisFileName)
Else
    ChDir SavePath ' вернуться в текущую папку
    MsgBox "Файл не установленного формата!"
    Exit Sub
End If
' отобразить сформированные данные
ChDir SavePath ' вернуться в текущую папку
Me.Label1.Visible = False
Unload frmOpenFiles
Exit Sub
FuckExit:
    If Err.Number = cdlCancel Then Exit Sub Else Resume Next
End Sub
'
'
' файл формирования данных для печати
Function MakeRezFile(CodeArh As Long, ThisFile As String) As Boolean
Select Case CodeArh
Case 11
    Call ReadHourVis(ThisFile) ' часовой архив ВЗЛЕТ ТСР
Case 12
    Call ReadDayVis(ThisFile) ' суточный архив ВЗЛЕТ ТСР
Case 21
    Call ReadHourHVod(ThisFile) ' часовой архив воды
Case 22
    Call ReadDayHVod(ThisFile) ' суточный архив Воды
Case 31
    Call ReadHourPar(ThisFile) ' часовой расход пара
Case 32
    Call ReadDayPar(ThisFile) ' суточный расход пара
Case 41, 42
    Call RepTSRV(ThisFile, CodeArh) ' обработать отчет-ТСРВ
Case 5
    Call OpenCurFile(ThisFile) ' обработать текущий файл
End Select
End Function
'
' функция расчета суточного расхода пара
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
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead ' брать построчно
            ' анализ взятой строки = дата
            str1 = Mid(sHead, 1, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' анализ взятой строки = время
            str1 = Mid(sHead, 11, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' температура Т1
            str1 = Mid(sHead, 60, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' давление Р1
            str1 = Mid(sHead, 70, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' масса М1
            str1 = Mid(sHead, 90, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' тепловая энергия W1
            str1 = Mid(sHead, 99, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' температура Т2
            str1 = Mid(sHead, 161, 14)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' давление Р2
            str1 = Mid(sHead, 175, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' масса М2
            str1 = Mid(sHead, 195, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' тепловая энергия W2
            str1 = Mid(sHead, 204, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' время отработки 1
            str1 = Mid(sHead, 117, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' время отказа 1
            dblX = Val(str1): dblX = 24 - dblX
            strRez = strRez & Trim(str(dblX)): strRez = strRez & ";"
            ' время отработки 2
            str1 = Mid(sHead, 222, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' время отказа 2
            dblX = Val(str1): dblX = 24 - dblX
            strRez = strRez & Trim(str(dblX))
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez, 2)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' функция расчета часового расхода пара
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
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead ' брать построчно
            ' анализ взятой строки = дата
            str1 = Mid(sHead, 1, 10)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' анализ взятой строки = время
            str1 = Mid(sHead, 11, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' температура Т1
            str1 = Mid(sHead, 24, 16)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' давление Р1
            str1 = Mid(sHead, 40, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' масса М1
            str1 = Mid(sHead, 60, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' тепловая энергия W1
            str1 = Mid(sHead, 69, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' температура Т2
            str1 = Mid(sHead, 101, 14)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' давление Р2
            str1 = Mid(sHead, 115, 20)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' масса М2
            str1 = Mid(sHead, 135, 9)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' тепловая энергия W2
            str1 = Mid(sHead, 144, 18)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            strRez = strRez & "-": strRez = strRez & ";" ' время отработки 1
            strRez = strRez & "-": strRez = strRez & ";" ' время отказа 1
            strRez = strRez & "-": strRez = strRez & ";" ' время отработки 2
            strRez = strRez & "-" ' время отказа 2
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez, 2)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' функция обработки отчета ТСРВ
Function RepTSRV(ThisFile As String, TipArh As Long) As Boolean
Dim sHead As String, FileNumber As Integer
Dim strRez As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim dblT As Double, dblT1 As Double, i As Integer
Dim pos As Long, pos1 As Long
Dim H As Integer, strV(12) As String
Dim dblt2 As Double, delta As Double, a As Double, b As Double ' для контроля темп-ры
'
On Error Resume Next
        a = KoefA: b = KoefB
        If TipArh = 41 Then H = 60 Else H = 1
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead ' брать построчно
            ' анализ взятой строки = дата-время
            pos = InStr(1, sHead, ";", 1) ' ищем первое вхождение
            str1 = Mid(sHead, 1, pos - 1)
            strV(0) = Trim(str1)
            ' тепло по прямому
            pos = InStr(1, sHead, ";", 1) ' ищем первое вхождение
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' искать предыдущий
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(1) = Format(str1, "0.00")
            dblW = CDbl(str1)
            ' тепло по обратке
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(2) = Format(str1, "0.00")
            dblX = CDbl(str1)
            ' вычисляем потребленную энергию' режим лето/зима
            If Me.swtchReg.Value Then dblW = dblW + dblX Else dblW = dblW - dblX
            ' масса по прямому
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(3) = Format(str1, "0.00")
            dblm = CDbl(str1)
            ' масса по обратке
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(4) = Format(str1, "0.00")
            dblX = CDbl(str1)
            ' вычисляем потребленную массу' режим лето/зима
            If Me.swtchReg.Value Then dblm = dblm + dblX Else dblm = dblm - dblX
            ' время останова
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 18
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT = H * CDbl(str1)
            str1 = Mid(sHead, pos + 1)
            dblT1 = H * CDbl(str1)
            strV(5) = Format(dblT, "0.00") & "-" & Format(dblT1, "0.00")
            ' давление по прямому
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 12
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(6) = Format(str1, "0.00")
            ' давление по обратке
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 13
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(7) = Format(str1, "0.00")
            ' запись потребленных значений
            strV(8) = Format(dblW, "0.00")
            strV(9) = Format(dblm, "0.00")
            ' здесь запись времени отработки
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 16
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT = H * CDbl(str1)
            pos1 = pos
            pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            dblT1 = H * CDbl(str1)
            strV(10) = Format(dblT, "0.00") & "-" & Format(dblT1, "0.00")
            ' температура по прямому
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 7
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            ' проверить есть ли контроль тем-ры
            strV(11) = Format(str1, "0.00")
            ' температура по обратке
            pos = InStr(1, sHead, ";", 1)  ' искать следующий
            For i = 1 To 9
                pos1 = pos
                pos = InStr(pos + 1, sHead, ";", 1) ' искать следующий
            Next
            str1 = Mid(sHead, pos1 + 1, pos - pos1 - 1)
            strV(12) = Format(str1, "0.00")
            ' проверить есть ли контроль тем-ры
            dblt2 = CDbl(strV(11)) * a + b:  delta = CDbl(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & _
                    Format(delta, "0.00")
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
End Function
'
' функция чтения суточного архива воды
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
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead ' брать построчно
            ' анализ взятой строки = дата-время
            str1 = Mid(sHead, 1, 13)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' объем потребленный
            str1 = Mid(sHead, 62, 24)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' время отказа
            str1 = Mid(sHead, 86, 32)
            strRez = strRez & Trim(str1) & " мм": strRez = strRez & ";" ' формирование выходной строки
            ' время работы
            lX = 24 - Val(Mid(str1, 1, InStr(1, str1, "чч", vbTextCompare) - 1))
            lM = 60 - Val(Mid(str1, InStr(1, str1, "чч", vbTextCompare) + 2))
            ' формирование выходной строки
            strRez = strRez & Trim(str(lX)) & " ч " & Trim(str(lM)) & " м"
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez, 1)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' функция чтения часового архива воды
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
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead ' брать построчно
            ' анализ взятой строки = дата-время
            str1 = Mid(sHead, 1, 15)
            strRez = strRez & Trim(str1): strRez = strRez & ";"
            ' объем потребленный
            str1 = Mid(sHead, 64, 24)
            strRez = strRez & Trim(str1): strRez = strRez & ";"  ' формирование выходной строки
            ' время отказа
            str1 = Mid(sHead, 88, 34)
            strRez = strRez & Trim(str1): strRez = strRez & ";" ' формирование выходной строки
            ' время работы
            lX = 3600 - Val(str1)
            strRez = strRez & Trim(str(lX))  ' формирование выходной строки
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez, 1)
            strRez = ""
        Loop
        Close #FileNumber
End Function
'
' функция чтения часового архива VISIKAL
Function ReadHourVis(ThisFile As String) As Boolean
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim strRez As String
Dim strV(12) As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim lngT As Long, bTipArh As Boolean
' для контроля темп-ры
Dim dblt2 As Double, delta As Double, str3 As String, a As Double, b As Double
'
On Error GoTo errOPA
        a = KoefA: b = KoefB: bTipArh = False
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        If InStr(1, sHead, "В-0", 1) > 0 Then bTipArh = True ' В-0=лето
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead1 ' брать построчно
            ' анализ взятой строки = дата-время
            str1 = Mid(sHead1, 1, 15)
            If InStr(1, str1, ":") > 0 Then str1 = Left(str1, InStr(1, str1, ":") - 1) & _
                    " " & Mid(str1, InStr(1, str1, ":") + 1) & ":00" Else str1 = str1 & " 00:00"
            strV(0) = Trim(str1)
            ' тепло по прямому
            str1 = Mid(sHead1, 16, 16)
            str2 = Mid(sHead, 16, 16)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            strV(1) = Format(dblX, "0.00")
            ' тепло по обратке
            str1 = Mid(sHead1, 32, 16)
            str2 = Mid(sHead, 32, 16)
            dblW = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            ' вычисляем потребленную энергию ' режим лето/зима
            If bTipArh Then dblW = dblW + dblX Else dblW = dblW - dblX
            strV(2) = Format(dblX, "0.00")
            ' масса по прямому
            str1 = Mid(sHead1, 48, 13)
            str2 = Mid(sHead, 48, 13)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            strV(3) = Format(dblX, "0.00")
            ' масса по обратке
            str1 = Mid(sHead1, 61, 13)
            str2 = Mid(sHead, 61, 13)
            dblm = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            ' вычисляем потребленную массу ' режим лето/зима
            If bTipArh Then dblm = dblm + dblX Else dblm = dblm - dblX
            strV(4) = Format(dblX, "0.00")
            ' время останова
            str1 = Mid(sHead1, 152, 14)
            str2 = Mid(sHead, 152, 14)
            lngT = CLng(str1) - CLng(str2)
            strV(5) = Format(lngT, "0.00")
            ' давление по прямому
            str1 = Mid(sHead1, 166, 13):  strV(6) = Format(Val(str1), "0.00")
            ' давление по обратке
            str1 = Mid(sHead1, 179, 14):  strV(7) = Format(Val(str1), "0.00")
            strV(8) = Format(dblW, "0.00"):   strV(9) = Format(dblm, "0.00")
            ' здесь часовой архив
            strV(10) = Format((60 - lngT), "0.00")
            ' температура по прямому
            str1 = Mid(sHead1, 100, 9):   strV(11) = Format(Val(str1), "0.00")
            ' температура по обратке
            str1 = Mid(sHead1, 109, 9):   strV(12) = Format(Val(str1), "0.00")
            ' контроль тем-ры
            dblt2 = Val(strV(11)) * a + b:   delta = Val(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & _
                    Format(delta, "0.00")
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
        Exit Function
errOPA:
    Resume Next
End Function
'
' функция чтения суточного архива VISIKAL
Function ReadDayVis(ThisFile As String) As Boolean
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim strRez As String
Dim strV(12) As String, str1 As String, str2 As String
Dim dblX As Double, dblW As Double, dblm As Double
Dim lngT As Long, bTipArh As Boolean
' для контроля темп-ры
Dim dblt2 As Double, delta As Double, str3 As String, a As Double, b As Double
'
On Error GoTo errOPAd
        a = KoefA: b = KoefB: bTipArh = False
        FileNumber = FreeFile
        Open ThisFile For Input As #FileNumber
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        If InStr(1, sHead, "В-0", 1) > 0 Then bTipArh = True
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Line Input #FileNumber, sHead
        Do While Not EOF(FileNumber)   ' пока не конец файла
            Line Input #FileNumber, sHead1 ' брать построчно
            ' анализ взятой строки = дата-время
            str1 = Mid(sHead1, 1, 12)
            If InStr(1, str1, ":") > 0 Then str1 = Left(str1, InStr(1, str1, ":") - 1) & _
                " " & Mid(str1, InStr(1, str1, ":") + 1) & ":00" Else str1 = str1 & " 00:00"
            strV(0) = Trim(str1)
            ' тепло по прямому
            str1 = Mid(sHead1, 13, 16)
            str2 = Mid(sHead, 13, 16)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            strV(1) = Format(dblX, "0.00")
            ' тепло по обратке
            str1 = Mid(sHead1, 29, 16)
            str2 = Mid(sHead, 29, 16)
            dblW = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Wn - Val(str2), Val(str1) - Val(str2))
            ' вычисляем потребленную энергию' режим лето/зима
            If bTipArh Then dblW = dblW + dblX Else dblW = dblW - dblX
            strV(2) = Format(dblX, "0.00")
            ' масса по прямому
            str1 = Mid(sHead1, 45, 13)
            str2 = Mid(sHead, 45, 13)
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            strV(3) = Format(dblX, "0.00")
            ' масса по обратке
            str1 = Mid(sHead1, 58, 13)
            str2 = Mid(sHead, 58, 13)
            dblm = dblX
            dblX = IIf(Val(str1) < Val(str2), Val(str1) + Vn - Val(str2), Val(str1) - Val(str2))
            ' вычисляем потребленную массу' режим лето/зима
            If bTipArh Then dblm = dblm + dblX Else dblm = dblm - dblX
            strV(4) = Format(dblX, "0.00")
            ' время останова
            str1 = Mid(sHead1, 133, 14)
            str2 = Mid(sHead, 133, 14)
            lngT = CLng(str1) - CLng(str2)
            strV(5) = Format(lngT / 60, "0.00")
            ' давление по прямому
            str1 = Mid(sHead1, 147, 13):     strV(6) = Format(Val(str1), "0.00")
            ' давление по обратке
            str1 = Mid(sHead1, 160, 14):     strV(7) = Format(Val(str1), "0.00")
            strV(8) = Format(dblW, "0.00"):  strV(9) = Format(dblm, "0.00")
            '  суточный  архив
            str(10) = Format((1440 - lngT) / 60, "0.00")
            ' температура по прямому
            str1 = Mid(sHead1, 97, 9):  strV(11) = Format(Val(str1), "0.00")
            ' температура по обратке
            str1 = Mid(sHead1, 106, 9): strV(12) = Format(Val(str1), "0.00")
            ' проверить есть ли контроль тем-ры
            dblt2 = Val(strV(11)) * a + b:   delta = Val(strV(12)) - dblt2
            strRez = strV(0) & ";" & strV(1) & ";" & strV(2) & ";" & strV(3) & ";" & strV(4) & ";" & _
                    strV(5) & ";" & strV(8) & ";" & strV(9) & ";" & strV(10) & ";" & strV(11) & ";" & _
                    strV(12) & ";;" & strV(6) & ";" & strV(7) & ";;" & Format(dblt2, "0.00") & ";" & Format(delta, "0.00")
            ' предварительная подготовка строки к записи (меняем [,] на [.])
            strRez = RemakeS(strRez)
           ' запись в итоговый файл
            Call MakeTempbase(strRez)
            sHead = sHead1: strRez = ""
        Loop
        Close #FileNumber
        Exit Function
errOPAd:
    Resume Next
End Function
'
' функция формирования итогового файла
Function MakeTempbase(strExit As String, Optional xHead As Integer) As Integer
Dim FileNumber As Integer
On Error Resume Next
If IsNull(xHead) Then xHead = 0
    FileNumber = FreeFile
    ChDir SavePath ' вернуться в текущую папку
    Open "tempgraf.csv" For Append As #FileNumber
    If LOF(FileNumber) = 0 Then
        If xHead = 0 Then
                Print #FileNumber, "ДатаВремя;W1;W2;m1;m2;TНС;Формула1;Формула2;Формула3;Формула4;" & _
                                                    "Формула5;Формула6;Формула7;Формула8;Формула9;t2r;dt"
        ElseIf xHead = 1 Then
            Print #FileNumber, "DateTime;W1;Формула1;Формула2"
        ElseIf xHead = 2 Then
            Print #FileNumber, "Date;Time;t1;P1;M1;W1;t2;P2;M2;W2;Tr1;To1;Tr2;To2"
       End If
    End If
    Print #FileNumber, strExit ' писать построчно
    Close #FileNumber
    ChDir NewPath
    MakeTempbase = 1
End Function
'
' функция предварительного анализа файла
Function PreAnalizeFiles(strNameFile As String) As Long
Dim sHead As String, sHead1 As String, FileNumber As Integer
Dim iNum As Integer, i As Integer, ExitCode As String
Dim k As Integer
On Error Resume Next
    k = 0: ExitCode = 0
    FileNumber = FreeFile
    Open strNameFile For Input As #FileNumber
    For i = 1 To 6 ' повторять цикл для нескольких первых строк
        Line Input #FileNumber, sHead ' брать построчно
        'sHead = ToOEM(sHead)
        sHead1 = ToAnsi(sHead)
        ' искать ключевые слова
        ' ищем основные архивы по VISIKAL PRO
        iNum = InStr(1, sHead, Me.txtVzljot, vbTextCompare)
        If iNum > 0 Then ExitCode = "1"
        ' ищем архивы по воде
        iNum = InStr(1, sHead, Me.txtVoda, vbTextCompare)
        If iNum > 0 Then ExitCode = "2"
        ' ищем архивы по пару
        iNum = InStr(1, sHead1, Me.txtPar, vbTextCompare)
        If iNum > 0 Then ExitCode = "3"
        ' ищем архивы по текущим значениям
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
        ' ищем архивы по ПО Отчет-ТСРВ
        iNum = InStr(1, sHead, ";", vbTextCompare): k = 1
        iNum = InStr(iNum + 1, sHead, ";", vbTextCompare): k = 2
        iNum = InStr(iNum + 1, sHead, ";", vbTextCompare): k = 3
        If iNum > 0 And i = 1 And k = 3 Then
            Line Input #FileNumber, sHead ' брать построчно
            iNum = InStr(1, sHead, "23:00", vbTextCompare)
            Line Input #FileNumber, sHead ' брать построчно
            k = InStr(1, sHead, "23:00", vbTextCompare)
            Close #FileNumber
            If iNum > 0 And k > 0 Then ExitCode = "42" Else ExitCode = "41"
            PreAnalizeFiles = CLng(ExitCode)
            Exit Function
        End If
        iNum = InStr(1, sHead & sHead1, "Час", vbTextCompare)
        If iNum > 0 Then ExitCode = ExitCode & "1"
        iNum = InStr(1, sHead & sHead1, "Сут", vbTextCompare)
        If iNum > 0 Then ExitCode = ExitCode & "2"
    Next
    Close #FileNumber
PreAnalizeFiles = CLng(ExitCode)
End Function
'
Private Sub Form_Load()
On Error Resume Next
Me.Dir1.Path = CurDir ' начать с текущей папки
Me.Label1.Visible = False
End Sub
' передает текущие полученные данные для построения графика
Function OpenCurFile(stFileName As String)
' получить данные для работы с архивом
Dim filenum ' имя файла для сохранения данных
Dim strX As String, strY As String
Dim priznak As String
On Error GoTo errPol
' открыть файл
If OpenCSV(stFileName) Then
    filenum = FreeFile ' дескриптор cвободного файла
    Open "tempbase.csv" For Input As #filenum ' открыть файл
    ' пропустить первые 3-и строки
    Line Input #filenum, strX '
    Line Input #filenum, strX '
    Line Input #filenum, strX '
    ' остальные строки записать как есть
    Do While Not EOF(filenum)
        Line Input #filenum, strX '
        strX = RemakeS(strX) ' меняем [,] на [.]
        strX = strX & ";;"
        Call MakeTempbase(strX)
    Loop
    Close #filenum ' закрыть файл
Else
    MsgBox "Файл не загружен"
End If
Exit Function
errPol:
Resume Next
End Function
' конверт кодировки текста из DOS в WIN
Function ToAnsi(S As String) As String
    Dim ss As String
    ss = S: OemToChar S, ss: ToAnsi = ss
End Function
' конверт кодировки текста из WIN в DOS
Function ToOEM(S As String) As String
    Dim ss As String
    ss = S: CharToOem S, ss: ToOEM = ss
End Function

