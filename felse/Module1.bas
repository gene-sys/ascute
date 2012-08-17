Attribute VB_Name = "Module1"
Public Const HKEY_CLASSES_ROOT = &H80000000
'
Declare Function OemToChar Lib "user32" Alias "OemToCharA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, lpReserved As Long, lptype As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)
'��������� API ��� ������
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' ��������� API ��� ������
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const ERROR_SUCCESS = 0
'
' ��������� �������� � ������� ����������� �����������
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
'
'Public Const LOCALE_SYSTEM_DEFAULT = &H800
'Public Const LOCALE_USER_DEFAULT = &H400 ' ������ �� ���������
Public Const LANG_RUSSIAN = &H19 '  ���� �� ���������
'Public Const LOCALE_SDECIMAL = &HE         '  ���������� �����������
Public Const LOCALE_ILANGUAGE = &H1 ' ID �����
Public Const LOCALE_SLANGUAGE = &H2 ' �������������� �������� �����
Public Const LOCALE_SENGLANGUAGE = &H1001 ' ���������� �������� �����
Public Const LOCALE_SABBREVLANGNAME = &H3 ' ������������ �����
Public Const LOCALE_SNATIVELANGNAME = &H4 ' ������ �������� �����
Public Const LOCALE_ICOUNTRY = &H5 ' ��� ������
Public Const LOCALE_SCOUNTRY = &H6 ' �������������� �������� ������
Public Const LOCALE_SENGCOUNTRY = &H1002 ' ���������� �������� ������
Public Const LOCALE_SABBREVCTRYNAME = &H7 ' ������������ �������� ������
Public Const LOCALE_SNATIVECTRYNAME = &H8 ' ������ �������� ������
Public Const LOCALE_IDEFAULTLANGUAGE = &H9 ' ID ����� �� ���������
Public Const LOCALE_IDEFAULTCOUNTRY = &HA ' ��� ������ �� ���������
Public Const LOCALE_IDEFAULTCODEPAGE = &HB ' ������� �������� �� ���������
Public Const LOCALE_SLIST = &HC ' list item separator
Public Const LOCALE_IMEASURE = &HD ' 0 = metric, 1 = US
Public Const LOCALE_SDECIMAL = &HE ' ����������� ���������� ��������
Public Const LOCALE_STHOUSAND = &HF ' ����������� �����
Public Const LOCALE_SGROUPING = &H10 ' digit grouping
Public Const LOCALE_IDIGITS = &H11 ' number of fractional digits
Public Const LOCALE_ILZERO = &H12 ' leading zeros For decimal
Public Const LOCALE_SNATIVEDIGITS = &H13 ' native ascii 0-9
Public Const LOCALE_SCURRENCY = &H14 ' local monetary symbol
Public Const LOCALE_SINTLSYMBOL = &H15 ' intl monetary symbol
Public Const LOCALE_SMONDECIMALSEP = &H16 ' monetary decimal separator
Public Const LOCALE_SMONTHOUSANDSEP = &H17 ' monetary thousand separator
Public Const LOCALE_SMONGROUPING = &H18 ' monetary grouping
Public Const LOCALE_ICURRDIGITS = &H19 ' # local monetary digits
Public Const LOCALE_IINTLCURRDIGITS = &H1A ' # intl monetary digits
Public Const LOCALE_ICURRENCY = &H1B ' positive currency mode
Public Const LOCALE_INEGCURR = &H1C ' negative currency mode
Public Const LOCALE_SDATE = &H1D ' date separator
Public Const LOCALE_STIME = &H1E ' time separator
Public Const LOCALE_SSHORTDATE = &H1F ' short date format String
Public Const LOCALE_SLONGDATE = &H20 ' long date format String
Public Const LOCALE_STIMEFORMAT = &H1003 ' time format String
Public Const LOCALE_IDATE = &H21 ' short date format ordering
Public Const LOCALE_ILDATE = &H22 ' Long date format ordering
Public Const LOCALE_ITIME = &H23 ' time format specifier
Public Const LOCALE_ICENTURY = &H24 ' century format specifier
Public Const LOCALE_ITLZERO = &H25 ' leading zeros in time field
Public Const LOCALE_IDAYLZERO = &H26 ' leading zeros in day field
Public Const LOCALE_IMONLZERO = &H27 ' leading zeros in month field
Public Const LOCALE_S1159 = &H28 ' AM designator
Public Const LOCALE_S2359 = &H29 ' PM designator
Public Const LOCALE_SDAYNAME1 = &H2A ' Long name For Monday
Public Const LOCALE_SDAYNAME2 = &H2B ' Long name For Tuesday
Public Const LOCALE_SDAYNAME3 = &H2C ' Long name For Wednesday
Public Const LOCALE_SDAYNAME4 = &H2D ' Long name For Thursday
Public Const LOCALE_SDAYNAME5 = &H2E ' Long name For Friday
Public Const LOCALE_SDAYNAME6 = &H2F ' Long name For Saturday
Public Const LOCALE_SDAYNAME7 = &H30 ' Long name For Sunday
Public Const LOCALE_SABBREVDAYNAME1 = &H31 ' abbreviated name For Monday
Public Const LOCALE_SABBREVDAYNAME2 = &H32 ' abbreviated name For Tuesday
Public Const LOCALE_SABBREVDAYNAME3 = &H33 ' abbreviated name For Wednesday
Public Const LOCALE_SABBREVDAYNAME4 = &H34 ' abbreviated name For Thursday
Public Const LOCALE_SABBREVDAYNAME5 = &H35 ' abbreviated name For Friday
Public Const LOCALE_SABBREVDAYNAME6 = &H36 ' abbreviated name For Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37 ' abbreviated name For Sunday
Public Const LOCALE_SMONTHNAME1 = &H38 ' Long name For January
Public Const LOCALE_SMONTHNAME2 = &H39 ' Long name For February
Public Const LOCALE_SMONTHNAME3 = &H3A ' Long name For March
Public Const LOCALE_SMONTHNAME4 = &H3B ' Long name For April
Public Const LOCALE_SMONTHNAME5 = &H3C ' Long name For May
Public Const LOCALE_SMONTHNAME6 = &H3D ' Long name For June
Public Const LOCALE_SMONTHNAME7 = &H3E ' Long name For July
Public Const LOCALE_SMONTHNAME8 = &H3F ' Long name For August
Public Const LOCALE_SMONTHNAME9 = &H40 ' Long name For September
Public Const LOCALE_SMONTHNAME10 = &H41 ' Long name For October
Public Const LOCALE_SMONTHNAME11 = &H42 ' Long name For November
Public Const LOCALE_SMONTHNAME12 = &H43 ' Long name For December
Public Const LOCALE_SABBREVMONTHNAME1 = &H44 ' abbreviated name For January
Public Const LOCALE_SABBREVMONTHNAME2 = &H45 ' abbreviated name For February
Public Const LOCALE_SABBREVMONTHNAME3 = &H46 ' abbreviated name For March
Public Const LOCALE_SABBREVMONTHNAME4 = &H47 ' abbreviated name For April
Public Const LOCALE_SABBREVMONTHNAME5 = &H48 ' abbreviated name For May
Public Const LOCALE_SABBREVMONTHNAME6 = &H49 ' abbreviated name For June
Public Const LOCALE_SABBREVMONTHNAME7 = &H4A ' abbreviated name For July
Public Const LOCALE_SABBREVMONTHNAME8 = &H4B ' abbreviated name For August
Public Const LOCALE_SABBREVMONTHNAME9 = &H4C ' abbreviated name For September
Public Const LOCALE_SABBREVMONTHNAME10 = &H4D ' abbreviated name For October
Public Const LOCALE_SABBREVMONTHNAME11 = &H4E ' abbreviated name For November
Public Const LOCALE_SABBREVMONTHNAME12 = &H4F ' abbreviated name For December
Public Const LOCALE_SABBREVMONTHNAME13 = &H100F

Public Const LOCALE_SYSTEM_DEFAULT& = &H800
Public Const LOCALE_USER_DEFAULT& = &H400

Const cMAXLEN = 255

'Private Declare Function apiGetLocaleInfo Lib "kernel32" 'Alias "GetLocaleInfoA" (ByVal Locale As Long, 'ByVal LCType As Long, ByVal lpLCData As String, 'ByVal cchData As Long) As Long

Private Declare Function apiSetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long

'Function GetLocaleInfo(lngLCType As Long) As String
'Dim lngLocale As Long
'Dim strLCData As String, lngData As Long
'Dim lngX As Long
'
'strLCData = String$(cMAXLEN, 0)
'lngData = cMAXLEN - 1
'lngX = apiGetLocaleInfo(LOCALE_USER_DEFAULT, lngLCType, 'strLCData, lngData)
'If lngX <> 0 Then
'GetLocaleInfo = Left$(strLCData, lngX - 1)
'End If
'End Function
'
'������ �������������:Call SetLocaleInfo(LOCALE_SDECIMAL, ".") - ������������� ����������� �������� - �����.'
Public Function SetLocaleInfo(lngLCType As Long, lValue) As String
On Error Resume Next
Dim lngLocale As Long
Dim strLCData As String
Dim lngX As Long

strLCData = String$(cMAXLEN, 0)
strLCData = CStr(lValue) & String(cMAXLEN - Len(CStr(lValue)), 0)
lngX = apiSetLocaleInfo(LOCALE_USER_DEFAULT, lngLCType, strLCData)
If lngX <> 0 Then
SetLocaleInfo = Left$(strLCData, lngX - 1)
End If
End Function
'
' ��� �������� ����� �������
Public Sub Encrypt()
    Dim sHead As String
    Dim sT As String
    Dim sA As String
    Dim cphX As New cipher
    Dim n As Long
    Open "Users" For Binary As #1
    'Load entire file into sA
    sA = Space$(LOF(1))
    Get #1, , sA
    Close #1
    ' ���������� � ���������� ����� ������� Hash
    sT = Hash(Date & str(Timer))
    sHead = "[Secret]" & sT & Hash(sT & "ThisIsPasswordForEncryption")
    ' ����������
    cphX.KeyString = sHead
    cphX.Text = sA
    cphX.DoXor
    cphX.Stretch
    sA = cphX.Text
    ' ������ �����������
    Open "Users" For Output As #1
    Print #1, sHead
    ' ���������
    n = 1
    Do
        Print #1, Mid(sA, n, 70)
        n = n + 70
    Loop Until n > Len(sA)
    Close #1
End Sub
' ��� ���������� ����� �������
Public Sub Decrypt()
    Dim sHead As String
    Dim sA As String
    Dim sT As String
    Dim cphX As New cipher
    Dim n As Long
    ' ��������� �� ������ [Secret] � ������ �����,
    ' ��� �� ���������� ����� �������
    Open "Users" For Input As #1
    Line Input #1, sHead
    Close #1
    ' ��������� ������
    sT = Mid(sHead, 9, 8)
    If InStr(sHead, Hash(sT & "ThisIsPasswordForEncryption")) <> 17 Then
        Beep
        Exit Sub
    End If
    ' ��������� ����
    Open "Users" For Input As #1
    Line Input #1, sHead
    Do Until EOF(1)
        Line Input #1, sT
        sA = sA & sT
    Loop
    Close #1
    ' ������������
    cphX.KeyString = sHead
    cphX.Text = sA
    cphX.Shrink
    cphX.DoXor
    sA = cphX.Text
    ' ������� ����
    Kill "Users"
    Open "Users" For Binary As #1
    Put #1, , sA
    Close #1
End Sub
' ������� �������� (�������) ��������
Public Function Hash(sA As String) As String
    Dim cphHash As New cipher
    cphHash.KeyString = sA & "123456" ' ����� �������� ���
    cphHash.Text = sA & "123456" ' ����� �����
    cphHash.DoXor ' �����
    cphHash.Stretch ' ��������
    cphHash.KeyString = cphHash.Text ' ��������� �����
    cphHash.Text = "123456" ' ������� ����� ���
    cphHash.DoXor ' �����
    cphHash.Stretch ' ��������
    Hash = cphHash.Text ' ������ ���������
End Function
' ��� �������� ������ ����� � ������� �����������
Public Sub EncFile()
    Dim sHead As String
    Dim sT As String
    Dim sA As String
    Dim cphX As New cipher
    Dim n As Long
    Dim filName As String
    filName = InputBox("������� ��� ����� ��� ����������:")
    Open filName For Binary As #1
    'Load entire file into sA
    sA = Space$(LOF(1))
    Get #1, , sA
    Close #1
    ' ���������� � ���������� ����� ������� Hash
    sT = Hash(Date & str(Timer))
    sHead = sT & Hash(sT & "ThisIsPasswordForEncryption")
    ' ����������
    cphX.KeyString = sHead
    cphX.Text = sA
    cphX.DoXor
    cphX.Stretch
    sA = cphX.Text
    ' ������ �����������
    Open filName For Output As #1
    Print #1, sHead
    ' ���������
    n = 1
    Do
        Print #1, Mid(sA, n, 70)
        n = n + 70
    Loop Until n > Len(sA)
    Close #1
End Sub
' ��� ���������� ������ ����� � ������� �����������
Public Sub DecFile()
    Dim sHead As String
    Dim sA As String
    Dim sT As String
    Dim cphX As New cipher
    Dim n As Long
    Dim filName As String
    filName = InputBox("������� ��� ����� ��� ����������:")
    Open filName For Input As #1
    Line Input #1, sHead
    Close #1
    ' ��������� ������
    sT = Mid(sHead, 1, 8)
    If InStr(sHead, Hash(sT & "ThisIsPasswordForEncryption")) <> 17 Then
    Beep
        Exit Sub
    End If
    ' ��������� ����
    Open filName For Input As #1
    Line Input #1, sHead
    Do Until EOF(1)
        Line Input #1, sT
        sA = sA & sT
    Loop
    Close #1
    ' ������������
    cphX.KeyString = sHead
    cphX.Text = sA
    cphX.Shrink
    cphX.DoXor
    sA = cphX.Text
    ' ������� ����
    Kill filName
    Open filName For Binary As #1
    Put #1, , sA
    Close #1
End Sub
'
Public Sub WriteParameters(NS As String, what As String)
Dim set0 As Boolean
On Error GoTo WP_err
set0 = False
    With DataEnvironment1.rsCommand2
        If .State <> adStateOpen Then
            .Open  ' ����������� ������
            set0 = True
        End If
        .MoveFirst
        .Find "NameSet = '" & NS & "'"
        .Fields("Set") = what
        .Update
        .Requery
        If set0 = True Then .Close
    End With
    Exit Sub
WP_err:
    MsgBox Err.Number & "->" & Err.Description
End Sub

Public Function ReadNParam(NameOfParam As String) As String
Dim strX As String
Dim setO As Boolean
On Error Resume Next
strX = "": setO = False
If DataEnvironment1.rsCommand2.State <> adStateOpen Then
    DataEnvironment1.rsCommand2.Open ', , adOpenDynamic, adLockOptimistic
    setO = True
End If
    With DataEnvironment1.rsCommand2
    .MoveFirst
    Do While Not .EOF
        If .Fields("NameSet") = NameOfParam Then
            strX = "" & .Fields("Set")
            Exit Do
        End If
        .MoveNext
    Loop
    End With
    If setO = True Then DataEnvironment1.rsCommand2.Close
    ReadNParam = strX
End Function
'
'������ � ������ [,] c [.]: napr=true => [.] �� [,]; napr=false => [,] �� [.]
Public Function RemakeS(str As String, Optional napr As Boolean) As String
Dim pos As Long, sX As String, sY As String
On Error Resume Next
If IsNull(napr) Or napr = False Then
    sX = "."
    sY = ","
Else
    sX = ","
    sY = "."
End If
pos = InStr(1, str, sY, 1) ' ���� ������ ���������
Do While pos
    str = Mid(str, 1, pos - 1) & sX & Mid(str, pos + 1) ' ���������
    pos = InStr(1, str, sY, 1) ' ������ ���������
Loop
RemakeS = str ' ������� ������������
End Function
'
' ������� ���������� ��� ����������� ���������� ����� (. ��� ,)
Public Function KindOfDecSep() As String
'
'
    Dim lBuffLen As Long
    Dim sBuffer As String
    Dim sDecimal As String
    Dim lResult As Long

    On Error GoTo vbErrorHandler

    lBuffLen = 128 ' ������ ������

    sBuffer = String$(lBuffLen, vbNullChar) ' ��� ����� � �����������
    ' ��������� ����������
    lResult = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBuffLen)
    sDecimal = Left$(sBuffer, lResult - 1) ' ��������� ���������� � ����������� ��������� �����

    KindOfDecSep = sDecimal ' �������� ���������� � ���������

    Exit Function

vbErrorHandler:
'
' Handle the errors here
'
End Function
'������� � ������������ � ������������ ������� ������ ����� ������
Public Function mon(x As Integer)
    mon = Switch(x = 1, "������", x = 2, "�������", x = 3, "����", x = 4, "������", _
    x = 5, "���", x = 6, "����", x = 7, "����", x = 8, "������", x = 9, "��������", _
    x = 10, "�������", x = 11, "������", x = 12, "�������")
End Function
' �������������� ���� ���� ��� ���������� �������
Public Function SQLDate(dData As Date, Optional tTime As Date) As String
If IsEmpty(tTime) Or IsNull(tTime) Or tTime = 0 Then
    SQLDate = "#" & Year(dData) & "/" & Month(dData) & "/" & Day(dData) & "#"
Else
    SQLDate = "#" & Year(dData) & "/" & Month(dData) & "/" & Day(dData) & " " & Format(tTime, "hh:mm") & "#"
End If
End Function
'
' �������������� ���� ���� ��� ���������� ������� MySQLDate
Public Function MySQLDate(dData As Date, Optional tTime As Date) As String
If IsEmpty(tTime) Or IsNull(tTime) Or tTime = 0 Then
    MySQLDate = Year(dData) & "-" & Month(dData) & "-" & Day(dData)
Else
    MySQLDate = Year(dData) & "-" & Month(dData) & "-" & Day(dData) & " " & Format(tTime, "hh:mm")
End If
End Function

'
' ������� �������� ����� ����� ������������� ������
Public Function OpenCSV(sFileName As String) As Boolean
Dim filenum ' ��� ����� ��� ���������� ������
Dim strX As String, strY As String
Dim pos As Long
On Error GoTo errOpenCSV
filenum = FreeFile ' ���������� c��������� �����
Open Mid(sFileName, 1, Len(sFileName) - 3) & "csv" For Input As #filenum
' ���������� ������ � ������ ������
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
Line Input #filenum, strX '
strY = strY & strX & vbCrLf
' �������������� ������ ���������
Line Input #filenum, strX '
pos = InStr(1, strX, "W3", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������1" & Mid(strX, pos + 2)
pos = InStr(1, strX, "m3", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������2" & Mid(strX, pos + 2)
pos = InStr(1, strX, "To", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������3" & Mid(strX, pos + 2)
pos = InStr(1, strX, "t1", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������4" & Mid(strX, pos + 2)
pos = InStr(1, strX, "t2", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������5" & Mid(strX, pos + 2)
pos = InStr(1, strX, "t3", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������6" & Mid(strX, pos + 2)
pos = InStr(1, strX, "P1", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������7" & Mid(strX, pos + 2)
pos = InStr(1, strX, "P2", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������8" & Mid(strX, pos + 2)
pos = InStr(1, strX, "T���", vbTextCompare)
If pos > 0 Then strX = Mid(strX, 1, pos - 1) & "�������9" & Mid(strX, pos + 4)
strY = strY & strX & vbCrLf
' ��������� ������ �������� ��� ����
Do While Not EOF(filenum)
   Line Input #filenum, strX '���������� ��������� �������� ������
   If InStr(1, strX, "�����", 1) = 0 Then strY = strY & strX & vbCrLf
Loop
Close #filenum ' ������� ����
ChDir App.Path
filenum = FreeFile ' ���������� c��������� �����
Open "tempbase.csv" For Output As #filenum ' ������� ����
Print #filenum, Mid(strY, 1, Len(strY) - 2) ' ��������� ����
Close #filenum
OpenCSV = True
Exit Function
errOpenCSV:
OpenCSV = False
End Function


