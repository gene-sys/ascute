Attribute VB_Name = "mdlExchgData"
' ������ ����������� ����� ������� �/� Access � MySQL
'
'strPath = ���� � �� � ������� ������������ �������
Public Function RemakeThisLink(strPath As String)
' ������������ ����� ADODC (��� ����� DAO) � spbase.mdb
Dim daoBase As DAO.Database
Dim tdfMySQL As DAO.TableDef
Dim arTables(9) As String
Dim strConnectString As String
Dim I As Integer ', pos As Long, pos1 As Long
'
On Error GoTo err_RemakeThisLink
If MsgBox("������������� ��������� ���������� ����� ������?", _
                        vbYesNo) = vbNo Then GoTo Exit_RemakeThisLink
Form1.stat.Caption = "�����..."
arTables(0) = "Par_h"
arTables(1) = "Par_s"
arTables(2) = "teplo_hn"
arTables(3) = "teplo_hr"
arTables(4) = "teplo_sn"
arTables(5) = "teplo_sr"
arTables(6) = "teplo_tek"
arTables(7) = "voda_h"
arTables(8) = "voda_s"
Set daoBase = OpenDatabase(strPath)
' ����� ����. �������������� ����� ODBC
' ������� � ������
For I = 0 To 8 '
    strConnectString = daoBase.TableDefs(arTables(I)).Connect ' ����� ���.�����
    If (Len(strConnectString & "") <> 0) Then
        ' ������� ����� �������������
        daoBase.TableDefs(arTables(I)).Connect = ";FileDSN=" & App.Path & "\first.dsn"
        'daoBase.TableDefs(arTables(I)).Connect = _
            "ODBC;DRIVER=MySQL ODBC 3.51 Driver;UID = root;PWD = 111;Charset = utf8;SERVER=" & _
            Form1.txtIP.Text & ";Port = 3306;OPTION=0;Database = askute"
        daoBase.TableDefs(arTables(I)).RefreshLink ' �������� �����
    End If
Next I
Set tdfMySQL = Nothing
daoBase.Close
Set daoBase = Nothing
Form1.stat.Caption = "���������� ���������"
Exit_RemakeThisLink:
      Exit Function
err_RemakeThisLink:
      Form1.stat.Caption = "������ ��� ���������� ����� " & Err.Description
      Resume Exit_RemakeThisLink
End Function
'
' ���������� ��������: strPath = ���� � �� � ������� ������������ �������
Public Function Performance(strPath As String)
' ������������ ����� ADODC (��� ����� DAO) � spbase.mdb
Dim daoBase As DAO.Database
Dim qryMySQL As DAO.QueryDef
Dim tdfMySQL As DAO.TableDef
Dim strZ As String, I As Long
'
'On Error Resume Next
On Error GoTo err_Performance
Form1.stat.Caption = "�����..."
Set daoBase = OpenDatabase(strPath)
For Each tdfMySQL In daoBase.TableDefs
    If InStr(1, tdfMySQL.Name, "��������", 1) > 0 Or _
            InStr(1, tdfMySQL.Name, "�������", 1) > 0 Then
        ' ����� ��������� �� ini-�����
        strZ = String$(255, " ")
        GetPrivateProfileString tdfMySQL.Name, "Update", _
                    "x", strZ, 255, App.Path & "/pattern.ini"
        ' ������� ������� ����� ������
        strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
        If Trim(strZ) <> "x" Then daoBase.Execute Trim(strZ)
    End If
Next
daoBase.Close
Set daoBase = Nothing
Form1.stat.Caption = "������ ��������"
Exit Function
err_Performance:
    'MsgBox Err.Number & ":" & Err.Description
    Resume Next
End Function
'
' ���������� ������� ������
Public Function PerformCur(strPath As String)
Dim daoBase As DAO.Database
Dim qryMySQL As DAO.QueryDef
Dim strZ As String, I As Long, kon As Long
Dim strX As String
'
On Error Resume Next
'On Error GoTo err_PerformCut
Form1.stat.Caption = "�����..."
Set daoBase = OpenDatabase(strPath)
strZ = String$(255, " ")
GetPrivateProfileString "�������", "�����", _
            "x", strZ, 255, App.Path & "/pattern.ini"
strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
kon = CLng(Trim(strZ))
For I = 1 To kon
    strZ = String$(255, " "): strX = Trim(str(I))
    GetPrivateProfileString "�������", strX, _
                "x", strZ, 255, App.Path & "/pattern.ini"
    strZ = Mid(strZ, 1, InStr(1, strZ, Chr(0)) - 1)
    If Trim(strZ) <> "x" Then daoBase.Execute Trim(strZ)
Next
daoBase.Close
Set daoBase = Nothing
Form1.stat.Caption = "������� ������ ��������"
'err_PerformCut:
'    MsgBox Err.Number & ":" & Err.Description
End Function
