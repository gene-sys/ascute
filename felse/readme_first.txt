

Public Sub main()

Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
' Step 1
conn.Open "DSN=pubs;uid=sa;pwd=;database=pubs"
' Step 2
Set cmd.ActiveConnection = conn
cmd.CommandText = "SELECT * from authors"
' Step 3
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockBatchOptimistic
' Step 4
rs!au_lname.Properties("Optimize") = True
rs.Sort = "au_lname"
rs.Filter = "phone LIKE '415 5*'"
rs.MoveFirst
Do While Not rs.EOF
    Debug.Print "Name: " & rs!au_fname & " "; rs!au_lname & _
        "Phone: "; rs!phone & vbCr
    rs!phone = "777" & Mid(rs!phone, 5, 11)
    rs.MoveNext
Loop

' Step 5
conn.BeginTrans

'Step 6, part A
On Error GoTo ConflictHandler
rs.UpdateBatch
On Error GoTo 0

conn.CommitTrans

Exit Sub

'Step 6, part B
ConflictHandler:

rs.Filter = adFilterConflictingRecords
rs.MoveFirst
Do While Not rs.EOF
    Debug.Print "Conflict: Name: " & rs!au_fname; " " & rs!au_lname
    rs.MoveNext
Loop
conn.Rollback
Resume Next

End Sub
















IP адрес сервера Взлет СП - 168.254.197.130

Когда Вы программируете с ADO API и MyODBC Вы должны обратить внимание на некоторые заданные по умолчанию свойства, которые пока не поддержаны сервером MySQL. Например, использование свойства CursorLocation Property как adUseServer возвратит для реквизита RecordCount Property -1. Чтобы иметь правильное значение, Вы должны установить это свойство в adUseClient, как показано в этом коде на VB: 
Dim myconn As New ADODB.Connection
Dim myrs As New Recordset
Dim mySQL As String
Dim myrows As Long

myconn.Open "DSN=MyODBCsample"
mySQL = "SELECT * from user"
myrs.Source = mySQL
Set myrs.ActiveConnection = myconn
myrs.CursorLocation = adUseClient
myrs.Open
myrows = myrs.RecordCount

myrs.Close
myconn.Close

Visual Basic 
Чтобы модифицировать таблицу, Вы должны определить первичный ключ для таблицы. Visual Basic с ADO не может обрабатывать большие целые числа. Это означает, что некоторые запросы подобно SHOW PROCESSLIST не будут работать правильно. Установите опцию OPTION=16834 в строке подключения ODBC или задайте параметр Change BIGINT columns to INT на экране соединения MyODBC. Вы можете также устанавливать опцию Return matching rows. 

Следующий пример для ADO (ActiveX Data Objects) создает таблицу my_ado и показывает использование rs.addNew, rs.delete и rs.update. 

Private Sub myodbc_ado_Click()
  Dim conn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim sql As String

  'connect to MySQL server using MySQL ODBC 3.51 Driver
  Set conn = New ADODB.Connection
  conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};"_
                        & "SERVER=localhost;"_
                        & " DATABASE=test;"_
                        & "UID=venu;PWD=venu; OPTION=35"
  conn.Open
  'create table
  conn.Execute "DROP TABLE IF EXISTS my_ado"
  conn.Execute "CREATE TABLE my_ado(id int not null primary key,
                name varchar(20)," _
                & "txt text, dt date, tm time, ts timestamp)"
  'direct insert
  conn.Execute "INSERT INTO my_ado(id,name,txt) values(1,100,'venu')"
  conn.Execute "INSERT INTO my_ado(id,name,txt) values(2,200,'MySQL')"
  conn.Execute "INSERT INTO my_ado(id,name,txt) values(3,300,'Delete')"
  Set rs = New ADODB.Recordset
  rs.CursorLocation = adUseServer
  'fetch the initial table ..
  rs.Open "SELECT * FROM my_ado", conn
    Debug.Print rs.RecordCount
    rs.MoveFirst
    Debug.Print String(50, "-") & "Initial my_ado Result Set "
                       & String(50, "-")
    For Each fld In rs.Fields
      Debug.Print fld.Name,
      Next
      Debug.Print
      Do Until rs.EOF
      For Each fld In rs.Fields
      Debug.Print fld.Value,
      Next
      rs.MoveNext
      Debug.Print
    Loop
  rs.Close
  'rs insert
  rs.Open "select * from my_ado", conn, adOpenDynamic, adLockOptimistic
  rs.AddNew
  rs!Name = "Monty"
  rs!txt = "Insert row"
  rs.Update
  rs.Close
  'rs update
  rs.Open "SELECT * FROM my_ado"
  rs!Name = "update"
  rs!txt = "updated-row"
  rs.Update
  rs.Close
  'rs update second time..
  rs.Open "SELECT * FROM my_ado"
  rs!Name = "update"
  rs!txt = "updated-second-time"
  rs.Update
  rs.Close
  'rs delete
  rs.Open "SELECT * FROM my_ado"
  rs.MoveNext
  rs.MoveNext
  rs.Delete
  rs.Close
  'fetch the updated table ..
  rs.Open "SELECT * FROM my_ado", conn
    Debug.Print rs.RecordCount
    rs.MoveFirst
    Debug.Print String(50, "-") & "Updated my_ado Result Set "
                       & String(50, "-")
    For Each fld In rs.Fields
      Debug.Print fld.Name,
      Next
      Debug.Print
      Do Until rs.EOF
      For Each fld In rs.Fields
      Debug.Print fld.Value,
      Next
      rs.MoveNext
      Debug.Print
    Loop
  rs.Close
  conn.Close
End Sub
'
The following table shows some recommended option values for various configurations
Configuration 						Option Value 
--------------------------------------------------------------------
Microsoft Access, Visual Basic 				3 
Driver trace generation (Debug mode) 			4 
Microsoft Access (with improved DELETE queries) 	35 
Large tables with too many rows 			2049 
Sybase PowerBuilder 					135168 
Query log generation (Debug mode) 			524288 
Generate driver trace as well as query log (Debug mode) 524292 
Large tables with no-cache results 			3145731 

Another workaround is to use a SELECT COUNT(*) statement for a similar query to get the correct row count. 

To find the number of rows affected by a specific SQL statement in ADO, use the RecordsAffected property in the ADO execute method. For more information on the usage of execute method, refer to http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdmthcnnexecute.asp. 

For information, see ActiveX Data Objects(ADO) Frequently Asked Questions. 

в разделе [mysqld] в /etc/mysql/my.cnf
   character_set_server = utf8
   collation_server = utf8_general_ci
Как обеспечить корректную работу MySQL с русскими символами при сортировке и выборке данных.
  В /etc/my.cnf вписать в блоке [mysqld]:
   default-character-set=koi8_ru (или cp1251)
При работе с базой можно выставить рабочую кодировку через:
   SET CHARACTER SET koi8_ru

Данный код вызывает окно "Установка связи" из "Удаленный доступ к сети". Естественно, вы должны знать имя текущего соединения с интернетом. 

Private Sub Form_Load()
Result = Shell("rundll32.exe rnaui.DLL,RnaDial " & "connection_name", 1)
End Sub 
 

