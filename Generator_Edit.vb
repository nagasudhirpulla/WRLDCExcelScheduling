Private Function FieldExistsInRS( _
   ByRef rs As ADODB.Recordset, _
   ByVal fieldName As String)
   Dim fld As ADODB.Field
    
   fieldName = UCase(fieldName)
    
   For Each fld In rs.Fields
      If UCase(fld.Name) = fieldName Then
         FieldExistsInRS = True
         Exit Function
      End If
   Next
    
   FieldExistsInRS = False
End Function
Private Function ClearSheet(ByVal sheetName As String)
    Sheets(sheetName).Cells.ClearContents
End Function

Private Function WriteLineConsole(ByVal str As String, ByRef ConsoleLineTick As Integer)
    Sheets("CONSOLE").Cells(ConsoleLineTick, 1).Value = str
    ConsoleLineTick = ConsoleLineTick + 1
End Function

Sub fetch_generator_data()
''https://www.reddit.com/r/excel/comments/2xpht7/vba_connection_string_to_google_cloud_sql/
''http://stackoverflow.com/questions/26369937/excel-vba-mysql-select-from-table-not-full-informaton
Dim cn As ADODB.Connection
Dim Server_Name As String
Dim Database_Name As String
Dim User_ID As String
Dim Password As String
Dim SQLStr As String
Server_Name = "localhost" ' Enter your server name here
Database_Name = "wrldc_schedule" ' Enter your database name here
User_ID = "root" ' enter your user ID here
Password = "123" ' Enter your password here
''SQLStr = "SELECT * FROM sudhir" ' Enter your SQL here
Set cn = New ADODB.Connection
cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & Server_Name & ";Database=" & Database_Name & _
";Uid=" & User_ID & ";Pwd=" & Password & ";"

Dim ConsoleLineTick As Integer
ConsoleLineTick = 1

' Create a recordset object.
Dim rsMaterialsdb As ADODB.Recordset
Set rsMaterialsdb = New ADODB.Recordset
ClearSheet "GENERATORS"
ClearSheet "CONSOLE"
With rsMaterialsdb
' Assign the Connection object.
.ActiveConnection = cn
' Extract the required records.
.Open "SELECT sname, name, ramp, nunits, unitcapacity FROM generator ORDER BY generator.sname ASC"

Dim FieldColTick As Integer
FieldColTick = 1
For Each fld In rsMaterialsdb.Fields
      Sheets("GENERATORS").Cells(1, FieldColTick).Value = fld.Name
      FieldColTick = FieldColTick + 1 'tick iterator
Next
Sheets("GENERATORS").Range("A2").CopyFromRecordset rsMaterialsdb
WriteLineConsole "Fetched the generated data", ConsoleLineTick
' Tidy up
.Close
End With

''rsMaterialsdb.Close 'close recordset
cn.Close    'close connect to db
''MARCH 21 2015 23:05 MONDAY
''TODO DO THE DATA VALIDATION AT SERVER
''INSERT INTO `constituent` (`idConstituent`, `name`, `sname`, `updated_at`, `connectedto`, `comments`) VALUES (NULL, 'DADRA&NAGAR HAVELI', 'DNH', CURRENT_TIMESTAMP, '', '');
''SELECT * FROM `constituent` ORDER BY `constituent`.`sname` ASC

''INSERT INTO `generator` (`idGenerator`, `name`, `sname`, `ramp`, `nunits`, `unitcapacity`, `ownedby`, `connectedto`, `comments`, `updated_at`) VALUES (NULL, 'VINDHYANCHAL STAGE 1', 'VSTPS1', '15', '6', '210', NULL, NULL, NULL, CURRENT_TIMESTAMP);
''INSERT INTO `generator` (`idGenerator`, `name`, `sname`, `ramp`, `nunits`, `unitcapacity`, `ownedby`, `connectedto`, `comments`, `updated_at`) VALUES (NULL, 'VINDHYANCHAL STAGE 2', 'VSTPS2', '35', '2', '500', NULL, NULL, NULL, CURRENT_TIMESTAMP), (NULL, 'VINDHYANCHAL STAGE 3', 'VSTPS3', '35', '2', '500', NULL, NULL, NULL, CURRENT_TIMESTAMP);
''SELECT * FROM `generator` ORDER BY `generator`.`sname` ASC

End Sub

Sub update_generator_data()
''https://www.reddit.com/r/excel/comments/2xpht7/vba_connection_string_to_google_cloud_sql/
''http://stackoverflow.com/questions/26369937/excel-vba-mysql-select-from-table-not-full-informaton
''https://dev.mysql.com/doc/connector-odbc/en/connector-odbc-examples-programming-vb-ado.html
''Execute SQL Reference
''https://msdn.microsoft.com/en-us/library/ms675023(VS.85).aspx
''Delete all rows
''sql = "TRUNCATE TABLE generator"
''Get a single field
''SELECT * FROM generator ASC LIMIT 1
Dim cn As ADODB.Connection
Dim Server_Name As String
Dim Database_Name As String
Dim User_ID As String
Dim Password As String
Dim SQLStr As String
Server_Name = "localhost" ' Enter your server name here
Database_Name = "wrldc_schedule" ' Enter your database name here
User_ID = "root" ' enter your user ID here
Password = "123" ' Enter your password here
''SQLStr = "SELECT * FROM sudhir" ' Enter your SQL here
Set cn = New ADODB.Connection
cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & Server_Name & ";Database=" & Database_Name & _
";Uid=" & User_ID & ";Pwd=" & Password & ";"

' Create a recordset object.
Dim rsMaterialsdb As ADODB.Recordset
Set rsMaterialsdb = New ADODB.Recordset

With rsMaterialsdb
' Assign the Connection object.
.ActiveConnection = cn
' Extract the required records.
.Open "SELECT * FROM generator LIMIT 1"

ClearSheet "CONSOLE"
Dim ConsoleLineTick As Integer
ConsoleLineTick = 1
Dim FieldColTick As Integer
Dim canProceed As Boolean
canProceed = True
FieldColTick = 1

''TODO find if 1st column is a unique type of column
Do While Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value) <> "" And canProceed = True
      If FieldExistsInRS(rsMaterialsdb, Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value)) Then
      Else
        canProceed = False
        MsgBox ("The field " & Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value) & " doesnot exist in database")
      End If
      FieldColTick = FieldColTick + 1 'tick iterator
Loop
''Now FieldColTick is equal to number of fields
FieldColTick = FieldColTick - 1

' Tidy up
.Close
End With

''MsgBox FieldColTick
''check if all fields exist
If canProceed = False Then
    MsgBox "All Fields donot exist"
    Exit Sub
End If

''TODO UPDATE IF EXISTS AND CREATE IF ABSENT
''Construct initial sql prefix
Dim sqlInitial As String
sqlInitial = "UPDATE generator SET "
Dim sql As String
Dim num As Long
Dim adExecuteNoRecords As Long

''UPDATE generator SET name = 'CGPL1', unitcapacity = '831' WHERE generator.idGenerator = 13;
Dim colIterator As Integer
Dim rowIterator As Integer
Dim cellVal As String
Dim msg As String
Dim numRowsAffected
numRowsAffected = 0
msg = "updated the rows "
rowIterator = 2
Do While Trim(Sheets("GENERATORS").Cells(rowIterator, 1).Value) <> "" And rowIterator < 100
    sql = sqlInitial
    For colIterator = 2 To FieldColTick
        cellVal = Trim(Sheets("GENERATORS").Cells(rowIterator, colIterator).Value)
        If cellVal = "" Then
            cellVal = "NULL"
        Else
            cellVal = "'" + cellVal + "'"
        End If
        sql = sql & Trim(Sheets("GENERATORS").Cells(1, colIterator).Value) & " = " & cellVal
        If colIterator = FieldColTick Then
            sql = sql & " "
        Else
            sql = sql & ", "
        End If
    Next
    sql = sql & "WHERE " & Trim(Sheets("GENERATORS").Cells(1, 1).Value) & " = '" & Trim(Sheets("GENERATORS").Cells(rowIterator, 1).Value) & "'"
    WriteLineConsole sql, ConsoleLineTick
    cn.Execute sql, num, adExecuteNoRecords
    ''MsgBox (num)
    If num = 1 Then
    numRowsAffected = numRowsAffected + 1
    msg = msg & CStr(rowIterator) & ", "
    ElseIf num > 1 Then
    ''not updating a unique row
    End If
    ''MsgBox (adExecuteNoRecords)
    rowIterator = rowIterator + 1
Loop

If numRowsAffected > 0 Then
    WriteLineConsole "Updated " & numRowsAffected & " rows, " & msg & "of the GENERATORS sheet", ConsoleLineTick
Else
    WriteLineConsole "Updated zero rows", ConsoleLineTick
End If

End Sub

Sub create_or_update_generator_data()
    Dim cn As ADODB.Connection
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Server_Name = "localhost" ' Enter your server name here
    Database_Name = "wrldc_schedule" ' Enter your database name here
    User_ID = "root" ' enter your user ID here
    Password = "123" ' Enter your password here
    Set cn = New ADODB.Connection
    cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & Server_Name & ";Database=" & Database_Name & _
";Uid=" & User_ID & ";Pwd=" & Password & ";"

    Dim FieldColTick As Integer
    FieldColTick = 1
    ClearSheet "CONSOLE"
    Dim ConsoleLineTick As Integer
    ConsoleLineTick = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Check if first column is blank and return if blank
    If Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value) = "" Then
        MsgBox "First Column header is blank"
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''fetch all the fields of the generator table and verify with the excel feild names. //TODO Check if the first column name is unique type column
    Dim rsMaterialsdb As ADODB.Recordset
    Set rsMaterialsdb = New ADODB.Recordset
    With rsMaterialsdb
        ' Assign the Connection object.
        .ActiveConnection = cn
        ' Extract the required records.
        .Open ("SELECT * FROM generator LIMIT 1")
        Dim canProceed As Boolean
        canProceed = True
        

        Do While Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value) <> "" And canProceed = True
            If FieldExistsInRS(rsMaterialsdb, Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value)) Then
            Else
                canProceed = False
                MsgBox ("The field " & Trim(Sheets("GENERATORS").Cells(1, FieldColTick).Value) & " doesnot exist in database")
            End If
            FieldColTick = FieldColTick + 1 'tick iterator
        Loop
        ''Now FieldColTick is equal to number of fields
        FieldColTick = FieldColTick - 1

        .Close
    End With

    ''MsgBox FieldColTick
    ''check if all fields exist
    If canProceed = False Then
        MsgBox ("All Fields do not exist")
        Exit Sub
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''fetched all the fields of the generator

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''now for each value row starting from row 2, update a row if exists or create it
    Dim firstFieldName As String
    Dim rowIterator As Integer
    Dim colIterator As Integer
    Dim sql As String
    Dim num As Long
    Dim adExecuteNoRecords As Long

    firstFieldName = Trim(Sheets("GENERATORS").Cells(1, 1).Value)

    Dim cellVal As String
    Dim updateMsg As String
    Dim numRowsAffected As Integer
    numRowsAffected = 0
    updateMsg = "updated the rows "
    rowIterator = 2
    Do While Trim(Sheets("GENERATORS").Cells(rowIterator, 1).Value) <> "" And rowIterator < 100
        ''Entered the row starting from row number 2
        ''First find if row exists using the first field
        Set rsMaterialsdb = New ADODB.Recordset
        rsMaterialsdb.ActiveConnection = cn
        rsMaterialsdb.Open ("SELECT * FROM generator WHERE " + firstFieldName + " = '" + Trim(Sheets("GENERATORS").Cells(rowIterator, 1).Value) + "'")
        
        If (rsMaterialsdb.BOF And rsMaterialsdb.EOF) Then
            ''rsMaterialsdb.Close
            ''The Row does not exist so INSERT THE ROW
            sql = "INSERT INTO generator ("
            For colIterator = 1 To FieldColTick
                sql = sql & Trim(Sheets("GENERATORS").Cells(1, colIterator).Value)
                If colIterator = FieldColTick Then
                    sql = sql & ") "
                Else
                    sql = sql & ", "
                End If
            Next
            sql = sql & "values ("
            For colIterator = 1 To FieldColTick
                cellVal = Trim(Sheets("GENERATORS").Cells(rowIterator, colIterator).Value)
                If cellVal = "" Then
                    cellVal = "NULL"
                Else
                    cellVal = "'" & cellVal & "'"
                End If
                sql = sql & cellVal
                If colIterator = FieldColTick Then
                    sql = sql & ") "
                Else
                    sql = sql & ", "
                End If
            Next
        Else
            ''rsMaterialsdb.Close
            ''The Row exists already so UPPDATE THE ROW
            sql = "UPDATE generator SET "
            For colIterator = 2 To FieldColTick
                cellVal = Trim(Sheets("GENERATORS").Cells(rowIterator, colIterator).Value)
                If cellVal = "" Then
                    cellVal = "NULL"
                Else
                    cellVal = "'" & cellVal & "'"
                End If
                sql = sql & Trim(Sheets("GENERATORS").Cells(1, colIterator).Value) & " = " & cellVal
                If colIterator = FieldColTick Then
                    sql = sql & " "
                Else
                    sql = sql & ", "
                End If
            Next
            sql = sql & "WHERE " & firstFieldName & " = '" & Trim(Sheets("GENERATORS").Cells(rowIterator, 1).Value) & "'"
        End If
        ''MsgBox (sql)
        WriteLineConsole sql, ConsoleLineTick
        cn.Execute sql, num, adExecuteNoRecords
        ''MsgBox (num)
        If num = 1 Then
            numRowsAffected = numRowsAffected + 1
            updateMsg = updateMsg & CStr(rowIterator) & ", "
        ElseIf num > 1 Then
            ''not updating a unique row
        End If
        rowIterator = rowIterator + 1
    Loop

    If numRowsAffected > 0 Then
        WriteLineConsole "Updated " & numRowsAffected & " rows, " & updateMsg & "of the GENERATORS sheet", ConsoleLineTick
    Else
        WriteLineConsole "Updated zero rows", ConsoleLineTick
    End If

End Sub
