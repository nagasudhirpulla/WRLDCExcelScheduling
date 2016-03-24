''TODO DECLARE GLOBAL VARIABLES LIKE TABLENAME SERVERNAME PASSWORD TABLENAME
Sub fetch_constituent_data()
    ''https://www.reddit.com/r/excel/comments/2xpht7/vba_connection_string_to_google_cloud_sql/
    ''http://stackoverflow.com/questions/26369937/excel-vba-mysql-select-from-table-not-full-informaton
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Connect_To_DB cn
    
    ClearConsole
    WriteLineConsole "From constituent info fetch method..."

    Dim rsMaterialsdb As ADODB.Recordset
    Dim FieldColTick As Integer
    
    ' Create a recordset object.
    Set rsMaterialsdb = New ADODB.Recordset

    With rsMaterialsdb
    ' Assign the Connection object.
    .ActiveConnection = cn
    ' Extract the required records.
    .Open "SELECT sname, name FROM constituent ORDER BY constituent.sname ASC"
    
    FieldColTick = 1
    For Each fld In rsMaterialsdb.Fields
          Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value = fld.Name
          FieldColTick = FieldColTick + 1 'tick iterator
     Next
    Sheets("CONSTITUENTS").Range("A2").CopyFromRecordset rsMaterialsdb
    WriteLineConsole "Fetched the constituents data"
    ' Tidy up
    .Close
    End With
    cn.Close    'close connect to db
End Sub

Sub update_constituent_data()
    ''https://www.reddit.com/r/excel/comments/2xpht7/vba_connection_string_to_google_cloud_sql/
    ''http://stackoverflow.com/questions/26369937/excel-vba-mysql-select-from-table-not-full-informaton
    ''https://dev.mysql.com/doc/connector-odbc/en/connector-odbc-examples-programming-vb-ado.html
    ''Execute SQL Reference
    ''https://msdn.microsoft.com/en-us/library/ms675023(VS.85).aspx
    ''Delete all rows
    ''sql = "TRUNCATE TABLE constituent"
    ''Get a single field
    ''SELECT * FROM constituent ASC LIMIT 1
    Dim colIterator As Integer
    Dim rowIterator As Integer
    Dim cellVal As String
    Dim msg As String
    Dim numRowsAffected
    
    Dim rsMaterialsdb As ADODB.Recordset
    
    Dim FieldColTick As Integer
    Dim canProceed As Boolean
    
    Dim sql As String
    Dim num As Long
    Dim adExecuteNoRecords As Long
    
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Connect_To_DB cn
    
    ClearConsole
    WriteLineConsole "From constituent info update method..."
    
    ' Create a recordset object.
    Set rsMaterialsdb = New ADODB.Recordset
    With rsMaterialsdb
    ' Assign the Connection object.
    .ActiveConnection = cn
    ' Extract the required records.
    .Open "SELECT * FROM constituent LIMIT 1"
    
    canProceed = True
    FieldColTick = 1
    ''TODO find if 1st column is a unique type of column
    Do While Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value) <> "" And canProceed = True
          If FieldExistsInRS(rsMaterialsdb, Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value)) Then
          Else
            canProceed = False
            MsgBox ("The field " & Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value) & " doesnot exist in database")
          End If
          FieldColTick = FieldColTick + 1 'tick iterator
    Loop
    ''Now FieldColTick is equal to number of fields
    FieldColTick = FieldColTick - 1
    ' Tidy up
    .Close
    End With
    
    ''check if all fields exist
    If canProceed = False Then
        MsgBox "All Fields donot exist"
        Exit Sub
    End If
    
    ''Construct sql string
    numRowsAffected = 0
    msg = "updated the rows "
    rowIterator = 2
    Do While Trim(Sheets("CONSTITUENTS").Cells(rowIterator, 1).Value) <> "" And rowIterator < 100
        sql = "UPDATE constituent SET "
        For colIterator = 2 To FieldColTick
            cellVal = Trim(Sheets("CONSTITUENTS").Cells(rowIterator, colIterator).Value)
            If cellVal = "" Then
                cellVal = "NULL"
            Else
                cellVal = "'" + cellVal + "'"
            End If
            sql = sql & Trim(Sheets("CONSTITUENTS").Cells(1, colIterator).Value) & " = " & cellVal
            If colIterator = FieldColTick Then
                sql = sql & " "
            Else
                sql = sql & ", "
            End If
        Next
        sql = sql & "WHERE " & Trim(Sheets("CONSTITUENTS").Cells(1, 1).Value) & " = '" & Trim(Sheets("CONSTITUENTS").Cells(rowIterator, 1).Value) & "'"
        WriteLineConsole sql
        cn.Execute sql, num, adExecuteNoRecords
        If num = 1 Then
            numRowsAffected = numRowsAffected + 1
            msg = msg & CStr(rowIterator) & ", "
        ElseIf num > 1 Then
            WriteLineConsole "Beware: not updating a unique row at row " & CStr(rowIterator) & " of CONSTITUENTS sheet"
        End If
        rowIterator = rowIterator + 1
    Loop
    
    If numRowsAffected > 0 Then
        WriteLineConsole "Updated " & numRowsAffected & " rows, " & msg & "of the CONSTITUENTS sheet"
    Else
        WriteLineConsole "Updated zero rows"
    End If

End Sub

Sub create_or_update_constituent_data()
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Connect_To_DB cn
    
    ClearConsole
    WriteLineConsole "From constituent info create or update method..."
    
    Dim FieldColTick As Integer
    Dim rsMaterialsdb As ADODB.Recordset
    Dim canProceed As Boolean
    Dim firstFieldName As String
    Dim rowIterator As Integer
    Dim colIterator As Integer
    Dim sql As String
    Dim num As Long
    Dim adExecuteNoRecords As Long
    Dim cellVal As String
    Dim updateMsg As String
    Dim numRowsAffected As Integer
    
    FieldColTick = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''Check if first column is blank and return if blank
    If Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value) = "" Then
        MsgBox "First Column header is blank"
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''fetch all the fields of the constituent table and verify with the excel feild names. //TODO Check if the first column name is unique type column
    Set rsMaterialsdb = New ADODB.Recordset
    With rsMaterialsdb
        ' Assign the Connection object.
        .ActiveConnection = cn
        ' Extract the required records.
        .Open ("SELECT * FROM constituent LIMIT 1")
        canProceed = True
        

        Do While Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value) <> "" And canProceed = True
            If FieldExistsInRS(rsMaterialsdb, Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value)) Then
            Else
                canProceed = False
                WriteLineConsole "The field " & Trim(Sheets("CONSTITUENTS").Cells(1, FieldColTick).Value) & " doesnot exist in database"
            End If
            FieldColTick = FieldColTick + 1 'tick iterator
        Loop
        ''Now FieldColTick is equal to number of fields
        FieldColTick = FieldColTick - 1

        .Close
    End With

    ''check if all fields exist
    If canProceed = False Then
        WriteLineConsole "All Fields do not exist"
        Exit Sub
    End If
    ''fetched all the fields of the constituent

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''now for each value row starting from row 2, update a row if exists or create it
    firstFieldName = Trim(Sheets("CONSTITUENTS").Cells(1, 1).Value)

    numRowsAffected = 0
    updateMsg = "updated the rows "
    rowIterator = 2
    Do While Trim(Sheets("CONSTITUENTS").Cells(rowIterator, 1).Value) <> "" And rowIterator < 100
        ''Entered the row starting from row number 2
        ''First find if row exists using the first field
        Set rsMaterialsdb = New ADODB.Recordset
        rsMaterialsdb.ActiveConnection = cn
        rsMaterialsdb.Open ("SELECT * FROM constituent WHERE " + firstFieldName + " = '" + Trim(Sheets("CONSTITUENTS").Cells(rowIterator, 1).Value) + "'")
        
        If (rsMaterialsdb.BOF And rsMaterialsdb.EOF) Then
            rsMaterialsdb.Close
            ''The Row does not exist so INSERT THE ROW
            sql = "INSERT INTO constituent ("
            For colIterator = 1 To FieldColTick
                sql = sql & Trim(Sheets("CONSTITUENTS").Cells(1, colIterator).Value)
                If colIterator = FieldColTick Then
                    sql = sql & ") "
                Else
                    sql = sql & ", "
                End If
            Next
            sql = sql & "values ("
            For colIterator = 1 To FieldColTick
                cellVal = Trim(Sheets("CONSTITUENTS").Cells(rowIterator, colIterator).Value)
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
            rsMaterialsdb.Close
            ''The Row exists already so UPPDATE THE ROW
            sql = "UPDATE constituent SET "
            For colIterator = 2 To FieldColTick
                cellVal = Trim(Sheets("CONSTITUENTS").Cells(rowIterator, colIterator).Value)
                If cellVal = "" Then
                    cellVal = "NULL"
                Else
                    cellVal = "'" & cellVal & "'"
                End If
                sql = sql & Trim(Sheets("CONSTITUENTS").Cells(1, colIterator).Value) & " = " & cellVal
                If colIterator = FieldColTick Then
                    sql = sql & " "
                Else
                    sql = sql & ", "
                End If
            Next
            sql = sql & "WHERE " & firstFieldName & " = '" & Trim(Sheets("CONSTITUENTS").Cells(rowIterator, 1).Value) & "'"
        End If
        WriteLineConsole sql
        cn.Execute sql, num, adExecuteNoRecords
        If num = 1 Then
            numRowsAffected = numRowsAffected + 1
            updateMsg = updateMsg & CStr(rowIterator) & ", "
        ElseIf num > 1 Then
            WriteLineConsole "Beware: not updating a unique column at row " & CStr(rowIterator) & " of CONSTITUENTS sheet"
        End If
        rowIterator = rowIterator + 1
    Loop

    If numRowsAffected > 0 Then
        WriteLineConsole "Updated " & numRowsAffected & " rows, " & updateMsg & "of the CONSTITUENTS sheet"
    Else
        WriteLineConsole "Updated zero rows"
    End If

End Sub
