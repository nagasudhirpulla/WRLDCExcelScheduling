''
''Module for common global variables and functions
''
''Difference between function and Sub at http://www.excel-easy.com/vba/function-sub.html
''Scope of Variable explained at http://www.excel-easy.com/vba/examples/variable-scope.html
Public ConsoleLineTick As Integer

Public Function ClearSheet(ByVal sheetName As String)
    Sheets(sheetName).Cells.ClearContents
End Function

Public Function ClearConsole()
    ClearSheet "CONSOLE"
    ConsoleLineTick = 1
End Function

Public Function WriteLineConsole(ByVal str As String)
    Sheets("CONSOLE").Cells(ConsoleLineTick, 1).Value = str
    ConsoleLineTick = ConsoleLineTick + 1
End Function

Public Function FieldExistsInRS( _
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

Public Function Connect_To_DB(ByRef cn As ADODB.Connection)
    ''https://www.reddit.com/r/excel/comments/2xpht7/vba_connection_string_to_google_cloud_sql/
    ''http://stackoverflow.com/questions/26369937/excel-vba-mysql-select-from-table-not-full-informaton
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Server_Name = "localhost" ' Enter your server name here
    Database_Name = "wrldc_schedule" ' Enter your database name here
    User_ID = "root" ' enter your user ID here
    Password = "123" ' Enter your password here
    cn.Open "Driver={MySQL ODBC 5.3 Unicode Driver};Server=" & Server_Name & ";Database=" & Database_Name & _
";Uid=" & User_ID & ";Pwd=" & Password & ";"
End Function
