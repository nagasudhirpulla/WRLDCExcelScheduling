Private Sub EntitlementGenCombo_Change()

End Sub

Sub populateGeneratorSNamesCombo()
    ''http://www.get-digital-help.com/2011/12/21/working-with-combo-boxes-form-control-using-vba/
    ''http://analysistabs.com/vba-code/activex-controls/combobox/
    getGeneratorSNamesArray
End Sub

Function getGeneratorSNamesArray()
    ''Connect to Schedule database
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    Connect_To_DB cn
    ' Create a recordset object.
    Dim rsMaterialsdb As ADODB.Recordset
    Set rsMaterialsdb = New ADODB.Recordset
    With rsMaterialsdb
        .ActiveConnection = cn
        .Open "SELECT sname FROM generator ORDER BY generator.sname ASC"
        'Ensure recordset is populated
        If Not .BOF And Not .EOF Then
            'not necessary but good practice
            Sheets("ENTITLEMENTS").EntitlementGenCombo.Clear
            While (Not .EOF)
                'print info from fields to the immediate window
                Sheets("ENTITLEMENTS").EntitlementGenCombo.AddItem rsMaterialsdb.Fields(0).Value
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    cn.Close
End Function
