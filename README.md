# VBA_SQL
Tenging milli VBA og SQL servers

Function ImportSQLtoRange(ByVal conString As String, ByVal query As String, _
    ByVal target As Range) As Integer

    On Error Resume Next

    ' Object type and CreateObject function are used instead of ADODB.Connection,
    ' ADODB.Command for late binding without reference to
    ' Microsoft ActiveX Data Objects 2.x Library

    ' ADO API Reference
    ' http://msdn.microsoft.com/en-us/library/ms678086(v=VS.85).aspx

    ' Dim con As ADODB.Connection
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")

    con.ConnectionString = conString

    ' Dim cmd As ADODB.Command
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")

    cmd.CommandText = query
    cmd.CommandType = 1         ' adCmdText

    ' The Open method doesn't actually establish a connection to the server
    ' until a Recordset is opened on the Connection object
    con.Open
    cmd.ActiveConnection = con

    ' Dim rst As ADODB.Recordset
    Dim rst As Object
    Set rst = cmd.Execute

    If rst Is Nothing Then
        con.Close
        Set con = Nothing

        ImportSQLtoRange = 1
        Exit Function
    End If

    Dim ws As Worksheet
    Dim col As Integer

    Set ws = target.Worksheet

    ' Column Names
    For col = 0 To rst.Fields.Count - 1
        ws.Cells(target.Row, target.Column + col).Value = rst.Fields(col).Name
    Next
    ws.Range(ws.Cells(target.Row, target.Column), _
        ws.Cells(target.Row, target.Column + rst.Fields.Count)).Font.Bold = True

    ' Data from Recordset
    ws.Cells(target.Row + 1, target.Column).CopyFromRecordset rst

    rst.Close
    con.Close

    Set rst = Nothing
    Set cmd = Nothing
    Set con = Nothing

    ImportSQLtoRange = 0

End Function

Sub RunSQL()

    Dim query           As String
    Dim conString       As String
    Dim target          As Range
    
    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    ' Þessu þarf að breyta fyrir hvert sheet/file
    ' SQL queryið er lesið úr t.d. Sheet1!B1
    '  - Líka hægt að skrifa það bara hér 
    '  - t.d. query = "SELECT * FROM [gagnagrunnur] ..."
    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    
    query = Worksheets("Sheet1").Range("B1").Value
    
    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    ' Target er síðan hvar taflan á að birtast úr SQL servernum
    '   - Annað hvort hægt að skrifa bara t.d. D10 í Sheet1!B2 og þá skrifast taflan út þar
    '   - Eða hægt að skrifa hér hvar taflan á að birtast
    '   - T.d. Set target = Worksheets("Sheet1").Range("D10")
    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    
    Set target = Worksheets("Sheet1").Range(Range("B2").Value)

    conString = GetConnectionString()
    
    Select Case ImportSQLtoRange(conString, query, target)
        Case 1
            MsgBox "Import database data error", vbCritical
        Case Else
    End Select

End Sub

Function OleDbConnectionString(ByVal Server As String, ByVal Database As String, _
    ByVal Username As String, ByVal Password As String) As String

    If Username = "" Then
        OleDbConnectionString = "Provider=SQLOLEDB.1;Data Source=" & Server _
            & ";Initial Catalog=" & Database _
            & ";Integrated Security=SSPI;Persist Security Info=False;"
    Else
        OleDbConnectionString = "Provider=SQLOLEDB.1;Data Source=" & Server _
            & ";Initial Catalog=" & Database _
            & ";User ID=" & Username & ";Password=" & Password & ";"
    End If

End Function

Function GetConnectionString() As String

    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    ' Hér þarf að slá inn upplýsingar um serverinn þegar Azure aðgangurinn er kominn í gang
    ' Fyrst kemur nafn á server, síðan nafn gagnagrunnsins, þarnæst notendandanafn og að lokum lykilorð
    ' ----------------------------------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------------------------------
    GetConnectionString = OleDbConnectionString("xxxxx.database.windows.net", "xxxx", "xxxxxx", "xxxxxxxxxxxx")
  
End Function


