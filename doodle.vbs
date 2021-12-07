Function doodle(pathToDb, sql)
    ' Connect to a database
    ' Build oConnection string
    Dim sConnectionString
    sConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & pathToDb

    ' Create oConnection object and open connection
    Dim oConnection
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open(sConnectionString)

    ' Get recordset object from SQL query    
    Dim objRecordset
    Set objRecordset = oConnection.Execute(sql)

    Set doodle = objRecordset
End Function


