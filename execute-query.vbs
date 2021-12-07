Function getCsvString(recordSet)
    Dim field
    Dim cursor
    cursor = 1

    ' Get the header rows
    For Each field in recordSet.Fields
        If field.Name = "" Then
            getCsvString = getCsvString & "(computed" & c & ")" & ";"
            cursor = cursor + 1
        Else
            getCsvString = getCsvString & field.Name & ";"
        End If
    Next
    
    ' remove the last semicolon from the result
    getCsvString = Left(getCsvString, Len(getCsvString) - 1)

    ' Get the data rows
    While Not(recordSet.EOF)
        ' add a new line
        getCsvString = getCsvString & vbLf
        For Each field In recordSet.Fields
            getCsvString = getCsvString & field.Value & ";"
        Next
        getCsvString = Left(getCsvString, Len(getCsvString) - 1)
        recordSet.MoveNext()
    Wend
End Function

Function queryDataFromDatabase(pathToDb, sql)
    ' Connect to a database
    ' Build oConnection string
    Dim sConnectionString
    sConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & pathToDb

    ' Create oConnection object and open connection
    Dim oConnection
    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Open(sConnectionString)

    ' Get recordset object from SQL query    
    Set queryDataFromDatabase = oConnection.Execute(sql)
End Function

Dim dbPath
dbPath = "./dvdrental.accdb"

Dim foobarz
foobarz = "SELECT * FROM film WHERE film_id = 4"

Dim queryResult
Set queryResult = queryDataFromDatabase(dbpath, foobarz)

Dim csvString
csvString = getCsvString(queryResult)

MsgBox(csvString)
