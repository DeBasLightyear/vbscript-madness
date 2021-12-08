Function getCsvString(recordSet)
    Dim field
    Dim cursor
    cursor = 1

    ' Get the header rows
    For Each field in recordSet.Fields
        getCsvString = getCsvString & field.Name & ";"
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
    MsgBox("Dit kan even duren. De volgende query wordt uitgevoerd:" & vbLf & sql)
    Set queryDataFromDatabase = oConnection.Execute(sql)
End Function

Function readTextFile(pathToFile)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set file = fso.OpenTextFile(pathToFile, 1)
    readTextFile = file.ReadAll()
End Function

Sub writeTextToFile(content, fileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim filePath
    filePath = "./" & fileName
    
    ' open a new file for writing and add the content
    Set file = fso.OpenTextFile(filePath, 2, True)
    file.Write(content)
    file.Close()
End Sub

' path to the MS Access database
Dim dbPath
dbPath = "./databases/dvdrental.mdb"

' your SQL
Dim sql
sql = readTextFile("./my-query.sql")

' execute the query and write the result to a file
Dim queryResult
Set queryResult = queryDataFromDatabase(dbpath, sql)

Dim csvString
csvString = getCsvString(queryResult)

Call writeTextToFile(csvString, "test-output.csv")

' Notify the user that things have happened.
MsgBox("Things have happened.")
