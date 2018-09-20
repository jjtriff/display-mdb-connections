

Sub ShowUserRosterMultipleUsers(sPath, dbPassword)
    Set cn = CreateObject ("ADODB.Connection")
    Set rs = CreateObject ("ADODB.Recordset")

    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open "Data Source='" & sPath & "';;Jet OLEDB:Database Password=" & dbPassword & ";"

    ' The user roster is exposed as a provider-specific schema rowset
    ' in the Jet 4 OLE DB provider.  You have to use a GUID to
    ' reference the schema, as provider-specific schemas are not
    ' listed in ADO's type library for schema rowsets

    Set rs = cn.OpenSchema(-1, _
    , "{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    'Output the list of all users in the current database.

    WScript.Echo rs.Fields(0).Name, VBTab,VBTab, rs.Fields(1).Name, VBTab _
        , rs.Fields(2).Name, rs.Fields(3).Name

    While Not rs.EOF
        WScript.Echo rs.Fields(0).Value, VBTab,VBTab, rs.Fields(1).Value, VBTab, _
        rs.Fields(2).Value, VBTab, rs.Fields(3).Value
        rs.MoveNext
    Wend

End Sub

Sub printUsage()
    WScript.Echo "Shows what users and what PCs are connected to a .MDB", VBCrLF, _
        VBCrLF, _
        "Usage: ", VBCrLF, _
        "cscript.exe .\display.vbs '<c:\full\path\accessDb.mdb>' ['password-for-db']"
End Sub


' variables q se usaran
sMdbPath = ""
sPass = ""
Set objArgs = Wscript.Arguments

' si no se pasan argumentos
If objArgs.Count = 0 Then
    ' imprimir el usage
    printUsage
Else
    ' sino
    ' imprimir quien esta usando el mdb
    sMdbPath = objArgs(0)
    If objArgs.Count = 2 Then
        sPass = objArgs(1)
    End If
    ShowUserRosterMultipleUsers sMdbPath, sPass
End If

WScript.Echo VBCrLF, _
    "Done"



