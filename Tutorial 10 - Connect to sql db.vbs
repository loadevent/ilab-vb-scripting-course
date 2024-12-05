'********************************************************
'          Tutorial 10 - Connect To DB
'********************************************************
Option Explicit

'Declare
    DIM strConnection
    DIM objConnection
    DIM rs 'Record set
    DIM strSelect, strResults

'Assign
    
    strConnection = "provider=sqloledb;server=ILAB-030\SQLEXPRESS;Database=NorthwindNovember;Integrated Security=SSPI"
    SET objConnection = CreateObject("ADODB.CONNECTION")
    SET rs = CreateObject("ADODB.RECORDSET")

    objConnection.Open strConnection

    strSelect = "SELECT * FROM Shippers order by CompanyName DESC"

    rs.Open strSelect, objConnection

    DO WHILE NOT rs.EOF
        strResults = strResults & rs.Fields.Item(0) & " - " & rs.Fields.Item(1) & vbNewLine
        rs.MoveNext
    LOOP

    MsgBox(strResults)

    objConnection.close



