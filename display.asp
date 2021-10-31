<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Show Data</title>
    <style>
        table, td{border: 1px solid blue}
    </style>
</head>
<body>
    
    <%
        Dim con ' for connection Object 
        Dim rec ' for Recordset Object
        Dim rs ' to hold pointer
        Dim x ' for loop counter variable

        ' create a connection object
        Set con = Server.createObject("ADODB.Connection")
        
        ' create a recordset Object
        Set rec = Server.createObject("ADODB.recordset")

        con.open "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = newpro; User Id = Joy; Password=1234"

        Set rs =  con.execute("Select * from Student")



    %>

    <table>
    <tr>
        <th>Roll</th>
        <th>Name</th>
    </tr>
        <%
            Do Until rs.EOF
                Response.write("<tr>")
                For Each x In rs.Fields
                Response.write("<td>"& x.value & "</td>")
                Next
                Response.write("</tr>")
                rs.movenext
            Loop
        %>
    </table>
</body>
</html>
