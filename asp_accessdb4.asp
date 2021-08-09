<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<html>
<body>


<%
 Dim strFile, strSQL, Conn, rs, rs2, rs3, rs4, rs5, data1, txtz, ccid, f1, f2, f3, f4 



SUB LoadDB_Query 'This sub queries the Access DB 
'Connect to Access DB===============================================
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    Conn.Open "C:\inetpub\wwwroot\ClassicASP\Database3.mdb"
'===================================================================



ccid=Request.Form("cid")
f1=ccid

if len(ccid)=0 then 
ccid="-1000"
f1=""
end if
Set data1 = Server.CreateObject("ADODB.Recordset")
Response.Write ("<br>")
strSQL = "SELECT Table1.FirstName, Table1.LastName, Table1.Address, Table1.ContactID FROM Table1 WHERE ContactID=" & ccid & ";"


data1.Open strSQL, Conn

Response.Write("Query Results:") 
Response.Write ("<br>")

if data1.eof then

Response.Write("No Records Found")
Response.Write ("<br>")
Response.Write ("<br>")
else

f2=data1.Fields("FirstName")
f3=data1.Fields("LastName")
f4=data1.Fields("Address")

Response.Write("Success! Record Found")
Response.Write ("<br>")
Response.Write ("<br>")
end if



data1.Close

END SUB

SUB UpdateDB 'This sub Inserts new records or Updates existing records in the Access DB 
'Connect to Access DB===============================================
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    Conn.Open "C:\inetpub\wwwroot\ClassicASP\Database3.mdb"
'===================================================================
ccid=Request.Form("cid")
f1=ccid



if len(ccid)=0 then 
Response.Write("Error! Record Not Created/Updated")
Response.Write ("<br>")
Response.Write ("<br>")
f1=""
exit sub
end if

Set data1 = Server.CreateObject("ADODB.Recordset")
Response.Write ("<br>")
strSQL = "SELECT Table1.FirstName, Table1.LastName, Table1.Address, Table1.ContactID FROM Table1 WHERE ContactID=" & ccid & ";"

data1.Open strSQL, Conn

f1=""
f2=""
f3=""
f4=""



if data1.eof then
f2=Request.Form("fname")
f3=Request.Form("lname")
f4=Request.Form("address")
f1=Request.Form("cid")



strSQL = "INSERT INTO Table1(FirstName, LastName, Address, ContactID) VALUES('"& f2 & "', '" & f3 & "','" & f4 & "','" & f1 & "');"



Conn.execute strSQL


Response.Write("Success! New Record Created")
Response.Write ("<br>")
Response.Write ("<br>")
else


f2=Request.Form("fname")
f3=Request.Form("lname")
f4=Request.Form("address")
f1=Request.Form("cid")


strSQL = "UPDATE Table1 SET FirstName='" & f2 & "', LastName='" & f3 & "', Address='" & f4 & "' WHERE ContactID="& f1 &";"

Conn.execute strSQL

Response.Write("Success! Record Updated")
Response.Write ("<br>")
Response.Write ("<br>")

end if

data1.Close
END SUB

SUB DeleteDB 'This sub deletes records the Access DB 
'Connect to Access DB===============================================
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    Conn.Open "C:\inetpub\wwwroot\ClassicASP\Database3.mdb"
'===================================================================
ccid=Request.Form("cid")
f1=ccid

if len(ccid)=0 then 
Response.Write("Record Does Not Exist")
Response.Write ("<br>")
Response.Write ("<br>")
f1=""
exit sub
end if

Set data1 = Server.CreateObject("ADODB.Recordset")
Response.Write ("<br>")
strSQL = "SELECT Table1.FirstName, Table1.LastName, Table1.Address, Table1.ContactID FROM Table1 WHERE ContactID=" & ccid & ";"

data1.Open strSQL, Conn

f1=""
f2=""
f3=""
f4=""



if data1.eof then
f2=Request.Form("fname")
f3=Request.Form("lname")
f4=Request.Form("address")
f1=Request.Form("cid")
Response.Write("Record Does Not Exist")
Response.Write ("<br>")
Response.Write ("<br>")
else

strSQL = "DELETE FROM Table1 WHERE ContactID=" & ccid & ";"

Conn.execute strSQL

Response.Write("Success! Record Deleted")
Response.Write ("<br>")
Response.Write ("<br>")

end if

data1.Close
END SUB



select case Request.Form("Query")
case "Query/Look-up"
LoadDB_Query

case "Add/Update"
UpdateDB

case "Delete"
DeleteDB

end select



%>

<label for="cid1">ContactID: </label><br>
  <input type="text" id="cid2" name="cid1" value="<%=f1%>"><br> 
<label for="cid2">First Name: </label><br>
  <input type="text" id="cid2" name="cid2" value="<%=f2%>"><br> 
<label for="cid3">Last Name: </label><br>
  <input type="text" id="cid3" name="cid3" value="<%=f3%>"><br> 
<label for="cid4">Address: </label><br>
  <input type="text" id="cid4" name="cid4" value="<%=f4%>"><br> 

<form action="html_form1.htm" method="put">
<p><input type="submit" value="Return"></p>
</form>
</body>
</html>
