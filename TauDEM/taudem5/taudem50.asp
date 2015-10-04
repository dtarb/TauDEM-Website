<% Response.Buffer=true  %>
<html>
<head>
<TITLE>TauDEM GP Toolbox Download</TITLE>
</head>
<body bgcolor="#FFFFFF">
<%
Response.Write("file")
Response.Write(Request.QueryString("personName"))
' Add Data to the DataBase
	set conn = server.createobject("adodb.connection")
	' **** BEGIN DSN-LESS CONNECTION CODE 
	DSNtemp="DRIVER={Microsoft Access Driver (*.mdb)}; "
    DSNtemp=dsntemp & "DBQ=" & server.mappath("taudem50.mdb")
    conn.Open DSNtemp
    ' **** END OF DSN-LESS CONNECTION CODE
    SQLstmt = "INSERT INTO download (ipaddress,action,when,personName,companyName,email,comments)" 
    SQLstmt = SQLstmt & " VALUES (" 
	SQLstmt = SQLstmt & "'" & Request.ServerVariables("REMOTE_ADDR") & "',"
	SQLstmt = SQLstmt & "'" & Request.Form("file") & "',"	
	SQLstmt = SQLstmt & "'" & NOW() & "',"
	SQLstmt = SQLstmt & "' " & Request.Form("personName") & "', "
	SQLstmt = SQLstmt & "' " & Request.Form("companyName") & "', "
	SQLstmt = SQLstmt & "' " & Request.Form("email") & "', "
      comments=replace(Request.Form("comments"),"'","") 
	SQLstmt = SQLstmt & "' " & comments & "' )"

  
  ' remember that numbers aren't enclosed 
  ' in single quotes but text must always be enclosed!
 ' Response.Write(dsntemp)
  'Response.Write(sqlstmt)   'Useful for debugging
  Set RS = conn.execute(SQLstmt)
  Conn.Close
  Response.Clear
  Response.Redirect(Request.Form("file"))
  Response.End
%>


</body></html>
