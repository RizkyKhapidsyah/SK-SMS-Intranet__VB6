<%@ LANGUAGE=VBScript %>

<html>

<head>
<title></title>
</head>

<body>


<%

	Dim textASP
	Dim BrojASP
	
	textASP=Request.Form("textfield")
	brojASP=Request.Form("EEM")
	
	if len(brojasp)<3 then
	response.write("<h2 ALIGN="& Center &"> Please select the user ! </h2>")
	else
	response.write("<h2 ALIGN=" & Center & ">SMS Message </h2></br>")  
	response.write("<h1 ALIGN=" & Center & ">" & textASP & " </h1></br></br>")
	response.write("<h2 ALIGN=" & Center & ">Is send to </h2></br>")  
	response.write("<h1 ALIGN=" & Center & ">" & BrojASP & " </h1>")
	
	
	Dim conn
    Dim str_con
    Dim str_ins 
	
    Set conn =server.createobject("ADODB.Connection")    

    conn.CommandTimeout = 2
    conn.Open "DSN=SMSIntranet;uid=sa;pwd="
  
    str_ins = "INSERT INTO Buffer( FromASP,ToASP, TextASP)" _
    & "VALUES (' Unknown ','" & BrojASP & " ',' "& textASP &"' )"
    conn.Execute str_ins
    

    conn.Close
end if

%>
</body>
</html>
