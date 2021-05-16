<%@ LANGUAGE=VBScript %>
<%response.buffer=true%>
<html>

<head>
<title></title>
</head>

<body>

<form name="form1" method="post" action="default1.asp">
  <div align="center"><center>
      <h1>SMS Intranet service</h1>
  </center></div><div align="center"><center>
      <p> 
        <input type="text" name="textfield" size="152" maxlength="159">
      </p>
      <p>
        <input type="submit" name="Submit"
  value="Submit">
      </p>
  </center></div><%


Dim SERVISI
Dim sqlObjekti
Dim rsObjekti
Dim eemObjekt
Dim i
i=1
set SERVISI =server.createobject("ADODB.Connection")
SERVISI.open "DSN=SMSIntranet;uid=sa;pwd=;"

	sqlObjekti="select Name from Users Order by Name"

	set rsObjekti=SERVISI.Execute(sqlObjekti)
	if rsObjekti.EOF=True then
	else
	rsObjekti.movefirst

	
	do while not rsObjekti.eof

response.write("<TABLE id=Table1 style=" & " WIDTH: 700px; HEIGHT: 54px " & "border=0 name=" & "Table1" & ">")  	
response.write("<TR>")

	if rsObjekti.eof=True then
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><BODY>--</BODY></TD>")
	i=i+1
	else		
	eemObjekt=Replace(rsObjekti("Name"),Chr(32),"&nbsp")
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><input type= radio name= EEM  value=" & eemObjekt & "> "& eemObjekt & " </BODY></TD>")

	i=i+1
	rsObjekti.MoveNext
	end if

	if rsObjekti.eof=True then
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><BODY>--</BODY></TD>")
	i=i+1
	else		
	eemObjekt=Replace(rsObjekti("Name"),Chr(32),"&nbsp")
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><input type= radio name= EEM  value=" & eemObjekt & "> "& eemObjekt & " </BODY></TD>")
	i=i+1
	rsObjekti.MoveNext
	end if

	if rsObjekti.eof=True then
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><BODY>--</BODY></TD>")
	i=i+1	
	else		
	eemObjekt=Replace(rsObjekti("Name"),Chr(32),"&nbsp")
	response.write("<TD WIDTH=30px><BODY ><B>" &  i & ".</B></BODY></TD><TD WIDTH=300px><input type= radio name= EEM  value=" & eemObjekt & "> "& eemObjekt & " </BODY></TD>")
	i=i+1
	rsObjekti.MoveNext
	end if

		response.write("</TR>")
		response.write("</TABLE>")

	loop
	
	rsObjekti.close
	end if
	set rsObjekti=nothing
	

SERVISI.close

%>

</form>
</body>
</html>
