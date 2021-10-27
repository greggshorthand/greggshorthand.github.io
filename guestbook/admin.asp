<!--#include file="function.asp"-->
<%
Dim WrongPass
WrongPass = "0"

If Request.Form("Name")= "" Or Request.Form("Password")=""  Then
	ShowForm
End If

If Request.Form("Name") <> "" and Request.Form("Password")<>"" Then
	CheckLog
End If


Sub CheckLog
  If Request.Form("Name") = AdminName and Request.Form("Password") = AdminPass Then
    Session("PassCheck") = "yes"
    Response.Redirect "adminedit.asp"
  Else
	  WrongPass = "1"
		ShowForm
  End If
End Sub


Sub ShowForm
%>
<html>
<head>
<title><%=gbName%></title>
</head>
<body>
<center>
<a href="http://www.devasp.com"><img src="banner1.gIf" border="0"></a>
<form Method="post" action="admin.asp"><center>
<H2><%=gbName%></H2>
<center>
  Please enter your User Name and Password.
</center>
<P>
<table border="2" cellpadding="15"  Cellspacing="0" BorderColor="#686898" width="200">
	<tr>
		<td Align="center">
			<B>Name:</b><br>
			<input type=text Name="Name" size=15><P>
			<b>Password:</B><BR>
			<input type="password" Name="Password" size=15><br><br>
			<center><input type="Submit" Value="Submit" id=Submit1 name=Submit1></center>
			</td>
	</tr>
	<%
	If WrongPass = "1" Then
	%>
	<tr>
	  <td>
		  <Font Face="arial" Color="Red">
		  Sorry Wrong Password Try again.
		  </font>
			</td>
	</tr>
	<%
	End If
	%>
</table>
</form>
</body>
</html>
<%
End Sub
%>