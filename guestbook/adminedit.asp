<%
If Session("PassCheck") <> "yes" Then
  Response.Redirect "admin.asp"
End If
%>
<!--#include FILE="function.asp"-->
<%

If Request.Form("Action")= "Update" Then
	UpdateRecord
End If

If Request.Form("Action")= "Delete" Then
	DeleteRecord
End If

Dim SearchText
Dim RS
Dim sql
Dim ShowSearch
Dim DisplayNumber
DisplayNumber = 0

' Open DataConnection with database
Call OpenDataConnection 

strTotalRecords = "5"

If Request("DisplayNumber") <> "" Then
  strTotalRecords =  Request.form("DisplayNumber")
End If

SearchText = Trim(Request("st"))
If SearchText <> "" Then
	sql = "select * from guest"
	sql = sql & " where name like '%" & SearchText & "%'"
	sql = sql & " or email like '%" & SearchText & "%'"
	sql = sql & " or title like '%" & SearchText & "%'"
	sql = sql & " or url  like '%" & SearchText & "%'"
	sql = sql & " or comments like '%" & SearchText & "%'"
	sql = sql & " order by entry_date desc"
Else
  sql = "Select * from guest order by entry_date desc"
End If
	
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.Open sql, gbConn , adOpenKeySet
%>
<html>
<head>
<title><%=gbName%></title>
</head>
<body>
<center>
<a href="http://www.devasp.com"><img src="banner1.gif" border="0"></a>
<font Face="arial" Size="2"><H2><%=gbName%></H2></font>
<table width="550" border="2" BorderColor="#686898" cellspacing="0" cellpadding="10" valign="top">
  <tr>
    <td bgColor="#96E1FF">
	    <table cellpadding="0" cellspacing="0" Border="0" Width="100%">
	      <tr>
	        <td>
	          <font Face="arial" Size="2"><B><a href="<%=gbURL%>"><%=gbTitle%></a></b></font>
	        </td>
	        <td Align="right">
	          <font Face="arial" Size="2"><B><a href="sign_guestbook.asp">Sign Our Guestbook</a></b></font>
	        </td>
	      </tr>
	     </table>
	   </td>
	 </tr>
	 <tr>
    <td Align="center">
      <form Method="Post" Action="adminedit.asp" id=form1 name=form1>
      <table cellspacing="0" cellpadding="5" border="0">
    	<tr>
    		<td>
    		    <font Face="arial" Size="2">Search Keywords: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input NAME="st" size="40" maxlength="100" value></font>
    		</td>
    	</tr>
    	<tr>
    		<td>
    		    <font Face="arial" Size="2">Number Of Records: &nbsp;&nbsp;&nbsp;<input NAME="DisplayNumber" size="5" maxlength="5" Value="5"></font>
    		    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="Submit" value="Search!" id=Submit1 name=Submit1></form>
    		</td>
    	</tr>
    </table>
   </td>
  </tr>
<%
If RS.EOF Then
  Response.Write " No Records Exist for Editing."
Else
  Do Until RS.EOF or DisplayNumber = Int(strTotalRecords)
  %>
  <tr>
	  <td>
		  <form method="post" action="adminedit.asp">
			<input type="hidden" Name="id" value="<%=RS("id")%>">
			<input type="hidden" Name="st" value="<%=request("st")%>">
			<table>
			  <tr>
				  <td>
					  <font face="arial" size="-1"><b>ID:</b></font>
					</td>
					<td>
  					<b><%=RS("id")%></b>
          </td>
        </tr>
        <tr>
          <td>
            <font face="arial" size="-1"><b>Name:</b></font>
          </td>
          <td>
            <input name="name"  value="<%=RS("name")%>" type="text" size="35">
          </td>
        </tr>
        <tr>
          <td>
            <font face="arial" size="-1"><b>Email:</b></font>
          </td>
          <td>
            <input name="email"  value="<%=RS("email")%>" type="text" size="35">
          </td>
        </tr>
        <tr>
          <td>
            <font face="arial" size="-1"><b>Title:</b></font>
          </td>
          <td>
            <input name="title"  value="<%=RS("title")%>" type="text" size="35">
          </td>
        </tr>
        <tr>
          <td>
            <font face="arial" size="-1"><b> URL:</b></font>
          </td>
          <td>
            <input type="text" name="url" value="<%=RS("url")%>" size="35">
          </td>
        </tr>
        <tr>
          <td>
            <font face="arial" size="-1"><b>Country:</b></font>
          </td>
          <td>
            <input type="text" name="country" value="<%=RS("country")%>" size="35">
          </td>
        </tr>
        <tr>
          <td Valign="top">
            <font face="arial" size="-1"><b>Comments:</b></font>
          </td>
          <td>
            <textarea rows="5" cols="49" Name="comments"><%=RS("comments")%></textarea>
          </td>
        </tr>
      </table>
    </td>
	</tr>
	<tr>
	  <td bgcolor="#96e1ff" align="center">
		  <input type="Submit" Name="Action" value="Update">&nbsp;&nbsp;&nbsp;
		  <input type="Submit" Name="Action" value="Delete">
    </td>
  </tr>
  </form>
	<%
	DisplayNumber = DisplayNumber + 1
	RS.MoveNext
	Loop
	End If
	%>
</table>
</body>
</html>
<%

Function strReplace(strValue)
  If not IsNull(strValue) or Trim(strValue) <> "" then
    strReplace =  Replace(strValue, "'","''")
  end if
End function
	

Sub UpdateRecord
  Call OpenDataConnection 
	Dim ID
	ID = Int(Request.Form("id"))
	sql = "update guest"
	sql = sql & " Set name ='" & strReplace(Trim(Request.Form("name"))) & "'"
	sql = sql & ", email ='" & strReplace(Trim(Request.Form("email"))) & "'"
	sql = sql & ", title ='" & strReplace(Trim(Request.Form("title"))) & "'"
	sql = sql & ", url ='" & strReplace(Trim(Request.Form("url"))) & "'"
	sql = sql & ", country ='" & strReplace(Trim(Request.Form("country"))) & "'"
	sql = sql & ", comments ='" & strReplace(Trim(Request.Form("comments"))) & "'"
	sql= sql & " where id =" & ID
	gbConn.Execute sql
	gbConn.Close
End Sub


Sub DeleteRecord
  Call OpenDataConnection 
	Dim ID
	ID = Int(Request.Form("id"))
	sql = "Delete from guest"
	sql= sql & " where id = "& ID
	gbConn.Execute sql
	gbConn.Close
End Sub
%>
