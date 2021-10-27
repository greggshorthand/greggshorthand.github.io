<%@ Language=VBScript %>
<!--#include file="function.asp"-->
<%
' Dimension Local variables
Dim RS                 ' Recordset Variable
Dim sql                ' variable for SQL statement
Dim strCount 
Dim GuestBookID
Dim TotalPages
Dim TotalRecords
Dim startNumber
Dim CurrentPage

SQL = "SELECT * from guest Order by entry_date desc"

''''''''''''''''''''''''''''''''
' Open Connection with database
''''''''''''''''''''''''''''''''
Call OpenDataConnection


Set RS = Server.CreateObject("ADODB.RecordSet")

RS.PageSize = Int(strPageSize)
RS.CacheSize = Int(strPageSize)

If Request("p") = "" THEN
	CurrentPage = 1
ELSE 
	CurrentPage =  Cint(Request("p"))
End If

RS.Open SQL, gbConn, adOpenKeyset , adLockOptimistic 
RS.AbsolutePage = CurrentPage
TotalPages = RS.PageCount
TotalRecords = RS.RecordCount
startNumber = TotalRecords - (strPagesize * (Currentpage -1))

%>
<html>
<head>
<title><%=gbName%></title>
</head>
<body>
<table width="948" border="0" cellpadding="0" cellspacing="0" class="text">
  <tr> 
    <td colspan="2" bgcolor="#CCCCCC" style="height: 53px; background-image: url(title3.gif); background-position: bottom right; background-repeat: no-repeat"><p>&nbsp;</p></td>
    <td width="182" bgcolor="#CCCCCC" style="background-image: url(gpco2.gif); background-position: top center; background-repeat: no-repeat"><div> 
        <div align="center"></div>
      </div></td>
  </tr>
  <tr> 
    <td colspan="2" valign="top" bgcolor="#CC3333" style="height: 10px; background-image: url(title2.gif); background-position: top right; background-repeat: no-repeat"> 
      <div class="text" style="color: #FFFFFF">&nbsp;A Web Site dedicated to the 
        perpetuation of Gregg&#8217;s Light-Line Phonography</div></td>
    <td width="182" bgcolor="#CC3333"><div align="center" style="color: #FFFFFF"> 
        - Anniversary Manual -</div></td>
  </tr>
  <tr> 
    <td width="182" valign="top" bgcolor="#CCCCCC" class="unnamed1"><table width="186" border="0" cellpadding="2" cellspacing="2" class="text">
        <tr> 
          <td width="178"> <!--#include file="nav.html" --> </td>
        </tr>
      </table></td>
<h2><%=gbName%></h2>
<table border="0" cellpadding="5"  cellspacing="0" width="640" >
	<tr>
	  <td bgcolor="#CC3333">
	    <table cellpadding="0" cellspacing="0" border="0" width="100%">
	      <tr>
	        <td>
	          <font face="Verdana, Arial, Helvetica" Size="2"><B><a href="<%=gbURL%>"><%=gbTitle%></a></b></font>
	        </td>
	        <td>
	          <font face="Verdana, Arial, Helvetica" Size="2">Total Guests: <B><%=TotalRecords%></b></font>
	          </td>
	        <td align="right">
	          <font face="Verdana, Arial, Helvetica" size="2"><b><a href="sign_guestbook.asp">Sign Our Guestbook</a></b></font>
	        </td>
	      </tr>
	     </table>
	  </td>
	</tr>
	<tr>
	  <td>
	  <Br>
    <%
    If RS.EOF THEN 
    	Response.Write "Sorry there are no Records to display."
    End If 

    strCount = Int(0)
       
    Do while strCount < strPageSize And NOT RS.EOF
      strcount = strcount + 1
      If NOT isNull(RS("name")) THEN
    	  strName = RS("name")
      ELSE
      	strName = " "
      End If
        
      If NOT isNull(RS("email")) THEN
    	  strEmail = RS("email")
      ELSE
    	  strEmail = " "
      End If


      If NOT isNull(RS("title")) THEN
      	StrTitle = RS("title")
      ELSE
      	strTitle = ""
      End If

      If NOT isNull(RS("url")) THEN
      	strURL = RS("url")
      ELSE
      	strURL = " "
      End If

      If NOT isNull(RS("country")) THEN
      	strCountry = RS("country")
      ELSE
      	strCountry = " "
      End If
        
      If NOT isNull(RS("reference")) THEN
      	strReference = RS("reference")
      ELSE
      	strReference = " "
      End If

      If NOT isNull(RS("comments")) THEN
      	strComments = RS("comments")
      ELSE
      	strComments = ":)"
      End If
      %>
      <b>Guest:</b>&nbsp;&nbsp;<%=startNumber%><br>
      <b>Name:</b>&nbsp;<a href="mailto:<%=strEmail%>"><%=strName%></a><br>
      <b>From:</b>&nbsp;<%=strCountry%><br>
      <b>Website:</b>&nbsp;<a href="<%=strURL%>" target="top"><%=strTitle%></a><br>
      <b>Referred By:</b>&nbsp;<%=strReference%><br>
      <b>Time:</b>&nbsp;<%=RS("entry_date")%><br>
      <b>Comments:</b>&nbsp;
      <center>
      <table width="600" border="0" cellpadding="5" cellspacing="0">
        <tr>
          <td>
            <%=FormatMessage(strComments)%>
          </td>
        </tr>
      </table>
      </center>
      <br>
      <hr width="400" color="Green" height="1">
      <P>
      <%
      startNumber = startNumber -1 
      RS.MoveNext
      Loop
      %>
      <center>
      <Table Border="0" Cellpadding="0" Cellspacing="0" >
	      <tr>
	        <%
	        If Cint(CurrentPage) > 1 THEN 
          %>
          <td>
	          <form Method="Post" Action="default.asp?p=<%=Currentpage - 1%>"><input Type="Submit" Value="Prev <%=strPageSize%> Guests"></form>
	        </td>
	        <%
          End If
          If Cint(CurrentPage) < Cint(TotalPages) THEN
          %>
          <td>
            <form Method="Post" Action="default.asp?p=<%=Currentpage + 1%>" Name="formNext"><input Type="Submit" Value="Next <%=strPageSize%> Guests"></form>
          </td>
          <%
          End If
          %>
        </tr>
      </table>
      </center>
    </td>
  </tr>
	<tr>
	  <td bgColor="#96E1FF">
	    <table cellpadding="0" cellspacing="0" Border="0" Width="100%">
	      <tr>
	        <td>
	          <font face="Verdana, Arial, Helvetica" size="2"><B><a href="<%=gbURL%>"><%=gbTitle%></a></b></font>
	        </td>
	        <td>
	          <font face="Verdana, Arial, Helvetica" size="2">Total Guests: <B><%=TotalRecords%></b></font>
	          </td>
	        <td Align="right">
	          <font face="Verdana, Arial, Helvetica" size="2"><B><a href="sign_guestbook.asp">Sign Our Guestbook</a></b></font>
	        </td>
	      </tr>
	     </table>
	  </td>
	</tr>
</table>
    <td width="182" valign="top" bgcolor="#CCCCCC" style="background-image: url(gpco3.gif); background-position: top center; background-repeat: no-repeat"><div> 
        <div align="center"> 
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="text">
            <tr> 
              <td valign="top">
                  <!--#include file="nav2.html" -->
                </td>
            </tr>
          </table>
        </div>
      </div></td>
  </tr>
  <tr bgcolor="#CC3333"> 
    <td height="6" colspan="3" valign="top"><div align="center" style="color: #FFFFFF">Design 
        Copyright &copy; 
        <!--#config timefmt="%Y" -->
        <!--#echo var="DATE_LOCAL" -->
        Andrew Owen. All Rights Reserved.</div></td>
  </tr>
</table>
</body>
</html>