<!--#include file="function.asp"-->
<html>
<head>
<title><%=gbName%></title>
<SCRIPT LANGUAGE="JavaScript">
<!--//
function Checkfields()
{
  if(document.frmGB.name.value == ""){
		alert("Please Enter your name.");
		return false;}
		
		if(document.frmGB.comments.value == ""){
		alert("Guestbook can't be sign without Comments.");
		return false;}
		
}
//-->
</script>
</head>
</body>
<FORM Name"GBsign" ACTION="process.asp" METHOD="post" ONSUBMIT="return Checkfields();" Name="frmGB">
<input Type="hidden" Name="FrmName" Value="GB">
<center>
<a href="http://www.devasp.com"><img src="banner.gif" border="0"></a>
<font Face="arial" Size="2"><H2>Sign <%=gbName%></H2></font>
<table border="2" cellpadding="15"  Cellspacing="0" Width="500" BorderColor="#686898">
	<tr>
	  <td bgColor="#96E1FF">
	    <table Border="0" Width="100%">
	      <tr>
	        <td>
	          <font Face="arial" Size="2"><B><a href="<%=gbURL%>"><%=gbTitle%></a></b></font>
	        </td>
	        <td Align="right">
	          <font Face="arial" Size="2"><B><a href="default.asp">View Our Guestbook</a></b></font>
	        </td>
	      </tr>
	     </table>
	  </td>
	</tr>
  <TR>
		<TD ALIGN="left">
		  <font Face="arial" Size="2">
		  <B>Name:</B><BR>
			<INPUT TYPE="text" NAME="name" SIZE="15"></INPUT><BR>
			<font Face="arial" Size="2"><B>E-mail Address:</B><BR>
			<INPUT TYPE="text" NAME="email" SIZE="35"></INPUT><BR>
			<font Face="arial" Size="2"><B>Title of your homepage:</B><BR>
			<INPUT TYPE="text" NAME="title" SIZE="35"></INPUT><BR>
			<B>Homepage URL:</B><BR>
			<INPUT TYPE="text" NAME="url" Value="http://"  SIZE="35"></INPUT><BR>
			<B>Country:</B><BR>
			<INPUT TYPE="text" NAME="country" SIZE="35"></INPUT><BR>
			<B>How did you find us?:</B><BR>
			<SELECT NAME="reference">
			<OPTION selected>Just Surfed On In
			<OPTION>From a Friend
			<OPTION>Banner Advertisement
			<OPTION>Newsletter
			<OPTION>Yahoo!
			<OPTION>Altavista
			<OPTION>AOL
			<OPTION>Tripod
			<OPTION>Lycos
			<OPTION>AngelFire
			<OPTION>Geocities
			<OPTION>Xoom
			<OPTION>Net Search
			<OPTION>NewsGroups
			<OPTION>Signing another Guestbook
			<OPTION>Viewing another Guestbook
			</SELECT><BR><BR>
			<B>Comments:</b><BR>
			<TEXTAREA NAME="comments" COLS=55 ROWS=8></textarea>
			</font>
	  </td>
  </tr>
  <Tr>
    <td bgColor="#96E1FF">
	    <center>
			<INPUT TYPE="Submit" VALUE="Submit!">&nbsp;&nbsp;&nbsp;<input type=reset value="Clear All">
			</center>
		</td>
	</tr>
</table>
</FORM>
</body>
</html>