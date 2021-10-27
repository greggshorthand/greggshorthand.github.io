<!--#include File="function.asp"-->
<%
If Request.Form("FrmName") = "GB" THEN
  Call AddGuest
ELSE
  Response.Redirect "default.asp"
End If

Sub CheckValid(MyStr)
	MyStr = Replace(MyStr, "'", "''")
	If MyStr = "" THEN
		MyStr = " "
	End If
End Sub

Sub AddGuest()

  Dim RS, MyStr, strName, strEmail, strTitle, strURL, strCountry, strReference, strComments
  strName = Server.HtmlEncode(Request.form("name"))
  strEmail = server.htmlEncode(Request.form("email"))
  strTitle = server.HTMLencode(Request.form("title"))
  strurl = server.htmlEncode(Request.form("url"))
  strCountry = server.HtmlEncode(Request.form("country"))
  strReference = server.HtmlEncode(Request.form("reference"))
  strComments = server.HtmlEncode(Request.form("comments"))

  ' You can add your default message here
  ' If no comments THEN add a Simple smile face for comments

  If Request.Form("comments")= "" THEN
  	strComments = ":)"
  End If
  If Len(Request.Form("comments")) > 1000 THEN
    strComments = Left(Request.Form("comments"),1000 ) & ".."
  End If


  ' This will check and replace single quotes with doubble quotes. 

  Call CheckValid(strName)
  Call Checkvalid(strEmail)
  Call Checkvalid(strTitle)
  Call Checkvalid(strURL)
  Call CheckValid(strCountry)
  Call CheckValid(strReference)
  Call CheckValid(strComments)

  Call OpenDataConnection 
  Set RS =  Server.CreateObject("ADODB.RecordSet")
  sql = "Insert into guest(name, email, title, url, country, reference, comments)"
  sql = sql & " Values('" 
  sql= sql & strName & "','"
  sql= sql & strEmail & "','"
  sql= sql & Server.HTMLEncode (strTitle) & "', '"
  sql= sql & strURL & "', '"
  sql= sql & strCountry & "', '"
  sql= sql & strReference & "', '"
  sql= sql & strComments & "')"
  
  gbConn.Execute sql
  
  ' Check If SendMail is set to true then send the mail.
 If LCase(SendMail) = "yes" and strEmail <> "" Then
	  Set objNewMail = Server.CreateObject("CDONTS.NewMail")
		objNewMail.From = "devasp@aol.com"
		objNewMail.To = strEmail
		
		'''''''''''''''''''''''''''''''''''
		' Change the Body text of mail here
		''''''''''''''''''''''''''''''''''' 
		strBody = "Thanks for visiting http://www.devasp.com" & VBCRLF 
    strBody = strBody & "We are glad that you signed our Guestbook" & VBCRLF 
    strBody = strBody & "Please visit DevASP.com again, as we are updating our site with new Samples and Articles regularly." & VBCRLF & VBCRLF 
    strBody = strBody & "Thanks" & VBCRLF 
    strBody = strBody & "Asif" & VBCRLF 
    strBody = strBody & "http://www.devasp.com" & VBCRLF 
      
		objNewMail.Subject = "Thanks for Signing our Guestbook!"
		objNewMail.Body = strBody
		objNewMail.Send
		Set objNewMail = Nothing
	End If			
	gbConn.Close
  Response.Redirect "default.asp"
End Sub
%>