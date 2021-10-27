<!--#include file="adovbs.inc"-->
<%
'*************************************************************
'
'           *********************************************
'           * DevASP GuestBook Version 1.0              *
'           * Date 12/12/99                             *
'           * Written by Asif (http://www.devasp.com)   *
'           * Copy Right www.devasp.com                 *
'           *********************************************
'
'
' You Can use DevASP GuestBook Version 1.0 as you like.
' DevASP GuestBook can be modified or changed by anyone as long as this
' copyright notice and the comments is not removed.
' Before using this GuestBook you are agree that 
' you are using this guestbook on your own risk without any kind of warranty.
' You agree that DevASP is not responsible for any demage 
' may cause with the Use of this GuestBook.
'                                                                            
' Selling the code for this program without prior written consent is expressly forbidden. 
'
' This guestbook is free to use, there is no charge for that. 
' All i ask is to link back to my site "http://www.devasp.com" 
' when you use this guestbook on your site, 
' drop me a mail and let me know where are you using this guestbook.
' Please send bug reports, comments, etc to devasp@aol.com
'------------------------------------------------------------------------------
' To get this guestbook working you only need to change the Constants on the bottom of this page
'
' Enjoy
' Asif
' devasp@aol.com
' http://www.devasp.com
'*************************************************************
%>

<%
' Change the name of your GuestBook here
Const gbName = "The Gregg Shorthand Guestbook"

' Your web site url
Const gbURL = "http://gregg.angelfishy.net/"

' Title of your web site
Const gbTitle = "Gregg Shorthand"

' How many records you want to display on one page
Const strPageSize = 10

' If you want to send an Auto response email to the guest 
' who sign your guestbook then change Sendmail to yes.
' You also need to change the default email message 
' from process.asp page.

Const SendMail = "No"

' Here you can change the username and password for Admin Page 
Const AdminName = "admin"
Const AdminPass = "garfield1"

' I keep the database in same directory
' If you move database in some other directory
' you need to change the path here
' You need to have write permission on the folder 
' where you keep the database or guestbook will not work

SUB OpenDataConnection
	Dim gbConn
	Set gbConn = Server.CreateObject("ADODB.Connection")
	conString = "DBQ=" & Server.MapPath("dbguestbook.mdb")
	gbConn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & conString
End Sub

Function FormatMessage(strMessage)
  Dim strTemp
	strTemp = Server.HTMLEncode(strMessage)
	strTemp = Replace(strTemp, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, vbCrLf, "<BR>" & vbCrLf, 1, -1, 1)
	FormatMessage = strTemp
End Function
%>