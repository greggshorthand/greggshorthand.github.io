<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><xsl:value-of select="title" /></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="766" border="0" cellpadding="0" cellspacing="0" class="text">
  <tr> 
    <td colspan="2" bgcolor="#CCCCCC"><p><img src="images/title3.gif" alt="Gregg Shorthand (Title)" width="760" height="53" align="right" /></p></td>
  </tr>
  <tr> 
    <td height="11" colspan="2" valign="top" bgcolor="#CC6666"><div align="right"> 
        <p align="left" class="text" style="color: #FFFFFF"><img src="images/title2.gif" alt="A Web Site dedicated to the perpetuation of pen stenography." width="243" height="15" align="right" />&nbsp;A 
          Web Site dedicated to the perpetuation of pen stenography</p>
      </div></td>
  </tr>
  <tr> 
    <td width="182" valign="top" bgcolor="#CCCCCC" class="unnamed1"><table width="100%" border="0" cellpadding="2" cellspacing="2" class="text">
        <tr> 
          <td width="174">
            <!--#include file="nav.html" -->
          </td>
        </tr>
      </table></td>
    <td width="585" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="2" cellspacing="2">
		<xsl:for-each select="paragraph">
        <tr> 
          <td colspan="2" class="text"><p><xsl:value-of select="paragraph" /></td>
        </tr>
    	</xsl:for-each>
      </table></td>
  </tr>
  <tr bgcolor="#CC6666"> 
    <td height="6" colspan="2" valign="top"><div align="center" style="color: #FFFFFF">Design 
        Copyright &copy; 2004 Andrew Owen. All Rights Reserved.</div></td>
  </tr>
</table>
</body>
</html>
</xsl:template>
</xsl:stylesheet>