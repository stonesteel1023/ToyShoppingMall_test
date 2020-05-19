<%@ Language=VBScript %>
<html>

<head>
<title></title>
<style type="text/css">  
 A            {text-decoration: none; color:white }  
 A:hover      {text-decoration: none }  
</style>
</head>

<body bgcolor="#00274f" text="#eefde3">

<table border="0" width="120" height="30" cellpadding="2" cellspacing="0">
  <tr>
    <td align="left" height="80" width="120"><p align="center"><font face="Comic Sans MS"
    color="#ffff00"><strong><big><big><big>&quot;NewMagazine&quot;</big></big></big></strong></font></td>
  </tr>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><strong><a href="Magazine/cart.asp" target="right">&nbsp; <font face="돋움" size="2"
    color="#00274f">구독내역</font></a></strong></td>
  </tr>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><font face="돋움" size="2"
    color="#00274f"><strong>&nbsp; 정기구독</strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="돋움" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#ffff80">&gt;</font><font face="돋움" size="2"><strong> <a
    href="Magazine/list.asp?part=Magazine1" target="right"><font color="#ffffff">잡지</font></a></strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="돋움" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#80ffff">&gt;</font><font face="돋움" size="2"><strong> <a
    href="Magazine/list.asp?part=Magazine2" target="right"><font color="#ffffff">주간지</font></a></strong></font></td>
  </tr>
  <tr>
    <td align="left"
    style="BORDER-BOTTOM: rgb(238,253,227) 1px dashed; BORDER-LEFT: rgb(238,253,227) 1px dashed; BORDER-RIGHT: rgb(238,253,227) 1px dashed"
    height="25" width="120"><font face="돋움" size="2"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </strong></font><font color="#80ff80">&gt;</font> <a href="Magazine/list.asp?part=Magazine3"
    target="right"><font color="#ffffff" face="돋움" size="2"><strong>월간지</strong></font></a></td>
  </tr>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
<% if session("id")<>"admin" then %>
  <tr>
    <td align="left" bgcolor="#eefde3" height="40" width="120"><strong><a
    href="admin/login.asp" target="right"><font face="돋움" size="2" color="#00274f">&nbsp; 
    관리자 페이지</font></a></strong></td>
  </tr>
<% else%>
  <tr>
    <td align="left" height="30" width="120"><strong><font face="돋움" size="2">&nbsp; <a
    href="admin/main.asp" target="right">관리자 구역</a></font></strong></td>
  </tr>
<% end if%>
  <tr>
    <td align="left" height="20" width="120">&nbsp; </td>
  </tr>
</table>
</body>
</html>
