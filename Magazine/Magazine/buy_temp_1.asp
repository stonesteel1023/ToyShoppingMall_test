<%
code = request("code")
ea = request("ea")

'buy_temp = code & "/" & sellprice & "/" & ea

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("Provider=SQLOLEDB; Data Source=sdmiweb; Initial Catalog=Magazine; User ID=n365qt; Password=n365qt;")


   sql = "insert into imsi_buy (imsi_memid,imsi_goodscode,imsi_ea) values "
   sql = sql & "('" & session.SessionID & "'"
   sql = sql & ",'" & code & "'"
   sql = sql & "," & ea & ")"
	
   db.Execute sql
  
 
 Response.Redirect "cart.asp"
%>