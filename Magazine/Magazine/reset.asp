<%

 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("Provider=SQLOLEDB; Data Source=sdmiweb; Initial Catalog=Magazine; User ID=n365qt; Password=n365qt;")
	
 sql = "delete * from imsi_buy where imsi_memid ='" & session.SessionID & "'"
 
 db.Execute sql
 
 Response.Redirect "cart.asp"
 %>
	