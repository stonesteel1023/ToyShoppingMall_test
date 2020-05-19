<%'@TRANSACTION="Required"%>
<!--#include file="../pkval.asp"//-->
<%
 sum = request("sum")
 name = request("name")
 address = request("address")
 sex = request("sex")
 tel = request("tel")
 bank = request("bank")
 dd = now()
  
 Set db = Server.CreateObject("ADODB.Connection")
 db.Open("Provider=SQLOLEDB; Data Source=sdmiweb; Initial Catalog=Magazine; User ID=n365qt; Password=n365qt;")
	
 sql = "select imsi_ea, g_code, g_name, g_sellprice from imsi_buy A, goods B  where A.imsi_memid ='" & session.SessionID & "'"
 sql = sql & " and A.imsi_goodscode = B.g_code"
 
 Set rs = Server.CreateObject("ADODB.RecordSet")
 rs.Open sql, db
 
 Do until rs.EOF
   price = rs("g_sellprice")
   ea = rs("imsi_ea")
   code = rs("g_code")
   g_name = rs("g_name")
   
   str = g_name & "(" & code & ") : " & ea & "개" & chr(13) & chr(10)
   summary = summary & str
   
   rs.MoveNext   
 Loop 
 
 rs.Close
 
 sql = "insert into customer (c_code,c_name,c_tel,c_sex,c_address) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & name & "'"
 sql = sql & ",'" & tel & "'"
 sql = sql & ",'" & sex & "'"
 sql = sql & ",'" & address & "')"
 db.Execute sql
 
 sql = "insert into buy (b_code,c_code,b_date, b_summary,b_totalprice,b_bank) values "
 sql = sql & "('" & primaryval & name & "'"
 sql = sql & ",'" & primaryval & name & "'"
 sql = sql & ",'" & dd & "'" 
 sql = sql & ",'" & summary & "'"
 sql = sql & "," & sum 
 sql = sql & ",'" & bank & "')" 
 db.Execute sql
 

 sql = "delete from imsi_buy where imsi_memid='" & session.SessionID & "'"
 
 db.Execute sql
 
Sub OnTransactionCommit()

 Response.Redirect "result.htm"

End Sub

Sub OnTransactionAbort()

    response.write  "입력에 문제발생... 다시 입력하세요"
%> 
<script language="javascript">
alert("오류가 발생해서 모든 작업을 취소했습니다.\n\n다시 시도해 주십시요\n\n다시 문제가 생기면 관리자에게 연락을 주시기 바랍니다");
history.back();
</script>
<%
End Sub
%>