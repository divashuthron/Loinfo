<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<%
Dim Dbcon, Rs
Set Dbcon = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Dim strConnect
strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect


Dim MemberID, StrSql, strResult

'파라미터
MemberID = trim(Request("input_text"))

'Response.write "MemberID: " & MemberID

'쿼리
StrSql = "select count(*) cnt from Test_Board_Member where MemberID = '"&MemberID&"'"

'실행
Rs.Open StrSql, Dbcon

'Response.write "open"

strResult = Rs("cnt")

'Response.write "strResult: " & strResult

Rs.close
Dbcon.close
set Rs = nothing
set Dbcon = nothing


%>
<itemlist>
    <datas><%= strResult %></datas>
</itemlist>
</Metissoft>
