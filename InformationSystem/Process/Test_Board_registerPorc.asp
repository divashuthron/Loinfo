<%@  codepage="65001" language="VBScript" %>
<%

Dim Dbcon, Rs
Set Dbcon = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Dim strConnect
strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect

Dim MemberID, MemberPW, MemberName, StrSql

memberName =trim(Request("memberName"))
memberID = trim(Request("memberID"))
memberPW = trim(Request("memberPW"))

StrSql="begin tran "
StrSql=StrSql & "insert into Test_Board_Member(memberName, MemberID, memberPWD) values('"
StrSql=StrSql & MemberName & "','"
StrSql=StrSql & MemberID & "','"
StrSql=StrSql & MemberPW & "') "
StrSql=StrSql & "commit tran"

Dbcon.Execute(StrSql)
msg = "가입이 완료되었습니다."
Response.Redirect "/Test_Board_login.asp"

Dbcon.close
set Rs = nothing
set Dbcon = nothing
%>