<%@  codepage="65001" language="VBScript" %>
<%
Dim Dbcon
Set Dbcon = Server.CreateObject("ADODB.Connection")
strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect

Dim state, url, INPT_USID, INPT_ADDR
state = Request("state")
USID = session("MemberID")
ADDR = Request.ServerVariables("REMOTE_ADDR")
IDX = Request("idx")

strSql =""
strSql = strSql & vbCrLf & "UPDATE Test_Board "
strSql = strSql & vbCrLf & "    SET"
strSql = strSql & vbCrLf & "        UPDT_USID='"&USID
strSql = strSql & vbCrLf & "       ', UPDT_ADDR='"&ADDR
strSql = strSql & vbCrLf & "       ', UPDT_DATE=getdate()"
strSql = strSql & vbCrLf & "       , UseType='N'"
strSql = strSql & vbCrLf & "WHERE idx = "&IDX

Dbcon.Execute(strSql)

Dbcon.close
set Dbcon = nothing
Response.Redirect "/Test_Board_list.asp"
%>