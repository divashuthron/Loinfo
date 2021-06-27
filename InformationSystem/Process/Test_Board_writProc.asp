<%@ codepage="65001" language="VBScript" %>
<%
Response.CharSet = "UTF-8"
Dim sessionId		:	sessionId = Session("MemberID")
Response.write sessionId

Dim Dbcon, Rs
Set Dbcon = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Dim strConnect
strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect

Dim Title, content1, INPT_USID, INPT_ADDR

Title = trim(Request.Form("title"))
content1 = trim(Request.Form("content1"))
INPT_ADDR = Request.ServerVariables("REMOTE_ADDR")


Dim StrSql
Dim idx     :   idx = Request("idx")
if idx = "" then
StrSql = ""
StrSql = StrSql & vbCrLf & "INSERT INTO Test_Board(Title, content1, INPT_USID, INPT_DATE,INPT_ADDR)values('"
StrSql = StrSql & vbCrLf & Title &"', '"
StrSql = StrSql & vbCrLf & content1 &"', '"
StrSql = StrSql & vbCrLf & sessionId &"', "
StrSql = StrSql & vbCrLf & "getdate(), '"
StrSql = StrSql & vbCrLf & INPT_ADDR & "')"

else 
StrSql = ""
StrSql = StrSql & vbCrLf & "UPDATE Test_Board "
StrSql = StrSql & vbCrLf & "    SET "
StrSql = StrSql & vbCrLf & "        Title = '"&Title
StrSql = StrSql & vbCrLf & "        ', content1 = '"&content1
StrSql = StrSql & vbCrLf & "        ', UPDT_USID = '"&sessionId
StrSql = StrSql & vbCrLf & "        ', UPDT_DATE = getdate()"
StrSql = StrSql & vbCrLf & "        , UPDT_ADDR = '"&INPT_ADDR
StrSql = StrSql & vbCrLf & "' WHERE idx ="&idx
end if

Dbcon.Execute(StrSql)

Response.Redirect "/Test_Board_list.asp"

Dbcon.close
set Rs = nothing
set Dbcon = nothing
%>