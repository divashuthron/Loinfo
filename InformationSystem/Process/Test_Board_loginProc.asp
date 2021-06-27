<%@  codepage="65001" language="VBScript" %>
<%
Dim Dbcon, Rs
Set Dbcon = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.RecordSet")
Dim strConnect
strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect

Dim MemberID, MemberPW, MemberName, StrSql

MemberID =trim(Request.Form("memberID"))
MemberPW =trim(Request.Form("memberPW"))

StrSql=""
StrSql=StrSql & vbCrLf & "SELECT top 1 "
StrSql=StrSql & vbCrLf & "  MemberID, MemberPWD, MemberName"
StrSql=StrSql & vbCrLf & "FROM "
StrSql=StrSql & vbCrLf & "  Test_Board_Member "
StrSql=StrSql & vbCrLf & "WHERE State = 'Y' "
StrSql=StrSql & vbCrLf & "  AND MemberID='"& MemberID &"'"

Rs.Open StrSql,Dbcon

if(rs.EOF or rs.BOF) then
%>
    <script language="javascript" charset="utf-8">
     alert("No found user id.")
     history.back();
    </script>
<%
else
    if rs("MemberPWD") <> MemberPW then
%>
    <script language="javascript" charset="utf-8">
     alert("Passwords do not match.")
     history.back();
    </script>
<%
    else
    session("MemberID") = (rs("MemberID"))
    session("MemberName") = (rs("MemberName"))

    Rs.close
    set Rs = nothing
    set Dbcon = nothing
    Response.Redirect "/Test_Board_list.asp"
    end if
end if%>