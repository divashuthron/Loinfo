<?xml version="1.0" encoding="utf-8"?>
<!--#InClude Virtual = "/Lostark/Include/Function.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere, strLogMSG

Dim LogDivision	: LogDivision = "Login"

Dim strResult : strResult = "false"
Dim UserID : UserID = fnR("ID", "")
Dim Passwd : Passwd = fnR("Password", "")
Dim Save : Save = fnR("Save", "")

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & VbCrLf & "Select"
SQL = SQL & VbCrLf & "	ID, Password"
SQL = SQL & VbCrLf & "From Member"
SQL = SQL & VbCrLf & "Where 1=1"
SQL = SQL & VbCrLf & "And State = 'Y'"
SQL = SQL & VbCrLf & "And OutDate Is Null"
SQL = SQL & VbCrLf & "And ID = '" & UserID & "'"
SQL = SQL & VbCrLf & "And Password = '" & Passwd & "'"

'Call objDB.sbSetArray("@ID", adVarchar, adParamInput, 25, UserID)
'Call objDB.sbSetArray("@Password", adVarchar, adParamInput, 25, Passwd)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'If Not IsNull(AryHash) Then
If IsArray(AryHash) then
	if Passwd = AryHash(0).Item("Password") then
		
		'// 기본설정 불러오기
		SQL = ""
        SQL = SQL & VbCrLf & "Select"
        SQL = SQL & VbCrLf & "	ClientCode, ClientLevel, NickName, JoinDate, LastDate, LastIP"
        SQL = SQL & VbCrLf & "From Member"
        SQL = SQL & VbCrLf & "Where 1=1"
        SQL = SQL & VbCrLf & "And State = 'Y'"
        SQL = SQL & VbCrLf & "And OutDate Is Null"
        SQL = SQL & VbCrLf & "And ID = '" & UserID & "'"
		SQL = SQL & VbCrLf & "And Password = '" & Passwd & "'"

        'Call objDB.sbSetArray("@ID", adVarchar, adParamInput, 25, UserID)
        'Call objDB.sbSetArray("@Password", adVarchar, adParamInput, 25, Passwd)

		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, Nothing)

		'// 사용자 정보
		Session("ID") = (AryHash(0).Item("ID"))

        Session("ClientCode") = (AryHash2(0).Item("ClientCode"))
        Session("ClientLevel") = (AryHash2(0).Item("ClientLevel"))
        Session("NickName") = (AryHash2(0).Item("NickName"))
        Session("JoinDate") = (AryHash2(0).Item("JoinDate"))
        Session("LastDate") = (AryHash2(0).Item("LastDate"))
        Session("LastIP") = (AryHash2(0).Item("LastIP"))
        
		'Response.Cookies("LoaRoom")("ClientCode") = AryHash(0).Item("ClientCode")
		'Response.Cookies("LoaRoom")("ClientLevel") = AryHash(0).Item("ClientLevel")

		'// 환경설정 정보
		If Not IsNull(AryHash2) then
			'Response.Cookies("InformationAdmin")("MYear") = AryHash2(0).Item("MYear")
		End if

		strResult = "true"

		strLogMSG = "로그인 > "& UserID &"가 로그인 하였습니다."
	Else
		strLogMSG = "로그인 > "& UserID &"가 로그인 하였지만 비밀번호가 틀렸습니다."
	end If
Else
	strLogMSG = "로그인 > "& UserID &"로 로그인 시도가 있었지만 등록된 ID가 아니거나 허용되지 않은 ID 입니다."
end If

Set objDB	= Nothing

'// 로그기록
'Call ActivityHistory(strLogMSG, LogDivision, UserID)
%>

<Loaroom>
	<Lists>
		<List>
			<Result><%= strResult %></Result>
		</List>
	</Lists>
</Loaroom>