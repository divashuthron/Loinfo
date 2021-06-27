<?xml version="1.0" encoding="utf-8"?>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere, strLogMSG

Dim LogDivision	: LogDivision = "Login"

Dim strResult : strResult = "false"
Dim UserID : UserID = fnR("UserID", "")
Dim Passwd : Passwd = fnR("Passwd", "")

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, ISNULL(PhoneNumber, '') as PhoneNumber, "
SQL = SQL & vbCrLf & "	ISNULL(Email, '') as Email, ISNULL(JoinDate, '') as JoinDate, ISNULL(OutDate, '') as OutDate, ISNULL(EmpInfo, '') as EmpInfo, "
SQL = SQL & vbCrLf & "	State, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName, "
SQL = SQL & vbCrLf & "	RegDate, RegID, EditDate, EditID "
SQL = SQL & vbCrLf & "From Employee AS A "
SQL = SQL & vbCrLf & "Where 1 = 1 "
'SQL = SQL & vbCrLf & "	AND (ClientLevel = 'Admin' or ClientLevel = 'SchoolAdmin') "	'관리자 / 학교관리자만 입장 가능
SQL = SQL & vbCrLf & "	AND State = 'Y' "	'상태 : 사용
SQL = SQL & vbCrLf & "	AND EmpID = ?; "

Call objDB.sbSetArray("@EmpID", adVarchar, adParamInput, 25, UserID)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'If Not IsNull(AryHash) Then
If IsArray(AryHash) then
	' EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, PhoneNumber
	' Email, JoinDate, OutDate, EmpInfo, State
	' StateName, RegDate, RegID, EditDate, EditID, ClientName
	
	if Passwd = AryHash(0).Item("EmpPWD") then
		
		'// 기본설정 불러오기
		SQL = ""
		SQL = SQL & vbCrLf & "SELECT top 1"
		SQL = SQL & vbCrLf & "	Idx, MYear, Division, Subject, ISNULL(Division1, '') AS Division1, ISNULL(Division2, '') AS Division2 "
		SQL = SQL & vbCrLf & "	, ISNULL(SchoolName, '') AS SchoolName, SchoolAddress, ISNULL(SchoolSmsNumber, '') AS SchoolSmsNumber, SchoolTelNumber "
		SQL = SQL & vbCrLf & "	, ApplyConfirm, ApplyPrintConfirm, InterviewConfirm "
		SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division', Division) AS DivisionName "
		SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', Subject) AS SubjectName "
		SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', Division1) AS Division1Name "
		'SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', Division2) AS Division2Name "
		SQL = SQL & vbCrLf & "	, '' AS Division2Name "
		SQL = SQL & vbCrLf & "	, State, (CASE  State "
		SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
		SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
		SQL = SQL & vbCrLf & "	END) AS StateName "
		SQL = SQL & vbCrLf & "	, RegDate, RegID "
		SQL = SQL & vbCrLf & "FROM ConfigTable AS A " 
		SQL = SQL & vbCrLf & "WHERE 1 = 1 "
		SQL = SQL & vbCrLf & "	AND State = 'Y' "
		SQL = SQL & vbCrLf & "ORDER BY IDX DESC; "

		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, Nothing)

		'Response.Cookies("InformationAdmin").Domain = "info.metissoft.co.kr"
		'// 사용자 정보
		'// 쿠기 위조 가능하기 때문에 ID는 세션에 입력
		'Response.Cookies("InformationAdmin")("EmpID") = base64_encode(AryHash(0).Item("EmpID"))
		Session("EmpID") = base64_encode(AryHash(0).Item("EmpID"))
		Response.Cookies("InformationAdmin")("EmpName") = AryHash(0).Item("EmpName")
		Response.Cookies("InformationAdmin")("ClientLevel") = AryHash(0).Item("ClientLevel")
		'// 환경설정 정보
		If Not IsNull(AryHash2) then
			Response.Cookies("InformationAdmin")("MYear") = AryHash2(0).Item("MYear")
			Response.Cookies("InformationAdmin")("Division") = AryHash2(0).Item("Division")
			'Response.Cookies("InformationAdmin")("Subject") = AryHash2(0).Item("Subject")
			'Response.Cookies("InformationAdmin")("Division1") = AryHash2(0).Item("Division1")
			'Response.Cookies("InformationAdmin")("Division2") = AryHash2(0).Item("Division2")
			'Response.Cookies("InformationAdmin")("SchoolName") = AryHash2(0).Item("SchoolName")
			'Response.Cookies("InformationAdmin")("SchoolSmsNumber") = AryHash2(0).Item("SchoolSmsNumber")
			'Response.Cookies("InformationAdmin")("ApplyConfirm") = AryHash2(0).Item("ApplyConfirm")
			'Response.Cookies("InformationAdmin")("ApplyPrintConfirm") = AryHash2(0).Item("ApplyPrintConfirm")
			'Response.Cookies("InformationAdmin")("InterviewConfirm") = AryHash2(0).Item("InterviewConfirm")
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
Call ActivityHistory(strLogMSG, LogDivision, UserID)
%>

<Metissoft>
	<Lists>
		<List>
			<Result><%= strResult %></Result>
		</List>
	</Lists>
</Metissoft>