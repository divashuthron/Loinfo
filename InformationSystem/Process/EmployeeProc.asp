<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim LogDivision		: LogDivision = "EmployeeProc"

Dim process			: process = fnR("process", "")
Dim ProcessType		: ProcessType = fnR("ProcessType", "")
Dim IDX				: IDX = fnR("IDX", 0)
Dim EmpID			: EmpID = fnRF("EmpID")
Dim ClientCode		: ClientCode = fnRF("ClientCode")
Dim ClientLevel		: ClientLevel = fnRF("ClientLevel")
Dim ClientLevel_OLD	: ClientLevel_OLD = fnRF("ClientLevel_OLD")
Dim EmpPWD			: EmpPWD = fnRF("EmpPWD")
Dim EmpName			: EmpName = fnRF("EmpName")
Dim PhoneNumber		: PhoneNumber = fnRF("PhoneNumber")
Dim Email			: Email = fnRF("Email")
Dim JoinDate		: JoinDate = fnRF("JoinDate")
Dim OutDate			: OutDate = fnRF("OutDate")
Dim EmpInfo			: EmpInfo = fnRF("EmpInfo")
Dim State			: State = fnRF("State")

Dim strResult		: strResult = "failure"
Dim returnMSG

Select Case process
	Case "CheckID"
		Call getIDCheckProc()
	Case "RegEmployee"
		Call setEmployee()
	Case ""
End Select

'============= 중복 아이디 검사 ===========
Sub getIDCheckProc()
	On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
	
	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB

	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "		EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, ISNULL(PhoneNumber, ''), "
	SQL = SQL & vbCrLf & "		ISNULL(Email, ''), ISNULL(JoinDate, ''), ISNULL(OutDate, ''), ISNULL(EmpInfo, ''), State, "
	SQL = SQL & vbCrLf & "		(CASE  State "
	SQL = SQL & vbCrLf & "			WHEN 'Y' THEN '사용' "
	SQL = SQL & vbCrLf & "			WHEN 'N' THEN '미사용' "
	SQL = SQL & vbCrLf & "		END) AS StateName, "
	SQL = SQL & vbCrLf & "		RegDate, RegID, EditDate, EditID "
	SQL = SQL & vbCrLf & "	From Employee AS A "
	SQL = SQL & vbCrLf & "	Where 1 = 1 "
	SQL = SQL & vbCrLf & "		AND EmpID = ?; "

	Call objDB.sbSetArray("@EmpID", adVarchar, adParamInput, 25, EmpID)

	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	aryList = objDB.fnExecSQLGetRows(SQL, arrParams)

	Set objDB = Nothing
	
	if IsArray(aryList) Then
		strRESULT = "true"
	else
		strRESULT = "false"
	end if
%>
<Lists>
	<List>
		<Result><%= strRESULT %></Result>
	</List>
</Lists>
<%
End Sub

'=============== 사용자 입력 ===============
Sub setEmployee()
	On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()
		
	if ProcessType = "EmployeeInsert" then
		'// Insert =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO Employee ( "
		SQL = SQL & vbCrLf & "		EmpID, ClientLevel, EmpPWD, EmpName, PhoneNumber,  "
		SQL = SQL & vbCrLf & "		Email, EmpInfo, State, RegID "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & " ) "

		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		arrParams = Array(_
			  Array("@EmpID",			adVarchar,			adParamInput, 25,			EmpID) _
			, Array("@ClientLevel",		adVarchar,			adParamInput, 50,			ClientLevel) _
			, Array("@EmpPWD",			adVarchar,			adParamInput, 25,			EmpPWD) _
			, Array("@EmpName",			adVarchar,			adParamInput, 25,			EmpName) _
			, Array("@PhoneNumber",		adVarchar,			adParamInput, 20,			PhoneNumber) _
			, Array("@Email",			adVarchar,			adParamInput, 255,			Email) _
			, Array("@EmpInfo",			adLongVarChar,		adParamInput, 1000,			StringToSQL(EmpInfo)) _
			, Array("@State",			adChar,				adParamInput, 1,			State) _
			, Array("@RegID",			adVarchar,			adParamInput, 25,			SessionUserID) _
		)

		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "사용자관리  > "& EmpID &" (ClientLevel : "& ClientLevel &")가 생성 되었습니다."
		InsertType = "Insert"
	else
		'// Update ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE Employee SET "
		SQL = SQL & vbCrLf & "		ClientLevel = ?, EmpPWD = ?, EmpName = ?, PhoneNumber = ?, Email = ?, "
		SQL = SQL & vbCrLf & "		EmpInfo = ?, State = ?, EditID = ?, EditDate = getdate() "
		SQL = SQL & vbCrLf & " WHERE EmpID = ?; "
		
		arrParams = Array(_
			  Array("@ClientLevel",		adVarchar,			adParamInput, 50,				ClientLevel) _
			, Array("@EmpPWD",			adVarchar,			adParamInput, 25,				EmpPWD) _
			, Array("@EmpName",			adVarchar,			adParamInput, 25,				EmpName) _
			, Array("@PhoneNumber",		adVarchar,			adParamInput, 20,				PhoneNumber) _
			, Array("@Email",			adVarchar,			adParamInput, 255,				Email) _
			, Array("@EmpInfo",			adLongVarChar,		adParamInput, 1000,				StringToSQL(EmpInfo)) _
			, Array("@State",			adChar,				adParamInput, 1,				State) _
			, Array("@EditID",			adVarchar,			adParamInput, 25,				SessionUserID) _
			, Array("@EmpID",			adVarchar,			adParamInput, 25,				EmpID) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "사용자관리  > "& SessionUserID & "이/가 " & EmpID &"의 정보를 수정 했습니다."
		If ClientLevel <> ClientLevel_OLD Then
			strLogMSG = strLogMSG & "(ClientLevel : "& ClientLevel_OLD &" -> "& ClientLevel &")"
		End if
		InsertType = "Update"
	end if
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "사용자 저장 완료"
		objDB.sbCommitTrans 
	End If	

	Set objDB  = Nothing

	'// 로그기록
	Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)
%>
<Lists>
	<List>
		<Result><%= strResult %></Result>
		<InsertType><%= InsertType %></InsertType>
		<ReturnMSG><%= returnMSG %></ReturnMSG>
	</List>
</Lists>
<%
End Sub
%>
</Metissoft>