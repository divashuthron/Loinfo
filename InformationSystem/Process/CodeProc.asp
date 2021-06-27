<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process 			: process = fnR("process", "")
Dim ProcessType			: ProcessType = fnR("ProcessType", "")
Dim LogDivision			: LogDivision = "Code"

Dim MasterCode			: MasterCode = fnR("MasterCode", "0")
Dim MasterCodeName		: MasterCodeName = fnRF("MasterCodeName")
Dim MasterCodeState		: MasterCodeState = fnRF("MasterCodeState")

Dim SubMasterCode		: SubMasterCode = fnR("SubMasterCode", "0")
Dim SubCode				: SubCode = fnRF("SubCode")
Dim SubCodeOld			: SubCodeOld = fnRF("SubCodeOld")
Dim SubCodeName			: SubCodeName = fnRF("SubCodeName")
Dim Temp1				: Temp1 = fnRF("Temp1")
Dim Temp2				: Temp2 = fnRF("Temp2")
Dim SubCodeStep			: SubCodeStep = fnRF("Step")
Dim State 				: State = fnRF("State")

Dim RecordCount			: RecordCount = 0
Dim PageCount			: PageCount = 0
Dim PageSize			: PageSize	= 8
Dim PageBlock			: PageBlock	= 5
Dim PageNum				: PageNum	= fnR("page", 1)

Dim strResult : strResult = "failure"
Dim returnMSG

Dim SQL, strWhere, arrParams, strLogMSG

Select Case process
	Case "RegMasterCode"
		Call setMasterCode()
	Case "RegSubCode"
		Call setSubCode()
	Case "getSubCodeList"
		Call getSubCodeList()
End Select

'=============== 마스터 코드 입력 ===============
Sub setMasterCode()
	On Error Resume Next

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()

	if ProcessType = "Insert" then
		'// Insert =================
		SQL = " INSERT INTO CodeMaster ( "
		SQL = SQL & vbCrLf &"		MasterCode, MasterCodeName, State, RegID "
		SQL = SQL & vbCrLf &" ) VALUES ( "
		SQL = SQL & vbCrLf &"		?, ?, ?, ? "
		SQL = SQL & vbCrLf &" ) "
		
		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		arrParams = Array(_
			  Array("@MasterCode",			adVarchar,			adParamInput, 50,			MasterCode) _
			, Array("@MasterCodeName",		adVarchar,			adParamInput, 255,			StringToSQL(MasterCodeName)) _
			, Array("@State",				adChar,				adParamInput, 1,			MasterCodeState) _
			, Array("@RegID",				adVarchar,			adParamInput, 25,			SessionUserID) _
		)

		'objDB.blnDebug = true
		objDB.sbExecSQL SQL, arrParams
		
		strLogMSG = "코드관리  > 마스터 코드 ["& MasterCode &"]이/가 생성 되었습니다."
		InsertType = "Insert"
	else
		'// Update ================
		SQL = " UPDATE CodeMaster SET "
		SQL = SQL & vbCrLf &"		MasterCode = ?, MasterCodeName = ?, State = ?, EditID = ?, EditDate = getdate() "
		SQL = SQL & vbCrLfL & " WHERE MasterCode = ?; "
		
		arrParams = Array(_
			  Array("@MasterCode",			adVarchar,			adParamInput, 50,			MasterCode) _
			, Array("@MasterCodeName",		adVarchar,			adParamInput, 255,			StringToSQL(MasterCodeName)) _
			, Array("@State",				adChar,				adParamInput, 1,			MasterCodeState) _
			, Array("@EditID",				adVarchar,			adParamInput, 25,			SessionUserID) _
			, Array("@MasterCode",			adVarchar,			adParamInput, 50,			MasterCode) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "코드관리  > 마스터 코드 ["& MasterCode &"]이/가 수정 되었습니다."
		InsertType = "Update"
	end if
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "마스터 코드 저장 완료"
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
		<ReturnMSG><![CDATA[<%= returnMSG %>]]></ReturnMSG>
	</List>
</Lists>
<%
End Sub

'=============== 서브 코드 입력 ===============
Sub setSubCode()
	On Error Resume Next

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()

	if ProcessType = "Insert" then
		'// Insert =================
		SQL = " INSERT INTO CodeSub ( "
		SQL = SQL &"		MasterCode, SubCode, SubCodeName, Step, Temp1, Temp2, UseYN, State, RegID "
		SQL = SQL &" ) VALUES ( "
		SQL = SQL &"		?, ?, ?, ?, ?, ?, 'Y', ?, ? "
		SQL = SQL &" ) "
		
		'adDate, adLongVarChar, adVarchar, adInteger, adChar
		
		arrParams = Array(_
			  Array("@MasterCode",			adVarchar,			adParamInput, 50,			SubMasterCode) _
			, Array("@SubCode",				adVarchar,			adParamInput, 25,			SubCode) _
			, Array("@SubCodeName",			adVarchar,			adParamInput, 255,			SubCodeName) _
			, Array("@Step",				adInteger,			adParamInput, 0,			SubCodeStep) _
			, Array("@Temp1",				adVarchar,			adParamInput, 255,			Temp1) _
			, Array("@Temp2",				adVarchar,			adParamInput, 255,			Temp2) _
			, Array("@State",				adChar,				adParamInput, 1,			State) _
			, Array("@RegID",				adVarchar,			adParamInput, 25,			SessionUserID) _
		)

		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "코드관리  > 마스터 코드 ["& SubMasterCode &"] -> 서브 코드 ["& SubCode &"]이/가 생성 되었습니다."
		InsertType = "Insert"
	else
		'// Update ================
		SQL = " UPDATE CodeSub SET "
		SQL = SQL &"		SubCode = ?, SubCodeName = ?, Step = ?, Temp1 = ?, Temp2 = ?, "
		SQL = SQL &"		State = ?, EditID = ?, EditDate = getdate() "
		SQL = SQL & " WHERE MasterCode = ? AND SubCode = ?; "
		
		arrParams = Array(_
			  Array("@SubCode",				adVarchar,			adParamInput, 25,			SubCode) _
			, Array("@SubCodeName",			adVarchar,			adParamInput, 255,			SubCodeName) _
			, Array("@Step",				adInteger,			adParamInput, 0,			SubCodeStep) _
			, Array("@Temp1",				adVarchar,			adParamInput, 255,			Temp1) _
			, Array("@Temp2",				adVarchar,			adParamInput, 255,			Temp2) _
			, Array("@State",				adChar,				adParamInput, 1,			State) _
			, Array("@EditID",				adVarchar,			adParamInput, 25,			SessionUserID) _
			, Array("@MasterCode",			adVarchar,			adParamInput, 50,			SubMasterCode) _
			, Array("@SubCode",				adVarchar,			adParamInput, 25,			SubCodeOld) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "코드관리  > 마스터 코드 ["& SubMasterCode &"] -> 서브 코드 ["& SubCode &"]이/가 수정 되었습니다."
		InsertType = "Update"
	end if
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "서브 코드  저장 완료"
		objDB.sbCommitTrans 
	End If	

	Set objCLS = Nothing
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

'=============== 서브 코드 리스트 ===============
Sub getSubCodeList()
	On Error Resume Next

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	'objDB.sbBeginTrans()


	SQL = " SELECT Count(*) "
	SQL = SQL & " FROM CodeSub AS A "
	SQL = SQL & " WHERE 1 = 1 "
	SQL = SQL & "		AND MasterCode = ? "
	
	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
	
	'objDB.blnDebug = True
	arrParams = objDB.fnGetArray			

	RecordCount = objDB.fnExecSQLGetRows(SQL, arrParams)(0,0)
	PageCount = int((RecordCount - 1) / PageSize) + 1
	intNUM = RecordCount - (PageSize * (PageNum - 1))

	SQL = " SELECT * FROM "
	SQL = SQL & "	( "
	SQL = SQL & "		SELECT "
	SQL = SQL & "			SubCode, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc "
	SQL = SQL & "			, UseYN, State "
	SQL = SQL &"			, (CASE  State "
	SQL = SQL &"				WHEN 'Y' THEN '사용' "
	SQL = SQL &"				WHEN 'N' THEN '미사용' "
	SQL = SQL &"			END) AS StateName "
	SQL = SQL & "			, RegDate, RegID, EditDate, EditID "
	SQL = SQL & "			, ROW_NUMBER() OVER (ORDER BY Step DESC) AS ROWNUM "
	SQL = SQL & "		FROM CodeSub AS A " 
	SQL = SQL & "		WHERE 1 = 1 "
	SQL = SQL & "			AND MasterCode = ? "
	SQL = SQL & "	) AS TBL_PAGELIST "
	SQL = SQL & "	WHERE ROWNUM BETWEEN "& (PageNum - 1) * PageSize + 1 &" AND "& PageNum * PageSize &";"
	'SQL = SQL & "ORDER BY IDX DESC;"
	
	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
	
	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		'objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "서브 코드 가져오기 완료"
		'objDB.sbCommitTrans 
	End If	

	Set objCLS = Nothing
	Set objDB  = Nothing
%>
<Lists>
	<Result><%= strResult %></Result>
	<ReturnMSG><![CDATA[<%= returnMSG %>]]></ReturnMSG>
	<PageInfo>
		<PageNum><%= PageNum %></PageNum>
		<PageSize><%= PageSize %></PageSize>
		<PageBlock><%= PageBlock %></PageBlock>
		<RecordCount><%= RecordCount %></RecordCount>
		<PageCount><%= PageCount %></PageCount>
	</PageInfo>
<%
	if IsArray(aryList) then
		for i = 0 to UBound(aryList, 2)
		' SubCode, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc,		0~7
		' UseYN, State, RegDate, RegID, EditDate, EditID 											8~13
%>
	<item>
		<Num><%= intNUM %></Num>
		<SubCode><![CDATA[<%= aryList(0, i) %>]]></SubCode>
		<SubCodeName><![CDATA[<%= StringToSQL(aryList(1, i)) %>]]></SubCodeName>
		<Step><%= aryList(2, i) %></Step>
		<Temp1><![CDATA[<%= aryList(3, i) %>]]></Temp1>
		<Temp2><![CDATA[<%= aryList(4, i) %>]]></Temp2>
		<Temp3><![CDATA[<%= aryList(5, i) %>]]></Temp3>
		<Temp4><![CDATA[<%= aryList(6, i) %>]]></Temp4>
		<TempEtc><![CDATA[<%= aryList(7, i) %>]]></TempEtc>
		<UseYN><![CDATA[<%= aryList(8, i) %>]]></UseYN>
		<State><%= aryList(9, i) %></State>
		<StateName><![CDATA[<%= aryList(10, i) %>]]></StateName>
		<RegDate><![CDATA[<%= aryList(11, i) %>]]></RegDate>
		<RegID><![CDATA[<%= aryList(12, i) %>]]></RegID>
		<EditDate><![CDATA[<%= aryList(13, i) %>]]></EditDate>
		<EditID><![CDATA[<%= aryList(14, i) %>]]></EditID>
	</item>
<%
		intNUM = intNUM - 1
		
		next
	end if
%>
</Lists>
<%
End Sub
%>
</Metissoft>