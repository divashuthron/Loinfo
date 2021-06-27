<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process'					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "DefaultConfig"

Dim IDX						: IDX = fnR("IDX", 0)
Dim MYear					: MYear = fnRF("MYear")
Dim Division				: Division = fnRF("Division0")
Dim Subject					: Subject = fnRF("Subject")
Dim Division1				: Division1 = fnRF("Division1")
Dim Division2				: Division2 = fnRF("Division2")

'// 기본정보
Dim SchoolName				: SchoolName = fnRF("SchoolName")
Dim SchoolAddress			: SchoolAddress = fnRF("SchoolAddress")
Dim SchoolSmsNumber			: SchoolSmsNumber = fnRF("SchoolSmsNumber")
Dim SchoolTelNumber			: SchoolTelNumber = fnRF("SchoolTelNumber")
Dim ApplyConfirm			: ApplyConfirm = fnRF("ApplyConfirm")
Dim ApplyPrintConfirm		: ApplyPrintConfirm = fnRF("ApplyPrintConfirm")

Dim BillConfirm				: BillConfirm = fnRF("BillConfirm")
Dim DemandsConfirm			: DemandsConfirm = fnRF("DemandsConfirm")
Dim ApplicationAddConfirm	: ApplicationAddConfirm = fnRF("ApplicationAddConfirm")
Dim ApplicantAddConfirm		: ApplicantAddConfirm = fnRF("ApplicantAddConfirm")
Dim InterviewConfirm		: InterviewConfirm = fnRF("InterviewConfirm")

'// T 사이즈
Dim TSize, PrepareQuantity, ApplyQuantity, ReceiveState
Dim LossQuantity, RemainderQuantity, ETC

'// 평가요소
Dim ItemOrder,ItemName, ItemRate
Dim ItemGrade_01, ItemGrade_02, ItemGrade_03, ItemGrade_04, ItemGrade_05
Dim ItemGrade_06, ItemGrade_07, ItemGrade_08, ItemGrade_09, ItemGrade_10
Dim ItemPoint_01, ItemPoint_02, ItemPoint_03, ItemPoint_04, ItemPoint_05
Dim ItemPoint_06, ItemPoint_07, ItemPoint_08, ItemPoint_09, ItemPoint_10
Dim InterviewGroundsCHK

Dim State					: State = fnRF("State")
Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

'=====================================
'// 파라미터값 자동으로 받아오기
'=====================================
Dim sElement
'// Form 파라미터값 받아오기
For Each sElement in Request.Form
	'// execute에서 변수 선언해도 되고 ...
	'// (기존 같은 이름의 변수가 이미 선언되어 있어도 에러 안남)
	'execute("Dim "& sElement &" : "& sElement &" = """& fnRF(sElement) &"""")
	'// 변수는 미리 선언해 놓고 변수에 값만 넣어도 됨
	execute(sElement &" = """& fnRF(sElement) &"""")
Next
'// QueryString 파라미터값 받아오기
For Each sElement in Request.QueryString
	execute(sElement &" = """& fnRQ(sElement) &"""")
Next
'=====================================

Select Case process
	Case "RegDefaultConfig"
		Call setRegDefaultConfig()
	Case "ReRegDefaultConfig"
		Call setReRegDefaultConfig()
	Case "RegTsizeConfig"
		Call setTsizeConfig()
	Case "RegInterviewItem"
		Call setInterviewItem()
	Case ""
End Select

'=============== 기본설정 입력 ===============
Sub setRegDefaultConfig()
	'On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()

	'// 입력 =================
	SQL = "UPDATE ConfigTable SET State = 'N'; "	'// 기존 입력 되어있던 환경설정 미사용으로 전환
	SQL = SQL & vbCrLf & "INSERT INTO ConfigTable ( "
	SQL = SQL & vbCrLf & "		MYear, Division, Subject, Division1, Division2 "
	SQL = SQL & vbCrLf & "		, SchoolName, SchoolAddress, SchoolSmsNumber, SchoolTelNumber "
	SQL = SQL & vbCrLf & "		, ApplyConfirm, ApplyPrintConfirm, InterviewConfirm "
	SQL = SQL & vbCrLf & "		, BillConfirm, DemandsConfirm, ApplicationAddConfirm, ApplicantAddConfirm "
	SQL = SQL & vbCrLf & "		, State, RegID "
	SQL = SQL & vbCrLf & " ) VALUES ( "
	SQL = SQL & vbCrLf & "		?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ? "
	SQL = SQL & vbCrLf & " ); "
	
	'adDate, adLongVarChar, adVarchar, adInteger, adChar

	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
		, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
		, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
		, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
		, Array("@SchoolName",				adVarchar,		adParamInput,		50,		SchoolName) _
		, Array("@SchoolAddress",			adVarchar,		adParamInput,		255,	SchoolAddress) _
		, Array("@SchoolSmsNumber",			adVarchar,		adParamInput,		50,		SchoolSmsNumber) _
		, Array("@SchoolTelNumber",			adVarchar,		adParamInput,		50,		SchoolTelNumber) _
		, Array("@ApplyConfirm",			adChar,			adParamInput,		1,		ApplyConfirm) _
		, Array("@ApplyPrintConfirm",		adChar,			adParamInput,		1,		ApplyPrintConfirm) _
		, Array("@InterviewConfirm",		adChar,			adParamInput,		1,		InterviewConfirm) _

		, Array("@BillConfirm",				adChar,			adParamInput,		1,		BillConfirm) _
		, Array("@DemandsConfirm",			adChar,			adParamInput,		1,		DemandsConfirm) _
		, Array("@ApplicationAddConfirm",	adChar,			adParamInput,		1,		ApplicationAddConfirm) _
		, Array("@ApplicantAddConfirm",		adChar,			adParamInput,		1,		ApplicantAddConfirm) _

		, Array("@State",					adChar,			adParamInput,		1,		"Y") _
		, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
	)

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, arrParams)

	'SQL = " SELECT @@IDENTITY; "
	'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
	'IDX = CInt(aryList(0, 0))
	
	strLogMSG = "기본환경설정  > " & SessionUserID & "이/가 기본환경설정이 수정 했습니다."
	InsertType = "Insert"
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "기본환경설정 저장 완료"
		objDB.sbCommitTrans

		'// 기본 설정값 설정
		Call setDefaultConfigValue(MYear, Division, Subject, Division1, Division2)
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

'// 기본 설정값 재설정
Sub setReRegDefaultConfig()
	'On Error Resume Next
	
	'// 기본 설정값 설정
	Call setDefaultConfigValue(MYear, Division, Subject, Division1, Division2)

	strResult = "Complete"
	returnMSG = "기본환경설정 변경 완료"
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

'// 기본 설정값 설정
Sub setDefaultConfigValue(MYear, Division, Subject, Division1, Division2)
	'On Error Resume Next

	'// 기본 설정값 설정
	Response.Cookies("InterviewAdmin")("MYear") = MYear
	Response.Cookies("InterviewAdmin")("Division") = Division
	Response.Cookies("InterviewAdmin")("Subject") = Subject
	Response.Cookies("InterviewAdmin")("Division1") = Division1
	Response.Cookies("InterviewAdmin")("Division2") = Division2

	strResult = "Complete"
	returnMSG = "기본환경설정 변경 완료"
End Sub


'// T셔츠관리
Sub setTsizeConfig()
	'On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()
	
	If ProcessType = "Insert" Then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO ConfigTableTsize ( "
		SQL = SQL & vbCrLf & "		MYear, Division, Subject, Division1, Division2 "
		SQL = SQL & vbCrLf & "		, TSize, PrepareQuantity, ApplyQuantity, ReceiveState, LossQuantity, RemainderQuantity, ETC "
		SQL = SQL & vbCrLf & "		, State, RegID "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ? "
		SQL = SQL & vbCrLf & " ) "

		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@TSize",					adVarchar,		adParamInput,		20,		TSize) _
			, Array("@PrepareQuantity",			adInteger,		adParamInput,		0,		PrepareQuantity) _
			, Array("@ApplyQuantity",			adInteger,		adParamInput,		0,		ApplyQuantity) _
			, Array("@ReceiveState",				adInteger,		adParamInput,		0,		ReceiveState) _
			, Array("@LossQuantity",				adInteger,		adParamInput,		0,		LossQuantity) _
			, Array("@RemainderQuantity",		adInteger,		adParamInput,		0,		RemainderQuantity) _
			, Array("@ETC",						adVarchar,		adParamInput,		255,	ETC) _
			, Array("@State",					adChar,			adParamInput,		1,		"Y") _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		SQL = " SELECT @@IDENTITY; "
		aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		IDX = CInt(aryList(0, 0))
		
		strLogMSG = "T셔츠관리  > "& MYear &" "& Division &" "& Subject &" "& IDX &" "& TSize &" 입력 되었습니다."
		InsertType = "Insert"
	ElseIf ProcessType = "Update" Then
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE ConfigTableTsize SET "
		SQL = SQL & vbCrLf & "		MYear = ?, Division = ?, Subject = ?, Division1 = ?, Division2 = ? "
		SQL = SQL & vbCrLf & "		, TSize = ?, PrepareQuantity = ?, ApplyQuantity = ?, ReceiveState = ? "
		SQL = SQL & vbCrLf & "		, LossQuantity = ?, RemainderQuantity = ?, ETC = ? "
		SQL = SQL & vbCrLf & "		, State = ?, EditID = ?,  EditDate = getdate() "
		SQL = SQL & vbCrLf & " WHERE IDX = ?; "
		
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@TSize",					adVarchar,		adParamInput,		20,		TSize) _
			, Array("@PrepareQuantity",			adInteger,		adParamInput,		0,		PrepareQuantity) _
			, Array("@ApplyQuantity",			adInteger,		adParamInput,		0,		ApplyQuantity) _
			, Array("@ReceiveState",				adInteger,		adParamInput,		0,		ReceiveState) _
			, Array("@LossQuantity",				adInteger,		adParamInput,		0,		LossQuantity) _
			, Array("@RemainderQuantity",		adInteger,		adParamInput,		0,		RemainderQuantity) _
			, Array("@ETC",						adVarchar,		adParamInput,		255,	ETC) _
			, Array("@State",					adChar,			adParamInput,		1,		"Y") _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
			, Array("@IDX",						adInteger,		adParamInput,		0,		IDX) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "T셔츠관리  > "& MYear &" "& Division &" "& Subject &" "& IDX &" "& TSize &" 수정 되었습니다."
		InsertType = "Update"
	ElseIf ProcessType = "Delete" Then
		'// 삭제 ================
		SQL = "DELETE FROM ConfigTableTsize WHERE IDX = ?; "

		Call objDB.sbSetArray("@IDX",	adInteger, adParamInput, 0,		IDX)

		'objDB.blnDebug = true
		arrParams = objDB.fnGetArray
		Call objDB.sbExecSQL(SQL, arrParams)

		strLogMSG = "T셔츠관리  > DIX : "& IDX &" 삭제 되었습니다."
		InsertType = "Delete"
	End If
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "T셔츠관리 저장 완료"
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

'// 평가요소관리
Sub setInterviewItem()
	'On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()
	
	If ProcessType = "Insert" Then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO ConfigTableInterviewItem ( "
		SQL = SQL & vbCrLf & "		MYear, Division, Subject, Division1, Division2 "
		SQL = SQL & vbCrLf & "		, ItemOrder, ItemName, ItemRate "
		SQL = SQL & vbCrLf & "		, ItemGrade_01, ItemGrade_02, ItemGrade_03, ItemGrade_04, ItemGrade_05 "
		SQL = SQL & vbCrLf & "		, ItemGrade_06, ItemGrade_07, ItemGrade_08, ItemGrade_09, ItemGrade_10 "
		SQL = SQL & vbCrLf & "		, ItemPoint_01, ItemPoint_02, ItemPoint_03, ItemPoint_04, ItemPoint_05 "
		SQL = SQL & vbCrLf & "		, ItemPoint_06, ItemPoint_07, ItemPoint_08, ItemPoint_09, ItemPoint_10 "
		SQL = SQL & vbCrLf & "		, InterviewGroundsCHK, ETC "
		SQL = SQL & vbCrLf & "		, State, RegID "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ? "
		SQL = SQL & vbCrLf & " ) "

		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@ItemOrder",				adInteger,		adParamInput,		0,		ItemOrder) _
			, Array("@ItemName",				adVarchar,		adParamInput,		255,	ItemName) _
			, Array("@ItemRate",				adVarchar,		adParamInput,		10,		ItemRate) _
			, Array("@ItemGrade_01",			adVarchar,		adParamInput,		10,		ItemGrade_01) _
			, Array("@ItemGrade_02",			adVarchar,		adParamInput,		10,		ItemGrade_02) _
			, Array("@ItemGrade_03",			adVarchar,		adParamInput,		10,		ItemGrade_03) _
			, Array("@ItemGrade_04",			adVarchar,		adParamInput,		10,		ItemGrade_04) _
			, Array("@ItemGrade_05",			adVarchar,		adParamInput,		10,		ItemGrade_05) _
			, Array("@ItemGrade_06",			adVarchar,		adParamInput,		10,		ItemGrade_06) _
			, Array("@ItemGrade_07",			adVarchar,		adParamInput,		10,		ItemGrade_07) _
			, Array("@ItemGrade_08",			adVarchar,		adParamInput,		10,		ItemGrade_08) _
			, Array("@ItemGrade_09",			adVarchar,		adParamInput,		10,		ItemGrade_09) _
			, Array("@ItemGrade_10",			adVarchar,		adParamInput,		10,		ItemGrade_10) _
			, Array("@ItemPoint_01",			adVarchar,		adParamInput,		10,		ItemPoint_01) _
			, Array("@ItemPoint_02",			adVarchar,		adParamInput,		10,		ItemPoint_02) _
			, Array("@ItemPoint_03",			adVarchar,		adParamInput,		10,		ItemPoint_03) _
			, Array("@ItemPoint_04",			adVarchar,		adParamInput,		10,		ItemPoint_04) _
			, Array("@ItemPoint_05",			adVarchar,		adParamInput,		10,		ItemPoint_05) _
			, Array("@ItemPoint_06",			adVarchar,		adParamInput,		10,		ItemPoint_06) _
			, Array("@ItemPoint_07",			adVarchar,		adParamInput,		10,		ItemPoint_07) _
			, Array("@ItemPoint_08",			adVarchar,		adParamInput,		10,		ItemPoint_08) _
			, Array("@ItemPoint_09",			adVarchar,		adParamInput,		10,		ItemPoint_09) _
			, Array("@ItemPoint_10",			adVarchar,		adParamInput,		10,		ItemPoint_10) _
			, Array("@InterviewGroundsCHK",		adChar,			adParamInput,		1,		InterviewGroundsCHK) _
			, Array("@ETC",						adVarchar,		adParamInput,		255,	ETC) _
			, Array("@State",					adChar,			adParamInput,		1,		"Y") _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		SQL = " SELECT @@IDENTITY; "
		aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		IDX = CInt(aryList(0, 0))
		
		strLogMSG = "평가요소관리  > "& MYear &" "& Division &" "& Subject &" "& IDX &" "& ItemName &" 입력 되었습니다."
		InsertType = "Insert"
	ElseIf ProcessType = "Update" Then
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE ConfigTableInterviewItem SET "
		SQL = SQL & vbCrLf & "		MYear = ?, Division = ?, Subject = ?, Division1 = ?, Division2 = ? "
		SQL = SQL & vbCrLf & "		, ItemOrder = ?, ItemName = ?, ItemRate = ? "
		SQL = SQL & vbCrLf & "		, ItemGrade_01 = ?, ItemGrade_02 = ?, ItemGrade_03 = ?, ItemGrade_04 = ?, ItemGrade_05 = ? "
		SQL = SQL & vbCrLf & "		, ItemGrade_06 = ?, ItemGrade_07 = ?, ItemGrade_08 = ?, ItemGrade_09 = ?, ItemGrade_10 = ? "
		SQL = SQL & vbCrLf & "		, ItemPoint_01 = ?, ItemPoint_02 = ?, ItemPoint_03 = ?, ItemPoint_04 = ?, ItemPoint_05 = ? "
		SQL = SQL & vbCrLf & "		, ItemPoint_06 = ?, ItemPoint_07 = ?, ItemPoint_08 = ?, ItemPoint_09 = ?, ItemPoint_10 = ? "
		SQL = SQL & vbCrLf & "		, InterviewGroundsCHK = ?, ETC = ? "
		SQL = SQL & vbCrLf & "		, State = ?, EditID = ?,  EditDate = getdate() "
		SQL = SQL & vbCrLf & " WHERE IDX = ?; "
		
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@ItemOrder",				adInteger,		adParamInput,		0,		ItemOrder) _
			, Array("@ItemName",				adVarchar,		adParamInput,		255,	ItemName) _
			, Array("@ItemRate",				adVarchar,		adParamInput,		10,		ItemRate) _
			, Array("@ItemGrade_01",			adVarchar,		adParamInput,		10,		ItemGrade_01) _
			, Array("@ItemGrade_02",			adVarchar,		adParamInput,		10,		ItemGrade_02) _
			, Array("@ItemGrade_03",			adVarchar,		adParamInput,		10,		ItemGrade_03) _
			, Array("@ItemGrade_04",			adVarchar,		adParamInput,		10,		ItemGrade_04) _
			, Array("@ItemGrade_05",			adVarchar,		adParamInput,		10,		ItemGrade_05) _
			, Array("@ItemGrade_06",			adVarchar,		adParamInput,		10,		ItemGrade_06) _
			, Array("@ItemGrade_07",			adVarchar,		adParamInput,		10,		ItemGrade_07) _
			, Array("@ItemGrade_08",			adVarchar,		adParamInput,		10,		ItemGrade_08) _
			, Array("@ItemGrade_09",			adVarchar,		adParamInput,		10,		ItemGrade_09) _
			, Array("@ItemGrade_10",			adVarchar,		adParamInput,		10,		ItemGrade_10) _
			, Array("@ItemPoint_01",			adVarchar,		adParamInput,		10,		ItemPoint_01) _
			, Array("@ItemPoint_02",			adVarchar,		adParamInput,		10,		ItemPoint_02) _
			, Array("@ItemPoint_03",			adVarchar,		adParamInput,		10,		ItemPoint_03) _
			, Array("@ItemPoint_04",			adVarchar,		adParamInput,		10,		ItemPoint_04) _
			, Array("@ItemPoint_05",			adVarchar,		adParamInput,		10,		ItemPoint_05) _
			, Array("@ItemPoint_06",			adVarchar,		adParamInput,		10,		ItemPoint_06) _
			, Array("@ItemPoint_07",			adVarchar,		adParamInput,		10,		ItemPoint_07) _
			, Array("@ItemPoint_08",			adVarchar,		adParamInput,		10,		ItemPoint_08) _
			, Array("@ItemPoint_09",			adVarchar,		adParamInput,		10,		ItemPoint_09) _
			, Array("@ItemPoint_10",			adVarchar,		adParamInput,		10,		ItemPoint_10) _
			, Array("@InterviewGroundsCHK",		adChar,			adParamInput,		1,		InterviewGroundsCHK) _
			, Array("@ETC",						adVarchar,		adParamInput,		255,	ETC) _
			, Array("@State",					adChar,			adParamInput,		1,		"Y") _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
			, Array("@IDX",						adInteger,		adParamInput,		0,		IDX) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "평가요소관리  > "& MYear &" "& Division &" "& Subject &" "& IDX &" "& ItemName &" 수정 되었습니다."
		InsertType = "Update"
	ElseIf ProcessType = "Delete" Then
		'// 삭제 ================
		SQL = "DELETE FROM ConfigTableInterviewItem WHERE IDX = ?; "

		Call objDB.sbSetArray("@IDX",	adInteger, adParamInput, 0,		IDX)

		'objDB.blnDebug = true
		arrParams = objDB.fnGetArray
		Call objDB.sbExecSQL(SQL, arrParams)

		strLogMSG = "평가요소관리  > DIX : "& IDX &" 삭제 되었습니다."
		InsertType = "Delete"
	End If
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "평가요소관리 저장 완료"
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