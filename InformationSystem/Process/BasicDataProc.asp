<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("BasicDataSetprocess", "")
'Dim ProcessType				: ProcessType = fnR("BasicDataSetProcessType", "")
Dim LogDivision				: LogDivision = "BasicDataSet"

'버튼번호
Dim BasicDataNum			: BasicDataNum = fnR("BasicDataBtn", "")

'=============== 평가비율 기본 데이터 변수 ===============

'생기부 면접 실기 수능 비율
Dim StudentRecordRatio		: StudentRecordRatio = getIntParameter(fnR("StudentRecordRatio", 0), 0) 
Dim InterviewerRatio		: InterviewerRatio = getIntParameter(fnR("InterviewerRatio", 0), 0) 
Dim PracticalRatio			: PracticalRatio = getIntParameter(fnR("PracticalRatio", 0), 0) 
Dim CSATRatio				: CSATRatio = getIntParameter(fnR("CSATRatio", 0), 0) 

'자격미달기준
Dim DrawStandard1			: DrawStandard1 = fnRF("DrawStandard1")
Dim DrawStandard2			: DrawStandard2 = fnRF("DrawStandard2")
Dim DrawStandard3			: DrawStandard3 = fnRF("DrawStandard3")
Dim DrawStandard4			: DrawStandard4 = fnRF("DrawStandard4")
Dim DrawStandard5			: DrawStandard5 = fnRF("DrawStandard5")
Dim DrawStandard6			: DrawStandard6 = fnRF("DrawStandard6")

'동석차기준
Dim UnqualifiedStandard1	: UnqualifiedStandard1 = fnRF("UnqualifiedStandard1")
Dim UnqualifiedStandard2	: UnqualifiedStandard2 = fnRF("UnqualifiedStandard2")
Dim UnqualifiedStandard3	: UnqualifiedStandard3 = fnRF("UnqualifiedStandard3")
Dim UnqualifiedStandard4	: UnqualifiedStandard4 = fnRF("UnqualifiedStandard4")
Dim UnqualifiedStandard5	: UnqualifiedStandard5 = fnRF("UnqualifiedStandard5")
Dim UnqualifiedStandard6	: UnqualifiedStandard6 = fnRF("UnqualifiedStandard6")

'가산점(사용 안 함-개별)
Dim ExtraPoint1				: ExtraPoint1 = fnRF("ExtraPoint1")
Dim ExtraPoint2				: ExtraPoint2 = fnRF("ExtraPoint2")
Dim ExtraPoint3				: ExtraPoint3 = fnRF("ExtraPoint3")
Dim ExtraPoint4				: ExtraPoint4 = fnRF("ExtraPoint4")
Dim ExtraPoint5				: ExtraPoint5 = fnRF("ExtraPoint5")
Dim ExtraPoint6				: ExtraPoint6 = fnRF("ExtraPoint6")

'장학(사용 안 함-개별)
Dim Scholarship1			: Scholarship1 = fnRF("Scholarship1")
Dim Scholarship2			: Scholarship2 = fnRF("Scholarship2")
Dim Scholarship3			: Scholarship3 = fnRF("Scholarship3")
Dim Scholarship4			: Scholarship4 = fnRF("Scholarship4")
Dim Scholarship5			: Scholarship5 = fnRF("Scholarship5")
Dim Scholarship6			: Scholarship6 = fnRF("Scholarship6")

'필수서류(사용 안 함)
Dim DocumentaryEvidence1	: DocumentaryEvidence1 = fnRF("DocumentaryEvidence1")
Dim DocumentaryEvidence2	: DocumentaryEvidence2 = fnRF("DocumentaryEvidence2")
Dim DocumentaryEvidence3	: DocumentaryEvidence3 = fnRF("DocumentaryEvidence3")
Dim DocumentaryEvidence4	: DocumentaryEvidence4 = fnRF("DocumentaryEvidence4")
Dim DocumentaryEvidence5	: DocumentaryEvidence5 = fnRF("DocumentaryEvidence5")
Dim DocumentaryEvidence6	: DocumentaryEvidence6 = fnRF("DocumentaryEvidence6")
Dim DocumentaryEvidence7	: DocumentaryEvidence7 = fnRF("DocumentaryEvidence7")
Dim DocumentaryEvidence8	: DocumentaryEvidence8 = fnRF("DocumentaryEvidence8")
Dim DocumentaryEvidence9	: DocumentaryEvidence9 = fnRF("DocumentaryEvidence9")
Dim DocumentaryEvidence10	: DocumentaryEvidence10 = fnRF("DocumentaryEvidence10")

'=============== 평가비율 기본 데이터 변수 끝 ===============

'=============== 지원자 기본 데이터 변수 (위반자 추가해야 함)===============

'가산점
Dim ExtraPoint				: ExtraPoint = fnRF("ExtraPoint")

'생기부, 검정, 수능 동의
Dim StudentRecord			: StudentRecord = fnRF("StudentRecord")
Dim Qualification			: Qualification = fnRF("Qualification")
Dim CSAT					: CSAT = fnRF("CSAT")

'면접, 실기 점수
Dim Interviewer				: Interviewer = fnRF("Interviewer")
Dim Practical				: Practical = fnRF("Practical")

'필수서류(1~8만 사용)
Dim document1				: document1 = fnRF("document1")
Dim document2				: document2 = fnRF("document2")
Dim document3				: document3 = fnRF("document3")
Dim document4				: document4 = fnRF("document4")
Dim document5				: document5 = fnRF("document5")
Dim document6				: document6 = fnRF("document6")
Dim document7				: document7 = fnRF("document7")
Dim document8				: document8 = fnRF("document8")
Dim document21				: document21 = fnRF("document21")
Dim document22				: document22 = fnRF("document22")
Dim document23				: document23 = fnRF("document23")
Dim document24				: document24 = fnRF("document24")

'=============== 지원자 기본 데이터 변수 끝 ===============

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM


Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()


'///////////////////////////////////////////////////////////////////////
'// insert or update 결정. 
'// BasicDataTable테이블에 있으면 Update, 없으면 Insert
'///////////////////////////////////////////////////////////////////////	
SQL = ""
SQL = SQL & vbCrLf & "Select BasicDataNum "
SQL = SQL & vbCrLf & "from BasicDataTable "
SQL = SQL & vbCrLf & "where DataType = ? "
SQL = SQL & vbCrLf & "And BasicDataNum = ? "
SQL = SQL & vbCrLf & "And UserId = ?; "

Call objDB.sbSetArray("@DataType", adVarchar, adParamInput, 60, process)
Call objDB.sbSetArray("@BasicDataNum", adVarchar, adParamInput, 50, BasicDataNum)
Call objDB.sbSetArray("@UserId", adVarchar, adParamInput, 60, SessionUserID)

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

If Not IsNull(AryHash) then
	ProcessType = "Update"
Else
	ProcessType = "Insert"
End If

Select Case process
	Case "RegAppraisalBasicDataSet" '평가비율
		Call setAppraisal()
	Case "RegApplicantBasicDataSet" '지원자
		Call setApplicant()
	Case "RegApplicantAddBasicDataSet" '지원자 서류체크(관리자 외)
		Call setApplicantAdd()
End Select

'=============== 평가비율 기본 데이터 입력 ===============
Sub setAppraisal()

	'On Error Resume Next

	'//////////////////////////////////////////////////////////
	'// 모집단위 별 평가비율관리
	'//////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO BasicDataTable ( "
		SQL = SQL & vbCrLf & "		DataType, BasicDataNum, UserId, StudentRecordRatio, InterviewerRatio, PracticalRatio, CSATRatio  "
		SQL = SQL & vbCrLf & "		,DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6  "
		SQL = SQL & vbCrLf & "		,UnqualifiedStandard1, UnqualifiedStandard2, UnqualifiedStandard3, UnqualifiedStandard4, UnqualifiedStandard5, UnqualifiedStandard6  "
		SQL = SQL & vbCrLf & "		,ExtraPoint1, ExtraPoint2, ExtraPoint3, ExtraPoint4, ExtraPoint5, ExtraPoint6 "
		SQL = SQL & vbCrLf & "		,Scholarship1, Scholarship2, Scholarship3, Scholarship4, Scholarship5, Scholarship6 "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 ,DocumentaryEvidence7, DocumentaryEvidence8, DocumentaryEvidence9, DocumentaryEvidence10 "
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
		SQL = SQL & vbCrLf & " ) "

		'insert일 때는 INPT입력  
		arrParams = Array(_
			  Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adInteger,		adParamInput,		0,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _
			, Array("@StudentRecordRatio",		adInteger,		adParamInput,		0,		StudentRecordRatio) _
			, Array("@InterviewerRatio",		adInteger,		adParamInput,		0,		InterviewerRatio) _
			, Array("@PracticalRatio",			adInteger,		adParamInput,		0,		PracticalRatio) _
			, Array("@CSATRatio",				adInteger,		adParamInput,		0,		CSATRatio) _
			, Array("@DrawStandard1",			adVarchar,		adParamInput,		50,		DrawStandard1) _
			, Array("@DrawStandard2",			adVarchar,		adParamInput,		50,		DrawStandard2) _
			, Array("@DrawStandard3",			adVarchar,		adParamInput,		50,		DrawStandard3) _
			, Array("@DrawStandard4",			adVarchar,		adParamInput,		50,		DrawStandard4) _
			, Array("@DrawStandard5",			adVarchar,		adParamInput,		50,		DrawStandard5) _
			, Array("@DrawStandard6",			adVarchar,		adParamInput,		50,		DrawStandard6) _
			, Array("@UnqualifiedStandard1",	adVarchar,		adParamInput,		50,		UnqualifiedStandard1) _
			, Array("@UnqualifiedStandard2",	adVarchar,		adParamInput,		50,		UnqualifiedStandard2) _
			, Array("@UnqualifiedStandard3",	adVarchar,		adParamInput,		50,		UnqualifiedStandard3) _
			, Array("@UnqualifiedStandard4",	adVarchar,		adParamInput,		50,		UnqualifiedStandard4) _
			, Array("@UnqualifiedStandard5",	adVarchar,		adParamInput,		50,		UnqualifiedStandard5) _
			, Array("@UnqualifiedStandard6",	adVarchar,		adParamInput,		50,		UnqualifiedStandard6) _
			, Array("@ExtraPoint1",				adVarchar,		adParamInput,		50,		ExtraPoint1) _
			, Array("@ExtraPoint2",				adVarchar,		adParamInput,		50,		ExtraPoint2) _
			, Array("@ExtraPoint3",				adVarchar,		adParamInput,		50,		ExtraPoint3) _
			, Array("@ExtraPoint4",				adVarchar,		adParamInput,		50,		ExtraPoint4) _
			, Array("@ExtraPoint5",				adVarchar,		adParamInput,		50,		ExtraPoint5) _
			, Array("@ExtraPoint6",				adVarchar,		adParamInput,		50,		ExtraPoint6) _
			, Array("@Scholarship1",			adVarchar,		adParamInput,		50,		Scholarship1) _
			, Array("@Scholarship2",			adVarchar,		adParamInput,		50,		Scholarship2) _
			, Array("@Scholarship3",			adVarchar,		adParamInput,		50,		Scholarship3) _
			, Array("@Scholarship4",			adVarchar,		adParamInput,		50,		Scholarship4) _
			, Array("@Scholarship5",			adVarchar,		adParamInput,		50,		Scholarship5) _
			, Array("@Scholarship6",			adVarchar,		adParamInput,		50,		Scholarship6) _
			, Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		DocumentaryEvidence1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		DocumentaryEvidence2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		DocumentaryEvidence3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		DocumentaryEvidence4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		DocumentaryEvidence5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		DocumentaryEvidence6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		DocumentaryEvidence7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		DocumentaryEvidence8) _
			, Array("@DocumentaryEvidence9",	adVarchar,		adParamInput,		50,		DocumentaryEvidence9) _
			, Array("@DocumentaryEvidence10",	adVarchar,		adParamInput,		50,		DocumentaryEvidence10) _
			, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
			, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		'SQL = " SELECT @@IDENTITY; "
		'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		'IDX = CInt(aryList(0, 0))

		strLogMSG = "평가기준관리 > " & BasicDataNum & "번 버튼의 평가비율 기본데이터가 등록되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE BasicDataTable SET "
		SQL = SQL & vbCrLf & "		StudentRecordRatio = ?, InterviewerRatio = ?, PracticalRatio = ?, CSATRatio = ?  "
		SQL = SQL & vbCrLf & "		,DrawStandard1 = ?, DrawStandard2 = ?, DrawStandard3 = ?, DrawStandard4 = ?, DrawStandard5 = ?, DrawStandard6 = ?  "
		SQL = SQL & vbCrLf & "		,UnqualifiedStandard1 = ?, UnqualifiedStandard2 = ?, UnqualifiedStandard3 = ?, UnqualifiedStandard4 = ?, UnqualifiedStandard5 = ?, UnqualifiedStandard6 = ?  "
		SQL = SQL & vbCrLf & "		,ExtraPoint1 = ?, ExtraPoint2 = ?, ExtraPoint3 = ?, ExtraPoint4 = ?, ExtraPoint5 = ?, ExtraPoint6 = ? "
		SQL = SQL & vbCrLf & "		,Scholarship1 = ?, Scholarship2 = ?, Scholarship3 = ?, Scholarship4 = ?, Scholarship5 = ?, Scholarship6 = ? "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1 = ?, DocumentaryEvidence2 = ?, DocumentaryEvidence3 = ?, DocumentaryEvidence4 = ?, DocumentaryEvidence5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 = ? ,DocumentaryEvidence7 = ?, DocumentaryEvidence8 = ?, DocumentaryEvidence9 = ?, DocumentaryEvidence10 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE DataType = ? "
		SQL = SQL & vbCrLf & " AND BasicDataNum = ? "
		SQL = SQL & vbCrLf & " AND UserId = ? "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@StudentRecordRatio",		adInteger,		adParamInput,		0,		StudentRecordRatio) _
			, Array("@InterviewerRatio",		adInteger,		adParamInput,		0,		InterviewerRatio) _
			, Array("@PracticalRatio",			adInteger,		adParamInput,		0,		PracticalRatio) _
			, Array("@CSATRatio",				adInteger,		adParamInput,		0,		CSATRatio) _
			, Array("@DrawStandard1",			adVarchar,		adParamInput,		50,		DrawStandard1) _
			, Array("@DrawStandard2",			adVarchar,		adParamInput,		50,		DrawStandard2) _
			, Array("@DrawStandard3",			adVarchar,		adParamInput,		50,		DrawStandard3) _
			, Array("@DrawStandard4",			adVarchar,		adParamInput,		50,		DrawStandard4) _
			, Array("@DrawStandard5",			adVarchar,		adParamInput,		50,		DrawStandard5) _
			, Array("@DrawStandard6",			adVarchar,		adParamInput,		50,		DrawStandard6) _
			, Array("@UnqualifiedStandard1",	adVarchar,		adParamInput,		50,		UnqualifiedStandard1) _
			, Array("@UnqualifiedStandard2",	adVarchar,		adParamInput,		50,		UnqualifiedStandard2) _
			, Array("@UnqualifiedStandard3",	adVarchar,		adParamInput,		50,		UnqualifiedStandard3) _
			, Array("@UnqualifiedStandard4",	adVarchar,		adParamInput,		50,		UnqualifiedStandard4) _
			, Array("@UnqualifiedStandard5",	adVarchar,		adParamInput,		50,		UnqualifiedStandard5) _
			, Array("@UnqualifiedStandard6",	adVarchar,		adParamInput,		50,		UnqualifiedStandard6) _
			, Array("@ExtraPoint1",				adVarchar,		adParamInput,		50,		ExtraPoint1) _
			, Array("@ExtraPoint2",				adVarchar,		adParamInput,		50,		ExtraPoint2) _
			, Array("@ExtraPoint3",				adVarchar,		adParamInput,		50,		ExtraPoint3) _
			, Array("@ExtraPoint4",				adVarchar,		adParamInput,		50,		ExtraPoint4) _
			, Array("@ExtraPoint5",				adVarchar,		adParamInput,		50,		ExtraPoint5) _
			, Array("@ExtraPoint6",				adVarchar,		adParamInput,		50,		ExtraPoint6) _
			, Array("@Scholarship1",			adVarchar,		adParamInput,		50,		Scholarship1) _
			, Array("@Scholarship2",			adVarchar,		adParamInput,		50,		Scholarship2) _
			, Array("@Scholarship3",			adVarchar,		adParamInput,		50,		Scholarship3) _
			, Array("@Scholarship4",			adVarchar,		adParamInput,		50,		Scholarship4) _
			, Array("@Scholarship5",			adVarchar,		adParamInput,		50,		Scholarship5) _
			, Array("@Scholarship6",			adVarchar,		adParamInput,		50,		Scholarship6) _
			, Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		DocumentaryEvidence1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		DocumentaryEvidence2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		DocumentaryEvidence3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		DocumentaryEvidence4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		DocumentaryEvidence5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		DocumentaryEvidence6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		DocumentaryEvidence7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		DocumentaryEvidence8) _
			, Array("@DocumentaryEvidence9",	adVarchar,		adParamInput,		50,		DocumentaryEvidence9) _
			, Array("@DocumentaryEvidence10",	adVarchar,		adParamInput,		50,		DocumentaryEvidence10) _
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adVarchar,		adParamInput,		50,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _			
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "평가기준관리 > " & BasicDataNum & "번 버튼의 평가비율 기본데이터가 수정되었습니다."
		InsertType = "Update"
	end If
	'=============== 평가비율 기본 데이터 입력 끝 ===============
End Sub


'=============== 지원자 기본 데이터 입력 ===============
Sub setApplicant()

	'On Error Resume Next

	'/////////////////////////////////////////////////////////////
	'// 지원자별 가산점, 동의, 필수서류 관리 (위반자 추가해야 함)
	'/////////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO BasicDataTable ( "
		SQL = SQL & vbCrLf & "		DataType, BasicDataNum, UserId, ExtraPoint1, StudentRecordRatio, Qualification, CSATRatio  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 ,DocumentaryEvidence7, DocumentaryEvidence8, Interviewer, Practical "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence21, DocumentaryEvidence22, DocumentaryEvidence23, DocumentaryEvidence24 "
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
		SQL = SQL & vbCrLf & " ) "

		'insert일 때는 INPT입력
		arrParams = Array(_
			  Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adInteger,		adParamInput,		0,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _
			, Array("@ExtraPoint1",				adInteger,		adParamInput,		0,		ExtraPoint) _
			, Array("@StudentRecordRatio",		adInteger,		adParamInput,		0,		StudentRecord) _
			, Array("@Qualification",			adInteger,		adParamInput,		0,		Qualification) _
			, Array("@CSATRatio",				adInteger,		adParamInput,		0,		CSAT) _
			, Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		document1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		document2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		document3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		document4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		document5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		document6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		document7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		document8) _
			, Array("@Interviewer",				adInteger,		adParamInput,		0,		Interviewer) _
			, Array("@Practical",				adInteger,		adParamInput,		0,		Practical) _
			, Array("@DocumentaryEvidence21",	adVarchar,		adParamInput,		50,		document21) _
			, Array("@DocumentaryEvidence22",	adVarchar,		adParamInput,		50,		document22) _
			, Array("@DocumentaryEvidence23",	adVarchar,		adParamInput,		50,		document23) _
			, Array("@DocumentaryEvidence24",	adVarchar,		adParamInput,		50,		document24) _
			, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
			, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		'SQL = " SELECT @@IDENTITY; "
		'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		'IDX = CInt(aryList(0, 0))

		strLogMSG = "지원자관리 > " & BasicDataNum & "번 버튼의 지원자 기본데이터가 등록되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE BasicDataTable SET "
		SQL = SQL & vbCrLf & "		ExtraPoint1 = ?, StudentRecordRatio = ?, Qualification = ?, CSATRatio = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1 = ?, DocumentaryEvidence2 = ?, DocumentaryEvidence3 = ?, DocumentaryEvidence4 = ?, DocumentaryEvidence5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 = ? ,DocumentaryEvidence7 = ?, DocumentaryEvidence8 = ?, Interviewer = ?, Practical = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence21 = ?, DocumentaryEvidence22 = ?, DocumentaryEvidence23 = ?, DocumentaryEvidence24 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE DataType = ? "
		SQL = SQL & vbCrLf & " AND BasicDataNum = ? "
		SQL = SQL & vbCrLf & " AND UserId = ? "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@ExtraPoint1",				adInteger,		adParamInput,		0,		ExtraPoint) _
			, Array("@StudentRecordRatio",		adInteger,		adParamInput,		0,		StudentRecord) _
			, Array("@Qualification",			adInteger,		adParamInput,		0,		Qualification) _
			, Array("@CSATRatio",				adInteger,		adParamInput,		0,		CSAT) _
			, Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		document1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		document2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		document3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		document4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		document5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		document6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		document7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		document8) _
			, Array("@Interviewer",				adInteger,		adParamInput,		0,		Interviewer) _
			, Array("@Practical",				adInteger,		adParamInput,		0,		Practical) _
			, Array("@DocumentaryEvidence21",	adVarchar,		adParamInput,		50,		document21) _
			, Array("@DocumentaryEvidence22",	adVarchar,		adParamInput,		50,		document22) _
			, Array("@DocumentaryEvidence23",	adVarchar,		adParamInput,		50,		document23) _
			, Array("@DocumentaryEvidence24",	adVarchar,		adParamInput,		50,		document24) _
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adVarchar,		adParamInput,		50,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "지원자관리 > " & BasicDataNum & "번 버튼의 지원자 기본데이터가 수정되었습니다."
		InsertType = "Update"
	end If
	'=============== 지원자 기본 데이터 입력 끝 ===============
End Sub

'=============== 필수서류 기본 데이터 입력(관리자 외용) ===============
Sub setApplicantAdd()

	'On Error Resume Next

	'/////////////////////////////////////////////////////////////
	'// 지원자별 자격미달, 필수서류 체크(관리자 외용)
	'/////////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO BasicDataTable ( "
		SQL = SQL & vbCrLf & "		DataType, BasicDataNum, UserId  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 ,DocumentaryEvidence7, DocumentaryEvidence8, Interviewer, Practical "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence21, DocumentaryEvidence22, DocumentaryEvidence23, DocumentaryEvidence24 "
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
		SQL = SQL & vbCrLf & " ) "

		'insert일 때는 INPT입력
		arrParams = Array(_
			  Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adInteger,		adParamInput,		0,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _
			, Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		document1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		document2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		document3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		document4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		document5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		document6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		document7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		document8) _
			, Array("@Interviewer",				adInteger,		adParamInput,		0,		Interviewer) _
			, Array("@Practical",				adInteger,		adParamInput,		0,		Practical) _
			, Array("@DocumentaryEvidence21",	adVarchar,		adParamInput,		50,		document21) _
			, Array("@DocumentaryEvidence22",	adVarchar,		adParamInput,		50,		document22) _
			, Array("@DocumentaryEvidence23",	adVarchar,		adParamInput,		50,		document23) _
			, Array("@DocumentaryEvidence24",	adVarchar,		adParamInput,		50,		document24) _
			, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
			, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		'SQL = " SELECT @@IDENTITY; "
		'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		'IDX = CInt(aryList(0, 0))

		strLogMSG = "지원자관리 > " & BasicDataNum & "번 버튼의 필수서류 체크 기본데이터가 등록되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE BasicDataTable SET "
		SQL = SQL & vbCrLf & "		DocumentaryEvidence1 = ?, DocumentaryEvidence2 = ?, DocumentaryEvidence3 = ?, DocumentaryEvidence4 = ?, DocumentaryEvidence5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 = ? ,DocumentaryEvidence7 = ?, DocumentaryEvidence8 = ?, Interviewer = ?, Practical = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence21 = ?, DocumentaryEvidence22 = ?, DocumentaryEvidence23 = ?, DocumentaryEvidence24 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE DataType = ? "
		SQL = SQL & vbCrLf & " AND BasicDataNum = ? "
		SQL = SQL & vbCrLf & " AND UserId = ? "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@DocumentaryEvidence1",	adVarchar,		adParamInput,		50,		document1) _
			, Array("@DocumentaryEvidence2",	adVarchar,		adParamInput,		50,		document2) _
			, Array("@DocumentaryEvidence3",	adVarchar,		adParamInput,		50,		document3) _
			, Array("@DocumentaryEvidence4",	adVarchar,		adParamInput,		50,		document4) _
			, Array("@DocumentaryEvidence5",	adVarchar,		adParamInput,		50,		document5) _
			, Array("@DocumentaryEvidence6",	adVarchar,		adParamInput,		50,		document6) _
			, Array("@DocumentaryEvidence7",	adVarchar,		adParamInput,		50,		document7) _
			, Array("@DocumentaryEvidence8",	adVarchar,		adParamInput,		50,		document8) _
			, Array("@Interviewer",				adInteger,		adParamInput,		0,		Interviewer) _
			, Array("@Practical",				adInteger,		adParamInput,		0,		Practical) _
			, Array("@DocumentaryEvidence21",	adVarchar,		adParamInput,		50,		document21) _
			, Array("@DocumentaryEvidence22",	adVarchar,		adParamInput,		50,		document22) _
			, Array("@DocumentaryEvidence23",	adVarchar,		adParamInput,		50,		document23) _
			, Array("@DocumentaryEvidence24",	adVarchar,		adParamInput,		50,		document24) _
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@DataType",				adVarchar,		adParamInput,		60,		process) _
			, Array("@BasicDataNum",			adVarchar,		adParamInput,		50,		BasicDataNum) _
			, Array("@UserId",					adVarchar,		adParamInput,		60,		SessionUserID) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "지원자관리 > " & BasicDataNum & "번 버튼의 필수서류 체크 기본데이터가 수정되었습니다."
		InsertType = "Update"
	end If
	'=============== 필수서류 기본 데이터 입력 끝(관리자 외용) ===============
End Sub

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "기본데이터 저장 완료"
	'objDB.sbCommitTrans 
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
</Metissoft>