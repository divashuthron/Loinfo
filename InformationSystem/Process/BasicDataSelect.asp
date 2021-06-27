<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("process", "")
'Dim ProcessType				: ProcessType = fnR("BasicDataSetProcessType", "")
'Dim LogDivision				: LogDivision = "BasicDataSet"

'버튼번호
Dim BasicDataNum			: BasicDataNum = fnR("BasicDataBtnNum", "")

'=============== 평가비율 기본 데이터 변수 ===============

'생기부 면접 실기 수능 비율
Dim StudentRecordRatio		
Dim InterviewerRatio		
Dim PracticalRatio			
Dim CSATRatio				

'자격미달기준
Dim DrawStandard1			
Dim DrawStandard2			
Dim DrawStandard3			
Dim DrawStandard4			
Dim DrawStandard5		
Dim DrawStandard6	

'동석차기준
Dim UnqualifiedStandard1	
Dim UnqualifiedStandard2	
Dim UnqualifiedStandard3	
Dim UnqualifiedStandard4	
Dim UnqualifiedStandard5	
Dim UnqualifiedStandard6

'가산점(사용 안 함-개별)
Dim ExtraPoint1				
Dim ExtraPoint2				
Dim ExtraPoint3				
Dim ExtraPoint4				
Dim ExtraPoint5				
Dim ExtraPoint6

'장학(사용 안 함-개별)
Dim Scholarship1			
Dim Scholarship2			
Dim Scholarship3			
Dim Scholarship4			
Dim Scholarship5	
Dim Scholarship6

'필수서류(사용 안 함)
Dim DocumentaryEvidence1	
Dim DocumentaryEvidence2	
Dim DocumentaryEvidence3	
Dim DocumentaryEvidence4	
Dim DocumentaryEvidence5	
Dim DocumentaryEvidence6	
Dim DocumentaryEvidence7	
Dim DocumentaryEvidence8	
Dim DocumentaryEvidence9	
Dim DocumentaryEvidence10	

'=============== 평가비율 기본 데이터 변수 끝 ===============

'=============== 지원자 기본 데이터 변수 (위반자 추가해야 함)===============

'가산점
Dim ExtraPoint				

'생기부, 검정, 수능 동의
Dim StudentRecord			
Dim Qualification			
Dim CSAT	

'실기, 면접 점수
Dim Interviewer
Dim Practical

'필수서류
Dim document1				
Dim document2				
Dim document3				
Dim document4				
Dim document5				
Dim document6				
Dim document7				
Dim document8				

'사용 안 함
Dim document21				
Dim document22				
Dim document23				
Dim document24

'=============== 지원자 기본 데이터 변수 끝 ===============

'입력, 수정
'Dim INPT_USID    			: INPT_USID = SessionUserID
'Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
'Dim UPDT_USID    			: UPDT_USID = SessionUserID
'Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

Select Case process
	Case "RegAppraisalBasicDataSet" '평가비율
		Call SelectAppraisal()
	Case "RegApplicantBasicDataSet" '지원자
		Call SelectApplicant()
	Case "RegApplicantAddBasicDataSet" '필수서류 체크
		Call SelectApplicantAdd()
End Select

'=============== 평가비율 기본 데이터 넣기 ===============
Sub SelectAppraisal()

	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "		BasicDataNum "
	SQL = SQL & vbCrLf & "		, StudentRecordRatio, InterviewerRatio, PracticalRatio, CSATRatio "
	SQL = SQL & vbCrLf & "		, DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6 "
	SQL = SQL & vbCrLf & "		, UnqualifiedStandard1, UnqualifiedStandard2, UnqualifiedStandard3, UnqualifiedStandard4, UnqualifiedStandard5, UnqualifiedStandard6 "
	SQL = SQL & vbCrLf & "		, ExtraPoint1, ExtraPoint2, ExtraPoint3, ExtraPoint4, ExtraPoint5, ExtraPoint6 "
	SQL = SQL & vbCrLf & "		, Scholarship1, Scholarship2, Scholarship3, Scholarship4, Scholarship5, Scholarship6 "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5 "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence6, DocumentaryEvidence7, DocumentaryEvidence8, DocumentaryEvidence9, DocumentaryEvidence10 "
	SQL = SQL & vbCrLf & "FROM BasicDataTable "
	SQL = SQL & vbCrLf & "WHERE DataType = ? "
	SQL = SQL & vbCrLf & "And BasicDataNum = ? "
	SQL = SQL & vbCrLf & "And UserId = ?; "

	Call objDB.sbSetArray("@DataType", adVarchar, adParamInput, 50, process)
	Call objDB.sbSetArray("@BasicDataNum", adVarchar, adParamInput, 50, BasicDataNum)
	Call objDB.sbSetArray("@UserId", adVarchar, adParamInput, 60, SessionUserID)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	Set objDB  = Nothing

	If isArray(AryHash) Then
		StudentRecordRatio		= AryHash(0).Item("StudentRecordRatio")
		InterviewerRatio		= AryHash(0).Item("InterviewerRatio")
		PracticalRatio			= AryHash(0).Item("PracticalRatio")
		CSATRatio				= AryHash(0).Item("CSATRatio")
		DrawStandard1			= AryHash(0).Item("DrawStandard1")
		DrawStandard2			= AryHash(0).Item("DrawStandard2")
		DrawStandard3			= AryHash(0).Item("DrawStandard3")
		DrawStandard4			= AryHash(0).Item("DrawStandard4")
		DrawStandard5			= AryHash(0).Item("DrawStandard5")
		DrawStandard6			= AryHash(0).Item("DrawStandard6")
		UnqualifiedStandard1	= AryHash(0).Item("UnqualifiedStandard1")
		UnqualifiedStandard2	= AryHash(0).Item("UnqualifiedStandard2")
		UnqualifiedStandard3	= AryHash(0).Item("UnqualifiedStandard3")
		UnqualifiedStandard4	= AryHash(0).Item("UnqualifiedStandard4")
		UnqualifiedStandard5	= AryHash(0).Item("UnqualifiedStandard5")
		UnqualifiedStandard6	= AryHash(0).Item("UnqualifiedStandard6")
		ExtraPoint1				= AryHash(0).Item("ExtraPoint1")
		ExtraPoint2				= AryHash(0).Item("ExtraPoint2")
		ExtraPoint3				= AryHash(0).Item("ExtraPoint3")
		ExtraPoint4				= AryHash(0).Item("ExtraPoint4")
		ExtraPoint5				= AryHash(0).Item("ExtraPoint5")
		ExtraPoint6				= AryHash(0).Item("ExtraPoint6")
		Scholarship1			= AryHash(0).Item("Scholarship1")
		Scholarship2			= AryHash(0).Item("Scholarship2")
		Scholarship3			= AryHash(0).Item("Scholarship3")
		Scholarship4			= AryHash(0).Item("Scholarship4")
		Scholarship5			= AryHash(0).Item("Scholarship5")
		Scholarship6			= AryHash(0).Item("Scholarship6")
		DocumentaryEvidence1	= AryHash(0).Item("DocumentaryEvidence1")
		DocumentaryEvidence2	= AryHash(0).Item("DocumentaryEvidence2")
		DocumentaryEvidence3	= AryHash(0).Item("DocumentaryEvidence3")
		DocumentaryEvidence4	= AryHash(0).Item("DocumentaryEvidence4")
		DocumentaryEvidence5	= AryHash(0).Item("DocumentaryEvidence5")
		DocumentaryEvidence6	= AryHash(0).Item("DocumentaryEvidence6")
		DocumentaryEvidence7	= AryHash(0).Item("DocumentaryEvidence7")
		DocumentaryEvidence8	= AryHash(0).Item("DocumentaryEvidence8")
		DocumentaryEvidence9	= AryHash(0).Item("DocumentaryEvidence9")
		DocumentaryEvidence10	= AryHash(0).Item("DocumentaryEvidence10")
	End If

	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		'objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "기본 데이터 가져오기"
		'objDB.sbCommitTrans 
	End If	
	%>
	<Lists>
		<List>
			<Result><%= strResult %></Result>
			<InsertType><%= InsertType %></InsertType>
			<ReturnMSG><%= returnMSG %></ReturnMSG>

			<StudentRecordRatio><%= StudentRecordRatio %></StudentRecordRatio>
			<InterviewerRatio><%= InterviewerRatio %></InterviewerRatio>
			<PracticalRatio><%= PracticalRatio %></PracticalRatio>
			<CSATRatio><%= CSATRatio %></CSATRatio>
			<DrawStandard1><%= DrawStandard1 %></DrawStandard1>
			<DrawStandard2><%= DrawStandard2 %></DrawStandard2>
			<DrawStandard3><%= DrawStandard3 %></DrawStandard3>
			<DrawStandard4><%= DrawStandard4 %></DrawStandard4>
			<DrawStandard5><%= DrawStandard5 %></DrawStandard5>
			<DrawStandard6><%= DrawStandard6 %></DrawStandard6>
			<UnqualifiedStandard1><%= UnqualifiedStandard1 %></UnqualifiedStandard1>
			<UnqualifiedStandard2><%= UnqualifiedStandard2 %></UnqualifiedStandard2>
			<UnqualifiedStandard3><%= UnqualifiedStandard3 %></UnqualifiedStandard3>
			<UnqualifiedStandard4><%= UnqualifiedStandard4 %></UnqualifiedStandard4>
			<UnqualifiedStandard5><%= UnqualifiedStandard5 %></UnqualifiedStandard5>
			<UnqualifiedStandard6><%= UnqualifiedStandard6 %></UnqualifiedStandard6>
			<ExtraPoint1><%= ExtraPoint1 %></ExtraPoint1>
			<ExtraPoint2><%= ExtraPoint2 %></ExtraPoint2>
			<ExtraPoint3><%= ExtraPoint3 %></ExtraPoint3>
			<ExtraPoint4><%= ExtraPoint4 %></ExtraPoint4>
			<ExtraPoint5><%= ExtraPoint5 %></ExtraPoint5>
			<ExtraPoint6><%= ExtraPoint6 %></ExtraPoint6>
			<Scholarship1><%= Scholarship1 %></Scholarship1>
			<Scholarship2><%= Scholarship2 %></Scholarship2>
			<Scholarship3><%= Scholarship3 %></Scholarship3>
			<Scholarship4><%= Scholarship4 %></Scholarship4>
			<Scholarship5><%= Scholarship5 %></Scholarship5>
			<Scholarship6><%= Scholarship6 %></Scholarship6>
			<DocumentaryEvidence1><%= DocumentaryEvidence1 %></DocumentaryEvidence1>
			<DocumentaryEvidence2><%= DocumentaryEvidence2 %></DocumentaryEvidence2>
			<DocumentaryEvidence3><%= DocumentaryEvidence3 %></DocumentaryEvidence3>
			<DocumentaryEvidence4><%= DocumentaryEvidence4 %></DocumentaryEvidence4>
			<DocumentaryEvidence5><%= DocumentaryEvidence5 %></DocumentaryEvidence5>
			<DocumentaryEvidence6><%= DocumentaryEvidence6 %></DocumentaryEvidence6>
			<DocumentaryEvidence7><%= DocumentaryEvidence7 %></DocumentaryEvidence7>
			<DocumentaryEvidence8><%= DocumentaryEvidence8 %></DocumentaryEvidence8>
			<DocumentaryEvidence9><%= DocumentaryEvidence9 %></DocumentaryEvidence9>
			<DocumentaryEvidence10><%= DocumentaryEvidence10 %></DocumentaryEvidence10>
		</List>
	</Lists>
	</Metissoft>
<%
End Sub

'=============== 지원자 기본 데이터 넣기 ===============
Sub SelectApplicant()

	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "		BasicDataNum "
	SQL = SQL & vbCrLf & "		, ExtraPoint1, StudentRecordRatio, Qualification, CSATRatio "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5 "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence6, DocumentaryEvidence7, DocumentaryEvidence8, Interviewer, Practical "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence21, DocumentaryEvidence22, DocumentaryEvidence23, DocumentaryEvidence24 "
	SQL = SQL & vbCrLf & "FROM BasicDataTable "
	SQL = SQL & vbCrLf & "WHERE DataType = ? "
	SQL = SQL & vbCrLf & "And BasicDataNum = ? "
	SQL = SQL & vbCrLf & "And UserId = ?; "

	Call objDB.sbSetArray("@DataType", adVarchar, adParamInput, 50, process)
	Call objDB.sbSetArray("@BasicDataNum", adVarchar, adParamInput, 50, BasicDataNum)
	Call objDB.sbSetArray("@UserId", adVarchar, adParamInput, 60, SessionUserID)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	Set objDB  = Nothing

	If isArray(AryHash) Then
		ExtraPoint				= AryHash(0).Item("ExtraPoint1")
		StudentRecord			= AryHash(0).Item("StudentRecordRatio")
		Qualification			= AryHash(0).Item("Qualification")
		CSAT					= AryHash(0).Item("CSATRatio")
		document1				= AryHash(0).Item("DocumentaryEvidence1")
		document2				= AryHash(0).Item("DocumentaryEvidence2")
		document3				= AryHash(0).Item("DocumentaryEvidence3")
		document4				= AryHash(0).Item("DocumentaryEvidence4")
		document5				= AryHash(0).Item("DocumentaryEvidence5")
		document6				= AryHash(0).Item("DocumentaryEvidence6")
		document7				= AryHash(0).Item("DocumentaryEvidence7")
		document8				= AryHash(0).Item("DocumentaryEvidence8")
		Interviewer				= AryHash(0).Item("Interviewer")
		Practical				= AryHash(0).Item("Practical")
		document21				= AryHash(0).Item("DocumentaryEvidence21")
		document22				= AryHash(0).Item("DocumentaryEvidence22")
		document23				= AryHash(0).Item("DocumentaryEvidence23")
		document24				= AryHash(0).Item("DocumentaryEvidence24")
	End If

	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		'objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "기본 데이터 가져오기"
		'objDB.sbCommitTrans 
	End If	
	%>
	<Lists>
		<List>
			<Result><%= strResult %></Result>
			<InsertType><%= InsertType %></InsertType>
			<ReturnMSG><%= returnMSG %></ReturnMSG>

			<ExtraPoint><%= ExtraPoint %></ExtraPoint>
			<StudentRecord><%= StudentRecord %></StudentRecord>
			<Qualification><%= Qualification %></Qualification>
			<CSAT><%= CSAT %></CSAT>

			<document1><%= document1 %></document1>
			<document2><%= document2 %></document2>
			<document3><%= document3 %></document3>
			<document4><%= document4 %></document4>
			<document5><%= document5 %></document5>
			<document6><%= document6 %></document6>
			<document7><%= document7 %></document7>
			<document8><%= document8 %></document8>
			<Interviewer><%= Interviewer %></Interviewer>
			<Practical><%= Practical %></Practical>
			<document21><%= document21 %></document21>
			<document22><%= document22 %></document22>
			<document23><%= document23 %></document23>
			<document24><%= document24 %></document24>
		</List>
	</Lists>
	</Metissoft>
<%
End Sub

'=============== 필수서류 기본 데이터 넣기 ===============
Sub SelectApplicantAdd()

	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "		BasicDataNum "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5 "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence6, DocumentaryEvidence7, DocumentaryEvidence8 "
	SQL = SQL & vbCrLf & "		, DocumentaryEvidence21, DocumentaryEvidence22, DocumentaryEvidence23, DocumentaryEvidence24 "
	SQL = SQL & vbCrLf & "FROM BasicDataTable "
	SQL = SQL & vbCrLf & "WHERE DataType = ? "
	SQL = SQL & vbCrLf & "And BasicDataNum = ? "
	SQL = SQL & vbCrLf & "And UserId = ?; "

	Call objDB.sbSetArray("@DataType", adVarchar, adParamInput, 50, process)
	Call objDB.sbSetArray("@BasicDataNum", adVarchar, adParamInput, 50, BasicDataNum)
	Call objDB.sbSetArray("@UserId", adVarchar, adParamInput, 60, SessionUserID)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	Set objDB  = Nothing

	If isArray(AryHash) Then
		document1				= AryHash(0).Item("DocumentaryEvidence1")
		document2				= AryHash(0).Item("DocumentaryEvidence2")
		document3				= AryHash(0).Item("DocumentaryEvidence3")
		document4				= AryHash(0).Item("DocumentaryEvidence4")
		document5				= AryHash(0).Item("DocumentaryEvidence5")
		document6				= AryHash(0).Item("DocumentaryEvidence6")
		document7				= AryHash(0).Item("DocumentaryEvidence7")
		document8				= AryHash(0).Item("DocumentaryEvidence8")
		document21				= AryHash(0).Item("DocumentaryEvidence21")
		document22				= AryHash(0).Item("DocumentaryEvidence22")
		document23				= AryHash(0).Item("DocumentaryEvidence23")
		document24				= AryHash(0).Item("DocumentaryEvidence24")
	End If

	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		'objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "기본 데이터 가져오기"
		'objDB.sbCommitTrans 
	End If	
	%>
	<Lists>
		<List>
			<Result><%= strResult %></Result>
			<InsertType><%= InsertType %></InsertType>
			<ReturnMSG><%= returnMSG %></ReturnMSG>

			<document1><%= document1 %></document1>
			<document2><%= document2 %></document2>
			<document3><%= document3 %></document3>
			<document4><%= document4 %></document4>
			<document5><%= document5 %></document5>
			<document6><%= document6 %></document6>
			<document7><%= document7 %></document7>
			<document8><%= document8 %></document8>
			<document21><%= document21 %></document21>
			<document22><%= document22 %></document22>
			<document23><%= document23 %></document23>
			<document24><%= document24 %></document24>
		</List>
	</Lists>
	</Metissoft>
<%
End Sub
%>

