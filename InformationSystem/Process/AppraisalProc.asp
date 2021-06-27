<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
'Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "Appraisal"

'모집단위
Dim IDX						: IDX = fnR("IDX", 0)
'Dim MYear                 	: MYear                =	fnRF("MYear")							'입력한 년도
Dim MYearHidden           	: MYearHidden = fnR("MyearHidden", "")
Dim MyearHiddenTemp
Dim SubjectCodehidden		: SubjectCodehidden = fnR("SubjectCodehidden", "")

'히스토리용(한글)
Dim Division0Name			: Division0Name = fnRF("Division0Name")
Dim SubjectName				: SubjectName = fnRF("SubjectName")
Dim Division1Name			: Division1Name = fnRF("Division1Name")
Dim Division2Name			: Division2Name = fnRF("Division2Name")
Dim Division3Name			: Division3Name = fnRF("Division3Name")

'히스토리용(카운트)
Dim InsertCnt				: InsertCnt = 1
Dim UpdateCnt				: UpdateCnt = 1

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

'가산점
Dim ExtraPoint1				: ExtraPoint1 = fnRF("ExtraPoint1")
Dim ExtraPoint2				: ExtraPoint2 = fnRF("ExtraPoint2")
Dim ExtraPoint3				: ExtraPoint3 = fnRF("ExtraPoint3")
Dim ExtraPoint4				: ExtraPoint4 = fnRF("ExtraPoint4")
Dim ExtraPoint5				: ExtraPoint5 = fnRF("ExtraPoint5")
Dim ExtraPoint6				: ExtraPoint6 = fnRF("ExtraPoint6")

'장학
Dim Scholarship1			: Scholarship1 = fnRF("Scholarship1")
Dim Scholarship2			: Scholarship2 = fnRF("Scholarship2")
Dim Scholarship3			: Scholarship3 = fnRF("Scholarship3")
Dim Scholarship4			: Scholarship4 = fnRF("Scholarship4")
Dim Scholarship5			: Scholarship5 = fnRF("Scholarship5")
Dim Scholarship6			: Scholarship6 = fnRF("Scholarship6")

'필수서류
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

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

'SubjectCodehidden 값을 하나씩 풀어, insert or update 결정. 
Dim SubjectCode				: SubjectCode = Split(SubjectCodehidden, ",")
Dim Myear					: Myear = Split(MYearHidden, ",")

Dim i, SubjectCodeTemp, MyearTemp
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, strLogMSG2

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'체크박스가 선택된 갯수만큼 반복
For i = 0 To Ubound(SubjectCode)
	'///////////////////////////////////////////////////////////////////////////////////////////////////////////
	'// insert or update 결정. 
	'// AppraisalTable테이블에 SubjectCode가 있으면 Update, 없으면 Insert
	'// *모집단위에서 년도 바꾼 뒤 평가비율 다시 입력할 때를 대비하여 년도를 업데이트시에는 년도를 뽑아서 사용.
	'///////////////////////////////////////////////////////////////////////////////////////////////////////////	
	SQL = ""
	SQL = SQL & vbCrLf & "Select MYear, SubjectCode "
	SQL = SQL & vbCrLf & "from AppraisalTable "
	SQL = SQL & vbCrLf & "where SubjectCode = ?; "

	Call objDB.sbSetArray("@SubjectCode", adVarchar, adParamInput, 50, SubjectCode(i))
	SubjectCodeTemp = SubjectCode(i)
	MyearTemp = Myear(i)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)
	
	If Not IsNull(AryHash) then
		ProcessType = "Update"
		MyearHiddenTemp = AryHash(0).Item("MYear")
	Else
		ProcessType = "Insert"
	End If

	'=============== 평가비율 입력 ===============

	'On Error Resume Next
	
	'//////////////////////////////////////////////////////////
	'// 모집단위 별 평가비율관리
	'//////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO AppraisalTable ( "
		SQL = SQL & vbCrLf & "		MYear, SubjectCode "
		SQL = SQL & vbCrLf & "		,StudentRecordRatio, InterviewerRatio, PracticalRatio, CSATRatio  "
		SQL = SQL & vbCrLf & "		,DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6  "
		SQL = SQL & vbCrLf & "		,UnqualifiedStandard1, UnqualifiedStandard2, UnqualifiedStandard3, UnqualifiedStandard4, UnqualifiedStandard5, UnqualifiedStandard6  "
		SQL = SQL & vbCrLf & "		,ExtraPoint1, ExtraPoint2, ExtraPoint3, ExtraPoint4, ExtraPoint5, ExtraPoint6 "
		SQL = SQL & vbCrLf & "		,Scholarship1, Scholarship2, Scholarship3, Scholarship4, Scholarship5, Scholarship6 "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 ,DocumentaryEvidence7, DocumentaryEvidence8, DocumentaryEvidence9, DocumentaryEvidence10 "
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?"
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
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
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MyearTemp) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		50,		SubjectCodeTemp) _
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

		'//////////////////////////////////////////////////////////////////////////////////////////////////
		'// 체크박스만 클릭하여 저장했을 때와 모집단위를 선택하여 저장했을 때를 비교하여 메세지 내용 등록
		'//////////////////////////////////////////////////////////////////////////////////////////////////
		If IsE(Division0Name) Then
			strLogMSG = "평가비율관리 > "& InsertCnt &" 건의 평가비율이 등록되었습니다."
		Else
			strLogMSG = "평가비율관리 > "& MYear(i) &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과 등 " & InsertCnt &" 건의 평가비율이 등록되었습니다."
		End If
		InsertType = "Insert"
		InsertCnt = InsertCnt + 1
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE AppraisalTable SET "
		SQL = SQL & vbCrLf & "		MYear = ?, StudentRecordRatio = ?, InterviewerRatio = ?, PracticalRatio = ?, CSATRatio = ?  "
		SQL = SQL & vbCrLf & "		,DrawStandard1 = ?, DrawStandard2 = ?, DrawStandard3 = ?, DrawStandard4 = ?, DrawStandard5 = ?, DrawStandard6 = ?  "
		SQL = SQL & vbCrLf & "		,UnqualifiedStandard1 = ?, UnqualifiedStandard2 = ?, UnqualifiedStandard3 = ?, UnqualifiedStandard4 = ?, UnqualifiedStandard5 = ?, UnqualifiedStandard6 = ?  "
		SQL = SQL & vbCrLf & "		,ExtraPoint1 = ?, ExtraPoint2 = ?, ExtraPoint3 = ?, ExtraPoint4 = ?, ExtraPoint5 = ?, ExtraPoint6 = ? "
		SQL = SQL & vbCrLf & "		,Scholarship1 = ?, Scholarship2 = ?, Scholarship3 = ?, Scholarship4 = ?, Scholarship5 = ?, Scholarship6 = ? "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence1 = ?, DocumentaryEvidence2 = ?, DocumentaryEvidence3 = ?, DocumentaryEvidence4 = ?, DocumentaryEvidence5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryEvidence6 = ? ,DocumentaryEvidence7 = ?, DocumentaryEvidence8 = ?, DocumentaryEvidence9 = ?, DocumentaryEvidence10 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE MYear = ? "
		SQL = SQL & vbCrLf & " And SubjectCode = ? "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		10,		MyearTemp) _
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
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@MYear",					adVarchar,		adParamInput,		10,		MyearHiddenTemp) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		50,		SubjectCodeTemp) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		'//////////////////////////////////////////////////////////////////////////////////////////////////
		'// 체크박스만 클릭하여 저장했을 때와 모집단위를 선택하여 저장했을 때를 비교하여 메세지 내용 등록
		'//////////////////////////////////////////////////////////////////////////////////////////////////
		If IsE(Division0Name) Then
			strLogMSG2 = "평가비율관리 > "& UpdateCnt &" 건의 평가비율이 수정되었습니다."
		Else
			strLogMSG2 = "평가비율관리 > "& Myear(i) &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과 등 " & UpdateCnt &" 건의 평가비율이 수정되었습니다."
		End If

		InsertType = "Update"
		UpdateCnt = UpdateCnt + 1
	end If
	'//////////////////////////////////////////////////////////
	
Next

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "평가비율 저장 완료"
	'objDB.sbCommitTrans 
End If	

Set objDB  = Nothing

'// 로그기록
If Not(IsE(strLogMSG)) Then
	Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)
End If
If Not(IsE(strLogMSG2)) Then
	Call ActivityHistory(strLogMSG2, LogDivision, SessionUserID)
End If
%>
<Lists>
	<List>
		<Result><%= strResult %></Result>
		<InsertType><%= InsertType %></InsertType>
		<ReturnMSG><%= returnMSG %></ReturnMSG>
	</List>
</Lists>
</Metissoft>