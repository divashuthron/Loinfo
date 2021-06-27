<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")

Dim IDX						: IDX = fnR("IDX", 0)
Dim MYear					: MYear = fnRF("MYear")
Dim Division				: Division = fnRF("Division")
Dim Subject					: Subject = fnRF("Subject")
Dim Division1				: Division1 = fnRF("Division1")
Dim Division2				: Division2 = fnRF("Division2")
Dim Degree					: Degree = fnRF("Degree")

Dim ApplyTitle				: ApplyTitle = fnRF("ApplyTitle")
Dim ApplyGroupCount			: ApplyGroupCount = fnRF("ApplyGroupCount")
Dim ApplyTotalNumber		: ApplyTotalNumber = fnRF("ApplyTotalNumber")

Dim ApplyStartDate			: ApplyStartDate = fnRF("ApplyStartDate")
Dim ApplyStartTime			: ApplyStartTime = fnRF("ApplyStartTime")
Dim ApplyEndDate			: ApplyEndDate = fnRF("ApplyEndDate")
Dim ApplyEndTime			: ApplyEndTime = fnRF("ApplyEndTime")

Dim GroupNameCnt			: GroupNameCnt = Request.Form("GroupName").Count
Dim RoomNameCnt				: RoomNameCnt = Request.Form("RoomName").Count

Dim ApplyPrintStartDate		: ApplyPrintStartDate = fnRF("ApplyPrintStartDate")
Dim ApplyPrintStartTime		: ApplyPrintStartTime = fnRF("ApplyPrintStartTime")
Dim ApplyPrintEndDate		: ApplyPrintEndDate = fnRF("ApplyPrintEndDate")
Dim ApplyPrintEndTime		: ApplyPrintEndTime = fnRF("ApplyPrintEndTime")

Dim InterviewStartDate		: InterviewStartDate = fnRF("InterviewStartDate")
Dim InterviewEndDate		: InterviewEndDate = fnRF("InterviewEndDate")
Dim InterviewDays			: InterviewDays = fnRF("InterviewDays")
'Dim InterviewDays			: InterviewDays = DateDiff("d",InterviewStartDate,InterviewEndDate)+1
Dim InterviewEtc			: InterviewEtc = fnRF("InterviewEtc")
Dim TShirtCheck				: TShirtCheck = fnRF("TShirtCheck")
Dim StandByRoom				: StandByRoom = fnRF("StandByRoom")

Dim InterviewDateCnt		: InterviewDateCnt = Request.Form("InterviewDate").Count
Dim TimeCodeCnt				: TimeCodeCnt = Request.Form("TimeCode").Count
Dim CheckTimeCnt			: CheckTimeCnt = Request.Form("CheckTime").Count
Dim InterviewStartTimeCnt	: InterviewStartTimeCnt = Request.Form("InterviewStartTime").Count
Dim InterviewEndTimeCnt		: InterviewEndTimeCnt = Request.Form("InterviewEndTime").Count
Dim QuorumCnt				: QuorumCnt = Request.Form("Quorum").Count

Dim State					: State = fnRF("State")
Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

If IsE(Degree) Then Degree = 0 End If
If IsE(ApplyGroupCount) Then ApplyGroupCount = 0 End If
If IsE(ApplyTotalNumber) Then ApplyTotalNumber = 0 End If
If IsE(InterviewDays) Then InterviewDays = 0 End if

Select Case process
	Case "RegSubject"
		Call setSubject()
	Case "getSubjectTimeSample"
		Call getSubjectTimeSample()
	Case ""
End Select

'=============== 학과 입력 ===============
Sub setSubject()
	'On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()
		
	'//////////////////////////////////////////////////////////
	'// 학과 기본 정보 관리
	'//////////////////////////////////////////////////////////
	if ProcessType = "SubjectInsert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO SubjectTable ( "
		SQL = SQL & vbCrLf & "		MYear, Division, Subject, Division1, Division2, Degree "
		SQL = SQL & vbCrLf & "		, ApplyTitle, ApplyGroupCount, ApplyTotalNumber "
		SQL = SQL & vbCrLf & "		, ApplyStartDate, ApplyStartTime, ApplyEndDate, ApplyEndTime "
		SQL = SQL & vbCrLf & "		, ApplyPrintStartDate, ApplyPrintStartTime, ApplyPrintEndDate, ApplyPrintEndTime "
		SQL = SQL & vbCrLf & "		, InterviewStartDate, InterviewEndDate, InterviewDays, InterviewEtc, TShirtCheck, StandByRoom "
		SQL = SQL & vbCrLf & "		, State, RegID "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ? "
		SQL = SQL & vbCrLf & " ) "

		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@Degree",					adInteger,		adParamInput,		0,		Degree) _
			, Array("@ApplyTitle",				adVarchar,		adParamInput,		255,	ApplyTitle) _
			, Array("@ApplyGroupCount",			adInteger,		adParamInput,		0,		ApplyGroupCount) _
			, Array("@ApplyTotalNumber",		adInteger,		adParamInput,		0,		ApplyTotalNumber) _
			, Array("@ApplyStartDate",			adVarchar,		adParamInput,		20,		ApplyStartDate) _
			, Array("@ApplyStartTime",			adVarchar,		adParamInput,		20,		ApplyStartTime) _
			, Array("@ApplyEndDate",			adVarchar,		adParamInput,		20,		ApplyEndDate) _
			, Array("@ApplyEndTime",			adVarchar,		adParamInput,		20,		ApplyEndTime) _
			, Array("@ApplyPrintStartDate",		adVarchar,		adParamInput,		20,		ApplyPrintStartDate) _
			, Array("@ApplyPrintStartTime",		adVarchar,		adParamInput,		20,		ApplyPrintStartTime) _
			, Array("@ApplyPrintEndDate",		adVarchar,		adParamInput,		20,		ApplyPrintEndDate) _
			, Array("@ApplyPrintEndTime",		adVarchar,		adParamInput,		20,		ApplyPrintEndTime) _
			, Array("@InterviewStartDate",		adVarchar,		adParamInput,		20,		InterviewStartDate) _
			, Array("@InterviewEndDate",		adVarchar,		adParamInput,		20,		InterviewEndDate) _
			, Array("@InterviewDays",			adInteger,		adParamInput,		0,		InterviewDays) _
			, Array("@InterviewEtc",			adVarchar,		adParamInput,		255,	InterviewEtc) _
			, Array("@TShirtCheck",				adChar,			adParamInput,		1,		TShirtCheck) _
			, Array("@StandByRoom",				adVarchar,		adParamInput,		50,		StandByRoom) _
			, Array("@State",					adChar,			adParamInput,		1,		State) _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		SQL = " SELECT @@IDENTITY; "
		aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		IDX = CInt(aryList(0, 0))
		
		strLogMSG = "학과관리  > "& MYear &"_"& Division &"_"& Subject &"_"& Degree &"차_"& ApplyTitle &" 학과가 입력 되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE SubjectTable SET "
		SQL = SQL & vbCrLf & "		MYear = ?, Division = ?, Subject = ?, Division1 = ?, Division2 = ?, Degree = ? "
		SQL = SQL & vbCrLf & "		, ApplyTitle = ?, ApplyGroupCount = ?, ApplyTotalNumber = ? "
		SQL = SQL & vbCrLf & "		, ApplyStartDate = ?, ApplyStartTime = ?, ApplyEndDate = ?, ApplyEndTime = ? "
		SQL = SQL & vbCrLf & "		, ApplyPrintStartDate = ?, ApplyPrintStartTime = ?, ApplyPrintEndDate = ?, ApplyPrintEndTime = ? "
		SQL = SQL & vbCrLf & "		, InterviewStartDate = ?, InterviewEndDate = ?, InterviewDays = ?, InterviewEtc = ?, TShirtCheck = ?, StandByRoom = ? "
		SQL = SQL & vbCrLf & "		, State = ?, EditID = ?,  EditDate = getdate() "
		SQL = SQL & vbCrLf & " WHERE IDX = ?; "
		
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@Degree",					adInteger,		adParamInput,		0,		Degree) _
			, Array("@ApplyTitle",				adVarchar,		adParamInput,		255,	ApplyTitle) _
			, Array("@ApplyGroupCount",			adInteger,		adParamInput,		0,		ApplyGroupCount) _
			, Array("@ApplyTotalNumber",		adInteger,		adParamInput,		0,		ApplyTotalNumber) _
			, Array("@ApplyStartDate",			adVarchar,		adParamInput,		20,		ApplyStartDate) _
			, Array("@ApplyStartTime",			adVarchar,		adParamInput,		20,		ApplyStartTime) _
			, Array("@ApplyEndDate",			adVarchar,		adParamInput,		20,		ApplyEndDate) _
			, Array("@ApplyEndTime",			adVarchar,		adParamInput,		20,		ApplyEndTime) _
			, Array("@ApplyPrintStartDate",		adVarchar,		adParamInput,		20,		ApplyPrintStartDate) _
			, Array("@ApplyPrintStartTime",		adVarchar,		adParamInput,		20,		ApplyPrintStartTime) _
			, Array("@ApplyPrintEndDate",		adVarchar,		adParamInput,		20,		ApplyPrintEndDate) _
			, Array("@ApplyPrintEndTime",		adVarchar,		adParamInput,		20,		ApplyPrintEndTime) _
			, Array("@InterviewStartDate",		adVarchar,		adParamInput,		20,		InterviewStartDate) _
			, Array("@InterviewEndDate",		adVarchar,		adParamInput,		20,		InterviewEndDate) _
			, Array("@InterviewDays",			adInteger,		adParamInput,		0,		InterviewDays) _
			, Array("@InterviewEtc",			adVarchar,		adParamInput,		255,	InterviewEtc) _
			, Array("@TShirtCheck",				adChar,			adParamInput,		1,		TShirtCheck) _
			, Array("@StandByRoom",				adVarchar,		adParamInput,		50,		StandByRoom) _
			, Array("@State",					adChar,			adParamInput,		1,		State) _
			, Array("@EditID",					adVarchar,		adParamInput,		50,		SessionUserID) _
			, Array("@IDX",						adInteger,		adParamInput,		0,		IDX) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "학과관리  > "& MYear &"_"& Division &"_"& Subject &"_"& Degree &"차_"& ApplyTitle &" 학과가 수정 되었습니다."
		InsertType = "Update"
	end If
	'//////////////////////////////////////////////////////////

	'//////////////////////////////////////////////////////////
	'// 평가조 관리
	'// 기존 입력되어 있던 내역 싹 지우고 다시 입력
	'//////////////////////////////////////////////////////////
	If IDX <> 0 Then
		'// 평가조 관리 기존 내역 삭제
		SQL = "DELETE FROM SubjectGroup WHERE 1 = 1 AND SubjectIDX = ?;"
		'SQL = SQL & vbCrLf & "	AND MYear = ? "
		'SQL = SQL & vbCrLf & "	AND Division = ? "
		'SQL = SQL & vbCrLf & "	AND Subject = ? "
		'SQL = SQL & vbCrLf & "	AND Division1 = ? "
		'SQL = SQL & vbCrLf & "	AND Division2 = ? "
		'SQL = SQL & vbCrLf & "	AND Degree = ? "

		Call objDB.sbSetArray("@SubjectIDX", adInteger, adParamInput, 0, IDX)
		
		'objDB.blnDebug = True
		arrParams = objDB.fnGetArray
		Call objDB.sbExecSQL(SQL, arrParams)

		'// 평가조 관리 신규 입력
		if GroupNameCnt > 0 then
			For intNUM = 1 To GroupNameCnt
				If Not(IsE(Request.Form("GroupName")(intNUM))) Then
					SQL = ""
					SQL = SQL & vbCrLf & "INSERT INTO SubjectGroup ( "
					SQL = SQL & vbCrLf & "		SubjectIDX, MYear, Division, Subject, Division1, Division2, Degree "
					SQL = SQL & vbCrLf & "		, GroupName, RoomName, RegID "
					SQL = SQL & vbCrLf & " ) VALUES ( "
					SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
					SQL = SQL & vbCrLf & "		, ?, ?, ? "
					SQL = SQL & vbCrLf & " ); "

					'adDate, adLongVarChar, adVarchar, adInteger, adChar

					arrParams = Array(_
						  Array("@SubjectIDX",				adInteger,		adParamInput,		0,		IDX) _
						, Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
						, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
						, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
						, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
						, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
						, Array("@Degree",					adInteger,		adParamInput,		50,		Degree) _
						, Array("@GroupName",				adVarchar,		adParamInput,		50,		Request.Form("GroupName")(intNUM)) _
						, Array("@RoomName",				adVarchar,		adParamInput,		50,		Request.Form("RoomName")(intNUM)) _
						, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
					)
					'objDB.blnDebug = True
					Call objDB.sbExecSQL(SQL, arrParams)
				End If
			Next
		End If
	End If
	'//////////////////////////////////////////////////////////

	'//////////////////////////////////////////////////////////
	'// 면접교시 관리
	'// 기존 입력되어 있던 내역 싹 지우고 다시 입력
	'//////////////////////////////////////////////////////////
	If IDX <> 0 Then
		'// 면접교시 관리 기존 내역 삭제
		SQL = "DELETE FROM SubjectTime WHERE 1 = 1 AND SubjectIDX = ?;"
		'SQL = SQL & vbCrLf & "	AND MYear = ? "
		'SQL = SQL & vbCrLf & "	AND Division = ? "
		'SQL = SQL & vbCrLf & "	AND Subject = ? "
		'SQL = SQL & vbCrLf & "	AND Division1 = ? "
		'SQL = SQL & vbCrLf & "	AND Division2 = ? "
		'SQL = SQL & vbCrLf & "	AND Degree = ? "

		Call objDB.sbSetArray("@SubjectIDX", adInteger, adParamInput, 0, IDX)
		
		'objDB.blnDebug = True
		arrParams = objDB.fnGetArray
		Call objDB.sbExecSQL(SQL, arrParams)

		'Dim InterviewDateCnt		: InterviewDateCnt = Request.Form("InterviewDate").Count
		'Dim TimeCodeCnt			: TimeCodeCnt = Request.Form("TimeCode").Count
		'Dim CheckTimeCnt			: CheckTimeCnt = Request.Form("CheckTime").Count
		'Dim InterviewStartTimeCnt	: InterviewStartTimeCnt = Request.Form("InterviewStartTime").Count
		'Dim InterviewEndTimeCnt	: InterviewEndTimeCnt = Request.Form("InterviewEndTime").Count
		'Dim QuorumCnt				: QuorumCnt = Request.Form("Quorum").Count

		'// 면접교시 관리 신규 입력
		if TimeCodeCnt > 0 then
			For intNUM = 1 To TimeCodeCnt
				If Not(IsE(Request.Form("TimeCode")(intNUM))) Then
					SQL = ""
					SQL = SQL & vbCrLf & "INSERT INTO SubjectTime ( "
					SQL = SQL & vbCrLf & "		SubjectIDX, MYear, Division, Subject, Division1, Division2, Degree "
					SQL = SQL & vbCrLf & "		, InterviewDate, TimeCode, CheckTime, InterviewStartTime, InterviewEndTime, Quorum, WeekName "
					SQL = SQL & vbCrLf & "		, RegID "
					SQL = SQL & vbCrLf & " ) VALUES ( "
					SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
					SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ?, ? "
					SQL = SQL & vbCrLf & "		, ? "
					SQL = SQL & vbCrLf & " ); "

					'adDate, adLongVarChar, adVarchar, adInteger, adChar

					arrParams = Array(_
						  Array("@SubjectIDX",				adInteger,		adParamInput,		0,		IDX) _
						, Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
						, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
						, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
						, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
						, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
						, Array("@Degree",					adInteger,		adParamInput,		0,		Degree) _
						, Array("@InterviewDate",			adVarchar,		adParamInput,		10,		Request.Form("InterviewDate")(intNUM)) _
						, Array("@TimeCode",				adVarchar,		adParamInput,		20,		Request.Form("TimeCode")(intNUM)) _
						, Array("@CheckTime",				adVarchar,		adParamInput,		5,		Request.Form("CheckTime")(intNUM)) _
						, Array("@InterviewStartTime",		adVarchar,		adParamInput,		5,		Request.Form("InterviewStartTime")(intNUM)) _
						, Array("@InterviewEndTime",		adVarchar,		adParamInput,		5,		Request.Form("InterviewEndTime")(intNUM)) _
						, Array("@Quorum",					adInteger,		adParamInput,		0,		Request.Form("Quorum")(intNUM)) _
						, Array("@WeekName",				adVarchar,		adParamInput,		10,		getWeekName(Request.Form("InterviewDate")(intNUM))) _
						, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
					)
					'objDB.blnDebug = True
					Call objDB.sbExecSQL(SQL, arrParams)
				End If
			Next
		End If
	End If
	'//////////////////////////////////////////////////////////
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "학과 저장 완료"
		objDB.sbCommitTrans 
	End If	

	Set objDB  = Nothing

	'// 로그기록
	Call ActivityHistory(strLogMSG, SessionUserID)
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

Sub getSubjectTimeSample()
	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB

	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "	TimeCode, CheckTime, InterviewStartTime, InterviewEndTime "
	SQL = SQL & vbCrLf & "FROM SubjectTimeSample AS A " 
	SQL = SQL & vbCrLf & "WHERE 1 = 1 "
	SQL = SQL & vbCrLf & "ORDER BY Step ASC;"

	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	Set objDB  = Nothing
%>
<Lists>
<%
	For i = 0 To datediff("D", InterviewStartDate, InterviewEndDate)
%>
	<InterviewtDate><%= DateAdd("d", i, InterviewStartDate) %></InterviewtDate>
<%
	Next

	If Not IsNull(AryHash) Then
		For i = 0 to ubound(AryHash,1)
%>
	<List>
		<TimeCode><%= AryHash(i).Item("TimeCode") %></TimeCode>
		<CheckTime><%= AryHash(i).Item("CheckTime") %></CheckTime>
		<InterviewStartTime><%= AryHash(i).Item("InterviewStartTime") %></InterviewStartTime>
		<InterviewEndTime><%= AryHash(i).Item("InterviewEndTime") %></InterviewEndTime>
	</List>
<%
		Next
	end if
%>
</Lists>
<%
End Sub
%>
</Metissoft>