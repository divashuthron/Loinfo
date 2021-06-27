<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "Student"

Dim IDX						: IDX = fnR("IDX", 0)
Dim MYear					: MYear = fnRF("MYear")
Dim Division				: Division = fnRF("Division")
Dim Subject					: Subject = fnRF("Subject")
Dim Division1				: Division1 = fnRF("Division1")
Dim Division2				: Division2 = fnRF("Division2")

Dim StudentNumber			: StudentNumber = fnRF("StudentNumber")
Dim StudentName				: StudentName = fnRF("StudentName")
Dim InterviewNumber			: InterviewNumber = fnRF("InterviewNumber")
Dim TSize					: TSize = fnRF("TSize")
Dim HighSchool				: HighSchool = fnRF("HighSchool")
Dim Birthday				: Birthday = fnRF("Birthday")
Dim Sex						: Sex = fnRF("Sex")
Dim Tel1					: Tel1 = fnRF("Tel1")
Dim Tel2					: Tel2 = fnRF("Tel2")
Dim Tel3					: Tel3 = fnRF("Tel3")
Dim EnglishPoint			: EnglishPoint = fnRF("EnglishPoint")

Dim State					: State = fnRF("State")
Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

Select Case process
	Case "RegStudnet"
		Call setStudent()
	Case ""
End Select

'=============== 학생 입력 ===============
Sub setStudent()
	'On Error Resume Next

	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	objDB.sbBeginTrans()
		
	'//////////////////////////////////////////////////////////
	'// 학생 정보 관리
	'//////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO StudentTable ( "
		SQL = SQL & vbCrLf & "		MYear, Division, Subject, Division1, Division2 "
		SQL = SQL & vbCrLf & "		, StudentNumber, StudentName, InterviewNumber, TSize "
		SQL = SQL & vbCrLf & "		, HighSchool, Birthday, Sex, Tel1, Tel2, Tel3, EnglishPoint "
		SQL = SQL & vbCrLf & "		, State, RegID "
		SQL = SQL & vbCrLf & "		, StudentClassify "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ? "
		SQL = SQL & vbCrLf & "		, ? "
		SQL = SQL & vbCrLf & " ) "

		'adDate, adLongVarChar, adVarchar, adInteger, adChar

		'// 학생마다 수험번호 외에 고유한 번호룰 부여 StudentTable > StudentClassify 입력 (StudentClassify : 이름_생년월일_전화번호1)
		'// 고유번호는 최초 지원자정보 입력 시 생성되므로 중간에 전화번호등의 정보가 바뀐다고 고유번호를 바꿀 필요 없음
		'// SessionStudentClassify 학생의 전체 지원목록을 가져오는 용도와 로그인체크 용도로만 사용
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@StudentNumber",			adVarchar,		adParamInput,		50,		StudentNumber) _
			, Array("@StudentName",				adVarchar,		adParamInput,		50,		StudentName) _
			, Array("@InterviewNumber",			adVarchar,		adParamInput,		50,		InterviewNumber) _
			, Array("@TSize",					adVarchar,		adParamInput,		10,		TSize) _
			, Array("@HighSchool",				adVarchar,		adParamInput,		255,	HighSchool) _
			, Array("@Birthday",				adVarchar,		adParamInput,		10,		Birthday) _
			, Array("@Sex",						adVarchar,		adParamInput,		4,		Sex) _
			, Array("@Tel1",					adVarchar,		adParamInput,		50,		Tel1) _
			, Array("@Tel2",					adVarchar,		adParamInput,		50,		Tel2) _
			, Array("@Tel3",					adVarchar,		adParamInput,		50,		Tel3) _
			, Array("@EnglishPoint",			adVarchar,		adParamInput,		50,		EnglishPoint) _
			, Array("@State",					adChar,			adParamInput,		1,		State) _
			, Array("@RegID",					adVarchar,		adParamInput,		50,		SessionUserID) _
			, Array("@StudentClassify",			adVarchar,		adParamInput,		50,		StudentName &"_"& Birthday &"_"& Tel1) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		SQL = " SELECT @@IDENTITY; "
		aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		IDX = CInt(aryList(0, 0))
		
		strLogMSG = "지원자관리  > "& StudentNumber &"_"& StudentName &" 학생이 입력 되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE StudentTable SET "
		SQL = SQL & vbCrLf & "		MYear = ?, Division = ?, Subject = ?, Division1 = ?, Division2 = ? "
		SQL = SQL & vbCrLf & "		, StudentNumber = ?, StudentName = ?, InterviewNumber = ?, TSize = ? "
		SQL = SQL & vbCrLf & "		, HighSchool = ?, Birthday = ?, Sex = ?, Tel1 = ?, Tel2 = ?, Tel3 = ?, EnglishPoint = ? "
		SQL = SQL & vbCrLf & "		, State = ?, EditID = ?,  EditDate = getdate() "
		SQL = SQL & vbCrLf & " WHERE IDX = ?; "
		
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@Division",				adVarchar,		adParamInput,		50,		Division) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		50,		Division2) _
			, Array("@StudentNumber",			adVarchar,		adParamInput,		50,		StudentNumber) _
			, Array("@StudentName",				adVarchar,		adParamInput,		50,		StudentName) _
			, Array("@InterviewNumber",			adVarchar,		adParamInput,		50,		InterviewNumber) _
			, Array("@TSize",					adVarchar,		adParamInput,		10,		TSize) _
			, Array("@HighSchool",				adVarchar,		adParamInput,		255,	HighSchool) _
			, Array("@Birthday",				adVarchar,		adParamInput,		10,		Birthday) _
			, Array("@Sex",						adVarchar,		adParamInput,		4,		Sex) _
			, Array("@Tel1",					adVarchar,		adParamInput,		50,		Tel1) _
			, Array("@Tel2",					adVarchar,		adParamInput,		50,		Tel2) _
			, Array("@Tel3",					adVarchar,		adParamInput,		50,		Tel3) _
			, Array("@EnglishPoint",			adVarchar,		adParamInput,		50,		EnglishPoint) _
			, Array("@State",					adChar,			adParamInput,		1,		State) _
			, Array("@EditID",					adVarchar,		adParamInput,		50,		SessionUserID) _
			, Array("@IDX",						adInteger,		adParamInput,		0,		IDX) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "지원자관리  > "& StudentNumber &"_"& StudentName &" 학생이 수정 되었습니다."
		InsertType = "Update"
	end If
	'//////////////////////////////////////////////////////////
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "지원자 저장 완료"
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