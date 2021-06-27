<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
'// ==============================================================================================================
'// 동석차순위 수정 순서
'// ==============================================================================================================
'// 1.		동석차 순위 업데이트
'// 2.		석차 생성(1순위 통합점수, 2순위 독석차 순위)
'// 3.		모집단위 별 모집인원으로 합/불여부 생성
'// 4.		불합격에 대한 예비석차 생성
'// ==============================================================================================================

'On Error Resume Next

'// ==============================================================================================================
'// 변수선언
'// ==============================================================================================================

'1.로그구분
Dim LogDivision				: LogDivision = "DrawStandingSet"
Dim strResult				: strResult = "failure"
Dim returnMSG
Dim InsertType
Dim strLogMSG

'2.대상구문
Dim BasicStudent
Dim BasicMYear						: BasicMYear = fnRF("DrawMyear")
Dim BasicDivision0					: BasicDivision0 = fnRF("DrawDivision0")
Dim BasicSubject					: BasicSubject = fnRF("DrawSubject")
Dim BasicDivision1					: BasicDivision1 = fnRF("DrawDivision1")
Dim DrawStudentNumber				: DrawStudentNumber = fnRF("DrawStudentNumber")
Dim DrawStanding					: DrawStanding = fnRF("DrawStanding")

'5.db변수
Dim objDB, SQL, AryHash, arrParams

'// ==============================================================================================================
'// 변수선언 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 1. 동석차 순위 업데이트 
'// ==============================================================================================================

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'// 수정 =================
SQL = ""
SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
SQL = SQL & vbCrLf & "SET	 DrawStanding=? "
SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

'Update일 때는 UPDT입력
arrParams = Array(_
	  Array("@DrawStanding",			adDouble,		adParamInput,		0,		DrawStanding) _

	, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
	, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
	, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		DrawStudentNumber) _
)

'objDB.blnDebug = True
Call objDB.sbExecSQL(SQL, arrParams)

'// ==============================================================================================================
'// 1. 동석차 순위 업데이트 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 2. 석차 생성(1순위 통합점수, 2순위 동석차 순위) 
'// ==============================================================================================================

'/////////// 학과 + 구분1 리스트 (리스트 별 정원) ///////////
SQL = ""
SQL = SQL & vbCrLf & " select SubjectCode, Subject, Division1, Quorum, QuorumFix "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Division0', Division0) AS DivisionName  "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Subject', Subject) AS SubjectName  "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Division1', Division1) AS Division1Name  "
SQL = SQL & vbCrLf & " from SubjectTable "
SQL = SQL & vbCrLf & " Where 1=1 "                 
SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"         
SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'" 
SQL = SQL & vbCrLf & " and Subject = '" & BasicSubject & "'" 
SQL = SQL & vbCrLf & " and Division1 = '" & BasicDivision1 & "'" 
SQL = SQL & vbCrLf & " group by SubjectCode, Division0, Subject, Division1, Quorum, QuorumFix "

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

If Not(isnull(AryHash)) Then
	Quorum				=	AryHash(0).Item("Quorum")		'// 모집인원
	QuorumFix			=	AryHash(0).Item("QuorumFix")	'// 입학정원

	'/////////// 모집단위별 석차 계산 ///////////
	'/////////// 1순위 totScore 2순위 DrawStanding ///////////		
	SQL = ""
	SQL = SQL & vbCrLf & " update ChangeScoreTable "
	SQL = SQL & vbCrLf & " set Standing = b.row_num "
	SQL = SQL & vbCrLf & " from ChangeScoreTable as a, "
	SQL = SQL & vbCrLf & "		(select row_number() over (order by totScore desc, DrawStanding desc) as row_num, StudentNumber "
	SQL = SQL & vbCrLf & "		 from ChangeScoreTable  "
	SQL = SQL & vbCrLf & "		 Where 1=1  "
	SQL = SQL & vbCrLf & "		 and MYear = '" & BasicMYear & "'"  
	SQL = SQL & vbCrLf & "		 and Division0 = '" & BasicDivision0 & "'" 
	SQL = SQL & vbCrLf & "		 and Subject = '" & BasicSubject & "'" 
	SQL = SQL & vbCrLf & "		 and Division1 = '" & BasicDivision1 & "') as b "
	SQL = SQL & vbCrLf & " where a.StudentNumber = b.StudentNumber "

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, null)

'// ==============================================================================================================
'// 2. 석차 생성(1순위 통합점수, 2순위 독석차 순위) 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 3. 모집단위 별 모집인원으로 합/불여부 생성
'// ==============================================================================================================

	'/////////// 모집단위별 합/불 계산 ///////////
	SQL = ""
	SQL = SQL & vbCrLf & " update ChangeScoreTable "
	SQL = SQL & vbCrLf & " set Result = '합격' "
	SQL = SQL & vbCrLf & " where Standing <= '" & Quorum & "'" 
	SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"  
	SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'" 
	SQL = SQL & vbCrLf & " and Subject = '" & BasicSubject & "'" 
	SQL = SQL & vbCrLf & " and Division1 = '" & BasicDivision1 & "'"

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, null)

	SQL = ""
	SQL = SQL & vbCrLf & " update ChangeScoreTable "
	SQL = SQL & vbCrLf & " set Result = '불합격' "
	SQL = SQL & vbCrLf & " where Standing > '" & Quorum & "'" 
	SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"  
	SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'" 
	SQL = SQL & vbCrLf & " and Subject = '" & BasicSubject & "'" 
	SQL = SQL & vbCrLf & " and Division1 = '" & BasicDivision1 & "'"

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, null)

'// ==============================================================================================================
'// 3. 모집단위 별 모집인원으로 합/불여부 생성 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 4. 불합격에 대한 예비석차 생성 
'// ==============================================================================================================

	'/////////// 불합격자 예비석차 계산 ///////////
	SQL = ""
	SQL = SQL & vbCrLf & " update ChangeScoreTable "
	SQL = SQL & vbCrLf & " set BackupStanding = b.row_num "
	SQL = SQL & vbCrLf & " from ChangeScoreTable as a, "
	SQL = SQL & vbCrLf & "		(select row_number() over (order by totScore desc, DrawStanding desc) as row_num, StudentNumber "
	SQL = SQL & vbCrLf & "		 from ChangeScoreTable  "
	SQL = SQL & vbCrLf & "		 Where 1=1  "
	SQL = SQL & vbCrLf & "		 and MYear = '" & BasicMYear & "'"  
	SQL = SQL & vbCrLf & "		 and Division0 = '" & BasicDivision0 & "'" 
	SQL = SQL & vbCrLf & "		 and Subject = '" & BasicSubject & "'" 
	SQL = SQL & vbCrLf & "		 and Division1 = '" & BasicDivision1 & "'"
	SQL = SQL & vbCrLf & "		 and Result = '불합격') as b "
	SQL = SQL & vbCrLf & " where a.StudentNumber = b.StudentNumber "
	
	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, null)
	
End If
Set objDB  = Nothing

'// ==============================================================================================================
'// 4. 불합격에 대한 예비석차 생성 끝 
'// ==============================================================================================================

'// ==============================================================================================================
'// 히스토리
'// ==============================================================================================================

InsertType = "insert"
strLogMSG = "합격자발표관리 > " & BasicMYear & "학년도 " & AryHash(0).Item("DivisionName") & "의 " & DrawStudentNumber & "지원자 동석차 순위가 수정 처리되었습니다."

'// 로그기록
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

'// ==============================================================================================================
'// 히스토리 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 트랜젝션 처리
'// ==============================================================================================================

If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "동석차 순위 수정 완료"
	'objDB.sbCommitTrans 
End If	

'// ==============================================================================================================
'// 트랜젝션 처리 끝
'// ==============================================================================================================
%>

<Lists>
	<List>
		<Result><%= strResult %></Result>
		<InsertType><%= InsertType %></InsertType>
		<ReturnMSG><%= returnMSG %></ReturnMSG>
	</List>
</Lists>
</Metissoft>