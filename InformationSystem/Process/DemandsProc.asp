<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "DemandsProc"

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

'고지서
Dim IDX							:	IDX						 =	fnRF("IDX")							'인덱스
Dim Myear						:	Myear					 =	fnRF("Myear")						'고지서 사용년도
Dim Title																							'고지서 제목
Dim Division0                   :   Division0                =  FnRF("Division0")                   '모집시기
Dim State						:   State					 =  FnRF("State")				        '사용여부
'Dim ReceiptDate					:	ReceiptDate				 =	fnRF("ReceiptDate")					'환불일자
'Dim CheckTime					:	CheckTime				 =	fnRF("CheckTime")					'환불시간
Dim content1					:	content1				 =	fnRF("content1")					'내용1
Dim content2					:	content2				 =	fnRF("content2")					'내용2

Dim Option1, Option4, Option5
Dim Option6, Option7, Option8, Option9, Option10

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, i 

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'모집시기 한글
SQL = ""
SQL = SQL & vbCrLf & "SELECT SubcodeName as DivisionName "
SQL = SQL & vbCrLf & "FROM codeSub  "
SQL = SQL & vbCrLf & "Where subcode = ? "

Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 50, Division0)

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'제목계산
Title = AryHash(0).Item("DivisionName")  & " 유의사항"


'=============== 고지서 입력 ===============

'On Error Resume Next

if ProcessType = "DemandsInsert" then
	'// 입력 =================
	SQL = ""
	SQL = SQL & vbCrLf & "INSERT INTO DemandsTable ( "
	SQL = SQL & vbCrLf & "		MYear,Division0,Title,State,option1,option2,option3,option4,option5,option6,option7,option8,option9,option10,content1,content2  "
	SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
	SQL = SQL & vbCrLf & " ) VALUES ( "
	SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
	SQL = SQL & vbCrLf & " ) "

	'insert일 때는 INPT입력
	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
		, Array("@Title",					adVarchar,		adParamInput,		200,	Title) _
		, Array("@State",					adVarchar,		adParamInput,		4,		State) _
		, Array("@option1",					adVarchar,		adParamInput,		60,		option1) _
		, Array("@option2",					adVarchar,		adParamInput,		255,	ReceiptDate) _
		, Array("@option3",					adVarchar,		adParamInput,		255,	CheckTime) _
		, Array("@option4",					adVarchar,		adParamInput,		4,		option4) _
		, Array("@option5",					adVarchar,		adParamInput,		60,		option5) _
		, Array("@option6",					adVarchar,		adParamInput,		60,		option6) _
		, Array("@option7",					adVarchar,		adParamInput,		60,		option7) _
		, Array("@option8",					adVarchar,		adParamInput,		60,		option8) _
		, Array("@option9",					adVarchar,		adParamInput,		60,		option9) _
		, Array("@option10",				adVarchar,		adParamInput,		60,		option10) _
		, Array("@content1",				adVarchar,		adParamInput,		5000,	content1) _
		, Array("@content2",				adVarchar,		adParamInput,		5000,	content2) _
		, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
		, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
	)

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, arrParams)

	'SQL = " SELECT @@IDENTITY; "
	'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
	'IDX = CInt(aryList(0, 0))

	'////////////////////////////////////
	'// 등록 히스토리 
	'////////////////////////////////////
	strLogMSG = "합격자발표관리 > 유의사항 설정 > "& MYear &"학년도_"& Title &"이 등록되었습니다."
	InsertType = "Insert"
else
	'// 수정 ================
	SQL = ""
	SQL = SQL & vbCrLf & "UPDATE DemandsTable SET "
	SQL = SQL & vbCrLf & "		MYear = ?,Division0 = ?,Title = ?,State = ? "
	SQL = SQL & vbCrLf & "		,option1 = ?,option2 = ?,option3 = ?,option4 = ?,option5 = ?  "	
	SQL = SQL & vbCrLf & "		,option6 = ?,option7 = ?,option8 = ?,option9 = ?,option10 = ?,content1 = ?,content2 = ? "
	SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(), UPDT_ADDR = ?, InsertTime = getdate() "
	SQL = SQL & vbCrLf & "WHERE IDX = ? "

	'update일 때는 UPDT입력
	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
		, Array("@Title",					adVarchar,		adParamInput,		100,	Title) _
		, Array("@State",					adVarchar,		adParamInput,		4,		State) _
		, Array("@option1",					adVarchar,		adParamInput,		60,		option1) _
		, Array("@option2",					adVarchar,		adParamInput,		255,	ReceiptDate) _
		, Array("@option3",					adVarchar,		adParamInput,		255,	CheckTime) _
		, Array("@option4",					adVarchar,		adParamInput,		4,		option4) _
		, Array("@option5",					adVarchar,		adParamInput,		60,		option5) _
		, Array("@option6",					adVarchar,		adParamInput,		60,		option6) _
		, Array("@option7",					adVarchar,		adParamInput,		60,		option7) _
		, Array("@option8",					adVarchar,		adParamInput,		60,		option8) _
		, Array("@option9",					adVarchar,		adParamInput,		60,		option9) _
		, Array("@option10",				adVarchar,		adParamInput,		60,		option10) _
		, Array("@content1",				adVarchar,		adParamInput,		5000,	content1) _
		, Array("@content2",				adVarchar,		adParamInput,		5000,	content2) _
		, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
		, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
		, Array("@IDX",						adVarchar,		adParamInput,		60,		IDX) _
	)

	'objDB.blnDebug = true
	Call objDB.sbExecSQL(SQL, arrParams)
	
	'////////////////////////////////////
	'// 수정 히스토리 
	'////////////////////////////////////
	strLogMSG = "합격자발표관리 > 유의사항 설정 > "& MYear &"학년도_"& Title &"이 수정되었습니다."
	InsertType = "Update"
end If
'//////////////////////////////////////////////////////////
	


'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "유의사항 설정 완료"
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