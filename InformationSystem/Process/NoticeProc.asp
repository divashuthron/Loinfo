<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "NoticeProc"

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM, i

'공지사항
Dim IDX							:	IDX						 =	fnRF("IDX")							'인덱스
Dim Myear						:	Myear					 =	fnRF("Myear")						'사용년도
Dim Division                    :   Division                 =  FnRF("Department")                  '부서
Dim Title						:   Title					 =  FnRF("Title") 						'제목
Dim content1					:	content1				 =	fnRF("content1")					'내용1

'첨부파일
Dim filenameCnt					:   filenameCnt				 =  Request.Form("filename").Count
Dim aryfilename()
ReDim aryfilename(5)

if filenameCnt > 0 then
	for i = 1 to filenameCnt
		aryfilename(i)	= getQueryFilter(Request.Form("filename")(i))
	next
end If


'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'=============== 공지사항 등록 ===============

'On Error Resume Next

if ProcessType = "Insert" then
	'// 입력 =================
	SQL = ""
	SQL = SQL & vbCrLf & "INSERT INTO NoticeTable ( "
	SQL = SQL & vbCrLf & "		MYear, Division, Title, content1, file1, file2, file3, file4, file5  "
	SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
	SQL = SQL & vbCrLf & " ) VALUES ( "
	SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ?, ?, ?"
	SQL = SQL & vbCrLf & "		, ?, getdate(), ? "
	SQL = SQL & vbCrLf & " ) "

	'insert일 때는 INPT입력
	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@Division",				adVarchar,		adParamInput,		60,		Division) _
		, Array("@Title",					adVarchar,		adParamInput,		200,	Title) _
		, Array("@content1",				adVarchar,		adParamInput,		5000,	content1) _
		, Array("@file1",					adVarchar,		adParamInput,		200,	aryfilename(1)) _
		, Array("@file2",					adVarchar,		adParamInput,		200,	aryfilename(2)) _
		, Array("@file3",					adVarchar,		adParamInput,		200,	aryfilename(3)) _
		, Array("@file4",					adVarchar,		adParamInput,		200,	aryfilename(4)) _
		, Array("@file5",					adVarchar,		adParamInput,		200,	aryfilename(5)) _
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
	strLogMSG = "공지사항 > "& MYear &"학년도_"& Title &"가 등록되었습니다."
	InsertType = "Insert"
elseif ProcessType = "Update" then
	'// 수정 ================
	SQL = ""
	SQL = SQL & vbCrLf & "UPDATE NoticeTable SET "
	SQL = SQL & vbCrLf & "		MYear = ?,Division = ?,Title = ?,content1 = ? , file1 = ?, file2 = ?, file3 = ?, file4 = ?, file5 = ? "
	SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(), UPDT_ADDR = ?, InsertTime = getdate() "
	SQL = SQL & vbCrLf & "WHERE IDX = ? "

	'update일 때는 UPDT입력
	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@Division",				adVarchar,		adParamInput,		60,		Division) _
		, Array("@Title",					adVarchar,		adParamInput,		100,	Title) _
		, Array("@content1",				adVarchar,		adParamInput,		5000,	content1) _
		, Array("@file1",					adVarchar,		adParamInput,		200,	aryfilename(1)) _
		, Array("@file2",					adVarchar,		adParamInput,		200,	aryfilename(2)) _
		, Array("@file3",					adVarchar,		adParamInput,		200,	aryfilename(3)) _
		, Array("@file4",					adVarchar,		adParamInput,		200,	aryfilename(4)) _
		, Array("@file5",					adVarchar,		adParamInput,		200,	aryfilename(5)) _
		, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
		, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
		, Array("@IDX",						adVarchar,		adParamInput,		60,		IDX) _
	)

	'objDB.blnDebug = true
	Call objDB.sbExecSQL(SQL, arrParams)
	
	'////////////////////////////////////
	'// 수정 히스토리 
	'////////////////////////////////////
	strLogMSG = "공지사항 > "& MYear &"학년도_"& Title &"가 수정되었습니다."
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
	returnMSG = "공지사항 저장 완료"
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