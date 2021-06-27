<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim LogDivision				: LogDivision = "SubjectExcelSave"

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim InsertType

'건수
Dim SubjectCount			: SubjectCount = fnRF("SubjectCount")

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'=======================================================
'=====모집학과 엑셀등록=================================
'=======================================================

'On Error Resume Next

'// 입력 =================
SQL = ""
SQL = SQL & vbCrLf & "INSERT INTO SubjectTable (MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix,RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11,INPT_USID,INPT_DATE,INPT_ADDR) "
SQL = SQL & vbCrLf & "Select MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix,RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11,INPT_USID,INPT_DATE,INPT_ADDR "
SQL = SQL & vbCrLf & "From ##SubjectTableTemp "

'objDB.blnDebug = True
Call objDB.sbExecSQL(SQL, arrParams)

'SQL = " SELECT @@IDENTITY; "
'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
'IDX = CInt(aryList(0, 0))

strLogMSG = "모집단위등록 > " & SubjectCount & "건의 모집단위가 엑셀로 등록되었습니다."
InsertType = "Insert"
	
'=======================================================
'=====모집학과 엑셀등록 끝==============================
'=======================================================

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "모집단위 엑셀 저장 완료"
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