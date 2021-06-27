<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
'Dim process					: process = fnR("process", "")
'Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim i
Dim LogDivision				: LogDivision = "StudentRecord"

'폼 카운터
Dim FormulaNumCnt			: FormulaNumCnt = Request.Form("FormulaNum").Count
Dim FormulaNameCnt			: FormulaNameCnt = Request.Form("FormulaName").Count
Dim Formula1Cnt				: Formula1Cnt = Request.Form("Formula1").Count
Dim Formula2Cnt				: Formula2Cnt = Request.Form("Formula2").Count
Dim Formula3Cnt				: Formula3Cnt = Request.Form("Formula3").Count
Dim Formula4Cnt				: Formula4Cnt = Request.Form("Formula4").Count
Dim Formula5Cnt				: Formula5Cnt = Request.Form("Formula5").Count
Dim INPT_USIDCnt			: INPT_USIDCnt = Request.Form("INPT_USID").Count

'폼 갯수만큼 배열생성
Dim aryFormulaNum, aryFormulaName, aryFormula1, aryFormula2, aryFormula3, aryFormula4, aryFormula5, aryINPT_USID
ReDim aryFormulaNum(FormulaNumCnt)
ReDim aryFormulaName(FormulaNameCnt)
ReDim aryFormula1(Formula1Cnt)
ReDim aryFormula2(Formula2Cnt)
ReDim aryFormula3(Formula3Cnt)
ReDim aryFormula4(Formula4Cnt)
ReDim aryFormula5(Formula5Cnt)
ReDim aryINPT_USID(INPT_USIDCnt)

'폼 넣기
if FormulaNumCnt > 0 Then
	for i = 1 to FormulaNumCnt
		aryFormulaNum(i)	= getQueryFilter(Request.Form("FormulaNum")(i)) 
		aryFormulaName(i)	= getQueryFilter(Request.Form("FormulaName")(i))
		aryFormula1(i)		= getQueryFilter(Request.Form("Formula1")(i))
		aryFormula2(i)		= getQueryFilter(Request.Form("Formula2")(i))
		aryFormula3(i)		= getQueryFilter(Request.Form("Formula3")(i))
		aryINPT_USID(i)		= getQueryFilter(Request.Form("INPT_USID")(i))
	next
end If

'예비4(사용 안 함)
if Formula4Cnt > 0 Then
	for i = 1 to Formula4Cnt
		aryFormula4(i)		= getQueryFilter(Request.Form("Formula4")(i))
	next
end If

'예비5(사용 안 함)
if Formula5Cnt > 0 Then
	for i = 1 to Formula5Cnt
		aryFormula5(i)		= getQueryFilter(Request.Form("Formula5")(i))
	next
end if

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

'DB, MSG
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, strLogMSG2

'DB open
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'///////////////////////////////////////////////////////////////////////
'// 공식종류가 있으면...(DB에서 1~6번 공식종류 고정으로 사용)
'// 공식종류가 늘어나거나, 줄어들 경우 DB수정
'///////////////////////////////////////////////////////////////////////
if FormulaNumCnt > 0 Then
	'공식종류 갯수만큼 반복(6번 고정)
	for i = 1 to FormulaNumCnt
		'최초입력이 있으면 update, 없으면 insert
		If aryINPT_USID(i) = null Or aryINPT_USID(i) = "" Then
			ProcessType = "Insert"
		Else
			ProcessType = "Update"
		End If

'=============== 생기부 공식 입력 ===============

	'On Error Resume Next
	
	'//////////////////////////////////////////////////////////
	'// 생기부 공식 입력
	'//////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "update StudentRecord "
		SQL = SQL & vbCrLf & "set Formula1 = ? "
		SQL = SQL & vbCrLf & ", Formula2 = ? "
		SQL = SQL & vbCrLf & ", Formula3 = ? "
		'예비(사용 안 함)
		'SQL = SQL & vbCrLf & ", Formula4 = ? "
		'SQL = SQL & vbCrLf & ", Formula5 = ? "
		SQL = SQL & vbCrLf & ", INPT_USID = ? "
		SQL = SQL & vbCrLf & ", INPT_DATE = getdate() "
		SQL = SQL & vbCrLf & ", INPT_ADDR = ? "
		SQL = SQL & vbCrLf & "where FormulaNum = ? "

		'insert일 때는 INPT입력
		arrParams = Array(_
			  Array("@Formula1",		adVarchar,		adParamInput,		80,		aryFormula1(i)) _
			, Array("@Formula2",		adVarchar,		adParamInput,		80,		aryFormula2(i)) _
			, Array("@Formula3",		adVarchar,		adParamInput,		80,		aryFormula3(i)) _
			, Array("@INPT_USID",		adVarchar,		adParamInput,		80,		INPT_USID) _
			, Array("@INPT_ADDR",		adVarchar,		adParamInput,		80,		INPT_ADDR) _
			, Array("@FormulaNum",		adInteger,		adParamInput,		0,		aryFormulaNum(i)) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		'SQL = " SELECT @@IDENTITY; "
		'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		'IDX = CInt(aryList(0, 0))
		
		strLogMSG = "생기부공식설정 > "& MYear &"_"& aryFormulaName(i) &"_공식1:"& aryFormula1(i) &"_공식2:"& aryFormula2(i) &"_공식3:"& aryFormula3(i) &"이 등록되었습니다."
		InsertType = "Insert"
	else
		'// 수정 =================
		SQL = ""
		SQL = SQL & vbCrLf & "update StudentRecord "
		SQL = SQL & vbCrLf & "set Formula1 = ? "
		SQL = SQL & vbCrLf & "	, Formula2 = ? "
		SQL = SQL & vbCrLf & "	, Formula3 = ? "
		'예비(사용 안 함)
		'SQL = SQL & vbCrLf & "	, Formula4 = ? "
		'SQL = SQL & vbCrLf & "	, Formula5 = ? "
		SQL = SQL & vbCrLf & "	, UPDT_USID = ? "
		SQL = SQL & vbCrLf & "	, UPDT_DATE = getdate() "
		SQL = SQL & vbCrLf & "	, UPDT_ADDR = ? "
		SQL = SQL & vbCrLf & "	, InsertTime = getdate() "
		SQL = SQL & vbCrLf & "where FormulaNum = ? "

		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@Formula1",		adVarchar,		adParamInput,		80,		aryFormula1(i)) _
			, Array("@Formula2",		adVarchar,		adParamInput,		80,		aryFormula2(i)) _
			, Array("@Formula3",		adVarchar,		adParamInput,		80,		aryFormula3(i)) _
			, Array("@UPDT_USID",		adVarchar,		adParamInput,		80,		UPDT_USID) _
			, Array("@UPDT_ADDR",		adVarchar,		adParamInput,		80,		UPDT_ADDR) _
			, Array("@FormulaNum",		adInteger,		adParamInput,		0,		aryFormulaNum(i)) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "생기부공식설정 > "& MYear &"_"& aryFormulaName(i) &"_공식1:"& aryFormula1(i) &"_공식2:"& aryFormula2(i) &"_공식3:"& aryFormula3(i) &"이 수정되었습니다."
		InsertType = "Update"
	end If
	'//////////////////////////////////////////////////////////		
	
	'// 로그기록
	Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

	next
end If

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "생기부공식설정 완료"
	'objDB.sbCommitTrans 
End If	

Set objDB  = Nothing
%>
<Lists>
	<List>
		<Result><%= strResult %></Result>
		<InsertType><%= InsertType %></InsertType>
		<ReturnMSG><%= returnMSG %></ReturnMSG>
	</List>
</Lists>
</Metissoft>