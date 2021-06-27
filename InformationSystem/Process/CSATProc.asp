<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
'Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "CSAT"

'모집단위
Dim IDX						: IDX = fnR("IDX", 0)
Dim MYear					: MYear = fnR("MYear", Year(Date()))
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

'수능 환산 공식
Dim Formula1				: Formula1 = fnRF("Formula1")
Dim Formula2				: Formula2 = fnRF("Formula2")
Dim Formula3				: Formula3 = fnRF("Formula3")

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

Dim i, SubjectCodeTemp
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, strLogMSG2

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'체크박스가 선택된 갯수만큼 반복
For i = 0 To Ubound(SubjectCode)
	'///////////////////////////////////////////////////////////////////////
	'// insert or update 결정. 
	'// csat테이블에 SubjectCode가 있으면 Update, 없으면 Insert
	'///////////////////////////////////////////////////////////////////////	
	SQL = ""
	SQL = SQL & vbCrLf & "Select SubjectCode "
	SQL = SQL & vbCrLf & "from csat "
	SQL = SQL & vbCrLf & "where SubjectCode = ?; "

	Call objDB.sbSetArray("@SubjectCode", adVarchar, adParamInput, 50, SubjectCode(i))
	SubjectCodeTemp = SubjectCode(i)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)
	
	If Not IsNull(AryHash) then
		ProcessType = "Update"
	Else
		ProcessType = "Insert"
	End If

	'=============== 수능 환산 공식 입력 ===============

	'On Error Resume Next
	
	'//////////////////////////////////////////////////////////
	'// 수능 환산 공식
	'//////////////////////////////////////////////////////////
	if ProcessType = "Insert" then
		'// 입력 =================
		SQL = ""
		SQL = SQL & vbCrLf & "INSERT INTO csat ( "
		SQL = SQL & vbCrLf & "		MYear, SubjectCode "
		SQL = SQL & vbCrLf & "		,Formula1, Formula2, Formula3 "
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR "
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?"
		SQL = SQL & vbCrLf & "		, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
		SQL = SQL & vbCrLf & " ) "

		'insert일 때는 INPT입력
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		50,		SubjectCodeTemp) _
			, Array("@Formula1",				adVarchar,		adParamInput,		80,		Formula1) _
			, Array("@Formula2",				adVarchar,		adParamInput,		80,		Formula2) _
			, Array("@Formula3",				adVarchar,		adParamInput,		80,		Formula3) _
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
			strLogMSG = "수능공식설정 > "& InsertCnt &" 건의 평가비율이 등록되었습니다."
		Else
			strLogMSG = "수능공식설정 > "& MYear &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과 등 " & InsertCnt &" 건의 수능공식설정이 등록되었습니다."
		End If
		InsertType = "Insert"
		InsertCnt = InsertCnt + 1
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE csat SET "
		SQL = SQL & vbCrLf & "		  Formula1 = ? "
		SQL = SQL & vbCrLf & "		, Formula2 = ? "
		SQL = SQL & vbCrLf & "		, Formula3 = ? "
		SQL = SQL & vbCrLf & "		, UPDT_USID = ? "
		SQL = SQL & vbCrLf & "		, UPDT_DATE = getdate() "
		SQL = SQL & vbCrLf & "		, UPDT_ADDR = ? "
		SQL = SQL & vbCrLf & "		, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE SubjectCode = ? "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@Formula1",				adVarchar,		adParamInput,		80,		Formula1) _
			, Array("@Formula2",				adVarchar,		adParamInput,		80,		Formula2) _
			, Array("@Formula3",				adVarchar,		adParamInput,		80,		Formula3) _
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		50,		SubjectCodeTemp) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		'//////////////////////////////////////////////////////////////////////////////////////////////////
		'// 체크박스만 클릭하여 저장했을 때와 모집단위를 선택하여 저장했을 때를 비교하여 메세지 내용 등록
		'//////////////////////////////////////////////////////////////////////////////////////////////////
		If IsE(Division0Name) Then
			strLogMSG2 = "수능공식설정 > "& UpdateCnt &" 건의 평가비율이 수정되었습니다."
		Else
			strLogMSG2 = "수능공식설정 > "& MYear &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과 등 " & UpdateCnt &" 건의 수능공식설정이 수정되었습니다."
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
	returnMSG = "수능공식설정 완료"
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