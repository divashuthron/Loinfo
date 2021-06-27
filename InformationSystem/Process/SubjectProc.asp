<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "Subject"

'모집단위 
Dim IDX						: IDX = fnR("IDX", 0)
Dim MYear					: MYear = fnRF("MYear")
Dim Division0				: Division0 = fnRF("Division0")
Dim Subject					: Subject = fnRF("Subject")
Dim Division1				: Division1 = fnRF("Division1")
Dim Division2				: Division2 = fnRF("Division2")
Dim Division3				: Division3 = fnRF("Division3")
Dim SubjectCode				: SubjectCode = Subject + Division0 + Division1 '오산대 모집단위코드 조합 : 학과코드 + 모집시기 + 전형
Dim Quorum    				: Quorum = getIntParameter(fnR("Quorum", 0), 0)    
Dim QuorumFix 				: QuorumFix = getIntParameter(fnR("QuorumFix", 0), 0)  

'히스토리용(한글)
Dim Division0Name			: Division0Name = fnRF("Division0Name")
Dim SubjectName				: SubjectName = fnRF("SubjectName")
Dim Division1Name			: Division1Name = fnRF("Division1Name")
Dim Division2Name			: Division2Name = fnRF("Division2Name")
Dim Division3Name			: Division3Name = fnRF("Division3Name")

'등록금
Dim RF1       				: RF1 = getIntParameter(fnR("RF1", 0), 0) 
Dim RF2       				: RF2 = getIntParameter(fnR("RF2", 0), 0)  
Dim RF4          			: RF4 = getIntParameter(fnR("RF4", 0), 0)
Dim RF6          			: RF6 = getIntParameter(fnR("RF6", 0), 0)
Dim RF7          			: RF7 = getIntParameter(fnR("RF7", 0), 0)
Dim RF8          			: RF8 = getIntParameter(fnR("RF8", 0), 0)
Dim RF9          			: RF9 = getIntParameter(fnR("RF9", 0), 0)

'등록금 계산
Dim RF3       				: RF3 = RF1 + RF2 
Dim RF10         			: RF10 = RF4 + RF7
Dim RF5          			: RF5 = RF3 + RF9 - RF8 - RF6	
Dim RF11         			: RF11 = RF5 + RF6 + RF10

'기타(예비)
Dim Etc1         			: Etc1 = fnRF("Etc1")
Dim Etc2         			: Etc2 = fnRF("Etc2")
Dim Etc3         			: Etc3 = fnRF("Etc3")
Dim Etc4         			: Etc4 = fnRF("Etc4")
Dim Etc5         			: Etc5 = fnRF("Etc5")
Dim Etc6         			: Etc6 = fnRF("Etc6")
Dim Etc7         			: Etc7 = fnRF("Etc7")
Dim Etc8         			: Etc8 = fnRF("Etc8")
Dim Etc9         			: Etc9 = fnRF("Etc9")
Dim Etc10        			: Etc10 = fnRF("Etc10")

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

Select Case process
	Case "RegSubject"
		Call setSubject()
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
		SQL = SQL & vbCrLf & "		MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix"
		SQL = SQL & vbCrLf & "		,RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11"
		SQL = SQL & vbCrLf & "		,INPT_USID,INPT_DATE,INPT_ADDR"
		SQL = SQL & vbCrLf & " ) VALUES ( "
		SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
		SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, getdate(), ? "
		SQL = SQL & vbCrLf & " ) "

		'insert일 때는 INPT입력
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		20,		SubjectCode) _
			, Array("@Division0",				adVarchar,		adParamInput,		20,		Division0) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		20,		Division2) _
			, Array("@Division3",				adVarchar,		adParamInput,		20,		Division3) _
			, Array("@Quorum",					adInteger,		adParamInput,		0,		Quorum) _
			, Array("@QuorumFix",				adInteger,		adParamInput,		0,		QuorumFix) _
			, Array("@RF1",						adInteger,		adParamInput,		0,		RF1) _
			, Array("@RF2",						adInteger,		adParamInput,		0,		RF2) _
			, Array("@RF3",						adInteger,		adParamInput,		0,		RF3) _
			, Array("@RF4",						adInteger,		adParamInput,		0,		RF4) _
			, Array("@RF5",						adInteger,		adParamInput,		0,		RF5) _
			, Array("@RF6",						adInteger,		adParamInput,		0,		RF6) _
			, Array("@RF7",						adInteger,		adParamInput,		0,		RF7) _
			, Array("@RF8",						adInteger,		adParamInput,		0,		RF8) _
			, Array("@RF9",						adInteger,		adParamInput,		0,		RF9) _
			, Array("@RF10",					adInteger,		adParamInput,		0,		RF10) _
			, Array("@RF11",					adInteger,		adParamInput,		0,		RF11) _
			, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
			, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)

		SQL = " SELECT @@IDENTITY; "
		aryList = objDB.fnExecSQLGetRows(SQL, nothing)
		IDX = CInt(aryList(0, 0))
		
		strLogMSG = "모집단위관리  > "& MYear &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과가 입력 되었습니다."
		InsertType = "Insert"
	else
		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE SubjectTable SET "
		SQL = SQL & vbCrLf & "		MYear = ?,SubjectCode = ?,Division0 = ?,Subject = ?,Division1 = ?,Division2 = ?,Division3 = ?,Quorum = ?,QuorumFix = ?"
		SQL = SQL & vbCrLf & "		,RF1 = ?,RF2 = ?,RF3 = ?,RF4 = ?,RF5 = ?,RF6 = ?,RF7 = ?,RF8 = ?,RF9 = ?,RF10 = ?,RF11 = ?"
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate()"
		SQL = SQL & vbCrLf & " WHERE IDX = ?; "
		
		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
			, Array("@SubjectCode",				adVarchar,		adParamInput,		20,		SubjectCode) _
			, Array("@Division0",				adVarchar,		adParamInput,		20,		Division0) _
			, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
			, Array("@Division1",				adVarchar,		adParamInput,		50,		Division1) _
			, Array("@Division2",				adVarchar,		adParamInput,		20,		Division2) _
			, Array("@Division3",				adVarchar,		adParamInput,		20,		Division3) _
			, Array("@Quorum",					adInteger,		adParamInput,		0,		Quorum) _
			, Array("@QuorumFix",				adInteger,		adParamInput,		0,		QuorumFix) _
			, Array("@RF1",						adInteger,		adParamInput,		0,		RF1) _
			, Array("@RF2",						adInteger,		adParamInput,		0,		RF2) _
			, Array("@RF3",						adInteger,		adParamInput,		0,		RF3) _
			, Array("@RF4",						adInteger,		adParamInput,		0,		RF4) _
			, Array("@RF5",						adInteger,		adParamInput,		0,		RF5) _
			, Array("@RF6",						adInteger,		adParamInput,		0,		RF6) _
			, Array("@RF7",						adInteger,		adParamInput,		0,		RF7) _
			, Array("@RF8",						adInteger,		adParamInput,		0,		RF8) _
			, Array("@RF9",						adInteger,		adParamInput,		0,		RF9) _
			, Array("@RF10",					adInteger,		adParamInput,		0,		RF10) _
			, Array("@RF11",					adInteger,		adParamInput,		0,		RF11) _
			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@IDX",						adInteger,		adParamInput,		0,		IDX) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
		
		strLogMSG = "모집단위관리  > "& MYear &"_"& Division0Name &"_"& SubjectName &"_"& Division1Name &"_"& Division2Name &"_"& Division3Name &" 학과가 수정 되었습니다."
		InsertType = "Update"
	end If
	
	'트랜젝션 처리
	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
		objDB.sbRollbackTrans
	Else 
		strResult = "Complete"
		returnMSG = "모집단위 저장 완료"
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