<%
'//=== 로그 기록 ============================================
Sub ActivityHistory(LogMSG, Division, UserID)
	Dim objDB, SQL, arrParams

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	
	SQL = "INSERT INTO ActivityHistory (ActivityContent, Division, RegID, MYear) VALUES (?, ?, ?, ?);"

	arrParams = Array(_
		  Array("@ActivityContent",		adVarchar,		adParamInput, 255,		LogMSG) _
		, Array("@Division",			adVarchar,		adParamInput, 50,		Division) _
		, Array("@RegID",				adVarchar,		adParamInput, 25,		UserID) _
		, Array("@MYear",				adVarchar,		adParamInput, 4,		SessionMYear) _
	)

	'objDB.blnDebug = true
	Call objDB.sbExecSQL(SQL, arrParams)

	Set objDB = Nothing
End Sub
'//==========================================================

'//=== 히스토리 읽기 기록(알람구분용) =======================
'//=== 해당 아이디로 해당 게시물을 읽지 않았으면 읽기처리 ===
Sub AlarmHistory(HistoryIDX, MYear, RegID)
	Dim objDB, SQL, arrParams, AryHash

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB

	SQL = ""
	SQL = SQL & vbCrLf & "Select * "
	SQL = SQL & vbCrLf & "from AlarmHistory "
	SQL = SQL & vbCrLf & "where HistoryIDX = ? "
	SQL = SQL & vbCrLf & "and MYear = ? "
	SQL = SQL & vbCrLf & "and RegID = ?; "

	Call objDB.sbSetArray("@HistoryIDX", adVarchar, adParamInput, 50, HistoryIDX)
	Call objDB.sbSetArray("@MYear", adVarchar, adParamInput, 50, MYear)
	Call objDB.sbSetArray("@RegID", adVarchar, adParamInput, 50, RegID)

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)
	
	If Not IsArray(AryHash) then
		SQL = ""
		SQL = "INSERT INTO AlarmHistory (HistoryIDX, MYear, RegID) VALUES (?, ?, ?);"

		arrParams = Array(_
			  Array("@HistoryIDX",			adInteger,		adParamInput, 0,		HistoryIDX) _
			, Array("@MYear",				adVarchar,		adParamInput, 25,		MYear) _
			, Array("@RegID",				adVarchar,		adParamInput, 25,		RegID) _
		)

		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
	End If

	Set objDB = Nothing
End Sub
'//==========================================================
%>