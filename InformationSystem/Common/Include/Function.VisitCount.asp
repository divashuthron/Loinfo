<%
'//=== 방문자수 기록하기 =====================
Sub setVisitCount(SchoolCode, objDB)
	Dim SQL, arrParams
	
	SQL = "INSERT INTO SchoolVisitCount (SchoolCode) VALUES (?);"
	Call objDB.sbSetArray("@SchoolCode", adVarchar, adParamInput, 50, SchoolCode)
	arrParams = objDB.fnGetArray
	Call objDB.sbExecSQL(SQL, arrParams)
End Sub
'//=========================================

'//=== 방문자수 가져오기 =====================
Function VisitCount(SchoolCode)
	Dim objDB, SQL, arrParams

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB

	'// index.asp 페이지에서만 방문자수 늘리기 실행
	If instr(ASP_SELF_URL, "/index.asp") <> 0 Then
		Call setVisitCount(SchoolCode, objDB)
	End If

	SQL = " SELECT Count(*), '전체' FROM SchoolVisitCount Where SchoolCode = ? "
	SQL = SQL & " UNION ALL "
	SQL = SQL & " SELECT Count(*), '오늘' FROM SchoolVisitCount Where SchoolCode = ? and convert(varchar(8), RegDate, 112) = convert(varchar(8), getdate(), 112) "
	
	Call objDB.sbSetArray("@SchoolCode", adVarchar, adParamInput, 50, SchoolCode)
	Call objDB.sbSetArray("@SchoolCode", adVarchar, adParamInput, 50, SchoolCode)
	
	arrParams = objDB.fnGetArray
	'objDB.blnDebug = True
	VisitCount = objDB.fnExecSQLGetRows(SQL, arrParams)

	Set objDB = Nothing
End Function
'//=========================================
%>