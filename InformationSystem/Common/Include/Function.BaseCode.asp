<%
'//=== 마스터 & 서브 코드 가져오기 =============
'// 마스터 코드 가져오기
Function MasterCodeList(MasterCode)
	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB

'	SQL = ""
'	SQL = SQL & vbCrLf & "SELECT "
'	SQL = SQL & vbCrLf & "	IDX,MasterCode,MasterCodeName,State,RegDate,RegID,EditDate,EditID "
'	SQL = SQL & vbCrLf & "	, (CASE  State "
'	SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
'	SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
'	SQL = SQL & vbCrLf & "	END) AS StateName "
'	SQL = SQL & vbCrLf & "From CodeMaster "
'	SQL = SQL & vbCrLf & "Where 1 = 1 "
'	SQL = SQL & vbCrLf & "	AND MasterCode = ? "
'	SQL = SQL & vbCrLf & "	AND State = 'Y'; "
'
'	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
'
'	'objDB.blnDebug = true
'	arrParams = objDB.fnGetArray
'	'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
'	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	SQL = "dbo.getCodeMaster"
	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)

	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	MasterCodeList = objDB.fnExecSPGetHashMap(SQL, arrParams)

	Set objDB = Nothing
End Function

'// 서브 코드 가져오기
Function SubCodeList(MasterCode)
	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere

	Set objDB = New clsDBHelper
	objDB.strConnectionString = strDBConnString
	objDB.sbConnectDB
	
'	SQL = ""
'	SQL = SQL & vbCrLf & "SELECT "
'	SQL = SQL & vbCrLf & "	IDX, SubCode, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc "
'	SQL = SQL & vbCrLf & "	, UseYN, State "
'	SQL = SQL & vbCrLf & "	, (CASE  State "
'	SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
'	SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
'	SQL = SQL & vbCrLf & "	END) AS StateName "
'	SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID "
'	SQL = SQL & vbCrLf & "From CodeSub "
'	SQL = SQL & vbCrLf & "Where 1 = 1 "
'	SQL = SQL & vbCrLf & "	AND MasterCode = ? "
'	SQL = SQL & vbCrLf & "	AND State = 'Y' "
'	SQL = SQL & vbCrLf & "ORDER BY Step ASC; "
'
'	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
'
'	'objDB.blnDebug = true
'	arrParams = objDB.fnGetArray
'	'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
'	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	SQL = "dbo.getCodeSub"
	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)

	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	SubCodeList = objDB.fnExecSPGetHashMap(SQL, arrParams)
	
	Set objDB = Nothing
End Function
'//============================================




'//=== 마스터코드 세팅 ========================
Sub MasterCodeSelectBox(Name, NameDesc, Value, AlterMsg, FirstNode, MasterCode)
'	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
'
'	Set objDB = New clsDBHelper
'	objDB.strConnectionString = strDBConnString
'	objDB.sbConnectDB
'	
'	SQL = "dbo.getCodeMaster"
'	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
'
'	'objDB.blnDebug = true
'	arrParams = objDB.fnGetArray
'	AryHash = objDB.fnExecSPGetHashMap(SQL, arrParams)
'
'	Set objDB = Nothing

	Dim AryHash : AryHash = MasterCodeList(MasterCode)
	Dim intNum
%>
<select name="<%= Name %>" class="form-control input-sm select2" alert="<%= AlterMsg %>">
	<option value=""><%= NameDesc %></option>
<%
	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
%>
	<option value="<%= AryHash(intNum).Item("MasterCode") %>" <%= setSelected(AryHash(intNum).Item("MasterCode"), Value) %>><%= AryHash(intNum).Item("MasterCodeName") %></option>
<%
		Next
	end if
%>
</select>
<%
End Sub
'//=========================================

'//=== 마스터코드 세팅 ======================
Sub MasterCodeSelectOptionBox(Value, AlterMsg, FirstNode, MasterCode)
	
	Dim AryHash : AryHash = MasterCodeList(MasterCode)
	Dim intNum
	
	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
%>
	<option value="<%= AryHash(intNum).Item("MasterCode") %>" <%= setSelected(AryHash(intNum).Item("MasterCode"), Value) %>><%= AryHash(intNum).Item("MasterCodeName") %></option>
<%
		Next
	end if
End Sub
'//==================================




'//=== 서브코드 세팅 ======================
Sub SubCodeSelectBox(Name, NameDesc, Value, AlterMsg, FirstNode, MasterCode)
'	Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
'
'	Set objDB = New clsDBHelper
'	objDB.strConnectionString = strDBConnString
'	objDB.sbConnectDB
'	
'	SQL = "dbo.getCodeSub"
'	Call objDB.sbSetArray("@MasterCode", adVarchar, adParamInput, 50, MasterCode)
'
'	'objDB.blnDebug = true
'	arrParams = objDB.fnGetArray
'	AryHash = objDB.fnExecSPGetHashMap(SQL, arrParams)
'	
'	Set objDB = Nothing

	Dim AryHash : AryHash = SubCodeList(MasterCode)
	Dim intNum
%>
<select name="<%= Name %>" class="form-control input-sm select2" alert="<%= AlterMsg %>">
<% 'If Not(IsE(NameDesc)) then %>
	<option value=""><%= NameDesc %></option>
<% 'End If %>
<% If FirstNode = "All" then %>
	<option value="<%= FirstNode %>" <%= setSelected(FirstNode, Value) %>>전체</option>
<% End If %>
<%
	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
%>
	<option value="<%= AryHash(intNum).Item("SubCode") %>" <%= setSelected(AryHash(intNum).Item("SubCode"), Value) %>><%= AryHash(intNum).Item("SubCodeName") %></option>
<%
		Next
	end if
%>
</select>
<%
End Sub
'//========================================

'//=== 학교별 서브코드 세팅 ===============
Sub SubCodeSelectBox_School(Name, NameDesc, Value, AlterMsg, FirstNode, MasterCode, SchoolCode)
	'// 학교코드(Temp1)가 입력받은 학교코드(SchoolCode)와 비교하여 같은것만 출력
	'// Temp1에 입력된 학교코드값이 있을때만 비교하여 처리

	Dim AryHash : AryHash = SubCodeList(MasterCode)
	Dim intNum
%>
<select name="<%= Name %>" class="form-control input-sm select2" alert="<%= AlterMsg %>">
<%
	'If Not(IsE(NameDesc)) then
%>
	<option value=""><%= NameDesc %></option>
<%
	'End If

	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
			If Not(IsE(AryHash(intNum).Item("Temp1"))) then
				'// 학교코드(Temp1)가 NULL이 아니면 입력받은 학교 코드와 비교하여 같은것만 출력
				If AryHash(intNum).Item("Temp1") = SchoolCode then
%>
	<option value="<%= AryHash(intNum).Item("SubCode") %>" <%= setSelected(AryHash(intNum).Item("SubCode"), Value) %>><%= AryHash(intNum).Item("SubCodeName") %></option>
<%
				End If
			Else
				'// 학교코드(Temp1)가 NULL이면 그냥 출력
%>
	<option value="<%= AryHash(intNum).Item("SubCode") %>" <%= setSelected(AryHash(intNum).Item("SubCode"), Value) %>><%= AryHash(intNum).Item("SubCodeName") %></option>
<%
			End if
		Next
	end if
%>
</select>
<%
End Sub
'//========================================



'//=== 서브코드 세팅 ======================
Sub SubCodeSelectOptionBox(Value, AlterMsg, FirstNode, MasterCode)
	
	Dim AryHash : AryHash = SubCodeList(MasterCode)
	Dim intNum
	
	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
%>
	<option value="<%= AryHash(intNum).Item("SubCode") %>" <%= setSelected(AryHash(intNum).Item("SubCode"), Value) %>><%= AryHash(intNum).Item("SubCodeName") %></option>
<%
		Next
	end if
End Sub
'//========================================

'//=== 학교별 서브코드 세팅 ===============
Sub SubCodeSelectOptionBox_School(Value, AlterMsg, FirstNode, MasterCode, SchoolCode)
	
	Dim AryHash : AryHash = SubCodeList(MasterCode)
	Dim intNum
	
	If IsArray(AryHash) then
		For intNum = 0 to ubound(AryHash,1)
			If AryHash(intNum).Item("Temp1") = SchoolCode then
%>
	<option value="<%= AryHash(intNum).Item("SubCode") %>" <%= setSelected(AryHash(intNum).Item("SubCode"), Value) %>><%= AryHash(intNum).Item("SubCodeName") %></option>
<%
			end if
		Next
	end if
End Sub
'//========================================


'//== 서브코드로 서브코드 이름 추출 ======
Function getSubCodeNameToSubCode(objDB, MasterCode, SubCode)
	Dim SQL, arrParams, aryList, AryHash, strWhere

	SQL = "dbo.getCodeSubCompareCode"
	Call objDB.sbSetArray("MasterCode",		adVarchar,	adParamInput, 50,	MasterCode)
	Call objDB.sbSetArray("@SubCodeName",	adVarchar,	adParamInput, 50,	SubCode)

	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSPGetHashMap(SQL, arrParams)
	
	If IsArray(AryHash) Then
		getSubCodeNameToSubCode = AryHash(0).Item("SubCodeName")
	Else
		getSubCodeNameToSubCode = SubCode
	End if
End Function

'//== 서브코드이름으로 서브코드 추출 ======
Function getSubCodeToSubCodeName(objDB, MasterCode, SubCodeName, SchoolCode)
	Dim SQL, arrParams, aryList, AryHash, strWhere

	SQL = "dbo.getCodeSubCompareName"
	Call objDB.sbSetArray("MasterCode",		adVarchar,	adParamInput, 50,	MasterCode)
	Call objDB.sbSetArray("@SubCodeName",	adVarchar,	adParamInput, 255,	SubCodeName)
	Call objDB.sbSetArray("@Temp1",			adVarchar,	adParamInput, 255,	SchoolCode)
	'objDB.blnDebug = true
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSPGetHashMap(SQL, arrParams)
	
	If IsArray(AryHash) Then
		getSubCodeToSubCodeName = AryHash(0).Item("SubCode")
	Else
		getSubCodeToSubCodeName = SubCodeName
	End if
End Function
'//========================================
%>