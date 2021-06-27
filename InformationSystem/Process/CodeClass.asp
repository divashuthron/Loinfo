<%
	Class clsCodeManager
		Public objDB, i, PageNum, PageSize, TotalCount, SearchType, SearchText, SearchState
		Public MasterCode, MasterCodeName
		Public SubCode, SubCodeOld, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc
		Public State, RegDate, RegID, EditDate, EditID

'=================================
'		Class 초기화 Proc
'=================================
		Private Sub Class_Initialize()
		End Sub

'=================================
'		Class 소멸 Proc
'=================================
		Private Sub Class_Terminate()
			Call sbObjectNothingProc()
		End Sub
		
'=================================
'		마스터 코드 COUNT 
'=================================
		Public Function fnSelectMasterCodeCountProc()
			Dim SQL, arrParams
			Dim arrCount
			Dim strWhere : strWhere = ""
			
			if (not(IsE(SearchText))) then
				strWhere = strWhere & " And MasterCodeName like '%' + ? + '%' "
				Call objDB.sbSetArray("@MasterCodeName", adVarchar, adParamInput, 255, SearchText)
			end If

			SQL = " SELECT Count(*) "
			SQL = SQL & " FROM CodeMaster AS A "
			SQL = SQL & " WHERE 1 = 1 "
			SQL = SQL & strWhere
			
			'objDB.blnDebug = True
			arrParams = objDB.fnGetArray			
			arrCount  = objDB.fnExecSQLGetRows(SQL, arrParams)
			fnSelectMasterCodeCountProc = arrCount(0,0)
		End Function

'=================================
'		마스터 코드 리스트
'=================================
		Public Function fnSelectMasterCodeProc()
			Dim SQL, arrParams
			Dim arrCount
			Dim strWhere : strWhere = ""
			
			if (not(IsE(SearchText))) then
				strWhere = strWhere & " And MasterCodeName like '%' + ? + '%' "
				Call objDB.sbSetArray("@MasterCodeName", adVarchar, adParamInput, 255, SearchText)
			end If
			
			' MasterCode, MasterCodeName, State, StateName, 		0~3
			' RegDate, RegID, EditDate, EditID								4~7

			SQL = " SELECT * FROM "
			SQL = SQL & "	( "
			SQL = SQL & "		SELECT "
			SQL = SQL & "			MasterCode, MasterCodeName, State, "
			SQL = SQL &"			(CASE  State "
			SQL = SQL &"				WHEN 'Y' THEN '사용' "
			SQL = SQL &"				WHEN 'N' THEN '미사용' "
			SQL = SQL &"			END) AS StateName, "
			SQL = SQL & "			RegDate, RegID, EditDate, EditID, "
			SQL = SQL & "			ROW_NUMBER() OVER (ORDER BY MasterCode desc) AS ROWNUM "
			SQL = SQL & "		FROM CodeMaster AS A " 
			SQL = SQL & "		WHERE 1 = 1 "
			SQL = SQL & strWhere
			SQL = SQL & "	) AS TBL_PAGELIST "
			SQL = SQL & "	WHERE ROWNUM BETWEEN "& (PageNum - 1) * PageSize + 1 &" AND "& PageNum * PageSize &";"
			
			'objDB.blnDebug = true
			arrParams = objDB.fnGetArray
			fnSelectMasterCodeProc = objDB.fnExecSQLGetRows(SQL, arrParams)
		End Function
		
'=================================
'		마스터 코드 정보 뷰
'=================================
		Public Function fnSelectMasterCodeViewProc()
			Dim SQL, arrParams
			
			' MasterCode, MasterCodeName, State, StateName, 		0~3
			' RegDate, RegID, EditDate, EditID								4~7
			
			SQL = " Select "
			SQL = SQL & "		MasterCode, MasterCodeName, State, "
			SQL = SQL &"		(CASE  State "
			SQL = SQL &"			WHEN 'Y' THEN '사용' "
			SQL = SQL &"			WHEN 'N' THEN '미사용' "
			SQL = SQL &"		END) AS StateName, "
			SQL = SQL & "		RegDate, RegID, EditDate, EditID "
			SQL = SQL & "	From CodeMaster AS A "
			SQL = SQL & "	Where 1 = 1 "
			SQL = SQL & "		AND MasterCode = ?; "
            
			Call objDB.sbSetArray("@MasterCode", adInteger, adParamInput, 0, MasterCode)

			'objDB.blnDebug = true			
			arrParams = objDB.fnGetArray
			fnSelectMasterCodeViewProc = objDB.fnExecSQLGetRows(SQL, arrParams)
		End Function
		
'=================================
'		마스터 코드 INSERT
'=================================
		Public Function fnInsertMasterCodeProc()
			Dim SQL, strWhere, arrParams, aryList
			
			SQL = " INSERT INTO CodeMaster ( "
			SQL = SQL &"		MasterCode, MasterCodeName, State, RegID "
			SQL = SQL &" ) VALUES ( "
			SQL = SQL &"		?, ?, ?, ? "
			SQL = SQL &" ) "
			
			'adDate, adLongVarChar, adVarchar, adInteger, adChar

			arrParams = Array(_
				  Array("@MasterCode",					adInteger,			adParamInput, 0,				MasterCode) _
				, Array("@MasterCodeName",			adVarchar,			adParamInput, 255,			MasterCodeName) _
				, Array("@State",							adChar,				adParamInput, 1,				State) _
				, Array("@RegID",							adVarchar,			adParamInput, 25,			RegID) _
			)

			'objDB.blnDebug = true
			Call objDB.sbExecSQL(SQL, arrParams)
		End Function
		
'=================================
'		마스터 코드 UPDATE
'=================================
		Public Function fnUpdateMasterCodeProc()
			Dim SQL, arrParams, aryList, nMAX, result
			
			SQL = " UPDATE CodeMaster SET "
			SQL = SQL &"		MasterCode = ?, MasterCodeName = ?, State = ?, EditID = ?, EditDate = getdate() "
			SQL = SQL & " WHERE MasterCode = ?; "
			
			arrParams = Array(_
				  Array("@MasterCode",					adInteger,			adParamInput, 0,				MasterCode) _
				, Array("@MasterCodeName",			adVarchar,			adParamInput, 255,			MasterCodeName) _
				, Array("@State",							adChar,				adParamInput, 1,				State) _
				, Array("@EditID",						adVarchar,			adParamInput, 25,			EditID) _
				, Array("@MasterCode",					adInteger,			adParamInput, 0,				MasterCode) _
			)
			
			'objDB.blnDebug = true
			Call objDB.sbExecSQL(SQL, arrParams)
		End Function
		
'=================================
'		서브 코드 COUNT 
'=================================
		Public Function fnSelectSubCodeCountProc()
			Dim SQL, arrParams
			Dim arrCount
			
			SQL = " SELECT Count(*) "
			SQL = SQL & " FROM CodeSub AS A "
			SQL = SQL & " WHERE 1 = 1 "
			SQL = SQL & "		AND MasterCode = ? "
			
			Call objDB.sbSetArray("@MasterCode", adInteger, adParamInput, 0, MasterCode)
			
			'objDB.blnDebug = True
			arrParams = objDB.fnGetArray			
			arrCount  = objDB.fnExecSQLGetRows(SQL, arrParams)
			fnSelectSubCodeCountProc = arrCount(0,0)
		End Function

'=================================
'		서브 코드 리스트
'=================================
		Public Function fnSelectSubCodeProc()
			Dim SQL, arrParams
			Dim arrCount

			' SubCode, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc,		0~7
			' UseYN, State, StateName, RegDate, RegID, EditDate, EditID 							8~14
			
			SQL = " SELECT * FROM "
			SQL = SQL & "	( "
			SQL = SQL & "		SELECT "
			SQL = SQL & "			SubCode, SubCodeName, Step, Temp1, Temp2, Temp3, Temp4, TempEtc,  "
			SQL = SQL & "			UseYN, State, "
			SQL = SQL &"			(CASE  State "
			SQL = SQL &"				WHEN 'Y' THEN '사용' "
			SQL = SQL &"				WHEN 'N' THEN '미사용' "
			SQL = SQL &"			END) AS StateName, "
			SQL = SQL & "			RegDate, RegID, EditDate, EditID, "
			SQL = SQL & "			ROW_NUMBER() OVER (ORDER BY Step ASC) AS ROWNUM "
			SQL = SQL & "		FROM CodeSub AS A " 
			SQL = SQL & "		WHERE 1 = 1 "
			SQL = SQL & "			AND MasterCode = ? "
			SQL = SQL & "	) AS TBL_PAGELIST "
			SQL = SQL & "	WHERE ROWNUM BETWEEN "& (PageNum - 1) * PageSize + 1 &" AND "& PageNum * PageSize &";"
			
			Call objDB.sbSetArray("@MasterCode", adInteger, adParamInput, 0, MasterCode)
			
			'objDB.blnDebug = true
			arrParams = objDB.fnGetArray
			fnSelectSubCodeProc = objDB.fnExecSQLGetRows(SQL, arrParams)
		End Function
		
'=================================
'		서브 코드 INSERT
'=================================
		Public Function fnInsertSubCodeProc()
			Dim SQL, strWhere, arrParams, aryList
			
			SQL = " INSERT INTO CodeSub ( "
			SQL = SQL &"		MasterCode, SubCode, SubCodeName, Step, Temp1, Temp2, UseYN, State, RegID "
			SQL = SQL &" ) VALUES ( "
			SQL = SQL &"		?, ?, ?, ?, ?, ?, 'Y', ?, ? "
			SQL = SQL &" ) "
			
			'adDate, adLongVarChar, adVarchar, adInteger, adChar
			
			arrParams = Array(_
				  Array("@MasterCode",				adInteger,			adParamInput, 0,				MasterCode) _
				, Array("@SubCode",					adVarchar,			adParamInput, 25,			SubCode) _
				, Array("@SubCodeName",			adVarchar,			adParamInput, 255,			SubCodeName) _
				, Array("@Step",						adInteger,			adParamInput, 0,				Step) _
				, Array("@Temp1",					adVarchar,			adParamInput, 255,			Temp1) _
				, Array("@Temp2",					adVarchar,			adParamInput, 255,			Temp2) _
				, Array("@State",						adChar,				adParamInput, 1,				State) _
				, Array("@RegID",						adVarchar,			adParamInput, 25,			RegID) _
			)

			'objDB.blnDebug = true
			Call objDB.sbExecSQL(SQL, arrParams)
		End Function
		
'=================================
'		서브 코드 UPDATE
'=================================
		Public Function fnUpdateSubCodeProc()
			Dim SQL, arrParams, aryList, nMAX, result
			
			SQL = " UPDATE CodeSub SET "
			SQL = SQL &"		SubCode = ?, SubCodeName = ?, Step = ?, Temp1 = ?, Temp2 = ?, "
			SQL = SQL &"		State = ?, EditID = ?, EditDate = getdate() "
			SQL = SQL & " WHERE MasterCode = ? AND SubCode = ?; "
			
			arrParams = Array(_
				  Array("@SubCode",					adVarchar,			adParamInput, 25,			SubCode) _
				, Array("@SubCodeName",			adVarchar,			adParamInput, 255,			SubCodeName) _
				, Array("@Step",						adInteger,			adParamInput, 0,				Step) _
				, Array("@Temp1",					adVarchar,			adParamInput, 255,			Temp1) _
				, Array("@Temp2",					adVarchar,			adParamInput, 255,			Temp2) _
				, Array("@State",						adChar,				adParamInput, 1,				State) _
				, Array("@EditID",					adVarchar,			adParamInput, 25,			EditID) _
				, Array("@MasterCode",				adInteger,			adParamInput, 0,				MasterCode) _
				, Array("@SubCode",					adVarchar,			adParamInput, 25,			SubCodeOld) _
			)
			
			'objDB.blnDebug = true
			Call objDB.sbExecSQL(SQL, arrParams)
		End Function
		
'=================================
'		객체 소멸 Proc
'=================================
		Public Sub sbObjectNothingProc()
			If IsObject(objDB) Then
				If Not objDB Is Nothing Then Set objDB = Nothing
			End if
		End Sub

	End Class
%>