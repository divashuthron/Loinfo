<%
'//*********************************************************
'//	dbhelper class
'//*********************************************************
Class clsDBHelper
	Private objConn																	'Connection 객체
	Private objCmd																	'Command 객체
	Private objRs																	'RecordSet 객체
	Private arrRows																	'Array 변수
	Private intZ

	Public strConnectionString														'DB Connection String 변수
	Public blnDebug																	'바인딩 된 쿼리 확인 Check 변수

'########## clsDBHelper 객체 생성시 자동 실행 Proc #######################
	Private Sub Class_Initialize()
		blnDebug = False			'기본적으로 바인딩 된 쿼리 확인을 False로 놓는다.
		intZ = 0

		Const MAX_TRIES = 10
		Dim intTries
		On Error Resume Next
		Do
			Err.Clear
			intTries = intTries + 1
		Loop While (Err.Number <> 0) And (intTries < MAX_TRIES)
		'response.wrtie "Start"
	End Sub
'#########################################################

'########## clsDBHelpter 객체 소명시 자동 실행 Proc ######################## 
	Private Sub Class_Terminate()
		Call sbCloseRs()			'Recordset 객체 소멸
		Call sbCloseCmd()		'Command 객체 소멸
		Call sbCloseDB()			'Connection 객체 소멸
	End Sub
'###########################################################

'########## DB Connection Proc ####################################
	Public Sub sbConnectDB()
		Call sbCloseDB()			'Connection 객체 소멸
		Call sbOpenDB()			'Connection 객체 생성
		Call sbOpenCmd()		'Command 객체 생성
	End Sub
'###########################################################

'########### 쿼리에 바인딩 될 파라미터를 Array로 생성 Proc ######################
	Public Sub sbSetArray(sVar, sType, sParam, sLen, sValue)
		If intZ = 0 Then
			ReDim arrRows(intZ)
		Else
			ReDim Preserve arrRows(intZ)
		End If
		arrRows(intZ)  =  Array(sVar, stype, sParam, sLen, sValue)
		intZ = intZ + 1
	End Sub
'#############################################################

'########### 쿼리에 바인딩 될 파라미터를 Return Proc ###########################
	Public Function fnGetArray()
		fnGetArray = arrRows
		intZ = 0
	End Function
'###############################################################

'################# Connection 객체 생성 Proc ##############################
	Private Sub sbOpenDB()
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open strConnectionString
	End Sub
'#################################################################

'################# Connection 객체 소멸 Proc ###############################
	Private Sub sbCloseDB()
		If IsObject(objConn) Then
			objConn.Close
			Set objConn = Nothing
		End If
	End Sub
'#################################################################

'################## Command 객체 생성 Proc ################################
	Private Sub sbOpenCmd()
		Set objCmd = Server.CreateObject("Adodb.Command")
		objCmd.ActiveConnection = objConn

		objCmd.CommandTimeOut = 100
	End Sub
'###################################################################

'################# Command 객체 소멸 Proc ###################################
	Private Sub sbCloseCmd()
		If IsObject(objCmd) Then Set objCmd = Nothing
	End Sub
'###################################################################

'################# Create 된 파라미터 Delete Proc #################################
	Private Sub sbRemoveParameters()
		If IsObject(objCmd) Then
			Dim intX
			For intX = objCmd.Parameters.Count To 1 Step -1
				objCmd.Parameters.Delete(intX - 1)
			Next
		End If
	End Sub
'#####################################################################

'################## Recordset 객체 생성 Proc ####################################
	Private Sub sbOpenRs()
		Set objRs = Server.CreateObject("ADODB.RecordSet")
	End Sub
'######################################################################

'################### Recordset 객체 소멸 Proc ###################################
	Private Sub sbCloseRs()
		If IsObject(objRs) Then Set objRs = Nothing
	End Sub
'######################################################################

'############ SP를 실행하고, RecordSet을 반환하는 Proc #################################
	Public Function fnExecSPReturnRs(spName, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(spName, params)

		Dim intX
		With objCmd
			.CommandText = spName
			.CommandType = adCmdStoredProc
			'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성 Command 객체에 추가한다
			Call sbCollectParams(objCmd, params)

			If Not IsObject(objRs) Then Call sbOpenRs()		'Recordset 객체 생성

			objRs.CursorLocation = adUseClient
			objRs.Open objCmd

			For intX = 0 To .Parameters.Count - 1
				If .Parameters(intX).Direction = adParamOutput _
					OR .Parameters(intX).Direction = adParamInputOutput _
					OR .Parameters(intX).Direction = adParamReturnValue Then
					If IsObject(params) Then
						If params is Nothing Then Exit For
					Else
						params(intX)(4) = .Parameters(intX).Value
					End If
				End If
			Next
			If IsArray(params) Then Call sbRemoveParameters()		'생성된 파라미터를 Delete 하는 Proc

			Set objRs.ActiveConnection = Nothing

			Set fnExecSPReturnRs = objRs											'생성된 Recordset 객체 Return
		End With
	End Function
'####################################################################

'################SP를 실행하고, RecordSet을 GetRows 배열로 반환하는 Proc ##################
	Public Function fnExecSPGetRows(spName, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(spName, params)

		Dim intX
		With objCmd
			.CommandText = spName
			.CommandType = adCmdStoredProc
			'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
			Call sbCollectParams(objCmd, params)

			If Not IsObject(objRs) Then Call sbOpenRs()		'Recordset 객체 생성

			objRs.Open objCmd

			'RW(IsObject(objRs))
			'RW(objRs.State)
			'RW(adStateOpen)

			'생성된 Recordset 객체를 GetRows로 Return
			If Not objRs.EOF Then fnExecSPGetRows = objRs.GetRows()
			'If Not objRs.EOF Then fnExecSPGetRows = objRs.GetRows(-1,0,Fields)
			objRs.Close

			For intX = 0 To .Parameters.Count - 1
				If .Parameters(intX).Direction = adParamOutput _
					OR .Parameters(intX).Direction = adParamInputOutput _
					OR .Parameters(intX).Direction = adParamReturnValue Then
					If IsObject(params) Then
						If params is Nothing Then Exit For
					Else
						params(intX)(4) = .Parameters(intX).Value
					End If
				End If
			Next
			If IsArray(params) Then Call sbRemoveParameters()				'생성된 파라미터를 Delete 하는 Proc
		End With
	End Function
'#######################################################################

'################SP를 실행하고 RecordSet을 해쉬맵으로 반환하는 Proc #######################
	Public Function fnExecSPGetHashMap(spName, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(spName, params)

		Dim intX
		With objCmd
			.CommandText = spName
			.CommandType = adCmdStoredProc
			'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
			Call sbCollectParams(objCmd, params)

			If Not IsObject(objRs) Then Call sbOpenRs()		'Recordset 객체 생성

			objRs.Open objCmd
			
			If objRs.EOF Or objRs.BOF Then
				fnExecSPGetHashMap = NULL
			Else
				Dim aryList, aryFieldName, objDic, aryDic, i, j
				' 필드명 추출
				Redim aryFieldName(objRS.Fields.Count - 1)
				For i = 0 To objRS.Fields.Count - 1
					aryFieldName(i) = objRS.Fields(i).Name
				Next
				'연괄배열을 이용해 필드명을 키값으로 사용하는 해쉬테이블 생성
				aryList = objRs.GetRows()
				Redim aryDic(UBound(aryList, 2))
				For i = 0 To UBound(aryList, 2)
					Set objDic = CreateObject("Scripting.Dictionary")
						For j = 0 To UBound(aryFieldName, 1)
							objDic.Add aryFieldName(j) , aryList(j, i)
						Next
					Set aryDic(i) = objDic
					Set objDic = Nothing
				Next
				'생성된 Recordset 객체를 해쉬맵으로 Return
				fnExecSPGetHashMap = aryDic
			End If

			objRs.Close

			For intX = 0 To .Parameters.Count - 1
				If .Parameters(intX).Direction = adParamOutput _
					OR .Parameters(intX).Direction = adParamInputOutput _
					OR .Parameters(intX).Direction = adParamReturnValue Then
					If IsObject(params) Then
						If params is Nothing Then Exit For
					Else
						params(intX)(4) = .Parameters(intX).Value
					End If
				End If
			Next
			If IsArray(params) Then Call sbRemoveParameters()				'생성된 파라미터를 Delete 하는 Proc
		End With
	End Function
'##################################################################################

'############SQL Query를 실행하고 RecordSet을 반환하는 Proc #############################
	Public Function fnExecSQLReturnRs(strSQL, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(strSQL, params)

		objCmd.CommandText = strSQL
		objCmd.CommandType = adCmdText
		'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
		Call sbCollectParams(objCmd, params)

		If Not IsObject(objRs) Then Call sbOpenRs()		'Recordset 객체 생성

 		objRs.CursorLocation = adUseClient
		objRs.Open objCmd
		Set objRs.ActiveConnection = Nothing

		Set fnExecSQLReturnRs = objRs							'생성된 Recordset 객체 Return

		If IsArray(params) Then Call sbRemoveParameters()				'생성된 파라미터를 Delete 하는 Proc
	End Function
'#######################################################################

'###############SQL Query를 실행하고 RecordSet을 배열로 반환하는 Proc #######################
	Public Function fnExecSQLGetRows(strSQL, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(strSQL, params)

		objCmd.CommandText = strSQL
		objCmd.CommandType = adCmdText
		'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
		Call sbCollectParams(objCmd, params)

		If Not IsObject(objRs) Then Call sbOpenRs()		'Recordset 객체 생성

		objRs.Open objCmd
		'생성된 Recordset 객체를 GetRows로 Return
		If Not objRs.EOF Then fnExecSQLGetRows = objRs.GetRows()
        
		objRs.Close

		If IsArray(params) Then Call sbRemoveParameters()				'생성된 파라미터를 Delete 하는 Proc
	End Function
'########################################################################

'###############SQL Query를 실행하고 RecordSet을 해쉬맵으로 반환하는 Proc #######################
	Public Function fnExecSQLGetHashMap(strSQL, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(strSQL, params)

		objCmd.CommandText = strSQL
		objCmd.CommandType = adCmdText
		'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
		Call sbCollectParams(objCmd, params)

		If Not IsObject(objRs) Then Call sbOpenRs() 'Recordset 객체 생성

		objRs.Open objCmd

		'RW(IsObject(objRs))
		'RW(objRs.State)
		'If IsObject(objRs) And objRs.State = adStateOpen Then
			If objRs.EOF Or objRs.BOF Then
				fnExecSQLGetHashMap = NULL
			Else
				Dim aryList, aryFieldName, objDic, aryDic, i, j
				' 필드명 추출
				Redim aryFieldName(objRS.Fields.Count - 1)
				For i = 0 To objRS.Fields.Count - 1
					aryFieldName(i) = objRS.Fields(i).Name
				Next
				'연괄배열을 이용해 필드명을 키값으로 사용하는 해쉬테이블 생성
				aryList = objRs.GetRows()
				Redim aryDic(UBound(aryList, 2))
				For i = 0 To UBound(aryList, 2)
					Set objDic = CreateObject("Scripting.Dictionary")
						For j = 0 To UBound(aryFieldName, 1)
							objDic.Add aryFieldName(j) , aryList(j, i)
						Next
					Set aryDic(i) = objDic
					Set objDic = Nothing
				Next
				'생성된 Recordset 객체를 해쉬맵으로 Return
				fnExecSQLGetHashMap = aryDic
			End If
		'End If

		objRs.Close

		If IsArray(params) Then Call sbRemoveParameters() '생성된 파라미터를 Delete 하는 Proc
	End Function
'##################################################################################

'##########SP를 실행하는 Proc (RecordSet 반환없음) #####################################
	Public Sub sbExecSP(strSP, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(strSP, params)

		Dim intX
		With objCmd
			.CommandText = strSP
			.CommandType = adCmdStoredProc
			'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성 Command 객체에 추가한다
			Call sbCollectParams(objCmd, params)

			.Execute , , adExecuteNoRecords

			For intX = 0 To .Parameters.Count - 1
				If .Parameters(intX).Direction = adParamOutput _
					OR .Parameters(intX).Direction = adParamInputOutput _
					OR .Parameters(intX).Direction = adParamReturnValue Then

					If IsObject(params) Then
						If params is Nothing Then
							Exit For
						End If
					Else
						params(intX)(4) = .Parameters(intX).Value
					End If
				End If
			Next

			If IsArray(params) Then Call sbRemoveParameters()		'생성된 파라미터를 Delete 하는 Proc
		End With
	End Sub
'##########################################################################

'#############Query를 실행하는 Proc (RecordSet 반환없음) ##################################
	Public Sub sbExecSQL(strSQL, params)
		'바인딩 된 쿼리 View Proc - blnDebug = True 일경우 화면에 쿼리 출력
		If blnDebug Then Call sbViewQuery(strSQL, params)

		With objCmd
			.CommandText = strSQL
			.CommandType = adCmdText
			'Array로 넘겨오는 파라미터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc
			Call sbCollectParams(objCmd, params)

			.Execute , , adExecuteNoRecords

			If IsArray(params) Then Call sbRemoveParameters()				'생성된 파라미터를 Delete 하는 Proc
		End With
	End Sub
'###########################################################################

'##################트랜잭션을 시작하고, Connetion 개체를 반환하는 Proc #########################
	Public Sub sbBeginTrans()
		If IsObject(objConn) Then objConn.BeginTrans
	End Sub
'###########################################################################

'###################활성화된 트랜잭션을 커밋하는 Proc #####################################
	Public Sub sbCommitTrans()
		If IsObject(objConn) Then objConn.CommitTrans
	End Sub
'############################################################################

'######################활성화된 트랜잭션을 롤백하는 Proc ##################################
	Public Sub sbRollbackTrans()
		If IsObject(objConn) Then objConn.RollbackTrans
	End Sub
'############################################################################

'####################매개변수 배열 내에서 지정된 이름의 매개변수 값을 반환하는 Proc ####################
	Public Function fnGetValue(params, paramName)
		Dim objParam
		For Each objParam in params
			If objParam(0) = paramName Then
				fnGetValue = objParam(4)
				Exit Function
			End If
		Next
	End Function
'#############################################################################

'# Array로 넘겨오는 파라메터를 Parsing 하여 Parameter 객체를 생성하여 Command 객체에 추가하는 Proc ###########
	Private Sub sbCollectParams(cmd, argparams)
		Dim objParams
		Dim intX
		Dim intL
		Dim intU
		Dim strValue

		If VarType(argparams) = 8192 _
			Or VarType(argparams) = 8204 _
			Or VarType(argparams) = 8209 Then			'파라미터로 넘오는 변수가 Array 인지를 검토 한다.

			objParams = argparams
			For intX = LBound(objParams) To UBound(objParams)
				intL = LBound(objParams(intX))
				intU = UBound(objParams(intX))

				' Check for nulls.
				If intU - intL = 4 Then
					If VarType(objParams(intX)(4)) = vbString Then
						If objParams(intX)(4) = "" Then
							strValue = Null
						Else
							strValue = objParams(intX)(4)
						End If
					ElseIf VarType(objParams(intX)(4)) = vbEmpty Then
						strValue = Null
					Else
						strValue = objParams(intX)(4)
					End If
					If blnDebug Then Response.Write("Parameter : "& intX &"<br>")
					cmd.Parameters.Append cmd.CreateParameter(objParams(intX)(0) _
										, objParams(intX)(1) _
										, objParams(intX)(2) _
										, objParams(intX)(3) _
										, strValue)						'Command 객체에 Parameter 추가 한다.
				End If
			Next
		End If
	End Sub
'###############################################################################

'###################바인딩된 내용과 함께 쿼리 보기 Proc #######################################
	Public Sub sbViewQuery(strQuery, arrParams)
		Dim intX

		Response.Write("<pre style='font-size:12px;'>")
		Response.Write("======================================================<br>")

		If IsArray(arrParams) Then
			For intX = 0 To UBound(arrParams)
				If arrParams(intX)(1) = adTinyInt Or _
					arrParams(intX)(1) = adSmallInt Or _
					arrParams(intX)(1) = adInteger Or _
					arrParams(intX)(1) = adBigInt Then
					strQuery = Replace(strQuery, "?", arrParams(intX)(4), 1, 1)
				Else
					strQuery = Replace(strQuery, "?", "'"& arrParams(intX)(4) &"'", 1, 1)
				End If

				Response.Write(intX &", "& arrParams(intX)(0) &", "& _
					arrParams(intX)(4) & _
					", length("& arrParams(intX)(3) &") : "& Len(arrParams(intX)(4)) & vbcrlf)
			Next
		End If

		Response.Write("======================================================<br>")
		Response.Write(strQuery &"<br>")
		Response.Write("======================================================<br>")
		Response.Write("</pre>")
	End Sub
'##################################################################################
End Class
%>