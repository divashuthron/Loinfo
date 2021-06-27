<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, i 

Dim MYear		: MYear = fnRF("MYear")   
Dim Division0	: Division0 = fnRF("Division0")  

'생기부/검정/수능 데이터, 최종 값
Dim StudentRecordDataStr, QualificationDataStr, CSATDataStr, ReslutStr
Dim InterviewerStr, PracticalStr

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'// ==============================================================================================================
'// 데이터, 최종 값 넣어주기
'// ==============================================================================================================
	
'생기부,검정,수능 데이터 쿼리
SQL = ""
SQL = SQL & vbCrLf & "SELECT   "
SQL = SQL & vbCrLf & "		e.StudentNumber as QualCheck, d.StudentNumber as CSATCheck, c.StudentNumber as StuRecCheck  "
SQL = SQL & vbCrLf & "		, c.Cors_Code, c.Majr_cs11, c.Majr_cs12, c.Majr_cs21, c.Majr_cs22, c.Majr_cs31, c.Majr_cs32 "
SQL = SQL & vbCrLf & "		, a.IDX, a.Myear, a.Division0, a.StudentNumber, a.StudentNameKor, a.StudentNameUsa, a.StudentNameChi  "
SQL = SQL & vbCrLf & "		, a.Citizen1, a.Citizen2, a.Sex, a.HighCode, a.HighSubject, a.HighGraduationYear, a.HighGraduationDivision  "
SQL = SQL & vbCrLf & "		, a.QualificationAreaCode, a.QualificationYear, a.Subject, a.Semester  "
SQL = SQL & vbCrLf & "		, a.UniversityName, a.AugScore, a.PerfectScore, a.Credit  "
SQL = SQL & vbCrLf & "		, a.Division1, a.Division2, a.Division3, a.Division4, a.HighDivision  "
SQL = SQL & vbCrLf & "		, a.RefundDivision, a.RefundAccountHolder, a.RefundBankCode, a.RefundAccount  "
SQL = SQL & vbCrLf & "		, a.DrawStandard, a.DrawMsg, a.ExtraPoint, a.InterviewerPoint, a.PracticalPoint, a.Tel1, a.Tel2, a.Tel3, a.Email, a.ZipCode, a.Address1, a.Address2  "
SQL = SQL & vbCrLf & "		, a.StudentNameAgreement, a.StudentAgreement, a.StudentRecordAgreement, a.QualificationAgreement  "
SQL = SQL & vbCrLf & "		, a.CSATAgreement, a.PersonalCollectionAgreement, a.UniquelyAgreement, a.PersonalTrustAgreement, a.PersonalofferAgreement  "
SQL = SQL & vbCrLf & "		, a.ReceiptDate, a.ReceiptTime  "
SQL = SQL & vbCrLf & "		, a.DocumentaryCheck1, a.DocumentaryCheck2, a.DocumentaryCheck3, a.DocumentaryCheck4 "
SQL = SQL & vbCrLf & "		, a.DocumentaryCheck5, a.DocumentaryCheck6, a.DocumentaryCheck7, a.DocumentaryCheck8 "
SQL = SQL & vbCrLf & "		, a.Document, a.DocumentMsg, a.Qualification, a.StudentRecordData, a.QualificationData, a.CSATData, a.Reslut "
SQL = SQL & vbCrLf & "		, a.DocumentaryCheck21, a.DocumentaryCheck22, a.DocumentaryCheck23, a.DocumentaryCheck24 "
SQL = SQL & vbCrLf & "		, a.INPT_USID, a.INPT_DATE, a.INPT_ADDR, a.UPDT_USID, a.UPDT_DATE, a.UPDT_ADDR, a.InsertTime  "
SQL = SQL & vbCrLf & "		, b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio "
SQL = SQL & vbCrLf & "		, b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5 "
SQL = SQL & vbCrLf & "		, b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5 "
SQL = SQL & vbCrLf & "		, b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5 "
SQL = SQL & vbCrLf & "		, b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5 "
SQL = SQL & vbCrLf & "		, b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5 "
SQL = SQL & vbCrLf & "		, dbo.getSubCodeName('Division0', a.Division0) AS DivisionName  "
SQL = SQL & vbCrLf & "		, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName  "
SQL = SQL & vbCrLf & "		, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name  "
SQL = SQL & vbCrLf & "		, dbo.getSubCodeName('HignSchoolDivision', a.HighDivision) AS HighDivisionName  "
'SQL = SQL & vbCrLf & "	/* 입학원서 */"
SQL = SQL & vbCrLf & "FROM ApplicationTable AS a "
'SQL = SQL & vbCrLf & "	/* 평가비율 */"
SQL = SQL & vbCrLf & "	LEFT OUTER JOIN AppraisalTable AS b "
SQL = SQL & vbCrLf & "		ON a.Myear = b.Myear "
SQL = SQL & vbCrLf & "		AND a.SubjectCode = b.SubjectCode "
'SQL = SQL & vbCrLf & "	/* 생기부 (IPSI215 : 인적사항(단일), IPSI213 : 교과성적(복수), IPSI217 : 학사사항(복수), IPSI237 : 출결사항(복수)) */"
SQL = SQL & vbCrLf & "	LEFT OUTER JOIN ( "
SQL = SQL & vbCrLf & "		SELECT "
SQL = SQL & vbCrLf & "			b.StudentRecordRatio, a.StudentNumber, c.EXAM_NUMB as StdentRecord1, c.Cors_Code, c.Majr_cs11, c.Majr_cs12, c.Majr_cs21, c.Majr_cs22, c.Majr_cs31, c.Majr_cs32, d.EXAM_NUMB as StdentRecord2, e.EXAM_NUMB as StdentRecord3, f.EXAM_NUMB as StdentRecord4  "
SQL = SQL & vbCrLf & "		FROM ApplicationTable AS a "
SQL = SQL & vbCrLf & "			LEFT OUTER JOIN AppraisalTable AS b "
SQL = SQL & vbCrLf & "				ON (a.Myear = b.Myear AND a.SubjectCode = b.SubjectCode)   "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( select b.EXAM_NUMB, b.Cors_Code, b.Majr_cs11, b.Majr_cs12, b.Majr_cs21, b.Majr_cs22, b.Majr_cs31, b.Majr_cs32 FROM ApplicationTable a JOIN IPSI215 b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB)) c  "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = c.EXAM_NUMB "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( select b.EXAM_NUMB FROM ApplicationTable a JOIN IPSI213 b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB) GROUP BY b.EXAM_NUMB) d  "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = d.EXAM_NUMB "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( select b.EXAM_NUMB FROM ApplicationTable a JOIN IPSI217 b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB) GROUP BY b.EXAM_NUMB) e  "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = e.EXAM_NUMB "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( select b.EXAM_NUMB FROM ApplicationTable a JOIN IPSI237 b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB) GROUP BY b.EXAM_NUMB) f  "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = f.EXAM_NUMB "
SQL = SQL & vbCrLf & "		WHERE 1 = 1    "
SQL = SQL & vbCrLf & "			AND (b.StudentRecordRatio is not null OR b.StudentRecordRatio <> '' OR b.StudentRecordRatio <> '0')  "
SQL = SQL & vbCrLf & "			AND (c.EXAM_NUMB IS NOT NULL AND d.EXAM_NUMB IS NOT NULL AND e.EXAM_NUMB IS NOT NULL AND f.EXAM_NUMB IS NOT NULL) "
SQL = SQL & vbCrLf & "	) AS c "
SQL = SQL & vbCrLf & "		ON a.StudentNumber = c.StudentNumber "
'SQL = SQL & vbCrLf & "	/* IPSICSAT 수능 */"
SQL = SQL & vbCrLf & "	LEFT OUTER JOIN ( "
SQL = SQL & vbCrLf & "		SELECT "
SQL = SQL & vbCrLf & "			b.CSATRatio, a.StudentNumber, c.EXAM_NUMB AS StdentRecord "
SQL = SQL & vbCrLf & "		FROM ApplicationTable AS a "
SQL = SQL & vbCrLf & "			LEFT OUTER JOIN AppraisalTable AS b "
SQL = SQL & vbCrLf & "				ON (a.Myear = b.Myear AND a.SubjectCode = b.SubjectCode) "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( SELECT b.EXAM_NUMB FROM ApplicationTable a JOIN IPSICSAT b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB)) AS c "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = c.EXAM_NUMB "
SQL = SQL & vbCrLf & "		WHERE 1 = 1 "
SQL = SQL & vbCrLf & "			AND (b.CSATRatio IS NOT NULL OR b.CSATRatio <> '' OR b.CSATRatio <> '0') "
SQL = SQL & vbCrLf & "			AND c.EXAM_NUMB IS NOT NULL "
SQL = SQL & vbCrLf & "	) AS d "
SQL = SQL & vbCrLf & "		ON a.StudentNumber = d.StudentNumber "
'SQL = SQL & vbCrLf & "	/* IPSI212 검정고시 */"
SQL = SQL & vbCrLf & "	LEFT OUTER JOIN ( "
SQL = SQL & vbCrLf & "		SELECT "
SQL = SQL & vbCrLf & "			b.StudentRecordRatio, a.Qualification, a.StudentNumber, c.EXAM_NUMB AS StdentRecord "
SQL = SQL & vbCrLf & "		FROM ApplicationTable AS a "
SQL = SQL & vbCrLf & "			LEFT OUTER JOIN AppraisalTable AS b "
SQL = SQL & vbCrLf & "				ON (a.Myear = b.Myear AND a.SubjectCode = b.SubjectCode) "
SQL = SQL & vbCrLf & "			FULL OUTER JOIN ( SELECT b.EXAM_NUMB FROM ApplicationTable a JOIN IPSI212 b ON (a.Myear = b.SCHL_YEAR AND a.StudentNumber = b.EXAM_NUMB) GROUP BY b.EXAM_NUMB) AS c "
SQL = SQL & vbCrLf & "				ON a.StudentNumber = c.EXAM_NUMB "
SQL = SQL & vbCrLf & "		WHERE 1 = 1 "
SQL = SQL & vbCrLf & "			AND (b.StudentRecordRatio IS NOT NULL OR b.StudentRecordRatio <> '' OR b.StudentRecordRatio <> '0') "
SQL = SQL & vbCrLf & "			AND c.EXAM_NUMB IS NOT NULL "
SQL = SQL & vbCrLf & "			AND a.Qualification = '1' "
SQL = SQL & vbCrLf & "	) AS e "
SQL = SQL & vbCrLf & "		ON a.StudentNumber = e.StudentNumber "
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & "	AND a.MYear = ? "
SQL = SQL & vbCrLf & "	AND a.Division0 = ? "

Call objDB.sbSetArray("@MYear", adVarchar, adParamInput, 4, MYear)
Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 60, Division0)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

If isArray(AryHash) Then
	For i = 0 to ubound(AryHash,1)		
		'면접 점수 입력 필요 여부
		'AryHash(i).Item("InterviewerRatio")    '면접 비율
		If isnull(AryHash(i).Item("InterviewerRatio")) Or AryHash(i).Item("InterviewerRatio") = "" Or AryHash(i).Item("InterviewerRatio") = "0" Then
			InterviewerStr = "-"
		Else
			if AryHash(i).Item("InterviewerPoint") <= "0" Or isnull(AryHash(i).Item("InterviewerPoint")) Then
				InterviewerStr = "Y"
			Else
				InterviewerStr = "N"
			End If
		End If

		'실기 점수 입력 필요 여부
		'AryHash(i).Item("PracticalRatio")      '실기 비율
		If isnull(AryHash(i).Item("PracticalRatio")) Or AryHash(i).Item("PracticalRatio") = "" Or AryHash(i).Item("PracticalRatio") = "0" Then
			PracticalStr = "-"
		Else
			if AryHash(i).Item("PracticalPoint") <= "0" Or isnull(AryHash(i).Item("PracticalPoint")) Then
				PracticalStr = "Y"
			Else
				PracticalStr = "N"
			End If
		End If		

		'생기부 데이터 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)									
		If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
			StudentRecordDataStr = "-"
		Else
			If isnull(AryHash(i).Item("StuRecCheck")) Then
				StudentRecordDataStr = "N"
			ElseIf Not(isnull(AryHash(i).Item("StuRecCheck"))) Then
				StudentRecordDataStr = "Y"
			End If
		End If

		'검정고시 데이터 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)
		If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
			QualificationDataStr = "-"
		Else
			'검정고시자면
			If AryHash(i).Item("Qualification") = "1" Then
				If isnull(AryHash(i).Item("QualCheck")) Then
					QualificationDataStr = "N"
				ElseIf Not(isnull(AryHash(i).Item("QualCheck"))) Then
					QualificationDataStr = "Y"
				End If	
			Else
				QualificationDataStr = "-"
			End If	
		End If

		'수능 데이터 체크(지원한 모집단위(모집시기+학과+전형)에 수능 비율이 있으면)
		If isnull(AryHash(i).Item("CSATRatio")) Or AryHash(i).Item("CSATRatio") = "" Or AryHash(i).Item("CSATRatio") = "0" Then
			CSATDataStr = "-"
		Else
			If isnull(AryHash(i).Item("CSATCheck")) Then
				CSATDataStr = "N"
			ElseIf Not(isnull(AryHash(i).Item("CSATCheck"))) Then
				CSATDataStr = "Y"
			End If
		End If
		
		'고교유형체크(일반고전형, 전문(직업)과정 전형)
		'0685 = 7차교육 과정(일반고)
		'***자율고 등 다른 코드도 추가해야 함
		'***아래 로직으로 생기부 등록 시 insert하여 그냥 값 받아와서 쓰는 걸로 변경 필요						

		Select case AryHash(i).Item("Semester")
			Case "1" : Semester = "Majr_cs11"
			Case "2" : Semester = "Majr_cs12"
			Case "3" : Semester = "Majr_cs21"
			Case "4" : Semester = "Majr_cs22"
			Case "5" : Semester = "Majr_cs31"
		End Select
		
		DrawStandard = "N"
		If AryHash(i).Item("Division1") = "X05041" Then '일반고전형
			If isnull(AryHash(i).Item("Cors_Code")) Then
				DrawStandard = "Y"
			Else
				If AryHash(i).Item("Cors_Code") <> "0685" Then
					'일반고, 자율고가 아님
					DrawStandard = "Y"						
				ElseIf AryHash(i).Item(Semester) <> " " Then
				'ElseIf Not(isnull(AryHash(i).Item(Semester))) Then
					'일반고, 자율고이나 선택한 학기에 직업과정 위탁교육을 이수하였음.
					DrawStandard = "Y"
				End If
			End If
		ElseIf AryHash(i).Item("Division1") = "X05042" Then '전문(직업)과정 전형
			If isnull(AryHash(i).Item("Cors_Code")) Then
				DrawStandard = "Y"
			Else
				If AryHash(i).Item("Cors_Code") = "0685" Then	
					If AryHash(i).Item(Semester) = " " Then
						'전문(직업)과정이 아님.
						DrawStandard = "Y"
					End If
				End If
			End If
		End If

		If DrawStandard <> "Y" Then
			DrawStandard = AryHash(i).Item("DrawStandard")
		End If

		'최종완료 체크(N이 있으면, 미완료)
		If StudentRecordDataStr = "N" Or QualificationDataStr = "N" Or CSATDataStr = "N" Or InterviewerStr = "Y" Or PracticalStr = "Y" Or DrawStandard <> "N" Then
			ReslutStr = "N(미완료)"
		Else
			ReslutStr = "Y(완료)"
		End If

		'// 데이터, 최종 값 넣어주기 ================
		SQL = ""
		SQL = SQL & vbCrLf & " UPDATE ApplicationTable "
		SQL = SQL & vbCrLf & " SET	  StudentRecordData = ?, QualificationData = ?, CSATData = ?, Reslut = ?  "
		SQL = SQL & vbCrLf & " WHERE MYear = ? "
		SQL = SQL & vbCrLf & " AND StudentNumber = ? "

		arrParams = Array(_
			  Array("@StudentRecordData",			adVarchar,		adParamInput,		10,		StudentRecordDataStr) _
			, Array("@QualificationData",			adVarchar,		adParamInput,		10,		QualificationDataStr) _
			, Array("@CSATData",					adVarchar,		adParamInput,		10,		CSATDataStr) _
			, Array("@Reslut",						adVarchar,		adParamInput,		10,		ReslutStr) _
			, Array("@MYear",						adVarchar,		adParamInput,		50,		AryHash(i).Item("Myear")) _
			, Array("@StudentNumber",				adVarchar,		adParamInput,		50,		AryHash(i).Item("StudentNumber")) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)
	Next
End If

'// ==============================================================================================================
'// 데이터, 최종 값 넣어주기 끝
'// ==============================================================================================================

'// ==============================================================================================================
'// 최종이 완료인 지원자만 뽑기 
'// ==============================================================================================================

SQL = ""
SQL = SQL & vbCrLf & " Select *  "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Division0', Division0) AS Division0Name "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & "		  , dbo.getSubCodeName('Division1', Division1) AS Division1Name "
SQL = SQL & vbCrLf & " from ApplicationTable "
SQL = SQL & vbCrLf & " where 1 = 1  "
SQL = SQL & vbCrLf & " and MYear = ? "
SQL = SQL & vbCrLf & " and Division0 = ? "
SQL = SQL & vbCrLf & " and Reslut = 'Y(완료)' "
SQL = SQL & vbCrLf & " order by StudentNumber "

Call objDB.sbSetArray("@MYear", adVarchar, adParamInput, 4, MYear)
Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 60, Division0)

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'// ==============================================================================================================
'// 최종이 완료인 지원자만 뽑기 끝
'// ==============================================================================================================

Set objDB  = Nothing

If Not IsNull(AryHash) Then
	intNUM = ubound(AryHash,1) + 1
	for i = 0 to ubound(AryHash,1)
		response.write "<tr class='viewDetail'>"
		response.write "<td>" & intNUM  & "</td>"
		response.write "<td>" & AryHash(i).Item("Myear")  & "</td>"
		response.write "<td>" & AryHash(i).Item("Division0Name")  & "</td>"
		response.write "<td>" & AryHash(i).Item("SubjectName")  & "</td>"
		response.write "<td>" & AryHash(i).Item("Division1Name")  & "</td>"
		response.write "<td>" & AryHash(i).Item("StudentNumber")  & "</td>"
		response.write "<td>" & AryHash(i).Item("StudentNameKor")  & "</td>"
		response.write "<td>" & AryHash(i).Item("Reslut")  & "</td>"
		response.write "</tr>"
		intNUM = intNUM - 1
	Next
	response.write "count" & ubound(AryHash,1) + 1
Else
	response.write "<tr class='viewDetail'>"
	response.write "<td colspan='8'>해당 년도, 모집시기에 최종결과가 완료인 지원자가 없습니다.</td>"
	response.write "</tr>"
	response.write "count0"	
End If
%>