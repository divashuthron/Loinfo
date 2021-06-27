<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 5
Dim LeftMenuCode : LeftMenuCode = "Applicant"
Dim LeftMenuName : LeftMenuName = "Home / 지원자관리 / 지원자 조회"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "지원자 조회"
Dim LogDivision	: LogDivision = "ApplicantList"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, AryHash2
Dim i, strMSG, intNUM, strTEMP, strRESULT

'검색조건
Dim SearchMYear		: SearchMYear = fnR("SearchMYear", "")
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", "")
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", "")
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", "")
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchType2		: SearchType2 = fnR("searchType2", "")
Dim SearchText2		: SearchText2 = fnR("searchText2", "")
Dim SearchTextTemp
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/ApplicationList.asp"
Dim SearchPart
Dim SearchPart2
Dim SearchPart3
Dim SearchPart4
Dim SearchPart5
Dim SearchPart6

'페이지설정(사이즈는 검색)
Dim PageSize		: PageSize = getIntParameter(FnR("PageSize", 5), 5)
Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'생기부/검정/수능 데이터, 최종 값
Dim AryHash3, AryHash4, AryHash5, AryHash6, AryHash7, AryHash8
Dim CSATCheck,QualCheck, StuRecCheck
Dim StudentRecordDataStr, QualificationDataStr, CSATDataStr, ReslutStr
Dim InterviewerStr, PracticalStr

'DBOpen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'모집시기 검색조건 쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And a.Division0 = ? "
	Call objDB.sbSetArray("@Division", adVarchar, adParamInput, 50, SearchDivision)
end If

'학과 검색조건 쿼리
if not(IsE(SearchSubject)) And SearchSubject <> "All" then
	strWhere = strWhere & " And a.Subject = ? "
	Call objDB.sbSetArray("@Subject", adVarchar, adParamInput, 50, SearchSubject)
end if

'학생 검생조건 쿼리
if (not(IsE(SearchText))) then
	if SearchType = "1" then
		SearchPart = "StudentNumber"
	elseif SearchType = "2" then
		 SearchPart = "StudentNameKor"
	Else
		SearchType = "1"
		SearchPart = "StudentNumber"
	end if
	
	strWhere = strWhere & " And a."& SearchPart &" like '%' + ? + '%' "
	Call objDB.sbSetArray("@SearchText", adVarchar, adParamInput, 255, SearchText)
end If

'여부 검색조건 서브쿼리 만들기
if (not(IsE(SearchText2))) then
	'서브 쿼리

	'1번 가산점 
	If SearchType2 = "1" Then 
		If SearchText2 = "1" Then 
			strWhere = strWhere & "And (a.ExtraPoint is not null And a.ExtraPoint <> '0')  "
		ElseIf SearchText2 = "2" Then									
			strWhere = strWhere & "And (a.ExtraPoint is null or  a.ExtraPoint = '' or  a.ExtraPoint = '0')  " 
		Else
			strWhere = strWhere & "And (a.ExtraPoint = '999') "
		End If
	'2번 자격미달 
	ElseIf SearchType2 = "2" Then 
		If SearchText2 = "1" Then 									
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'Y')  "
		ElseIf SearchText2 = "2" Then									
			strWhere = strWhere & "And (DrawStandard is null or  DrawStandard = '' or  DrawStandard = 'N')  " 
		ElseIf SearchText2 = "3" Then
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'C')  " 
		ElseIf SearchText2 = "4" Then
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'D')  " 
		ElseIf SearchText2 = "5" Then
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'E')  " 
		ElseIf SearchText2 = "6" Then
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'F')  " 
		ElseIf SearchText2 = "7" Then
			strWhere = strWhere & "And (DrawStandard is not null and DrawStandard = 'G')  " 
		Else
			strWhere = strWhere & "And (DrawStandard = '999') "
		End If
	'3번 위반자는 추가 예정

	'4번 생기부 동의 
	ElseIf SearchType2 = "4" Then	
		If SearchText2 = "1" Then 
			strWhere = strWhere & "And  (B.StudentRecordRatio Is not null And B.StudentRecordRatio <> '' And B.StudentRecordRatio <> '0' )  "
			strWhere = strWhere & "And A.StudentRecordAgreement  = '1'  "
		ElseIf SearchText2 = "2" Then     	
			strWhere = strWhere & "And  (B.StudentRecordRatio <> '0' )  "
			strWhere = strWhere & "And (A.StudentRecordAgreement is null Or A.StudentRecordAgreement  = '0' )  "
		Else
			strWhere = strWhere & "And B.StudentRecordRatio = '999' "
		End If
	'5번 검정 동의 
	ElseIf SearchType2 = "5" Then 
		If SearchText2 = "1" Then 										   
			strWhere = strWhere & "And (B.StudentRecordRatio Is not null and B.StudentRecordRatio <> '' and B.StudentRecordRatio <> '0' )    "
			strWhere = strWhere & "And ((A.QualificationYear is not null And A.QualificationYear <> '') Or (A.QualificationAreaCode is not null And A.QualificationAreaCode <> ''))  "
			strWhere = strWhere & "And A.QualificationAgreement  = '1'    "
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "And (B.StudentRecordRatio Is not null and B.StudentRecordRatio <> '' and B.StudentRecordRatio <> '0' )    "
			strWhere = strWhere & "And ((A.QualificationYear is not null And A.QualificationYear <> '') Or (A.QualificationAreaCode is not null And A.QualificationAreaCode <> ''))  "
			strWhere = strWhere & "And A.QualificationAgreement is null Or A.QualificationAgreement  = '0'    "
		Else
			strWhere = strWhere & "And B.StudentRecordRatio = '999' "
		End If
	'6번 수능 동의 
	ElseIf SearchType2 = "6" Then	
		If SearchText2 = "1" Then 
			strWhere = strWhere & "And  (B.CSATRatio Is not null And B.CSATRatio <> '' And B.CSATRatio <> '0' )  "
			strWhere = strWhere & "And A.SDSN_AGYN  = '1'  "
		ElseIf SearchText2 = "2" Then     	
			strWhere = strWhere & "And  (B.CSATRatio <> '0' )  "
			strWhere = strWhere & "And (A.SDSN_AGYN is null Or A.SDSN_AGYN  = '0' )  "
		Else
			strWhere = strWhere & "And B.StudentRecordRatio = '999' "
		End If
	'7번 수동입력 
	ElseIf SearchType2 = "7" Then 
		If SearchText2 = "1" Then 
			strWhere = strWhere & "And  (((B.StudentRecordRatio Is not null And B.StudentRecordRatio <> '' And B.StudentRecordRatio <> '0' )   "
			strWhere = strWhere & "		And (A.StudentRecordAgreement is null Or A.StudentRecordAgreement  = '0' ))    "
			strWhere = strWhere & "Or  ((B.StudentRecordRatio Is not null And B.StudentRecordRatio <> '' And B.StudentRecordRatio <> '0' )   "
			strWhere = strWhere & "		And (((A.HighCode is null Or A.HighCode = '') And (A.HighGraduationYear is null Or A.HighGraduationYear = ''))   "
			strWhere = strWhere & "		And ((A.QualificationYear is not null And A.QualificationYear <> '') Or (A.QualificationAreaCode is not null And A.QualificationAreaCode <> '')))  "
			strWhere = strWhere & "		And (A.QualificationAgreement is null Or A.QualificationAgreement  = '0' ))   "
			strWhere = strWhere & "Or  ((B.CSATRatio Is not null And B.CSATRatio <> '' And B.CSATRatio <> '0' )   "
			strWhere = strWhere & "		And (A.SDSN_AGYN is null Or A.SDSN_AGYN  = '0' )))    "
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "And  not(((B.StudentRecordRatio Is not null And B.StudentRecordRatio <> '' And B.StudentRecordRatio <> '0' )   "
			strWhere = strWhere & "		And (A.StudentRecordAgreement is null Or A.StudentRecordAgreement  = '0' ))    "
			strWhere = strWhere & "Or  ((B.StudentRecordRatio Is not null And B.StudentRecordRatio <> '' And B.StudentRecordRatio <> '0' )   "
			strWhere = strWhere & "		And (((A.HighCode is null Or A.HighCode = '') And (A.HighGraduationYear is null Or A.HighGraduationYear = ''))   "
			strWhere = strWhere & "		And ((A.QualificationYear is not null And A.QualificationYear <> '') Or (A.QualificationAreaCode is not null And A.QualificationAreaCode <> '')))  "
			strWhere = strWhere & "		And (A.QualificationAgreement is null Or A.QualificationAgreement  = '0' ))   "
			strWhere = strWhere & "Or  ((B.CSATRatio Is not null And B.CSATRatio <> '' And B.CSATRatio <> '0' )   "
			strWhere = strWhere & "		And (A.SDSN_AGYN is null Or A.SDSN_AGYN  = '0' )))    "
		Else
			strWhere = strWhere & "And B.StudentRecordRatio = '999' "
		End If	
	'8번 생기부데이터 
	ElseIf SearchType2 = "8" Then 
		If SearchText2 = "1" Then 	
			strWhere = strWhere & "and c.StudentNumber is not null "		
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "and c.StudentNumber is null "
			strWhere = strWhere & "and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0') "
		End If
	'9번 검정데이터
	ElseIf SearchType2 = "9" Then 
		If SearchText2 = "1" Then 	
			strWhere = strWhere & "and e.StudentNumber is not null "		
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "and e.StudentNumber is null "
			strWhere = strWhere & "and a.Qualification = '1' "
			strWhere = strWhere & "and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0') "
		End If
	'10번 수능데이터
	ElseIf SearchType2 = "10" Then 
		If SearchText2 = "1" Then 	
			strWhere = strWhere & "and d.StudentNumber is not null "		
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "and d.StudentNumber is null "
			strWhere = strWhere & "and (b.CSATRatio is not null Or b.CSATRatio <> '' Or b.CSATRatio <> '0') "
		End If
	'11번 최종완료
	ElseIf SearchType2 = "11" Then 
		If SearchText2 = "1" Then 	
			strWhere = strWhere & "and not((c.StudentNumber is null and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0'))   "
			strWhere = strWhere & "	Or (d.StudentNumber is null and (b.CSATRatio is not null Or b.CSATRatio <> '' Or b.CSATRatio <> '0'))  "
			strWhere = strWhere & "	Or (e.StudentNumber is null and a.Qualification = '1' and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0'))  "
			strWhere = strWhere & "	Or a.DrawStandard <> 'N'  "
			strWhere = strWhere & "	Or a.Document <> 'N') "
		ElseIf SearchText2 = "2" Then
			strWhere = strWhere & "and ((c.StudentNumber is null and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0'))   "
			strWhere = strWhere & "	Or (d.StudentNumber is null and (b.CSATRatio is not null Or b.CSATRatio <> '' Or b.CSATRatio <> '0'))  "
			strWhere = strWhere & "	Or (e.StudentNumber is null and a.Qualification = '1' and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0'))  "
			strWhere = strWhere & "	Or a.DrawStandard <> 'N'  "
			strWhere = strWhere & "	Or a.Document <> 'N') "
		End If
	End If
end If

'리스트 쿼리
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
SQL = SQL & vbCrLf & "		, a.SDSN_AGYN, a.SDAG_AAYN, a.SDAG_BBYN, a.SDAG_CCYN, a.SDAG_DDYN  "
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
SQL = SQL & vbCrLf & strWhere
SQL = SQL & vbCrLf & "ORDER BY a.IDX DESC; "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'개인정보가 있는 입학원서는 조회도 기록함
strLogMSG = "지원자관리  > " & SessionUserID  &"가/이 지원자 리스트를 조회 했습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

if IsArray(AryHash) Then
	'// 페이지 계산
	TotalCount = ubound(AryHash,1) + 1
	PageCount = int((TotalCount - 1) / PageSize) + 1
	StartNum = (PageNum * PageSize) - PageSize
	EndNum = StartNum + PageSize - 1
	intNUM = TotalCount - (PageNum * PageSize) + PageSize

	If EndNum > TotalCount - 1 Then
		EndNum = TotalCount - 1
	End If
End If
%>


<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 검색조건 -->
			<div class="ibox-title">
				<h5>검색정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div>
					<form id="SearchForm" method="get">
					<input type="hidden" name="Page" value="<%= PageNum %>">

						<div class="row show-grid">
							<div class="col-md-1 col-xs-1 grid_sub_title">
								모집시기
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchDivision", "모집시기선택", SearchDivision, "", "All", "Division0") %>
							</div>
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								학과
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchSubject", "학과명선택", SearchSubject, "", "All", "Subject") %>
							</div>

							<div class="col-md-1 col-xs-1 grid_sub_title2">
								학생
							</div>
							<div class="col-md-2 col-xs-2 grid_sub_title">
								<select name="searchType" id="searchType" class="form-control input-sm">
									<option value="">구분</option>
									<option value="1" <%= setSelected(searchType, "1") %>>수험번호</option>
									<option value="2" <%= setSelected(searchType, "2") %>>이름(한글)</option>
								</select>
							</div>
							<div class="col-md-2 col-xs-2">
								<input type="text" name="searchText" id="searchText" value="<%= SearchText %>" class="form-control input-sm"/>
							</div>

							<div class="col-md-1 col-xs-1 grid_sub_title2">
								여부검색
							</div>
							<div class="col-md-2 col-xs-2 grid_sub_title">
								<select name="searchType2" id="searchType2" class="form-control input-sm">
									<option value="">구분</option>
									<option value="1" <%= setSelected(searchType2, "1") %>>가산점여부</option>
									<option value="2" <%= setSelected(searchType2, "2") %>>자격미달여부</option>
									<option value="3" <%= setSelected(searchType2, "3") %>>위반자여부</option>
									<option value="4" <%= setSelected(searchType2, "4") %>>생기부동의</option>
									<option value="5" <%= setSelected(searchType2, "5") %>>검정동의</option>
									<option value="6" <%= setSelected(searchType2, "6") %>>수능동의</option>
									<option value="7" <%= setSelected(searchType2, "7") %>>수동입력</option>
									<option value="8" <%= setSelected(searchType2, "8") %>>생기부데이터</option>
									<option value="9" <%= setSelected(searchType2, "9") %>>검정데이터</option>
									<option value="10" <%= setSelected(searchType2, "10") %>>수능데이터</option>
									<option value="11" <%= setSelected(searchType2, "11") %>>최종완료</option>
								</select>
							</div>
							<div class="col-md-2 col-xs-2">
								<select name="searchText2" id="searchText2" class="form-control input-sm">
									<option value="">구분</option>
									<option value="1" <%= setSelected(searchText2, "1") %>>Y</option>
									<option value="2" <%= setSelected(searchText2, "2") %>>N</option>
									<option value="3" <%= setSelected(searchText2, "3") %>>C (서류필요)</option>
									<option value="4" <%= setSelected(searchText2, "4") %>>D (면접/실기 점수 미입력)</option>
									<option value="5" <%= setSelected(searchText2, "5") %>>E (면접 점수 미입력)</option>
									<option value="6" <%= setSelected(searchText2, "6") %>>F (실기 점수 미입력)</option>
									<option value="7" <%= setSelected(searchText2, "7") %>>G (평가비율미입력)</option>
								</select>
							</div>
						</div>
						<div class="pad_t10 pad_r10 text-right">
							<span class="btnBasic btnSubmit">지원자 조회</span>
						</div>
					<!--</form>-->
				</div>
			</div>
			<!-- 검색조건 끝 -->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div style="float:right;">
					<!-- 게시물 갯수 선택 -->
					<!--<a class="collapse-link">-->
					<!--<form id="PageSizeForm" method="get">-->
						<a href="#"><span class="btnBasic btnTypeComplete" onclick="alert('생기부 엑셀로 업로드가 지원되지 않습니다.');return false;">생기부 엑셀샘플</span></a>
						<a href="/Download/수능샘플.xlsx"><span class="btnBasic btnTypeComplete">수능 엑셀샘플</span></a>
						<span class="btnBasic btnTypeExcel" onclick="alert('생기부 엑셀로 업로드가 지원되지 않습니다.');return false;">생기부 엑셀로 등록</span>
						<span class="btnBasic btnTypeExcel" id="btnExcel" onClick="window.open('./CSATUpload.asp','CSATUpload','toolbar=no,menubar=no,scrollbars=no,resizable=no,width=1200 height=615'); return false;">수능성적 엑셀로 등록</span>
						<span class="btnBasic btnTypeAccept" onclick="alert('중간테이블과 연결되어 있지 않습니다.');return false;">생기부 가져오기</span>
						<span class="btnBasic btnTypeAccept" onclick="alert('중간테이블과 연결되어 있지 않습니다.');return false;">수능성적 가져오기</span>
						<select name = "PageSize" style="margin-left:10px;" onChange="SearchForm.submit();">
							<option value="5" <% If PageSize = 5 then response.write "selected" end if%>>5개씩 보기</option>
							<option value="15" <% If PageSize = 15 then response.write "selected" end if%>>15개씩 보기</option>
							<option value="30" <% If PageSize = 30 then response.write "selected" end if%>>30개씩 보기</option>
							<option value="50" <% If PageSize = 50 then response.write "selected" end if%>>50개씩 보기</option>
							<option value="100" <% If PageSize = 100 then response.write "selected" end if%>>100개씩 보기</option>
							<option value="200" <% If PageSize = 200 then response.write "selected" end if%>>200개씩 보기</option>
						</select>
					</form>
						<!--<i class="fa fa-chevron-up"></i>-->
					<!--</a>-->
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<colgroup>
								<col width="3%"></col>
								<col width="4%"></col>
								<col width="6%"></col>
								<col width="11%"></col>
								<col width="18%"></col>
								<col width="3%"></col>
								<col width="5%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="7%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="7%"></col>
							</colgroup>
							<thead>			                
								<tr>
									<!--<th colspan="1" style="padding: 7px 0px 6px 0px; text-align: center;">
										<input type="checkbox" id="checkall"/>
									</th>-->
									<th data-hide="phone">No.</th>
									<th data-hide="phone">년도</th>
									<th>시기</th>
									<th>학과</th>
									<th>전형</th>
									<th data-hide="phone,tablet">수험번호</th>
									<th data-hide="phone">이름</th>
									<th data-hide="phone">가산점</th>
									<th data-hide="phone">미달</th>
									<th data-hide="phone">위반자</th>									
									<th data-hide="phone">생동의</th>
									<th data-hide="phone">검동의</th>
									<th data-hide="phone">수동의</th>
									<th data-hide="phone">수동입력</th>
									<th data-hide="phone">생데이터</th>
									<th data-hide="phone">검데이터</th>
									<th data-hide="phone">수데이터</th>
									<th data-hide="phone">최종완료</th>
								</tr>
							</thead>
							<tbody>
							<%
								Dim StudentRecordStr, QualificationStr, CSATStr, ManualStr
								Dim DrawTemp, DocumentTemp, DrawStandard, Semester						

								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then								
									For i = StartNum to EndNum
									' e.StudentNumber as QualCheck, d.StudentNumber as CSATCheck, c.StudentNumber as StuRecCheck 
									' , a.IDX, a.Myear, a.Division0, a.StudentNumber, a.StudentNameKor, a.StudentNameUsa, a.StudentNameChi  
									' , a.Citizen1, a.Citizen2, a.Sex, a.HighCode, a.HighSubject, a.HighGraduationYear, a.HighGraduationDivision  
									' , a.QualificationAreaCode, a.QualificationYear, a.Subject, a.Semester  
									' , a.UniversityName, a.AugScore, a.PerfectScore, a.Credit  
									' , a.Division1, a.Division2, a.Division3, a.Division4, a.HighDivision  
									' , a.RefundDivision, a.RefundAccountHolder, a.RefundBankCode, a.RefundAccount  
									' , a.DrawStandard, a.DrawMsg ,a.ExtraPoint, a.InterviewerPoint, a.PracticalPoint, a.Tel1, a.Tel2, a.Tel3, a.Email, a.ZipCode, a.Address1, a.Address2  
									' , a.StudentNameAgreement, a.StudentAgreement, a.StudentRecordAgreement, a.QualificationAgreement  
									' , a.SDSN_AGYN, a.SDAG_AAYN, a.SDAG_BBYN, a.SDAG_CCYN, a.SDAG_DDYN  
									' , a.ReceiptDate, a.ReceiptTime  
									' , a.DocumentaryCheck1, a.DocumentaryCheck2, a.DocumentaryCheck3, a.DocumentaryCheck4 "
									' , a.DocumentaryCheck5, a.DocumentaryCheck6, a.DocumentaryCheck7, a.DocumentaryCheck8 "
									' , a.INPT_USID, a.INPT_DATE, a.INPT_ADDR, a.UPDT_USID, a.UPDT_DATE, a.UPDT_ADDR, a.InsertTime  
									' , dbo.getSubCodeName('Division0', a.Division0) AS DivisionName  
									' , dbo.getSubCodeName('Subject', a.Subject) AS SubjectName  
									' , dbo.getSubCodeName('Division1', a.Division1) AS Division1Name  
									' , dbo.getSubCodeName('HignSchoolDivision', a.HighDivision) AS HighDivisionName  
									' , b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio 
									' , b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5 
									' , b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5 
									' , b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5 
									' , b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5 
									' , b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5 
									
									'생기부 동의 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)									
									If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
										StudentRecordStr = "-"
									Else
										if AryHash(i).Item("StudentRecordAgreement") = "1" Then
											StudentRecordStr = "Y"
										Else
											StudentRecordStr = "N"
										End If
									End If

									'검정고시 동의 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)
									If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
										QualificationStr = "-"
									Else
										'검정고시 자면
										If AryHash(i).Item("Qualification") = "1" Then
											if AryHash(i).Item("QualificationAgreement") = "1" Then
												QualificationStr = "Y"
											Else
												QualificationStr = "N"
											End If
										Else
											QualificationStr = "-"
										End If	
									End If

									'수능 동의 체크(지원한 모집단위(모집시기+학과+전형)에 수능 비율이 있으면)
									If isnull(AryHash(i).Item("CSATRatio")) Or AryHash(i).Item("CSATRatio") = "" Or AryHash(i).Item("CSATRatio") = "0" Then
										CSATStr = "-"
									Else
										if AryHash(i).Item("SDSN_AGYN") = "1" Then
											CSATStr = "Y"
										Else
											CSATStr = "N"
										End If
									End If				
									
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
									DrawStandard = AryHash(i).Item("DrawStandard")
									DrawTemp = AryHash(i).Item("DrawMsg")							

									Select case AryHash(i).Item("Semester")
										Case "1" : Semester = "Majr_cs11"
										Case "2" : Semester = "Majr_cs12"
										Case "3" : Semester = "Majr_cs21"
										Case "4" : Semester = "Majr_cs22"
										Case "5" : Semester = "Majr_cs31"
									End Select

									If AryHash(i).Item("Division1") = "X05041" Then '일반고전형
										If isnull(AryHash(i).Item("Cors_Code")) Then
											DrawStandard = "Y"
											If isnull(DrawTemp) Then
											'If LEN(DrawMsg) < 1 Then
												DrawTemp = "<b>일반고전형체크-생기부가 등록되지 않았습니다.</b>"
											Else
												DrawTemp = DrawTemp & "= <b>일반고전형체크-생기부가 등록되지 않았습니다.</b>"
											End If
										Else
											If AryHash(i).Item("Cors_Code") <> "0685" Then
												'일반고, 자율고가 아님
												DrawStandard = "Y"
												If isnull(DrawTemp) Then
												'If LEN(DrawMsg) < 1 Then
													DrawTemp = "<b>일반고전형에 해당하는 교육과정을 이수하지 않았습니다.</b>"
												Else
													DrawTemp = DrawTemp & "= <b>일반고전형에 해당하는 교육과정을 이수하지 않았습니다.</b>"
												End If							
											ElseIf AryHash(i).Item(Semester) <> " " Then
											'ElseIf Not(isnull(AryHash(i).Item(Semester))) Then
												'일반고, 자율고이나 선택한 학기에 직업과정 위탁교육을 이수하였음.
												DrawStandard = "Y"
												If isnull(DrawTemp) Then
												'If LEN(DrawMsg) < 1 Then
													DrawTemp = "<b>선택한 학기에 직업과정 위탁교육을 이수하였습니다.</b>"
												Else
													DrawTemp = DrawTemp & "= <b>선택한 학기에 직업과정 위탁교육을 이수하였습니다.</b>"
												End If
											End If
										End If
									ElseIf AryHash(i).Item("Division1") = "X05042" Then '전문(직업)과정 전형
										If isnull(AryHash(i).Item("Cors_Code")) Then
											DrawStandard = "Y"
											If isnull(DrawTemp) Then
											'If LEN(DrawMsg) < 1 Then
												DrawTemp = "<b>전문(직업)과정전형체크-생기부가 등록되지 않았습니다.</b>"
											Else
												DrawTemp = DrawTemp & "= <b>전문(직업)과정전형체크-생기부가 등록되지 않았습니다.</b>"
											End If
										Else
											If AryHash(i).Item("Cors_Code") = "0685" Then	
												If AryHash(i).Item(Semester) = " " Then
													'전문(직업)과정이 아님.
													DrawStandard = "Y"
													If isnull(DrawTemp) Then
													'If LEN(DrawMsg) < 1 Then
														DrawTemp = "<b>전문(직업)과정전형에 해당하는 교육과정을 이수하지 않았습니다.</b>"
													Else
														DrawTemp = DrawTemp & "= <b>전문(직업)과정전형에 해당하는 교육과정을 이수하지 않았습니다.</b>"
													End If
												End If
											End If
										End If
									End If

									'수동입력 필요 체크(동의에 N이 있으면, 필요)
									If StudentRecordStr = "N" Or QualificationStr = "N" Or CSATStr = "N" Or InterviewerStr = "Y" Or PracticalStr = "Y" Then
										'동의하지 않았어도 데이터가 N이 없으면 불필요(수동입력 했기 때문에)
										If StudentRecordDataStr = "N" Or QualificationDataStr = "N" Or CSATDataStr = "N" Or InterviewerStr = "Y" Or PracticalStr = "Y" Then
											ManualStr = "Y(필요)"
										Else
											ManualStr = "N(불필요)"
										End If
									Else
										ManualStr = "N(불필요)"
									End If	

									'최종완료 체크(N이 있으면, 미완료)
									If StudentRecordDataStr = "N" Or QualificationDataStr = "N" Or CSATDataStr = "N" Or InterviewerStr = "Y" Or PracticalStr = "Y" Or DrawStandard <> "N" Then
										ReslutStr = "N(미완료)"
									Else
										ReslutStr = "Y(완료)"
									End If									

									'DrawMsg의 =를 <br>로 변경
									If Not isnull(DrawTemp) Then
										DrawTemp = replace(DrawTemp, "=", "<br>") 
									Else
										DrawTemp = "<b>목록이 없습니다.</b>"
									End If
							%>
								<tr class="viewDetail_SetDate_2">
									<!--<td colspan="1" style="background-color: <%=BGColor%>; padding: 8px 0px 0px 0px; text-align: center;">
										<input class="CheckboxCheck" type="Checkbox" name="chk" ID="Checkbox<%=i%>" value="<%=AryHash(i).Item("StudentNumber")%>" Myear="<%= AryHash(i).Item("Myear") %>">
									</td>-->
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("Myear")  %></td>
									<td><%= AryHash(i).Item("DivisionName") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("StudentNumber") %></td>
									<td><%= AryHash(i).Item("StudentNameKor") %></td>
									<td><% if AryHash(i).Item("ExtraPoint") > "0" Then%> Y <%Else%> N <%End If%></td>
									<td><%= DrawStandard %></td>
									<td></td>
									<td><%= StudentRecordStr %></td>
									<td><%= QualificationStr %></td>
									<td><%= CSATStr %></td>
									<td><b style="color:red;"><%=ManualStr%></b></td>
									<td><%= StudentRecordDataStr %></td>
									<td><%= QualificationDataStr %></td>
									<td><%= CSATDataStr %></td>
									<td>
										<b style="color:red;"> <%= ReslutStr %></b>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint")) %>"						ColumnName="ExtraPoint"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordAgreement")) %>"			ColumnName="StudentRecord"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationAgreement")) %>"			ColumnName="Qualification"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SDSN_AGYN")) %>"							ColumnName="CSAT"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck1")) %>"					ColumnName="document1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck2")) %>"					ColumnName="document2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck3")) %>"					ColumnName="document3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck4")) %>"					ColumnName="document4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck5")) %>"					ColumnName="document5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck6")) %>"					ColumnName="document6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck7")) %>"					ColumnName="document7"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck8")) %>"					ColumnName="document8"></li>
											<li Columnvalue="<%= DrawTemp %>"													ColumnName="DrawMsg"></li>

											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck1")) %>"					ColumnName="Check1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck2")) %>"					ColumnName="Check2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck3")) %>"					ColumnName="Check3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck4")) %>"					ColumnName="Check4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck5")) %>"					ColumnName="Check5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck6")) %>"					ColumnName="Check6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck7")) %>"					ColumnName="Check7"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck8")) %>"					ColumnName="Check8"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck21")) %>"				ColumnName="Check21"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck22")) %>"				ColumnName="Check22"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck23")) %>"				ColumnName="Check23"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck24")) %>"				ColumnName="Check24"></li>
											<li Columnvalue="<%= AryHash(i).Item("Myear") %>"									ColumnName="Myear"></li>
											<li Columnvalue="<%= AryHash(i).Item("StudentNumber") %>"							ColumnName="StudentNumber"></li>

											<li Columnvalue="<%= StudentRecordStr %>"											ColumnName="StudentRecordStr"></li>
											<li Columnvalue="<%= QualificationStr %>"											ColumnName="QualificationStr"></li>
											<li Columnvalue="<%= CSATStr %>"													ColumnName="CSATStr"></li>

											<li Columnvalue="<%= AryHash(i).Item("InterviewerRatio") %>"						ColumnName="InterviewerRatio"></li>
											<li Columnvalue="<%= AryHash(i).Item("PracticalRatio") %>"							ColumnName="PracticalRatio"></li>
											<li Columnvalue="<%= InterviewerStr %>"												ColumnName="InterviewerStr"></li>
											<li Columnvalue="<%= PracticalStr %>"												ColumnName="PracticalStr"></li>

											<li Columnvalue="<%= AryHash(i).Item("InterviewerPoint") %>"						ColumnName="InterviewerTemp"></li>
											<li Columnvalue="<%= AryHash(i).Item("PracticalPoint") %>"							ColumnName="PracticalTemp"></li>
										</div>
									</td>
								</tr>
							<%
										intNUM = intNUM - 1
									Next
								Else
							%>
								<tr>
									<td colspan="19" style="height:50px; vertical-align: middle;">검색된 자료가 없습니다.</td>
								</tr>
							<%
								end If
								
							Set objDB = Nothing
							%>
							</tbody>
						</table>
					</form>

					<div class="paging pad_r10">&nbsp;</div>
				</div>
				
				
			</div>
			<!-- 테이블 끝 -->			

			<div class="pad_t10"></div>

			<!-- 상세보기 -->
			<div class="ibox-title">
				<h5>상세정보</h5>
				<div style="float:right;">
					<span class="btnBasic btnTypeEdit" id="BasicDataSet">기본 데이터 설정</span>
					<span class="btnBasic btnTypeSave" id="BasicData1" onclick="BasicDataBtn(1)">1. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData2" onclick="BasicDataBtn(2)">2. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData3" onclick="BasicDataBtn(3)">3. 기본 데이터 넣기</span>
					<form id="BasicDataBtnFrom" method="post" action="/Process/BasicDataSelect.asp">
						<div style="display:none;">
							<input type="text" name="process" id="process" value="RegApplicantBasicDataSet">
							<input type="text" name="BasicDataBtnNum" id="BasicDataBtnNum" value="">
						</div>
					</form>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/ApplicantProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegApplicant">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="text" name="StudentNumberHidden" id="StudentNumberHidden" value="">
						<input type="text" name="MyearHidden" id="MyearHidden" value="">
						<input type="text" name="Check1" id="Check1" value="">
						<input type="text" name="Check2" id="Check2" value="">
						<input type="text" name="Check3" id="Check3" value="">
						<input type="text" name="Check4" id="Check4" value="">
						<input type="text" name="Check5" id="Check5" value="">
						<input type="text" name="Check6" id="Check6" value="">
						<input type="text" name="Check7" id="Check7" value="">
						<input type="text" name="Check8" id="Check8" value="">		
						<input type="text" name="Check21" id="Check21" value="">
						<input type="text" name="Check22" id="Check22" value="">
						<input type="text" name="Check23" id="Check23" value="">
						<input type="text" name="Check24" id="Check24" value="">	
						<input type="text" name="Myear" id="Myear" value="">
						<input type="text" name="StudentNumber" id="StudentNumber" value="">	
						<input type="text" name="StudentRecordStr" id="StudentRecordStr" value="">
						<input type="text" name="QualificationStr" id="QualificationStr" value="">
						<input type="text" name="CSATStr" id="CSATStr" value="">
						<input type="text" name="InterviewerRatio" id="InterviewerRatio" value="">
						<input type="text" name="PracticalRatio" id="PracticalRatio" value="">
						<input type="text" name="InterviewerStr" id="InterviewerStr" value="">
						<input type="text" name="PracticalStr" id="PracticalStr" value="">
						<input type="text" name="InterviewerTemp" id="InterviewerTemp" value="">
						<input type="text" name="PracticalTemp" id="PracticalTemp" value="">
					</div>
					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							가산점
						</div>
						<div class="col-md-3 col-xs-7 grid_sub_title">
							<% Call SubCodeSelectBox("ExtraPoint", "가산점 선택", "", "", "", "ExtraPoint") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							위반자
						</div>
						<div class="col-md-1 col-xs-7">
							<input type="text" name="HighCodeTemp" value="<%=HighCodeTemp%>" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							생기부동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<select name="StudentRecord" id="StudentRecord" class="form-control input-sm">
								<option value="">구분</option>
								<option value="1" <%= setSelected(StudentRecord, "1") %>>Y</option>
								<option value="0" <%= setSelected(StudentRecord, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							검정동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<select name="Qualification" id="Qualification" class="form-control input-sm">
								<option value="">구분</option>
								<option value="1" <%= setSelected(Qualification, "1") %>>Y</option>
								<option value="0" <%= setSelected(Qualification, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수능동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<select name="CSAT" id="CSAT" class="form-control input-sm">
								<option value="">구분</option>
								<option value="1" <%= setSelected(CSAT, "1") %>>Y</option>
								<option value="0" <%= setSelected(CSAT, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							면접점수
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<select name="Interviewer" id="Interviewer" class="form-control input-sm">
								<option value="">구분</option>								
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							실기점수
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<select name="Practical" id="Practical" class="form-control input-sm">
								<option value="">구분</option>								
							</select>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 제출서류 &nbsp;&nbsp;&nbsp;&nbsp;
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document1" id="document1" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							정시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document2" id="document2" class="form-control input-sm" value="1">
						</div>
						<!-- 면접은 서류 없음. 점수만 있으면 됨.
						<div class="col-md-1 col-xs-2 grid_sub_title">
							(공통)면접서류
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document3" id="document3" class="form-control input-sm" value="1">
						</div>
						-->
						<div class="col-md-1 col-xs-2 grid_sub_title">
							전문대이상
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document8" id="document8" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌1유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document4" id="document4" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌2유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document5" id="document5" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							기초수급자
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document6" id="document6" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							차상위계층
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document7" id="document7" class="form-control input-sm" value="1">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 목록
						</div>
						<div class="col-md-11 col-xs-7 grid_sub_title">
							<input type="hidden" name="DrawMsg" id="DrawMsg" class="form-control input-sm">
							<div name="DrawMsgUser" id="DrawMsgUser"></div>
						</div>
					</div>			

					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">신 규</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>

				</form>
			</div>
			<!-- 상세보기 끝 -->

			<!-- 기본 데이터 설정 -->
			<div id="BasicDataSetModal" style="width:100%; margin:5px; display:none;">
				<form name="BasicDataSetForm" id="BasicDataSetForm" method="post" action="/Process/BasicDataProc.asp">
				<input type="hidden" name="BasicDataSetprocess" value="RegApplicantBasicDataSet">
				<input type="hidden" name="BasicDataSetProcessType" id="BasicDataSetProcessType" value="Insert">
				<div class="ibox-content">		
					<!-- 버튼 번호 -->
					<div class="row show-grid " style="text-align:left;">
						<div class="col-md-1 grid_sub_title">
							버튼번호
						</div>
						<div class="col-md-2" style="text-align:left;">
							<% Call SubCodeSelectBox("BasicDataBtn", "버튼번호 선택", "", "", "", "BasicDataBtn") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title" style="text-align:left;">
							가산점
						</div>
						<div class="col-md-3 col-xs-7 grid_sub_title" style="text-align:left;">
							<% Call SubCodeSelectBox("ExtraPoint", "가산점 선택", "", "", "", "ExtraPoint") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title" style="text-align:left;">
							위반자
						</div>
						<div class="col-md-1 col-xs-7" style="text-align:left;">
							<input type="text" name="HighCodeTemp" value="<%=HighCodeTemp%>" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title" style="text-align:left;">
							생기부동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title" style="text-align:left;">
							<select name="StudentRecord" class="form-control input-sm">
								<option value="">-</option>
								<option value="1" <%= setSelected(StudentRecord, "1") %>>Y</option>
								<option value="0" <%= setSelected(StudentRecord, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title" style="text-align:left;">
							검정동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title" style="text-align:left;">
							<select name="Qualification" class="form-control input-sm">
								<option value="">-</option>
								<option value="1" <%= setSelected(Qualification, "1") %>>Y</option>
								<option value="0" <%= setSelected(Qualification, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title" style="text-align:left;">
							수능동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title" style="text-align:left;">
							<select name="CSAT" class="form-control input-sm">
								<option value="">-</option>
								<option value="1" <%= setSelected(CSAT, "1") %>>Y</option>
								<option value="0" <%= setSelected(CSAT, "0") %>>N</option>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							면접점수
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<select name="Interviewer" class="form-control input-sm">
								<option value="">구분</option>
								<% Dim conut 
								For conut = 0 To 100 %>
								<option value="<%= conut%>" <%= setSelected(Interviewer, conut) %>><%= conut%></option>
								<% Next %>
							</select>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							실기점수
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<select name="Practical" class="form-control input-sm">
								<option value="">구분</option>
								<% For conut = 0 To 100 %>
								<option value="<%= conut%>" <%= setSelected(Practical, conut) %>><%= conut%></option>
								<% Next %>
							</select>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 제출서류 &nbsp;&nbsp;&nbsp;&nbsp;
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document1" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							정시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document2" class="form-control input-sm" value="1">
						</div>
						<!-- 면접은 서류 없음. 점수만 있으면 됨.
						<div class="col-md-1 col-xs-2 grid_sub_title">
							(공통)면접서류
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document3" class="form-control input-sm" value="1">
						</div>
						-->
						<div class="col-md-1 col-xs-2 grid_sub_title">
							전문대이상
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document8" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌1유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document4" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌2유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document5" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							기초수급자
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document6" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							차상위계층
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document7" class="form-control input-sm" value="1">
						</div>
					</div>		

				</form>

				<br>
				<div class="row show-grid grid_sub_button" >					
					<div class="col-md-12" >
						<span class="btnBasic btnTypeSave" id="RegBasicDataSet" style="width:80px;">저장</span>
						<span class="btnBasic btnTypeClose SelfCloseDIV" style="width:80px;">취소</span>
					</div>
				</div>
				</div>
			</div>
			<!-- 기본 데이터 설정 -->

		</div>		
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->
<script type="text/javascript">
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// 전체선택 체크박스
    $("#checkall").click(function(){
        if($("#checkall").prop("checked")){
            $("input[name=chk]").prop("checked",true);
        }else{
            $("input[name=chk]").prop("checked",false);
        }
    });

	// tr선택 시 
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
		//1. 면접이나 실기비율이 있는 곳만 선택 가능하도록 설정
		//2. 면접/실기 비율(100점/100% 기준)만큼 점수를 넣을 수 있도록 생성
		//* 100점/100%가 아닐 경우, 비율에 해당하는 최대 점수 계산하여 Interviewer/PracticalNumber를 셋팅해주면 됨.
		var InterviewerNumber = Number($("#InterviewerRatio").val());
		var PracticalNumber = Number($("#PracticalRatio").val());
		InterviewerNumber = InterviewerNumber+1;
		PracticalNumber = PracticalNumber+1;

		// 셀렉트박스 초기화
		$("#Interviewer > option").not(":eq(0)").remove();
		$("#Practical > option").not(":eq(0)").remove();

		// 면접점수
		if ($("#InterviewerStr").val() === '-') {
			$("#Interviewer").prop("disabled", true).children("option").eq(0).attr("selected", true);
		} else {			
			$("#Interviewer").prop("disabled", false);
			
			// option 생성
			for (var conut=0; conut<InterviewerNumber; conut++) {
				$("#Interviewer").append("<option value='"+conut+"'>"+conut+"점</option>");
			}
			$("#Interviewer").val($("#InterviewerTemp").val());
		}

		// 실기점수
		if ($("#PracticalStr").val() === '-') {
			$("#Practical").prop("disabled", true).children("option").eq(0).attr("selected", true);
		} else {
			$("#Practical").prop("disabled", false);

			// option 생성
			for (var conut=0; conut<PracticalNumber; conut++) {
				$("#Practical").append("<option value='"+conut+"'>"+conut+"점</option>");
			}
			$("#Practical").val($("#PracticalTemp").val());
		}

		//동의해야 하는 항목만 선택 가능하도록 설정
		if ($("#StudentRecordStr").val() === '-') {
			$("#StudentRecord").prop("disabled", true);
		}else{
			$("#StudentRecord").prop("disabled", false);
		}
		if ($("#QualificationStr").val() === '-') {
			$("#Qualification").prop("disabled", true);
		}else{
			$("#Qualification").prop("disabled", false);
		}
		if ($("#CSATStr").val() === '-') {
			$("#CSAT").prop("disabled", true);
		}else{
			$("#CSAT").prop("disabled", false);
		}
		//서류체크를 해야하는 자격미달 체크박스만 선택 가능하도록 설정
		for (var i=1; i<25; i++) {
			if ($("#Check" + i).val() == "0" || $("#Check" + i).val() == "1" || $("#Check" + i).val() == "10"  ) {
				$("#document" + i ).prop("disabled", true);
			}else if($("#Check" + i).val() != "0") {
				$("#document" + i ).prop("disabled", false);
			}
		}	
		//서류체크를 해야하는 자격미달목록출력
		$("#DrawMsgUser").html($("#DrawMsg").val());
		//tr선택 시 체크박스 체크
		var $Checkbox = $(this).find("input[type='Checkbox']")
		if ($Checkbox.is(":checked")) {
			$Checkbox.prop("checked", false); 
		} else {
			$Checkbox.prop("checked", true); 
		}
	});

	// 체크박스 선택 시 체크 
	$(document).on("click", "input.CheckboxCheck", function(){
		var $Checkbox = $(this)
		if ($Checkbox.is(":checked")) {
			$Checkbox.prop("checked", false); 
		} else {
			$Checkbox.prop("checked", true); 
		}
	});	

	// 기본 데이터 설정(모달) 오픈
	$("#BasicDataSet").click(function() {
		$("#BasicDataSetProcessType").val("Update");
		$.openMadal($("#BasicDataSetModal"), "2");
	});

	// 기본 데이터 저장
	$("#RegBasicDataSet").click(function() {
		if (!$.chkInputValue($("select[name=BasicDataBtn]"),		"버튼번호를 선택해 주시기 바랍니다.")) { return; }
		
		if (confirm("기본데이터를 저장 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setBasicData(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/BasicDataProc.asp";
			$.Ajax4Form("#BasicDataSetForm", objOpt);
			$("#BasicDataSetForm").submit();
		}
	});

	// 기본 데이터 저장 결과
	$.setBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {
			alert("기본 데이터가 저장 되었습니다.");
				
		} else {
			alert("기본 데이터가 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 기본 데이터 넣기
	$.PutBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
			
		if ($objList.find("Result").text() == "Complete") {
			 
			 $("select[name=ExtraPoint]").val($objList.find("ExtraPoint").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=StudentRecord]").val($objList.find("StudentRecord").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Qualification]").val($objList.find("Qualification").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=CSAT]").val($objList.find("CSAT").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=Interviewer]").val($objList.find("Interviewer").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Practical]").val($objList.find("Practical").text()).prop("selected", true).trigger("chosen:updated");
				
			 for (var j=1; j<25; j++) {
				 if ($objList.find("document" + j).text() == "1") {
			 		 $("input[name=document" + j + "]").prop("checked", true).trigger("chosen:updated");
				 }
			 }
		} else {
			alert("기본 데이터 넣기 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 신규
	$("#btnNew").click(function () {
		if (confirm("입력되어 있던 내용이 초기화 됩니다.\n신규로 입력 하시겠습니까?")) {
			$.FormReset($("#InputForm"));
		}
	});

	// 저장
	$("#btnSave").click(function () { 
		if ($.setValidation($("#InputForm"))) {
			if (!$("#DrawMsgUser").html()){
				alert("선택된 지원자가 없습니다.");
				return;
			}

			if (confirm("입력하신 내용을 저장 하시겠습니까?")) {
				/////////////////////////////////////////////////////////////////////////////////////////
				//선택된 체크박스의 값(수험번호,년도)을 가져와서 input hidden에 넣어준 후 submit
				/////////////////////////////////////////////////////////////////////////////////////////
				var checkBoxArr = [];
				var checkBoxArr2 = [];
				$("input[name=chk]:checked").each(function(i){
					checkBoxArr.push($(this).val());
					checkBoxArr2.push($(this).attr("Myear"));
				});
				$("#StudentNumberHidden").val(checkBoxArr);
				$("#MyearHidden").val(checkBoxArr2);
				$.Ajax4FormSubmit($("#InputForm"), "입력하신 정보 저장이 완료되었습니다.");
			}
		}
	});

	/*
	// 삭제
	$("#btnDelete").click(function () {
		if (!$.chkInputValue($("#InputForm input[name='IDX']"),	"삭제할 항목을 선택해 주세요.")) { return; }

		if (confirm("선택된 항목을 삭제 하시겠습니까?")) {
			$("#InputForm input[name='ProcessType']").val("Delete");
			$.Ajax4FormSubmit($("#InputForm"), "선택된 항목이 삭제되었습니다.");
		}
	});
	*/
});
//기본 데이터 가져오기
function BasicDataBtn(num) {
	if (confirm(num + "번 기본데이터를 넣으시겠습니까?")) {
		$("#BasicDataBtnNum").val(num);

		var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.PutBasicData(datas)','complete':'','clear':'','reset':''};
		objOpt["url"] = "/Process/BasicDataSelect.asp";
		$.Ajax4Form("#BasicDataBtnFrom", objOpt);
		$("#BasicDataBtnFrom").submit();
	}
}
</script>