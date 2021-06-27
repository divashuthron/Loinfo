<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 100
Dim LeftMenuCode : LeftMenuCode = ""
Dim LeftMenuName : LeftMenuName = "Home / 계산 테스트"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "계산 테스트"
%>
<!-- #include virtual="/Common/Header.asp" -->
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			<div class="pad_t10"></div>

			<div class="ibox-title">				

<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, AryHash2, arrParams2, intNum, AryHash3, arrParams3

Dim CompleteUnit, Grade									'// 이수단위, 등급
Dim ConvertGrade, ConvertCompleteGrade					'// 환산등급, 환산이수*등급
Dim GEDScore											'// 과목별 점수
Dim GetScore			: GetScore = 0					'// 취득점수

Dim YearType											'// 학년학기구분
Dim CompleteUnit_1_1	: CompleteUnit_1_1 = 0			'// 1학년 1학기 이수단위
Dim CompleteGrade_1_2	: CompleteGrade_1_2 = 0			'// 1학년 2학기 이수*등급
Dim CompleteUnit_2_1	: CompleteUnit_2_1 = 0			'// 2학년 1학기 이수단위
Dim CompleteGrade_2_2	: CompleteGrade_2_2 = 0			'// 2학년 2학기 이수*등급
Dim CompleteUnit_3_1	: CompleteUnit_3_1 = 0			'// 3학년 1학기 이수단위
Dim CompleteGrade_3_2	: CompleteGrade_3_2 = 0			'// 3학년 2학기 이수*등급
Dim GradeCalculation_1_1: GradeCalculation_1_1 = 0		'// 1학년 1학기 등급
Dim GradeCalculation_1_2: GradeCalculation_1_2 = 0		'// 1학년 2학기 등급
Dim GradeCalculation_2_1: GradeCalculation_2_1 = 0		'// 2학년 1학기 등급
Dim GradeCalculation_2_2: GradeCalculation_2_2 = 0		'// 2학년 2학기 등급
Dim GradeCalculation_3_1: GradeCalculation_3_1 = 0		'// 3학년 1학기 등급
Dim GradeCalculation_3_2: GradeCalculation_3_2 = 0		'// 3학년 2학기 등급

Dim CompleteUnitValueCheck 								'// 값 모두 입력 했는지 체크
Dim ResultMSG											'// 계산 불가 시 메시지
Dim CompleteUnitCnt		: CompleteUnitCnt = 0			'// 학생별 과목수

Dim t, Max, One, OneMax, Two, TwoMax, OneTwoAug			'// 수능 점수 환산 변수
ReDim one(3)											'// 국어,수학,영어 배열
ReDim two(5)											'// 탐구, 제2외국어 배열
Max = 0

Dim UniversityName, AugScore, PerfectScore, Credit		'// 대학명, 평균학점, 평균학점만점, 대학이수학점
Dim ConvertGradetot, Check, DivistionCheck				'// 검정고시 등급 합, Grade 숫자와 문자 비교용 Check, 수시정시 구분
Dim FormulaNum, Formula1, Formula2, Formula3			'// 생기부 환산공식 번호, 수시공식, 정시공식, 기타공식
Dim StudentDivision, ScoreDim, Score, OriginalScore		'// 학생유형, 수식, 검정고시 점수, 원점수
Dim AveScore, Deviation, Ranking, EnrollmentCount		'// 평균점수, 표준편차, 석차, 재적수

Dim YearTypeTemp, CompleteUnitTemp, GradeTemp			'// 위에 변수를 function으로 보낼 배열변수
Dim OriginalScoreTemp, AveScoreTemp, DeviationTemp		
Dim RankingTemp, EnrollmentCountTemp		
Dim AryInterviewerScore, AryStudentRecordAverage, AryCreditSum
Dim AryChoiceSemester, AryMinor, AryUniversityCredit
Dim AryKorLanScore, AryEnglishScore, AryMathematicsScore, AryStudentNumber
ReDim YearTypeTemp(30)
ReDim CompleteUnitTemp(30)
ReDim GradeTemp(30)
ReDim OriginalScoreTemp(30)
ReDim AveScoreTemp(30)
ReDim DeviationTemp(30)
ReDim RankingTemp(30)
ReDim EnrollmentCountTemp(30)
ReDim Score(30)
ReDim AryStudentNumber(30)
ReDim AryInterviewerScore(30)
ReDim AryStudentRecordAverage(30)
ReDim AryCreditSum(30)
ReDim AryChoiceSemester(30)
ReDim AryMinor(30)
ReDim AryUniversityCredit(30)
ReDim AryKorLanScore(30)
ReDim AryEnglishScore(30)
ReDim AryMathematicsScore(30)

Dim BasicMYear, BasicDivision0, BasicSubject, BasicDivision1						'// 사정 구분
'테스트용 사정구분
BasicMYear = 2019
BasicDivision0 = "X03021"

Dim MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3, SubjectCode		'// 입력 기본데이터 (년도, 수험번호, 모집시기, 학과, 구분1, 구분2, 구분3)
Dim StudentRecordScore, StudentRecordAverage, CreditSum, ChoiceSemester							'// 생기부 기본데이터 (생기부성적, 교과 성적평균, 선택한 학기 이수단위 합, 반영학기)
Dim SubjectCnt, InUpDivision, InUpDivisionCheck													'// 과목수, InsertUpdate구분, InsertUpdate구분체크
Dim ExtraPointScore, StudentRecordScoreTemp, InterviewerScoreTemp								'// 가산점, 생기부통합용, 면접통합용
Dim QualificationScoreTemp, UniversityScoreTemp, CSATScoreTemp									'// 검정고시통합용, 대학통합용, 수능통합용
Dim StudentRecordAverageTemp, CreditSumTemp, ChoiceSemesterTemp, MinorTemp						'// 교과성적평균동석차용, 이수단위합동석차용, 반영학기동석차용
Dim KorLanScoreTemp, EnglishScoreTemp, MathematicsScoreTemp, UniversityCreditTemp				'// 국어성적동석차용, 영어성적동석차용, 수학성적동석차용, 대학졸업이수학점동석차용\
Dim totScore, DrawStanding, DrawScore, y, Minor, Minor2, Quorum, QuorumFix						'// 통합점수, 동석차순위, 동점인 점수, 반복수, 나이(연소자), 성별구분, 모집인원, 입학정원
Dim UnqualifiedStandard1, UnqualifiedStandardCheck, totScoreCount								'// 자격미달기준, 자격미달 체크, 동석차명수

Dim StandardNum, StandardNum2, StandardNumTemp, StudentNumberTemp								'// 동석차 기준 변수
Dim DrawRanking, DrawRankingNum, DrawRankingNum2, DrawRankingTemp
ReDim DrawRanking(30)

Dim OldDrawRanking
Dim TxtStudentNumber, TxtInterviewerScore, TxtStudentRecordAverage, TxtCreditSum, TxtChoiceSemester, TxtMinor, TxtDrawRanking
Dim ArrTxtStudentNumber, ArrTxtxtInterviewerScore, ArrTxtxtStudentRecordAverage, ArrTxtxtCreditSum, ArrTxtxtChoiceSemester, ArrTxtxtMinor, ArrTxtDrawRanking

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & " Select * , dbo.getSubCodeName('PerfectScore', PerfectScore) AS PerfectScoreName "
SQL = SQL & vbCrLf & "			, dbo.getSubCodeTemp1('ExtraPoint', ExtraPoint) AS ExtraPointScore "
SQL = SQL & vbCrLf & " from ApplicationTable "
SQL = SQL & vbCrLf & " WHERE 1 = 1  "
'SQL = SQL & vbCrLf & " AND MYear = " & BasicMYear
'SQL = SQL & vbCrLf & " AND Division0 = '" & BasicDivision0 & "'"
SQL = SQL & vbCrLf & " AND Reslut = 'Y(완료)' "

'Call objDB.sbSetArray("@MYear", adVarchar, adParamInput, 4, BasicMYear)
'Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 60, BasicDivision0)

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'// ==============================================================================================================
'// 학생 유형 구분하기
'// ==============================================================================================================
'// "1"		'수시/ 전문대이상 졸업자 전형(출신대학 성적 100%)
'// "2"		'수시/ 일반전형 / 검정고시
'// "3"		'수시/ 일반전형 / 2008 ~ 현재 졸업예정(학생부 40% + 면접 60%)
'// "4"		'수시/ 일반전형 / 1998 ~ 2007년 졸업자(학생부 40% + 면접 60%)
'// "5"		'수시/ 일반전형 / 1997년 이전 졸업자(학생부 40% + 면접 60%)
'// "6"		'수시/ 일반전형,전문대이상 외 / 검정고시
'// "7"		'수시/ 일반전형,전문대이상 외 / 2008 ~ 현재 졸업예정(학생부 100%)
'// "8"		'수시/ 일반전형,전문대이상 외 / 1998 ~ 2007년 졸업자(학생부 100%)
'// "9"		'수시/ 일반전형,전문대이상 외 / 1997년 이전 졸업자(학생부 100%)
'// "10"	'정시/ 면접전형 / 기술드론부사관과, 기술행정부사관과, 조리부사관과, 항공서비스과 (수능 40% + 면접 60%)
'// "11"	'정시/ 비면접전형 / 검정고시
'// "12"	'정시/ 비면접전형 / 2008 ~ 현재 졸업예정(수능 70% + 학생부 30%)
'// "13"	'정시/ 비면접전형 / 1998 ~ 2007년 졸업자(수능 70% + 학생부 30%)
'// "14"	'정시/ 비면접전형 / 1997년 이전 졸업자(수능 70% + 학생부 30%)
'// ==============================================================================================================

If Not IsNull(AryHash) then
	for i = 0 to ubound(AryHash,1)
		'수시일 때
		If AryHash(i).Item("Division0") = "X03021" Or AryHash(i).Item("Division0") = "X03022" Then
			'전문대이상 졸업자 전형일 때(출신대학 성적 100%)
			If AryHash(i).Item("Division1") = "X05130" Then
				StudentDivision = "1"				
			'일반전형일 때(학생부 40% + 면접 60%)
			ElseIf AryHash(i).Item("Division1") = "X05010" Then
				'수시/ 일반졍형/ 검정고시
				If AryHash(i).Item("Qualification") = "1" Then
					StudentDivision = "2"
				Else 
					'수시/ 일반졍형/ 2008 ~ 현재 졸업예정(학생부 100%)
					If AryHash(i).Item("HighGraduationYear") >= "2008" Then
						StudentDivision = "3"
					'수시/ 일반졍형/ 1998 ~ 2007년 졸업자(학생부 100%)
					ElseIf AryHash(i).Item("HighGraduationYear") >= "1998" Then
						StudentDivision = "4"
					'수시/ 일반졍형/ 1997년 이전 졸업자(학생부 100%)
					ElseIf AryHash(i).Item("HighGraduationYear") <= "1997" Then
						StudentDivision = "5"
					End If
				End If
			Else
				'수시/ 검정고시
				If AryHash(i).Item("Qualification") = "1" Then
					StudentDivision = "6"
				Else 
					'수시/ 2008 ~ 현재 졸업예정(학생부 100%)
					If AryHash(i).Item("HighGraduationYear") >= "2008" Then
						StudentDivision = "7"
					'수시/ 1998 ~ 2007년 졸업자(학생부 100%)
					ElseIf AryHash(i).Item("HighGraduationYear") >= "1998" Then
						StudentDivision = "8"
					'수시/ 1997년 이전 졸업자(학생부 100%)
					ElseIf AryHash(i).Item("HighGraduationYear") <= "1997" Then
						StudentDivision = "9"
					End If
				End If
			End If
		End If
		'정시/추가일 때
		If AryHash(i).Item("Division0") = "X03031" Or AryHash(i).Item("Division0") = "X03050" Then
			'면접 기술드론부사관과, 기술행정부사관과, 조리부사관과, 항공서비스과 (수능 40% + 면접 60%)
			If AryHash(i).Item("Subject") = "170" Or AryHash(i).Item("Subject") = "220" Or AryHash(i).Item("Subject") = "310" Or AryHash(i).Item("Subject") = "040" Then
				StudentDivision = "10"
			'비면접 (수능 70% + 학생부 30%)
			Else
				'정시/ 검정고시
				If AryHash(i).Item("Qualification") = "1" Then
					StudentDivision = "11"
				Else 
					'정시/ 2008 ~ 현재 졸업예정(수능 70% + 학생부 30%)
					If AryHash(i).Item("HighGraduationYear") >= "2008" Then
						StudentDivision = "12"
					'정시/ 1998 ~ 2007년 졸업자(수능 70% + 학생부 30%)
					ElseIf AryHash(i).Item("HighGraduationYear") >= "1998" Then
						StudentDivision = "13"
					'정시/ 1997년 이전 졸업자(수능 70% + 학생부 30%)
					ElseIf AryHash(i).Item("HighGraduationYear") <= "1997" Then
						StudentDivision = "14"
					End If
				End If
			End If
		End If
		'위탁과정과 전공심화는 차후 추가해야 함

		'테스트용 학생유형구분
		StudentDivision = 7

		'// =================================================================
		'// 생기부, 검정고시 공식 가져오기
		'// =================================================================
		Select Case StudentDivision
			'2. 수시/ 일반전형 / 검정고시	
			'6. 수시/ 일반전형, 전문대이상 외 / 검정고시
			'11. 정시/ 비면접전형 / 검정고시
			Case "2", "6", "11"
				FormulaNum = 4 '// 검정고시 출신자

			'3. 수시/ 일반전형 / 2008 ~ 현재 졸업예정(학생부 40% + 면접 60%)
			'7. 수시/ 일반전형, 전문대이상 외 / 2008 ~ 현재 졸업예정(학생부 100%)
			'12. 정시/ 비면접전형 / 2008 ~ 현재 졸업예정(수능 70% + 학생부 30%)
			Case "3", "7", "12"					
				FormulaNum = 1 '// 2008~현재 졸업(예정)자
			
			'4. 수시/ 일반전형 / 1998 ~ 2007년 졸업자(학생부 40% + 면접 60%)
			'8. 수시/ 일반전형, 전문대이상 외 / 1998 ~ 2007년 졸업자(학생부 100%)
			'13. 정시/ 비면접전형 / 1998 ~ 2007년 졸업자(수능 70% + 학생부 30%)
			Case "4", "8", "13"					
				FormulaNum = 2 '// 1998년~2007년 졸업자

			'5. 수시/ 일반전형 / 1997년 이전 졸업자(학생부 40% + 면접 60%)
			'9. 수시/ 일반전형, 전문대이상 외 / 1997년 이전 졸업자(학생부 100%)
			'14. 정시/ 비면접전형 / 1997년 이전 졸업자(수능 70% + 학생부 30%)
			Case "5", "9", "14"				
				FormulaNum = 3 '// 1997년 이전 졸업자

			'1. 전문대이상
			Case Else
				FormulaNum = 5
		End Select

		SQL = ""
		SQL = SQL & vbCrLf & " Select * "
		SQL = SQL & vbCrLf & " from StudentRecord  "
		SQL = SQL & vbCrLf & " WHERE 1 = 1  "
		SQL = SQL & vbCrLf & " And FormulaNum = " & FormulaNum

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

		If Not(isnull(AryHash2)) Then
			Formula1 = AryHash2(0).Item("Formula1") '수시공식
			Formula2 = AryHash2(0).Item("Formula2") '정시공식
			Formula3 = AryHash2(0).Item("Formula3") '기타공식
		End If

		'Response.write Formula1 & " / "
		'Response.write Formula2 & " / "
		'Response.write Formula3 & " / "

		'// =================================================================
		'// 반영학기 구하기
		'// =================================================================

		Select case AryHash(i).Item("Semester")
			Case "1" 
				Semester1 = "1"
				Semester2 = "1"
			Case "2" 
				Semester1 = "1"
				Semester2 = "2"
			Case "3" 
				Semester1 = "2"
				Semester2 = "1"
			Case "4" 
				Semester1 = "2"
				Semester2 = "2"
			Case "5" 
				Semester1 = "3"
				Semester2 = "1"
		End Select

		'// =================================================================
		'// 모집시기 구하기
		'// =================================================================

		Select case AryHash(i).Item("Division0")
			Case "X03021", "X03022" '수시1, 수시2'
				DivistionCheck = "1"
			Case "X03031", "X03050" '정시, 추가
				DivistionCheck = "2"
		End Select

		'// =================================================================
		'// 기본데이터 변수 저장
		'// =================================================================

		MYear				=	AryHash(i).Item("Myear")
		StudentNumber		=	AryHash(i).Item("StudentNumber")
		Division0			=	AryHash(i).Item("Division0")
		Subject				=	AryHash(i).Item("Subject")
		Division1			=	AryHash(i).Item("Division1")
		Division2			=	AryHash(i).Item("Division2")
		Division3			=	AryHash(i).Item("Division3")
		ChoiceSemester		=	AryHash(i).Item("Semester")
		ExtraPointScore		=	AryHash(i).Item("ExtraPointScore")
		Minor				=	AryHash(i).Item("Citizen1")
		Minor2				=	AryHash(i).Item("Citizen2")
		InUpDivisionCheck	=	False
		
		'// 나이 구하기
		'// 주민1 앞 2자리
		'// 주민2 뒷 1자리
		'// 주민2가 1,2면 1900년대 3,4면 2000년대
		Minor = Left(Minor,2)
		Minor2 = Left(Minor2,1)
		If Minor2 = 1 Or Minor2 = 2 Then
			Minor = Year(Date()) - (1900 + Minor) + 1
		ElseIf Minor2 = 3 Or Minor2 = 4 Then
			Minor = Year(Date()) - (2000 + Minor) + 1
		End If

		'Response.write "나이 : " & Minor & "<br>"

		'Response.write StudentNumber & "<br>"
		'Response.write Division0 & "<br>"
		'Response.write Subject & "<br>"
		'Response.write Division1 & "<br>"
		'Response.write Division2 & "<br>"
		'Response.write Division3 & "<br>"
		'Response.write ChoiceSemester & "<br>"

		'StudentRecordScore	'생기부성적
		'StudentRecordAverage	'교과 성적평균
		'CreditSum	'선택한 학기 이수단위 합
		'ChoiceSemester '반영학기

		'// =================================================================
		'// ChangeScoreTable / Insert & Update 최초구분
		'// =================================================================
		SQL = ""
		SQL = SQL & vbCrLf & " select *  "
		SQL = SQL & vbCrLf & " from ChangeScoreTable  "
		SQL = SQL & vbCrLf & " Where StudentNumber =  " & StudentNumber

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

		If Not(isnull(AryHash2)) Then
			InUpDivision = "Update"
		Else
			InUpDivision = "Insert"
		End If

		'// =================================================================
		'// 성적 가져오기
		'// =================================================================
		'// CSAT - 수능
		'// interviewmng.dbo.EvaluationRecord - 면접서버 만들어지면 연결해야 함
		'// ApplicationTable - 전문대
		'// 212 - 검정고시
		'// 213 - 생기부		
		'// =================================================================
		
		'/////////// 수능 ///////////
		SQL = ""
		SQL = SQL & vbCrLf & " select EXAM_NUMB, LGFD_SDSC, MTFD_SDSC, FLFD_GRAD, RSFD_SCR1, RSFD_SCR2, RSFD_SCR3, RSFD_SCR4, SCFL_SDSC  "
		SQL = SQL & vbCrLf & " from IPSICSAT  "
		SQL = SQL & vbCrLf & " Where EXAM_NUMB =  " & StudentNumber

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)
	
		If Not(isnull(AryHash2)) Then
			EXAM_NUMB	=	AryHash2(0).Item("EXAM_NUMB")
			LGFD_SDSC	=	AryHash2(0).Item("LGFD_SDSC") '언어영역표준점수
			MTFD_SDSC	=	AryHash2(0).Item("MTFD_SDSC") '수리영역표준점수
			FLFD_GRAD	=	AryHash2(0).Item("FLFD_GRAD") '외국어영역등급
			RSFD_SCR1	=	AryHash2(0).Item("RSFD_SCR1") '탐구영역표준점수1
			RSFD_SCR2	=	AryHash2(0).Item("RSFD_SCR2") '탐구영역표준점수2
			RSFD_SCR3	=	AryHash2(0).Item("RSFD_SCR3") '탐구영역표준점수3
			RSFD_SCR4	=	AryHash2(0).Item("RSFD_SCR4") '탐구영역표준점수4
			SCFL_SDSC	=	AryHash2(0).Item("SCFL_SDSC") '제2외국어표준점수
		Else
			EXAM_NUMB	=	null
			LGFD_SDSC	=	null
			MTFD_SDSC	=	null
			FLFD_GRAD	=	null
			RSFD_SCR1	=	null
			RSFD_SCR2	=	null
			RSFD_SCR3	=	null
			RSFD_SCR4	=	null
			SCFL_SDSC	=	null
		End IF

		'Response.write "<br>"
		'Response.write EXAM_NUMB & " / "
		'Response.write LGFD_SDSC & " / "
		'Response.write MTFD_SDSC & " / "
		'Response.write FLFD_GRAD & " / "
		'Response.write RSFD_SCR1 & " / "
		'Response.write RSFD_SCR2 & " / "
		'Response.write RSFD_SCR3 & " / "
		'Response.write RSFD_SCR4 & " / "
		'Response.write SCFL_SDSC & " / "
		'Response.write "<Br>"

		'/////////// 면접 ///////////
		SQL = ""
		SQL = SQL & vbCrLf & " select StudentNumber  "
		SQL = SQL & vbCrLf & " , ItemPoint_01, ItemPoint_02, ItemPoint_03, ItemPoint_04, ItemPoint_05  "
		SQL = SQL & vbCrLf & " , ItemPoint_06, ItemPoint_07, ItemPoint_08, ItemPoint_09, ItemPoint_10  "
		SQL = SQL & vbCrLf & " from interviewmng.dbo.EvaluationRecord  "
		SQL = SQL & vbCrLf & " Where StudentNumber =  " & StudentNumber

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

		If Not(isnull(AryHash2)) Then
			ItemPoint_01		=	AryHash2(0).Item("ItemPoint_01")	'// 1번항목 점수
			ItemPoint_02		=	AryHash2(0).Item("ItemPoint_02")	'// 2번항목 점수
			ItemPoint_03		=	AryHash2(0).Item("ItemPoint_03")	'// 3번항목 점수
			ItemPoint_04		=	AryHash2(0).Item("ItemPoint_04")	'// 4번항목 점수
			ItemPoint_05		=	AryHash2(0).Item("ItemPoint_05")	'// 5번항목 점수
			ItemPoint_06		=	AryHash2(0).Item("ItemPoint_06")	'// 6번항목 점수
			ItemPoint_07		=	AryHash2(0).Item("ItemPoint_07")	'// 7번항목 점수
			ItemPoint_08		=	AryHash2(0).Item("ItemPoint_08")	'// 8번항목 점수
			ItemPoint_09		=	AryHash2(0).Item("ItemPoint_09")	'// 9번항목 점수
			ItemPoint_10		=	AryHash2(0).Item("ItemPoint_10")	'// 10번항목 점수
		Else
			ItemPoint_01		=	null
			ItemPoint_02		=	null
			ItemPoint_03		=	null
			ItemPoint_04		=	null
			ItemPoint_05		=	null
			ItemPoint_06		=	null
			ItemPoint_07		=	null
			ItemPoint_08		=	null
			ItemPoint_09		=	null
			ItemPoint_10		=	null
		End if

		'Response.write "<br>"
		'Response.write ItemPoint_01 & " / "
		'Response.write ItemPoint_02 & " / "
		'Response.write ItemPoint_03 & " / "
		'Response.write ItemPoint_04 & " / "
		'Response.write ItemPoint_05 & " / "
		'Response.write ItemPoint_06 & " / "
		'Response.write ItemPoint_07 & " / "
		'Response.write ItemPoint_08 & " / "
		'Response.write ItemPoint_09 & " / "
		'Response.write ItemPoint_10 & " / "
		'Response.write "<Br>"

		'/////////// 전문대이상 ///////////
		UniversityName	=	AryHash(i).Item("UniversityName")
		AugScore		=	AryHash(i).Item("AugScore")
		PerfectScore	=	AryHash(i).Item("PerfectScoreName")
		If Not(isnull(PerfectScore)) Then
			PerfectScore	=	Cdbl(left(PerfectScore, 3))
		End If
		Credit			=	AryHash(i).Item("Credit")

		'Response.write "<br>"
		'Response.write UniversityName & " / "
		'Response.write AugScore & " / "
		'Response.write PerfectScore & " / "
		'Response.write Credit & " / "
		'Response.write "<Br>"

		'/////////// 검정고시 ///////////
		SQL = ""
		SQL = SQL & vbCrLf & " Select SCHL_YEAR, COLL_FLAG, EXAM_NUMB, WORK_SEQN, SBJT_NAME, SBJT_SCOR, SBJT_GRAD   "
		SQL = SQL & vbCrLf & " from IPSI212   "
		SQL = SQL & vbCrLf & " WHERE 1 = 1    "
		SQL = SQL & vbCrLf & " And Exam_Numb =  " & StudentNumber

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

		If Not(isnull(AryHash2)) Then
			for j = 0 to ubound(AryHash2,1)
				Score(j)	=	AryHash2(j).Item("SBJT_SCOR")	'// 점수
			Next
			ScoreCnt		=	ubound(AryHash2,1)				'// 과목수 카운터
		Else
			for j = 0 to 20
				Score(j)	=	null
			Next
			ScoreCnt		=	null
		End If

		'Response.write "<br>"
		'Response.write Score(0) & " / "
		'Response.write ScoreCnt & " / "
		'Response.write "<Br>"

		'/////////// 생기부 ///////////
		SQL = ""
		SQL = SQL & vbCrLf & " Select CORS_NAME, ADPT_AVRG, RANK_GRAD, CMPT_UNIT, STDD_DEVI, ORGL_SCOR, ADPT_INDX, STHS_RANK, ENRL_CONT  "
		SQL = SQL & vbCrLf & " from IPSI213  "
		SQL = SQL & vbCrLf & " WHERE 1 = 1  "
		SQL = SQL & vbCrLf & " And Exam_Numb =  " & StudentNumber
		SQL = SQL & vbCrLf & " And STDT_YEAR =  " & Semester1
		SQL = SQL & vbCrLf & " And SCHL_TERM =  " & Semester2
		SQL = SQL & vbCrLf & " AND ((ISNULL(RANK_GRAD,'0') != '이수'   AND ISNULL(ADPT_INDX,'0') != '이수'  ) "
		SQL = SQL & vbCrLf & " AND  (ISNULL(RANK_GRAD,'0') != '미이수' AND ISNULL(ADPT_INDX,'0') != '미이수') "
		SQL = SQL & vbCrLf & " AND  (ISNULL(RANK_GRAD,'0') != '우수'   AND ISNULL(ADPT_INDX,'0') != '우수'  ) "
		SQL = SQL & vbCrLf & " AND  (ISNULL(RANK_GRAD,'0') != '보통'   AND ISNULL(ADPT_INDX,'0') != '보통'  ) "
		SQL = SQL & vbCrLf & " AND  (ISNULL(RANK_GRAD,'0') != '미흡'   AND ISNULL(ADPT_INDX,'0') != '미흡'  ) "
		SQL = SQL & vbCrLf & " AND  (ISNULL(RANK_GRAD,'0') != 'P'      AND ISNULL(ADPT_INDX,'0') != 'P'     )) "
		SQL = SQL & vbCrLf & " AND CMPT_UNIT is not null "
		SQL = SQL & vbCrLf & " AND (RANK_GRAD is not null Or (ORGL_SCOR is not null and ADPT_AVRG is not null and STDD_DEVI is not null)) "

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)
		
		If Not(isnull(AryHash2)) Then	
			for j = 0 to ubound(AryHash2,1)
				YearType			=	Semester1 & "_" & Semester2		'// 학년학기
				CompleteUnit		=	AryHash2(j).Item("CMPT_UNIT")	'// 이수단위
				Grade				=	AryHash2(j).Item("RANK_GRAD")	'// 석차등급
				OriginalScore		=	AryHash2(j).Item("ORGL_SCOR")	'// 원점수
				AveScore			=	AryHash2(j).Item("ADPT_AVRG")	'// 평균점수
				Deviation			=	AryHash2(j).Item("STDD_DEVI")	'// 표준편차
				Ranking				=	AryHash2(j).Item("STHS_RANK")	'// 석차 
				EnrollmentCount		=	AryHash2(j).Item("ENRL_CONT")	'// 재적수
					
				'Grade 숫자와 문자 비교 중 어떤 것이 올 지 모르므로 ElseIf가 아닌 Check사용하여 분기
				Check = True
					
				'1.등급이 1~9등급
				'그냥 등급을 씀
				If Check And isNumeric(Grade) Then				
					If Not(isnull(CompleteUnit)) And Grade > 0 And Grade <= 9 Then
						Check = false
				'		Response.write "1. 등급이 1~9등급 : " & Grade 
					End If	
				End If
				'2.등급이 A~E등급
				'원점수, 평균점수, 표준편차
				If Check And Not(isnull(CompleteUnit)) And Trim(Grade) >= "A" And Trim(Grade) <= "E" Then
					If Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
						Check = False						
						ScoreDim = "Dim executeTemp : executeTemp = (" & OriginalScore & "-" & AveScore & ")/" & Deviation 
						execute(ScoreDim) 
						executeTemp = FormatNumber(executeTemp - 0.005,2)		
						Dim valueTemp
						valueTemp = NORMDIST(executeTemp, 0, 1, 1)
						valueTemp = FormatNumber(valueTemp - 0.000005, 5)
						valueTemp = (1 - valueTemp) * 100
						Grade = PercentageGrade(valueTemp)
				'		Response.write "2. 등급이 A~E등급 : " & Grade
					End If
				End If
				'3.석차가 1 이상이고, 재적수가 1 이상이면
				'석차, 재적수
				If Check And Not(isnull(CompleteUnit)) And Ranking > 0 And EnrollmentCount > 0 Then
					Check = False
					ScoreDim = "Dim executeTemp : executeTemp = " & Ranking & "* 100 /" & EnrollmentCount
					execute(ScoreDim) 
					Grade = PercentageGrade(executeTemp)
				'	Response.write "3. 석차가 1 이상이고, 재적수가 1 이상 : " & Grade
				End If
				'4.석차가 0이거나 null이고, 재적수가 0이거나 null이고 원점수가 null이 아니고, 표준편차 null이 아니면
				'원점수, 평균점수, 표준편차
				If Check And Not(isnull(CompleteUnit)) And (Ranking = 0 Or isnull(Ranking)) And (EnrollmentCount = 0 Or isnull(EnrollmentCount)) And Not(isnull(OriginalScore)) And Not(isnull(Deviation)) Then
					If Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
						Check = False
						ScoreDim = "Dim executeTemp : executeTemp = (" & OriginalScore & "-" & AveScore & ")/" & Deviation 
						execute(ScoreDim) 
						executeTemp = FormatNumber(executeTemp - 0.005,2)
						valueTemp = NORMDIST(executeTemp, 0, 1, 1)
						valueTemp = FormatNumber(valueTemp - 0.000005, 5)
						valueTemp = (1 - valueTemp) * 100
						Grade = PercentageGrade(valueTemp)
				'		Response.write "4. 석차 null, 재적수 null 원점수 있고, 표준점차 있음 : " & Grade
					End If
				End If
				'5.석차가 0이거나 null이고, 재적수가 1이상이고, 등급이 0보다 작거나, 9보다 크다.
				'원점수, 평균점수, 표준편차
				If Check And isNumeric(Grade) Then
					If Not(isnull(CompleteUnit)) And (Ranking = 0 Or isnull(Ranking)) And Not(isnull(EnrollmentCount)) And EnrollmentCount > 0 And (Grade < 0 Or Grade > 9) Then
						If Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
							Check = False
							ScoreDim = "Dim executeTemp : executeTemp = (" & OriginalScore & "-" & AveScore & ")/" & Deviation 
							execute(ScoreDim) 
							executeTemp = FormatNumber(executeTemp - 0.005,2)
							valueTemp = NORMDIST(executeTemp, 0, 1, 1)
							valueTemp = FormatNumber(valueTemp - 0.000005, 5)
							valueTemp = (1 - valueTemp) * 100
							Grade = PercentageGrade(valueTemp)
				'			Response.write "5. 석차 null, 재적수 1이상 등급이 0보다 작고 9보다 크다 : " & Grade 
						End If
					End If
				End If
				'6.등급이 0이거나 null이고, 원점수가 null이 아니고, 평균점수가 null이 아니고, 표준편차도 null이 아니면				
				'원점수, 평균점수, 표준편차	
				If Check And (isNumeric(Grade) Or isnull(Grade)) Then
					If Not(isnull(CompleteUnit)) And (Grade = 0 Or isnull(Grade)) And Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
						If Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
							Check = False
							ScoreDim = "Dim executeTemp : executeTemp = (" & OriginalScore & "-" & AveScore & ")/" & Deviation 
							execute(ScoreDim) 
							executeTemp = FormatNumber(executeTemp - 0.005,2)
							valueTemp = NORMDIST(executeTemp, 0, 1, 1)
							valueTemp = FormatNumber(valueTemp - 0.000005, 5)
							valueTemp = (1 - valueTemp) * 100
							Grade = PercentageGrade(valueTemp)
				'			Response.write "6. 등급 null, 원점수,평균점수,표준편차 있음 : " & Grade
						End If
					End If
				End If
				'7. 등급이 .이면
				'원점수, 평균점수, 표준편차
				If Check And Not(isnull(CompleteUnit)) And Grade = "." Then
					If Not(isnull(OriginalScore)) And Not(isnull(AveScore)) And Not(isnull(Deviation)) Then
						Check = False
						ScoreDim = "Dim executeTemp : executeTemp = (" & OriginalScore & "-" & AveScore & ")/" & Deviation 
						execute(ScoreDim) 
						executeTemp = FormatNumber(executeTemp - 0.005,2)
						valueTemp = NORMDIST(executeTemp, 0, 1, 1)
						valueTemp = FormatNumber(valueTemp - 0.000005, 5)
						valueTemp = (1 - valueTemp) * 100
						Grade = PercentageGrade(valueTemp)
				'		Response.write "7. 등급이 .이면 : " & Grade
					End If
				End If
				'8. 그 외는 모두 9등급
				If Check Then
					Check = False
					Grade = 9
				'	Response.write "8. 그 외는 모두 9등급 : " & Grade
				End If

				YearTypeTemp(j)			=	YearType					'// 학년학기
				CompleteUnitTemp(j)		=	CompleteUnit				'// 이수단위
				GradeTemp(j)			=	Grade						'// 석차등급
				OriginalScoreTemp(j)	=	OriginalScore				'//	원점수
			Next
			CompleteUnitCnt				=	ubound(AryHash2,1)			'// 학생 별 과목수
		Else
			for j = 0 to 20
				YearTypeTemp(j)			=	null
				CompleteUnitTemp(j)		=	null
				GradeTemp(j)			=	null
				OriginalScoreTemp(j)	=	null
			Next
			CompleteUnitCn	t			=	null
		End If
		
		'Response.write "<br>"
		'Response.write YearTypeTemp(0) & " / "
		'Response.write CompleteUnitTemp(0) & " / "
		'Response.write GradeTemp(0) & " / "
		'Response.write CompleteUnitCn & " / "
		'Response.write "<Br>"

		'// =================================================================
		'// 학생 유형별 환산하기
		'// =================================================================
		'// GradeCalculation_C() '// 생기부    
		'// GradeCalculation_D() '// 면접점수 가져오기   
		'// GradeCalculation_E() '// 검정고시 출신자
		'// GradeCalculation_F() '// 전문대학이상 졸업자
		'// GradeCalculation_G() '// 수능점수 환산
		'// =================================================================
		
		Select Case StudentDivision
			'1. 수시/ 전문대이상 졸업자 전형(출신대학 성적 100%)
			Case "1"	
				Call GradeCalculation_F() '// 전문대학이상 졸업자

			'2. 수시/ 일반전형 / 검정고시
			Case "2"	
				Call GradeCalculation_E() '// 검정고시 출신자
				Call GradeCalculation_D() '// 면접점수 가져오기

			'3. 수시/ 일반전형 / 2008 ~ 현재 졸업예정(학생부 40% + 면접 60%)
			'4. 수시/ 일반전형 / 1998 ~ 2007년 졸업자(학생부 40% + 면접 60%)
			'5. 수시/ 일반전형 / 1997년 이전 졸업자(학생부 40% + 면접 60%)
			Case "3", "4", "5"
				Call GradeCalculation_C() '// 생기부
				Call GradeCalculation_D() '// 면접점수 가져오기

			'6. 수시/ 일반전형, 전문대이상 외 / 검정고시
			Case "6"	
				Call GradeCalculation_E() '// 검정고시 출신자

			'7. 수시/ 일반전형, 전문대이상 외 / 2008 ~ 현재 졸업예정(학생부 100%)
			'8. 수시/ 일반전형, 전문대이상 외 / 1998 ~ 2007년 졸업자(학생부 100%)
			'9. 수시/ 일반전형, 전문대이상 외 / 1997년 이전 졸업자(학생부 100%)
			Case "7", "8", "9"	
				Call GradeCalculation_C() '// 생기부

			'10. 정시/ 면접전형 / 기술드론부사관과, 기술행정부사관과, 조리부사관과, 항공서비스과 (수능 40% + 면접 60%)
			Case "10"	
				Call GradeCalculation_G() '// 수능점수 환산
				Call GradeCalculation_D() '// 면접점수 가져오기

			'11. 정시/ 비면접전형 / 검정고시
			Case "11"	
				Call GradeCalculation_G() '// 수능점수 환산
				Call GradeCalculation_E() '// 검정고시 출신자

			'12. 정시/ 비면접전형 / 2008 ~ 현재 졸업예정(수능 70% + 학생부 30%)
			'13. 정시/ 비면접전형 / 1998 ~ 2007년 졸업자(수능 70% + 학생부 30%)
			'14. 정시/ 비면접전형 / 1997년 이전 졸업자(수능 70% + 학생부 30%)
			Case "12", "13", "14"
				Call GradeCalculation_G() '// 수능점수 환산
				Call GradeCalculation_C() '// 생기부
		End Select

		'// =================================================================
		'// 환산 후 서비스
		'// =================================================================
		'// 1. 전형별 통합점수 생성 (가산점 + 전형별 %)
		'// 2. 같은 점수 찾아 동석차 순위 생성
		'// 3. 석차 생성(1순위 통합점수, 2순위 독석차 순위)
		'// 4. 모집단위 별 모집인원으로 합/불여부 생성
		'// 5. 불합격에 대한 예비석차 생성
		'// =================================================================

		'/////////// 가산점 + 전형별 %적용 점수 합치기 ///////////
		SQL = ""
		SQL = SQL & vbCrLf & " Select *   "
		SQL = SQL & vbCrLf & " from ChangeScoreTable   "
		SQL = SQL & vbCrLf & " WHERE 1 = 1    "
		SQL = SQL & vbCrLf & " And StudentNumber =  " & StudentNumber

		'objDB.blnDebug = TRUE
		arrParams2 = objDB.fnGetArray
		AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)		

		If Not(isnull(AryHash2)) Then
			for j = 0 to ubound(AryHash2,1)
				StudentRecordScoreTemp		=	AryHash2(j).Item("StudentRecordScore")
				InterviewerScoreTemp		=	AryHash2(j).Item("InterviewerScore")
				QualificationScoreTemp		=	AryHash2(j).Item("QualificationScore")
				UniversityScoreTemp			=	AryHash2(j).Item("UniversityScore")
				CSATScoreTemp				=	AryHash2(j).Item("CSATScore")

				StudentRecordAverageTemp	=	AryHash2(j).Item("StudentRecordAverage")
				CreditSumTemp				=	AryHash2(j).Item("CreditSum")
				ChoiceSemesterTemp			=	AryHash2(j).Item("ChoiceSemester")

				KorLanScoreTemp				=	AryHash2(j).Item("KorLanScore")
				EnglishScoreTemp			=	AryHash2(j).Item("EnglishScore")
				MathematicsScoreTemp		=	AryHash2(j).Item("MathematicsScore")

				UniversityCreditTemp		=	AryHash2(j).Item("UniversityCredit")

				'Response.write "<Br>"
				'Response.write	"수험번호 : " & StudentNumber & "<br>"
				'Response.write	"가산점 : " & ExtraPointScore & "<br>"
				'Response.write	"나이 : " & Minor & "<br>"

				'Response.write	"생기부 : " & StudentRecordScoreTemp & "<br>"
				'Response.write	"면접 : " & InterviewerScoreTemp & "<br>"
				'Response.write	"검정고시 : " & QualificationScoreTemp & "<br>"
				'Response.write	"대학점수 : " & UniversityScoreTemp & "<br>"
				'Response.write	"수능점수 : " & CSATScoreTemp & "<br>"

				'Response.write	"교과 성적 평균 : " & StudentRecordAverageTemp & "<br>"
				'Response.write	"선택한 학기 이수단위 합 : " & CreditSumTemp & "<br>"
				'Response.write	"반영학기 : " & ChoiceSemesterTemp & "<br>"

				'Response.write	"국어영역 성적 : " & KorLanScoreTemp & "<br>"
				'Response.write	"영어영역 성적 : " & EnglishScoreTemp & "<br>"
				'Response.write	"수학영역 성적 : " & MathematicsScoreTemp & "<br>"

				'Response.write	"졸업(취득)학점 : " & UniversityCreditTemp & "<br>"
				'Response.write "<Br>"

				Select Case StudentDivision
					'1. 수시/ 전문대이상 졸업자 전형(출신대학 성적 100%)
					Case "1"	
						'// 가산점 + 대학점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(UniversityScoreTemp)
							'Response.write "<Br>"
							'Response.write	"1. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(UniversityScoreTemp)
							'Response.write "<Br>"
							'Response.write	"1. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'2. 수시/ 일반전형 / 검정고시
					Case "2"	
						'// 가산점 + 검정고시점수 + 면접점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(QualificationScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"2. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(QualificationScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"2. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'3. 수시/ 일반전형 / 2008 ~ 현재 졸업예정(학생부 40% + 면접 60%)
					'4. 수시/ 일반전형 / 1998 ~ 2007년 졸업자(학생부 40% + 면접 60%)
					'5. 수시/ 일반전형 / 1997년 이전 졸업자(학생부 40% + 면접 60%)
					Case "3", "4", "5"
						'// 가산점 + 생기부점수 + 면접점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(StudentRecordScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"3.4.5. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(StudentRecordScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"3.4.5. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'6. 수시/ 일반전형, 전문대이상 외 / 검정고시
					Case "6"	
						'// 가산점 + 검정고시
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(QualificationScoreTemp)
							'Response.write "<Br>"
							'Response.write	"6. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(QualificationScoreTemp)
							'Response.write "<Br>"
							'Response.write	"6. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'7. 수시/ 일반전형, 전문대이상 외 / 2008 ~ 현재 졸업예정(학생부 100%)
					'8. 수시/ 일반전형, 전문대이상 외 / 1998 ~ 2007년 졸업자(학생부 100%)
					'9. 수시/ 일반전형, 전문대이상 외 / 1997년 이전 졸업자(학생부 100%)
					Case "7", "8", "9"	
						'// 가산점 + 생기부점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(StudentRecordScoreTemp)
							'Response.write "<Br>"
							'Response.write	"7.8.9. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(StudentRecordScoreTemp)
							'Response.write "<Br>"
							'Response.write	"7.8.9. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'10. 정시/ 면접전형 / 기술드론부사관과, 기술행정부사관과, 조리부사관과, 항공서비스과 (수능 40% + 면접 60%)
					Case "10"	
						'// 가산점 + 수능점수 + 면접점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(CSATScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"10. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(CSATScoreTemp) + CDbl(InterviewerScoreTemp)
							'Response.write "<Br>"
							'Response.write	"10. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'11. 정시/ 비면접전형 / 검정고시
					Case "11"	
						'// 가산점 + 수능점수 + 검정고시점수
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(CSATScoreTemp) + CDbl(QualificationScoreTemp)
							'Response.write "<Br>"
							'Response.write	"11. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(CSATScoreTemp) + CDbl(QualificationScoreTemp)
							'Response.write "<Br>"
							'Response.write	"11. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

					'12. 정시/ 비면접전형 / 2008 ~ 현재 졸업예정(수능 70% + 학생부 30%)
					'13. 정시/ 비면접전형 / 1998 ~ 2007년 졸업자(수능 70% + 학생부 30%)
					'14. 정시/ 비면접전형 / 1997년 이전 졸업자(수능 70% + 학생부 30%)
					Case "12", "13", "14"
						'// 가산점 + 수능점수 + 생기부
						If Not(isnull(ExtraPointScore)) Then
							totScore = CDbl(ExtraPointScore) + CDbl(CSATScoreTemp) + CDbl(StudentRecordScoreTemp)
							'Response.write "<Br>"
							'Response.write	"12.13.14. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						Else
							totScore = CDbl(CSATScoreTemp) + CDbl(StudentRecordScoreTemp)
							'Response.write "<Br>"
							'Response.write	"12.13.14. 통합점수 : " & totScore & "<br>"
							'Response.write "<Br>"
						End If

				End Select

				'Response.write "<br>"
				'Response.write totScore
				'Response.write "<Br>"

				'테스트용 동석차 - 전부 100점, 학과100 , 전형X05130
				totScore = 100
				Subject = "100"
				Division1 = "X05130"
				
				If totScore <> "" Then
					'// totScore db에 저장 =================
					SQL = ""
					SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
					SQL = SQL & vbCrLf & "SET	 totScore=?, Minor=?  "
					SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
					SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

					'Update일 때는 UPDT입력
					arrParams = Array(_
						  Array("@totScore",				adDouble,		adParamInput,		0,		totScore) _
						, Array("@Minor",					adInteger,		adParamInput,		0,		Minor) _

						, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
						, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
						, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
					)

					'objDB.blnDebug = True
					Call objDB.sbExecSQL(SQL, arrParams)
				End If		
			Next
		End If
	Next

	'/////////// 학과 + 구분1 리스트 (리스트 별 정원) ///////////
	SQL = ""
	SQL = SQL & vbCrLf & " select SubjectCode, Subject, Division1, Quorum, QuorumFix "
	SQL = SQL & vbCrLf & " from SubjectTable "
	SQL = SQL & vbCrLf & " Where 1=1 "                 
	SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"         
	SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'" 
	SQL = SQL & vbCrLf & " group by SubjectCode, Subject, Division1, Quorum, QuorumFix "

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	If Not(isnull(AryHash)) Then
		for i = 0 to ubound(AryHash,1)
			SubjectCode			=	AryHash(i).Item("SubjectCode")
			BasicSubject		=	AryHash(i).Item("Subject")
			BasicDivision1		=	AryHash(i).Item("Division1")
			Quorum				=	AryHash(i).Item("Quorum")		'// 모집인원
			QuorumFix			=	AryHash(i).Item("QuorumFix")	'// 입학정원

			'Response.write "<br>"
			'Response.write "학과 : " & BasicSubject & "<br>"
			'Response.write "구분1 : " & BasicDivision1 & "<br>"
			'Response.write "정원 : " & Quorum & "<br>"
			'Response.write "모집인원 : " & QuorumFix
			'Response.write "<Br>"

			'/////////// 동석차 기준 ///////////
			SQL = "" 
			SQL = SQL & vbCrLf & " select UnqualifiedStandard1 "         
			SQL = SQL & vbCrLf & " from AppraisalTable "
			SQL = SQL & vbCrLf & " Where 1=1 "                 
			SQL = SQL & vbCrLf & " and SubjectCode = '" & SubjectCode & "'"         
			
			'objDB.blnDebug = TRUE
			arrParams2 = objDB.fnGetArray
			AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

			If Not(isnull(AryHash2)) Then
				UnqualifiedStandard1 = AryHash2(0).Item("UnqualifiedStandard1")
			End If

			'Response.write "동석차 기준" & UnqualifiedStandard1 & "<br>"

			'/////////// 학과 + 구분1 리스트 별 동석차인 점수 리스트 ///////////
			SQL = "" 
			SQL = SQL & vbCrLf & " select totScore ,count(totScore) totScoreCount "         
			SQL = SQL & vbCrLf & " from ChangeScoreTable "
			SQL = SQL & vbCrLf & " Where 1=1 "                 
			SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"         
			SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'"  
			SQL = SQL & vbCrLf & " and Subject = '" & BasicSubject & "'"        
			SQL = SQL & vbCrLf & " and Division1 = '" & BasicDivision1 & "'"   
			SQL = SQL & vbCrLf & " group by totScore "           
			SQL = SQL & vbCrLf & " having count(totScore)>=2 " 
			
			'objDB.blnDebug = TRUE
			arrParams2 = objDB.fnGetArray
			AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

			If Not(isnull(AryHash2)) Then
				for j = 0 to ubound(AryHash2,1)
					'Response.write "<br>"
					'Response.write "동점인 점수 : " & AryHash2(j).Item("totScore") & "<br>"
					'Response.write "해당점수 개수 : " & AryHash2(j).Item("totScoreCount")
					'Response.write "<Br>"

					DrawScore = AryHash2(j).Item("totScore")
					totScoreCount = AryHash2(j).Item("totScoreCount")
					totScoreCount = totScoreCount-1

					'/////////// 학과 + 구분1 리스트 별 동석차인 점수별로 동석차 순위 생성 ///////////
					SQL = "" 
					SQL = SQL & vbCrLf & " select *  "        
					SQL = SQL & vbCrLf & " from ChangeScoreTable  "       
					SQL = SQL & vbCrLf & " Where 1=1  "                   
					SQL = SQL & vbCrLf & " and MYear = '" & BasicMYear & "'"        
					SQL = SQL & vbCrLf & " and Division0 = '" & BasicDivision0 & "'"   
					SQL = SQL & vbCrLf & " and Subject = '" & BasicSubject & "'"          
					SQL = SQL & vbCrLf & " and Division1 = '" & BasicDivision1 & "'"    
					SQL = SQL & vbCrLf & " and totScore = '" & DrawScore & "'" 
					
					'objDB.blnDebug = TRUE
					arrParams3 = objDB.fnGetArray
					AryHash3 = objDB.fnExecSQLGetHashMap(SQL, arrParams3)

					If Not(isnull(AryHash3)) Then
						for y= 0 to ubound(AryHash3,1)
							StudentNumber				=	AryHash3(y).Item("StudentNumber")
							Minor						=	AryHash3(y).Item("Minor")

							InterviewerScoreTemp		=	AryHash3(y).Item("InterviewerScore")

							StudentRecordAverageTemp	=	AryHash3(y).Item("StudentRecordAverage")
							CreditSumTemp				=	AryHash3(y).Item("CreditSum")
							ChoiceSemesterTemp			=	AryHash3(y).Item("ChoiceSemester")

							KorLanScoreTemp				=	AryHash3(y).Item("KorLanScore")
							EnglishScoreTemp			=	AryHash3(y).Item("EnglishScore")
							MathematicsScoreTemp		=	AryHash3(y).Item("MathematicsScore")

							UniversityCreditTemp		=	AryHash3(y).Item("UniversityCredit")

							AryStudentNumber(y)			=	StudentNumber
							AryInterviewerScore(y)		=	InterviewerScoreTemp
							AryStudentRecordAverage(y)	=	StudentRecordAverageTemp
							AryCreditSum(y)				=	CreditSumTemp
							AryChoiceSemester(y)		=	ChoiceSemesterTemp
							AryMinor(y)					=	Minor
							AryUniversityCredit(y)		=	UniversityCreditTemp
							AryKorLanScore(y)			=	KorLanScoreTemp
							AryEnglishScore(y)			=	EnglishScoreTemp
							AryMathematicsScore(y)		=	MathematicsScoreTemp
						Next
					End If
							
					UnqualifiedStandardCheck = True

					'테스트용 동석차기준 
					UnqualifiedStandard1 = "1"
					
					Dim asdf
					For asdf = 0 To totScoreCount
						'Response.write "<Br>"
						'Response.write	"수험번호 : " & AryStudentNumber(asdf) & "<br>"
						'Response.write	"나이 : " & AryMinor(asdf) & "<br>"

						'Response.write	"면접 : " & AryInterviewerScore(asdf) & "<br>"

						'Response.write	"교과 성적 평균 등급 : " & AryStudentRecordAverage(asdf) & "<br>"
						'Response.write	"선택한 학기 이수단위 합 : " & AryCreditSum(asdf) & "<br>"
						'Response.write	"반영학기 : " & AryChoiceSemester(asdf) & "<br>"

						'Response.write	"국어영역 성적 : " & AryKorLanScore(asdf) & "<br>"
						'Response.write	"영어영역 성적 : " & AryEnglishScore(asdf) & "<br>"
						'Response.write	"수학영역 성적 : " & AryMathematicsScore(asdf) & "<br>"

						'Response.write	"졸업(취득)학점 : " & AryUniversityCredit(asdf) & "<br>"
						'Response.write	"동석차 기준 : " & UnqualifiedStandard1 & "<br>"	
						'Response.write "<Br>"
					Next

					'// 1. 수시/면접 - 일반전형
					'// 1순위 면접고사 성적 상위자
					'// 2순위 교과 성적 평균등급 상위자
					'// 3순위 선택한 학기 이수단위 합이 높은자
					'// 4순위 학생부 성적 고학년 고학기 선택자
					'// 5순위 연소자
					If UnqualifiedStandard1 = "1" Then
						UnqualifiedStandardCheck = UnqualifiedStandard_A(AryStudentNumber, AryInterviewerScore, AryStudentRecordAverage, AryCreditSum, AryChoiceSemester, AryMinor, totScoreCount)		
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null

							AryStudentNumber(DrawRankingNum) = null
							AryInterviewerScore(DrawRankingNum) = null
							AryStudentRecordAverage(DrawRankingNum) = null
							AryCreditSum(DrawRankingNum) = null
							AryChoiceSemester(DrawRankingNum) = null
							AryMinor(DrawRankingNum) = null
						Next
							
					'// 2. 수시/비면접 - 일반고,전문,농어촌,기초및차상위
					'// 1순위 교과 성적 평균등급 상위자
					'// 2순위 선택한 학기 이수단위 합이 높은자
					'// 3순위 학생부 성적 고학년 고학기 선택자
					'// 4순위 연소자
					ElseIf UnqualifiedStandard1 = "2" Then
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryStudentRecordAverage, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryCreditSum, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryChoiceSemester, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Down(AryStudentNumber, AryMinor, totScoreCount)
						End If	
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryStudentRecordAverage(DrawRankingNum) = null
							AryCreditSum(DrawRankingNum) = null
							AryChoiceSemester(DrawRankingNum) = null
							AryMinor(DrawRankingNum) = null
						Next

					'// 3. 수시 - 전문대졸이상졸업자전형
					'// 1순위 대학 백분위 성적 우수자 - 성적이 같은 자가 동점이므로 생략
					'// 2순위 졸업(취득) 학점이 많은 자
					ElseIf UnqualifiedStandard1 = "3" Then
						UnqualifiedStandardCheck = UnqualifiedStandard_C(AryStudentNumber, AryUniversityCredit, totScoreCount)
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryUniversityCredit(DrawRankingNum) = null
						Next

					'// 4. 정시/면접 - 일반전형
					'// 1순위 면접고사 성적 상위자
					'// 2순위 국어영역 성적 상위자
					'// 3순위 영어역역 성적 상위자
					'// 4순위 수학영역 성적 상위자
					'// 5순위 연소자
					ElseIf UnqualifiedStandard1 = "4" Then
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryInterviewerScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryKorLanScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryEnglishScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryMathematicsScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Down(AryStudentNumber, AryMinor, totScoreCount)
						End If	
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryInterviewerScore(DrawRankingNum) = null
							AryKorLanScore(DrawRankingNum) = null
							AryEnglishScore(DrawRankingNum) = null
							AryMathematicsScore(DrawRankingNum) = null
							AryMinor(DrawRankingNum) = null
						Next

					'// 5. 정시/비면접 - 일반전형
					'// 1순위 국어영역 성적 상위자
					'// 2순위 영어역역 성적 상위자
					'// 3순위 수학영역 성적 상위자
					'// 4순위 교과성적 평균등급 상위자
					'// 5순위 연소자
					ElseIf UnqualifiedStandard1 = "5" Then
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryKorLanScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryEnglishScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryMathematicsScore, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryStudentRecordAverage, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Down(AryStudentNumber, AryMinor, totScoreCount)
						End If	
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryKorLanScore(DrawRankingNum) = null
							AryEnglishScore(DrawRankingNum) = null
							AryMathematicsScore(DrawRankingNum) = null
							AryStudentRecordAverage(DrawRankingNum) = null
							AryMinor(DrawRankingNum) = null
						Next

					'// 6. 정시 - 농어촌,기초및차상위
					'// 1순위 교과 성적 평균등급 상위자
					'// 2순위 선택한 학기 이수단위 합이 높은자
					'// 3순위 학생부 성적 고학년 고학기 선택자
					'// 4순위 연소자
					ElseIf UnqualifiedStandard1 = "6" Then
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryStudentRecordAverage, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryCreditSum, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryChoiceSemester, totScoreCount)
						End If
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Down(AryStudentNumber, AryMinor, totScoreCount)
						End If	
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryStudentRecordAverage(DrawRankingNum) = null
							AryCreditSum(DrawRankingNum) = null
							AryChoiceSemester(DrawRankingNum) = null
							AryMinor(DrawRankingNum) = null
						Next

					'// 7. 정시 - 전문대졸이상졸업자전형
					'// 1순위 대학 백분위 성적 우수자 - 성적이 같은 자가 동점이므로 생략
					'// 2순위 졸업(취득) 학점이 많은 자
					ElseIf UnqualifiedStandard1 = "7" Then
						If UnqualifiedStandardCheck Then
							UnqualifiedStandardCheck = UnqualifiedStandard_Up(AryStudentNumber, AryUniversityCredit, totScoreCount)
						End If	
						'배열 초기화
						For DrawRankingNum = 0 To totScoreCount							
							DrawRanking(DrawRankingNum) = null
							AryStudentNumber(DrawRankingNum) = null
							AryUniversityCredit(DrawRankingNum) = null
						Next
					End If
				Next
			End If
		Next
	End If
End If

'생기부 환산
Sub GradeCalculation_C()
	'On Error Resume Next

	'변수 초기화
	GradeCalculation_1_1 = 0
	CompleteUnit_1_1 = 0
	CompleteGrade_1_1 = 0
	GradeCalculation_1_2 = 0
	CompleteUnit_1_2 = 0
	CompleteGrade_1_2 = 0
	GradeCalculation_2_1 = 0
	CompleteUnit_2_1 = 0
	CompleteGrade_2_1 = 0
	GradeCalculation_2_2 = 0
	CompleteUnit_2_2 = 0
	CompleteGrade_2_2 = 0
	GradeCalculation_3_1 = 0
	CompleteUnit_3_1 = 0
	CompleteGrade_3_1 = 0
	StudentRecordAverage = 0
	SubjectCnt = 0

	if CompleteUnitCnt > 0 Then
		'// 학생별 과목수 만큼 Loop
		for intNUM = 0 to CompleteUnitCnt
			YearType		= YearTypeTemp(intNUM)		'// 학년학기
			CompleteUnit	= CompleteUnitTemp(intNUM)	'// 이수단위
			Grade			= GradeTemp(intNUM)			'// 석차등급
			OriginalScore	= OriginalScoreTemp(intNUM)	'// 원점수

			'Response.write "<Br>"
			'Response.write "<Br>"
			'Response.write "반영학기 : " & YearType & " / "
			'Response.write "이수단위 : " & CompleteUnit & " / "
			'Response.write "등급 : " & Grade & " / "	
			'Response.write "원점수 : " & OriginalScore & " / "
			'Response.write "과목수 : " & CompleteUnitCnt + 1 & " / "

			'// 이수단위와 등급이 있을 때
			If (CompleteUnit <> 0 And Grade <> 0) Then
				'// 이수*등급 구하기
				'// 이수*등급 = 이수단위*석차등급
				ConvertCompleteGrade = CompleteUnit * Grade

				'Response.write "<Br>"
				'Response.write "<Br>"
				'Response.write "이수*등급 : " & ConvertCompleteGrade 

				'// 이수단위 총합 구하기
				'// 성적 총합 구하기
				'// 이수*등급 총합 구하기
				execute("CompleteUnit_"& YearType &" = CompleteUnit_"& YearType &" + "& CompleteUnit)
				If Not(isnull(OriginalScore)) Then
					execute("StudentRecordAverage = StudentRecordAverage + "& OriginalScore)
					SubjectCnt = SubjectCnt + 1
				End If
				execute("CompleteGrade_"& YearType &" = CompleteGrade_"& YearType &" + "& ConvertCompleteGrade)
			End if
		Next

		'Response.write "<Br>"
		'Response.write "<Br>"
		'Response.write "이수단위 총합 : " & CompleteUnit_2_1 & " / "
		'Response.write "이수*등급 총합 : " & CompleteGrade_2_1 

		'// 입력값 체크 설정
		CompleteUnitValueCheck = True

		If CInt(eval("CompleteUnit_"& YearType)) <> 0 then
			'// 학기별 등급 계산
			execute("GradeCalculation_"& YearType &" = CompleteGrade_"& YearType &" / CompleteUnit_"& YearType)
			'// 성적 평균 구하기
			execute("StudentRecordAverage = StudentRecordAverage / "& SubjectCnt)
			StudentRecordAverage = FormatNumber(StudentRecordAverage, 5)
			If YearType = "1_1" Then
				GradeCalculation_1_1 = FormatNumber(GradeCalculation_1_1 - 0.0000005, 6)
				GradeCalculation_1_1 = FormatNumber(CDbl(GradeCalculation_1_1),5)
				CreditSum = CompleteUnit_1_1
				StudentRecordGradeAverage = GradeCalculation_1_1
			ElseIf YearType = "1_2" Then
				GradeCalculation_1_2 = FormatNumber(GradeCalculation_1_2 - 0.0000005, 6)
				GradeCalculation_1_2 = FormatNumber(CDbl(GradeCalculation_1_2),5)
				CreditSum = CompleteUnit_1_2
				StudentRecordGradeAverage = GradeCalculation_1_2
			ElseIf YearType = "2_1" Then
				GradeCalculation_2_1 = FormatNumber(GradeCalculation_2_1 - 0.0000005, 6)
				GradeCalculation_2_1 = FormatNumber(CDbl(GradeCalculation_2_1),5)
				CreditSum = CompleteUnit_2_1
				StudentRecordGradeAverage = GradeCalculation_2_1
			ElseIf YearType = "2_2" Then
				GradeCalculation_2_2 = FormatNumber(GradeCalculation_2_2 - 0.0000005, 6)
				GradeCalculation_2_2 = FormatNumber(CDbl(GradeCalculation_2_2),5)
				CreditSum = CompleteUnit_2_2
				StudentRecordGradeAverage = GradeCalculation_2_2
			ElseIf YearType = "3_1" Then
				GradeCalculation_3_1 = FormatNumber(GradeCalculation_3_1 - 0.0000005, 6)
				GradeCalculation_3_1 = FormatNumber(CDbl(GradeCalculation_3_1),5)
				CreditSum = CompleteUnit_3_1
				StudentRecordGradeAverage = GradeCalculation_3_1
			End IF			

			'Response.write "<Br>"
			'Response.write "<Br>"
			'Response.write "선택학기 평균등급 : " & StudentRecordGradeAverage

			'// 환산점수 구하기
			'// 수시 312+((9-선택학기 평균등급)*11)
			'// 정시 236+((9-선택학기 평균등급)*11)
			If DivistionCheck = "1" Then '수시1, 수시2
				'//수시공식 가져와서 치환하여 환산
				FormulaTemp = Replace(Replace(Formula1, "A", "GradeCalculation_"& YearType), "Z", "Dim StudentRecordScore : StudentRecordScore")
				'Response.write Formula1 & "<br>"
				'Response.write FormulaTemp
				'Response.End
				execute(FormulaTemp)
			ElseIf DivistionCheck = "2" Then '정시, 추가
				'//정시공식 가져와서 치환하여 환산
				FormulaTemp = Replace(Replace(Formula2, "A", "GradeCalculation_"& YearType), "Z", "Dim StudentRecordScore : StudentRecordScore")
				'Response.write Formula2 & "<br>"
				'Response.write FormulaTemp
				'Response.End
				execute(FormulaTemp)
			End If

			'StudentRecordScore	'생기부성적	
			'StudentRecordAverage	'교과 성적평균
			'CreditSum	'선택한 학기 이수단위 합

			'Response.write "<Br>"
			'Response.write "<Br>"
			'Response.write "수험번호 : " & StudentNumber & "<br>"
			'Response.write "모집시기 : " & Division0 & "<br>"
			'Response.write "학과명 : " & Subject & "<br>"
			'Response.write "구분1 : " & Division1 & "<br>"
			'Response.write "구분2 : " & Division2 & "<br>"
			'Response.write "구분3 : " & Division3 & "<br>"

			'Response.write "생기부성적 : " & StudentRecordScore & "<br>"
			'Response.write "교과성적평균 : " & StudentRecordAverage & "<br>"
			'Response.write "선택한 학기 이수단위 합 : " & CreditSum & "<br>"
			'Response.write "반영학기 : " & ChoiceSemester & "<br>"

			'Response.write "입력한 유저 : " & SessionUserID & "<br>"
			'Response.write "유저 아이피 : " & ASP_USER_IP & "<br>"

			'Response.write "<Br>"
			'Response.write "<Br>"

			'// Insert & Update 구분
			If InUpDivisionCheck Then
				InUpDivision = "Update"
			Else
				If InUpDivision = "Update" Then
					InUpDivision = "Update"
				Else
					InUpDivision = "Insert"
				End If
			End If

			'Response.write "<Br>"
			'Response.write "<Br>"
			'Response.write "수험번호 : " & StudentNumber
			'Response.write "insert구분 : " & InUpDivision
			'Response.write "<Br>"
			'Response.write "<Br>"

			InUpDivisionCheck = true

			If InUpDivision = "Insert" Then
				'// 입력 =================
				SQL = ""
				SQL = SQL & vbCrLf & "INSERT INTO ChangeScoreTable ( "
				SQL = SQL & vbCrLf & "		MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3  "
				SQL = SQL & vbCrLf & "		 , StudentRecordScore, StudentRecordAverage, CreditSum, ChoiceSemester  "
				SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
				SQL = SQL & vbCrLf & " ) VALUES ( "
				SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
				SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
				SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
				SQL = SQL & vbCrLf & " ) "

				'insert일 때는 INPT입력
				arrParams = Array(_
					  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
					, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
					, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
					, Array("@Subject",					adVarchar,		adParamInput,		60,		Subject) _
					, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
					, Array("@Division2",				adVarchar,		adParamInput,		60,		Division2) _
					, Array("@Division3",				adVarchar,		adParamInput,		60,		Division3) _

					, Array("@StudentRecordScore",		adDouble,		adParamInput,		0,		StudentRecordScore) _
					, Array("@StudentRecordAverage",	adDouble,		adParamInput,		0,		StudentRecordGradeAverage) _
					, Array("@CreditSum",				adInteger,		adParamInput,		0,		CreditSum) _
					, Array("@ChoiceSemester",			adInteger,		adParamInput,		0,		ChoiceSemester) _

					, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
					, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
				)

				'objDB.blnDebug = True
				Call objDB.sbExecSQL(SQL, arrParams)
			ElseIf InUpDivision = "Update" Then
				'// 수정 =================
				SQL = ""
				SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
				SQL = SQL & vbCrLf & "SET	 StudentRecordScore=?, StudentRecordAverage=?, CreditSum=?, ChoiceSemester=?  "
				SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
				SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

				'Update일 때는 UPDT입력
				arrParams = Array(_
					  Array("@StudentRecordScore",		adDouble,		adParamInput,		0,		StudentRecordScore) _
					, Array("@StudentRecordAverage",	adDouble,		adParamInput,		0,		StudentRecordGradeAverage) _
					, Array("@CreditSum",				adInteger,		adParamInput,		0,		CreditSum) _
					, Array("@ChoiceSemester",			adInteger,		adParamInput,		0,		ChoiceSemester) _

					, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
					, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
					, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
				)

				'objDB.blnDebug = True
				Call objDB.sbExecSQL(SQL, arrParams)
			End If
		Else
			'// 제대로 입력되지 않았다면 오류 처리
			CompleteUnitValueCheck = CompleteUnitValueCheck And False
			ResultMSG = "생기부 계산실패"
		End If
	End If

	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
	Else 
		strResult = "Complete"
		returnMSG = "2008~현재졸업(예정)자 내신성적산출완료"
	End If
End Sub

'// 면접접수 환산
Sub GradeCalculation_D()
	'On Error Resume Next

	'Response.write "<br>"
	'Response.write ItemPoint_01 & " / "
	'Response.write ItemPoint_02 & " / "
	'Response.write ItemPoint_03 & " / "
	'Response.write ItemPoint_04 & " / "
	'Response.write ItemPoint_05 & " / "
	'Response.write ItemPoint_06 & " / "
	'Response.write ItemPoint_07 & " / "
	'Response.write ItemPoint_08 & " / "
	'Response.write ItemPoint_09 & " / "
	'Response.write ItemPoint_10 & " / "
	'Response.write "<Br>"

	'// 입력받은 값이 모두 있을때 계산
	If (ItemPoint_01 <> 0 And ItemPoint_02 <> 0 And ItemPoint_03 <> 0 And ItemPoint_04 <> 0 And ItemPoint_05 <> 0 And ItemPoint_06 <> 0 And ItemPoint_07 <> 0 And ItemPoint_08 <> 0 And ItemPoint_09 <> 0 And ItemPoint_10 <> 0) Then
		'// 모두 더한 값 
		GetScore = CInt(ItemPoint_01) + CInt(ItemPoint_02) + CInt(ItemPoint_03) + CInt(ItemPoint_04) + CInt(ItemPoint_05) + CInt(ItemPoint_06) + CInt(ItemPoint_07) + CInt(ItemPoint_08) + CInt(ItemPoint_09) + CInt(ItemPoint_10)

		'Response.write "<br>"
		'Response.write GetScore & " / "
		'Response.write "<Br>"

		'// insert & Update 구분
		If InUpDivisionCheck Then
			InUpDivision = "Update"
		Else
			If InUpDivision = "Update" Then
				InUpDivision = "Update"
			Else
				InUpDivision = "Insert"
			End If
		End If

		'Response.write "<Br>"
		'Response.write "<Br>"
		'Response.write "수험번호 : " & StudentNumber & "<br>"
		'Response.write "모집시기 : " & Division0 & "<br>"
		'Response.write "학과명 : " & Subject & "<br>"
		'Response.write "구분1 : " & Division1 & "<br>"
		'Response.write "구분2 : " & Division2 & "<br>"
		'Response.write "구분3 : " & Division3 & "<br>"

		'Response.write "면접점수 : " & InterviewerScore & "<br>"
		'Response.write "insert구분 : " & InUpDivision & "<br>"
		'Response.write "<Br>"
		'Response.write "<Br>"

		InUpDivisionCheck = true

		If InUpDivision = "Insert" Then
			'// 입력 =================
			SQL = ""
			SQL = SQL & vbCrLf & "INSERT INTO ChangeScoreTable ( "
			SQL = SQL & vbCrLf & "		MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3  "
			SQL = SQL & vbCrLf & "		 , InterviewerScore "
			SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
			SQL = SQL & vbCrLf & " ) VALUES ( "
			SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
			SQL = SQL & vbCrLf & "		, ? "
			SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
			SQL = SQL & vbCrLf & " ) "

			'insert일 때는 INPT입력
			arrParams = Array(_
				  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
				, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
				, Array("@Subject",					adVarchar,		adParamInput,		60,		Subject) _
				, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
				, Array("@Division2",				adVarchar,		adParamInput,		60,		Division2) _
				, Array("@Division3",				adVarchar,		adParamInput,		60,		Division3) _

				, Array("@InterviewerScore",		adDouble,		adParamInput,		0,		GetScore) _

				, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		ElseIf InUpDivision = "Update" Then
			'// 수정 =================
			SQL = ""
			SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
			SQL = SQL & vbCrLf & "SET	 InterviewerScore=?  "
			SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
			SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

			'Update일 때는 UPDT입력
			arrParams = Array(_
				  Array("@InterviewerScore",		adDouble,		adParamInput,		0,		GetScore) _

				, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		End If
	Else
		ResultMSG = "면접점수 가져오기 실패"
		'Response.write "<br>"
		'Response.write ResultMSG & " / "
		'Response.write "<Br>"
	End if

	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
	Else 
		strResult = "Complete"
		returnMSG = "면접점수 가져오기 완료"
	End If
End Sub

'// 검정고시 출신자
Sub GradeCalculation_E()
	'On Error Resume Next
	
	'변수 초기화
	ConvertGradetot = 0

	'검정고시 과목이 있으면 계산
	if ScoreCnt > 0 Then
		for intNUM = 0 to ScoreCnt
			'// 과목별 점수
			GEDScore = CInt(Score(intNUM))

			'Response.write GEDScore & "<br>"

			'// 입력받은 값이 있을때 계산
			If GEDScore <> 0 Then
				'//등급 구하기
				ConvertGrade = QualificationGrade(GEDScore)
				'Response.write ConvertGrade & "<br>"
				'//등급 더하기
				ConvertGradetot = ConvertGradetot + ConvertGrade				
				'Response.write ConvertGradetot & "<br>"
			End if
		Next

		'// 평균등급 구하기
		'// 총등급 / 과목수
		'// 총등급이 0이 아닐때 계산
		If ConvertGrade <> 0 Then
			execute("GradeCalculation = " & ConvertGradetot & " / " & ScoreCnt)
			GradeCalculation = FormatNumber(GradeCalculation - 0.0000005, 6)
			GradeCalculation = FormatNumber(CDbl(GradeCalculation),5)

			'// 환산점수 구하기
			'// 수시 312+((9-선택학기 평균등급)*11)
			'// 정시 236+((9-선택학기 평균등급)*11)		
			If DivistionCheck = "1" Then '수시1, 수시2
				'//수시공식 가져와서 치환하여 환산
				FormulaTemp = Replace(Replace(Formula1, "A", GradeCalculation), "Z", "Dim QualificationScore : QualificationScore")
				'Response.write Formula1 & "<br>"
				'Response.write FormulaTemp
				'Response.End
				execute(FormulaTemp)
			ElseIf DivistionCheck = "2" Then '정시, 추가
				'//정시공식 가져와서 치환하여 환산
				FormulaTemp = Replace(Replace(Formula2, "A", GradeCalculation), "Z", "Dim QualificationScore : QualificationScore")
				'Response.write Formula2 & "<br>"
				'Response.write FormulaTemp
				'Response.End
				execute(FormulaTemp)
			End If
		End If

		'// insert & Update 구분
		If InUpDivisionCheck Then
			InUpDivision = "Update"
		Else
			If InUpDivision = "Update" Then
				InUpDivision = "Update"
			Else
				InUpDivision = "Insert"
			End If
		End If

		'Response.write "<Br>"
		'Response.write "<Br>"
		'Response.write "수험번호 : " & StudentNumber & "<br>"
		'Response.write "모집시기 : " & Division0 & "<br>"
		'Response.write "학과명 : " & Subject & "<br>"
		'Response.write "구분1 : " & Division1 & "<br>"
		'Response.write "구분2 : " & Division2 & "<br>"
		'Response.write "구분3 : " & Division3 & "<br>"

		'Response.write "검정점수 : " & QualificationScore & "<br>"
		'Response.write "insert구분 : " & InUpDivision & "<br>"
		'Response.write "<Br>"
		'Response.write "<Br>"

		InUpDivisionCheck = true

		If InUpDivision = "Insert" Then
			'// 입력 =================
			SQL = ""
			SQL = SQL & vbCrLf & "INSERT INTO ChangeScoreTable ( "
			SQL = SQL & vbCrLf & "		MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3  "
			SQL = SQL & vbCrLf & "		 , QualificationScore "
			SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
			SQL = SQL & vbCrLf & " ) VALUES ( "
			SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
			SQL = SQL & vbCrLf & "		, ? "
			SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
			SQL = SQL & vbCrLf & " ) "

			'insert일 때는 INPT입력
			arrParams = Array(_
				  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
				, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
				, Array("@Subject",					adVarchar,		adParamInput,		60,		Subject) _
				, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
				, Array("@Division2",				adVarchar,		adParamInput,		60,		Division2) _
				, Array("@Division3",				adVarchar,		adParamInput,		60,		Division3) _

				, Array("@QualificationScore",		adDouble,		adParamInput,		0,		QualificationScore) _

				, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		ElseIf InUpDivision = "Update" Then
			'// 수정 =================
			SQL = ""
			SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
			SQL = SQL & vbCrLf & "SET	 QualificationScore=?  "
			SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
			SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

			'Update일 때는 UPDT입력
			arrParams = Array(_
				  Array("@QualificationScore",		adDouble,		adParamInput,		0,		QualificationScore) _

				, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		End If
	Else
		ResultMSG = "검정고시 계산 실패"
	End If

	'Response.write "QualificationScore : " & QualificationScore & "<br>"

	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
	Else 
		strResult = "Complete"
		returnMSG = "검정고시 출신자 내신성적산출완료"
	End If
End Sub

'// 전문대학이상 졸업자
Sub GradeCalculation_F()
	'On Error Resume Next

	'Response.write "<br>"
	'Response.write "평균점수 : " & AugScore & " / "
	'Response.write "만점 : " & PerfectScore & " / "
	'Response.write "취득학점 : " & Credit 
	'Response.write "<Br>"

	'// 입력받은 값이 모두 있을때 계산
	If (PerfectScore <> 0 And AugScore <> 0) Then
		'// 평균점수 / 만점 * 100 
		GetScore = (AugScore / PerfectScore) * 100
		GetScore = FormatNumber(GetScore - 0.0005, 3)

		'Response.write "<br>"
		'Response.write GetScore & " / "
		'Response.write "<Br>"

		'// insert & Update 구분
		If InUpDivisionCheck Then
			InUpDivision = "Update"
		Else
			If InUpDivision = "Update" Then
				InUpDivision = "Update"
			Else
				InUpDivision = "Insert"
			End If
		End If

		'Response.write "<Br>"
		'Response.write "<Br>"
		'Response.write "수험번호 : " & StudentNumber & "<br>"
		'Response.write "모집시기 : " & Division0 & "<br>"
		'Response.write "학과명 : " & Subject & "<br>"
		'Response.write "구분1 : " & Division1 & "<br>"
		'Response.write "구분2 : " & Division2 & "<br>"
		'Response.write "구분3 : " & Division3 & "<br>"

		'Response.write "환산점수 : " & GetScore & "<br>"
		'Response.write "취득학점 : " & Credit & "<br>"
		'Response.write "insert구분 : " & InUpDivision & "<br>"
		'Response.write "<Br>"
		'Response.write "<Br>"

		InUpDivisionCheck = true

		If InUpDivision = "Insert" Then
			'// 입력 =================
			SQL = ""
			SQL = SQL & vbCrLf & "INSERT INTO ChangeScoreTable ( "
			SQL = SQL & vbCrLf & "		MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3  "
			SQL = SQL & vbCrLf & "		 , UniversityScore, UniversityCredit "
			SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
			SQL = SQL & vbCrLf & " ) VALUES ( "
			SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
			SQL = SQL & vbCrLf & "		, ?, ? "
			SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
			SQL = SQL & vbCrLf & " ) "

			'insert일 때는 INPT입력
			arrParams = Array(_
				  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
				, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
				, Array("@Subject",					adVarchar,		adParamInput,		60,		Subject) _
				, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
				, Array("@Division2",				adVarchar,		adParamInput,		60,		Division2) _
				, Array("@Division3",				adVarchar,		adParamInput,		60,		Division3) _

				, Array("@UniversityScore",			adDouble,		adParamInput,		0,		GetScore) _
				, Array("@UniversityCredit",		adInteger,		adParamInput,		0,		Credit) _

				, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		ElseIf InUpDivision = "Update" Then
			'// 수정 =================
			SQL = ""
			SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
			SQL = SQL & vbCrLf & "SET	 UniversityScore=?, UniversityCredit=?  "
			SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
			SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

			'Update일 때는 UPDT입력
			arrParams = Array(_
				  Array("@UniversityScore",			adDouble,		adParamInput,		0,		GetScore) _
				, Array("@UniversityCredit",		adInteger,		adParamInput,		0,		Credit) _

				, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		End If
	Else
		ResultMSG = "전문대학이상 계산 실패"
	End if


	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
	Else 
		strResult = "Complete"
		returnMSG = "전문대학이상 졸업자 내신성적산출완료"
	End If
End Sub

'// 수능
Sub GradeCalculation_G()
	'On Error Resume Next

	'Response.write "<br>"
	'Response.write EXAM_NUMB & " / "
	'Response.write LGFD_SDSC & " / "
	'Response.write MTFD_SDSC & " / "
	'Response.write FLFD_GRAD & " / "
	'Response.write RSFD_SCR1 & " / "
	'Response.write RSFD_SCR2 & " / "
	'Response.write RSFD_SCR3 & " / "
	'Response.write RSFD_SCR4 & " / "
	'Response.write SCFL_SDSC & " / "
	'Response.write "<Br>"

	'구분1,2에 성적이 각각 1개 이상 있으면 계산
	If ((LGFD_SDSC <> 0 Or MTFD_SDSC <> 0 Or FLFD_GRAD <> 0) And (RSFD_SCR1 <> 0 Or RSFD_SCR2 <> 0 Or RSFD_SCR3 <> 0 Or RSFD_SCR4 <> 0 Or SCFL_SDSC <> 0)) Then
		FLFD_GRAD = CSATGrad(FLFD_GRAD)
	'	Response.write FLFD_GRAD & " / "

		'구분1 : 국어, 수학, 영어
		one(0) = LGFD_SDSC
		one(1) = MTFD_SDSC
		one(2) = FLFD_GRAD
		'구분2 : 선택, 제2외국어
		two(0) = RSFD_SCR1
		two(1) = RSFD_SCR2
		two(2) = RSFD_SCR3
		two(3) = RSFD_SCR4
		two(4) = SCFL_SDSC

		'구분1,2 가장 큰 점수 구하기
		For t = 0 To UBound(one)
           If one(t) > one(Max) Then
                Max = t
           End If
		Next
		For t = 0 To UBound(two)
           If two(t) > two(Max) Then
                Max = t
           End If
		Next

		'구분1,2의 가장 큰 점수 2개로 평균구하기
		OneMax = one(Max)
		twoMax = two(Max)
		OneTwoAug = (OneMax + twoMax) / 2

		'Response.write "<br>"
		'Response.write OneMax & " / "
		'Response.write twoMax & " / "
		'Response.write OneTwoAug & " / "
		'Response.write "<Br>"

		'Insert &Update 구분
		If InUpDivisionCheck Then
			InUpDivision = "Update"
		Else
			If InUpDivision = "Update" Then
				InUpDivision = "Update"
			Else
				InUpDivision = "Insert"
			End If
		End If

		'Response.write "<Br>"
		'Response.write "<Br>"
		'Response.write "수험번호 : " & StudentNumber
		'Response.write "insert구분 : " & InUpDivision
		'Response.write "<Br>"
		'Response.write "<Br>"

		InUpDivisionCheck = true

		If InUpDivision = "Insert" Then
			'// 입력 =================
			SQL = ""
			SQL = SQL & vbCrLf & "INSERT INTO ChangeScoreTable ( "
			SQL = SQL & vbCrLf & "		MYear, StudentNumber, Division0, Subject, Division1, Division2, Division3  "
			SQL = SQL & vbCrLf & "		 , CSATScore, KorLanScore, EnglishScore, MathematicsScore  "
			SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
			SQL = SQL & vbCrLf & " ) VALUES ( "
			SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
			SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
			SQL = SQL & vbCrLf & "		, ?, getdate(), ? "
			SQL = SQL & vbCrLf & " ) "

			'insert일 때는 INPT입력
			arrParams = Array(_
				  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
				, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
				, Array("@Subject",					adVarchar,		adParamInput,		60,		Subject) _
				, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
				, Array("@Division2",				adVarchar,		adParamInput,		60,		Division2) _
				, Array("@Division3",				adVarchar,		adParamInput,		60,		Division3) _

				, Array("@CSATScore",				adDouble,		adParamInput,		0,		OneTwoAug) _
				, Array("@KorLanScore",				adDouble,		adParamInput,		0,		one(0)) _
				, Array("@EnglishScore",			adDouble,		adParamInput,		0,		one(2)) _
				, Array("@MathematicsScore",		adDouble,		adParamInput,		0,		one(1)) _

				, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		ElseIf InUpDivision = "Update" Then
			'// 수정 =================
			SQL = ""
			SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
			SQL = SQL & vbCrLf & "SET	 CSATScore=?, KorLanScore=?, EnglishScore=?, MathematicsScore=?  "
			SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
			SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

			'Update일 때는 UPDT입력
			arrParams = Array(_
				  Array("@CSATScore",				adDouble,		adParamInput,		0,		OneTwoAug) _
				, Array("@KorLanScore",				adDouble,		adParamInput,		0,		one(0)) _
				, Array("@EnglishScore",			adDouble,		adParamInput,		0,		one(2)) _
				, Array("@MathematicsScore",		adDouble,		adParamInput,		0,		one(1)) _

				, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
				, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
				, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
			)

			'objDB.blnDebug = True
			Call objDB.sbExecSQL(SQL, arrParams)
		End If
	Else
		ResultMSG = "수능점수 계산 실패"
	End if


	If Err.Number <> 0 Then 
		strResult = "Error"
		returnMSG = Err.Number&":"&Err.Description
	Else 
		strResult = "Complete"
		returnMSG = "수능 성적산출완료"
	End If
End Sub

'/////////////////////////////////////////////
'// 동석차 기준
'/////////////////////////////////////////////

Dim DrawRankingNum3, DrawRankingNum4, DrawRankingNum5

'1. 수시/면접 - 일반전형
'// 1순위 면접고사 성적 상위자
'// 2순위 교과 성적 평균등급 상위자
'// 3순위 선택한 학기 이수단위 합이 높은자
'// 4순위 학생부 성적 고학년 고학기 선택자
'// 5순위 연소자
Function UnqualifiedStandard_A(AryStudentNumber, AryInterviewerScore, AryStudentRecordAverage, AryCreditSum, AryChoiceSemester, AryMinor, count)		
	'For DrawRankingNum2 = 0 To count
		'Response.write "수험번호 : " & AryStudentNumber(DrawRankingNum2) & "<br>"
		'Response.write "면접점수 : " & AryInterviewerScore(DrawRankingNum2) & "<br>"
		'Response.write "평균등급 : " & AryStudentRecordAverage(DrawRankingNum2) & "<br>"
		'Response.write "이수합 : " & AryCreditSum(DrawRankingNum2) & "<br>"
		'Response.write "반영학기 : " & AryChoiceSemester(DrawRankingNum2) & "<br>"
		'Response.write "나이 : " & AryMinor(DrawRankingNum2) & "<br>"
		'Response.write "명수 : " & count & "<br>"
		'Response.write "<br>"
	'Next

	'// 1순위 면접고사 성적 상위자
	'1칸씩밖에 교체가 안 되므로, 동점자있을 경우 추가로 해야 내림차순이 됨
	For DrawRankingNum2 = 0 To count
		For StandardNum = 0 To count - 1
			For StandardNum2 = StandardNum + 1 To count
				If AryInterviewerScore(StandardNum) < AryInterviewerScore(StandardNum2) Then
					StandardNumTemp = AryInterviewerScore(StandardNum)
					AryInterviewerScore(StandardNum)	= AryInterviewerScore(StandardNum2)
					AryInterviewerScore(StandardNum2) = StandardNumTemp

					StudentRecordAverageTemp = AryStudentRecordAverage(StandardNum)
					AryStudentRecordAverage(StandardNum) = AryStudentRecordAverage(StandardNum2)
					AryStudentRecordAverage(StandardNum2) = StudentRecordAverageTemp

					CreditSumTemp = AryCreditSum(StandardNum)
					AryCreditSum(StandardNum) = AryCreditSum(StandardNum2)
					AryCreditSum(StandardNum2) = CreditSumTemp

					ChoiceSemesterTemp = AryChoiceSemester(StandardNum)
					AryChoiceSemester(StandardNum) = AryChoiceSemester(StandardNum2)
					AryChoiceSemester(StandardNum2) = ChoiceSemesterTemp

					MinorTemp = AryMinor(StandardNum)
					AryMinor(StandardNum) = AryMinor(StandardNum2)
					AryMinor(StandardNum2) = MinorTemp

					StudentNumberTemp = AryStudentNumber(StandardNum)
					AryStudentNumber(StandardNum) = AryStudentNumber(StandardNum2)
					AryStudentNumber(StandardNum2) = StudentNumberTemp

					DrawRankingTemp = DrawRanking(StandardNum)
					DrawRanking(StandardNum) = DrawRanking(StandardNum2)
					DrawRanking(StandardNum2) = DrawRankingTemp
				ElseIf AryInterviewerScore(StandardNum) = AryInterviewerScore(StandardNum2) Then
					DrawRanking(StandardNum) = 0
					DrawRanking(StandardNum2) = 0
				End If
			Next			
		Next
	Next

	'석차 정하기-동석차는 제외
	For DrawRankingNum3 = 0 To count		
		If DrawRanking(DrawRankingNum3) <> "0" Then
			DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
		End If

		'입력
	'	SQL = ""
	'	SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
	'	SQL = SQL & vbCrLf & "SET	 DrawStanding=?  "
	'	SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
	'	SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

		'Update일 때는 UPDT입력
	'	arrParams = Array(_
	'		  Array("@DrawStanding",			adInteger,		adParamInput,		0,		DrawRanking(DrawRankingNum3)) _

	'		, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
	'		, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
	'		, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		AryStudentNumber(DrawRankingNum3)) _
	'	)

		'objDB.blnDebug = True
	'	Call objDB.sbExecSQL(SQL, arrParams)

		'Response.write DrawRanking(DrawRankingNum3) & " / "
		'Response.write AryStudentNumber(DrawRankingNum3) & " / "
		'Response.write AryInterviewerScore(DrawRankingNum3) & " / "
		'Response.write AryStudentRecordAverage(DrawRankingNum3) & " / "
		'Response.write AryCreditSum(DrawRankingNum3) & " / "
		'Response.write AryChoiceSemester(DrawRankingNum3) & " / "
		'Response.write AryMinor(DrawRankingNum3) & " / "
		'Response.write "<br>"
	Next

	'1순위 석차 계산 후 동석차가 있는지 체크
	For DrawRankingNum = 0 To count
		If DrawRanking(DrawRankingNum) = 0 Then
			Draw = True
		Else
			Draw = false
		End If
	Next

	'1순위 석차 계산 후 동석차가 있으면
	If Draw Then
		For DrawRankingNum = 0 To count
			If DrawRanking(DrawRankingNum) = 0 Then
				'Response.write "수험번호 : " & AryStudentNumber(DrawRankingNum) & "<br>"
				'Response.write "면접점수 : " & AryInterviewerScore(DrawRankingNum) & "<br>"
				'Response.write "평균등급 : " & AryStudentRecordAverage(DrawRankingNum) & "<br>"
				'Response.write "이수합 : " & AryCreditSum(DrawRankingNum) & "<br>"
				'Response.write "반영학기 : " & AryChoiceSemester(DrawRankingNum) & "<br>"
				'Response.write "나이 : " & AryMinor(DrawRankingNum) & "<br>"
				'Response.write "명수 : " & count + 1 & "<br>"
				'Response.write "<br>"

				TxtStudentNumber = TxtStudentNumber & "," & AryStudentNumber(DrawRankingNum)
				TxtStudentRecordAverage = TxtStudentRecordAverage & "," & AryStudentRecordAverage(DrawRankingNum)
				TxtCreditSum = TxtCreditSum & "," & AryCreditSum(DrawRankingNum)
				TxtChoiceSemester = TxtChoiceSemester & "," & AryChoiceSemester(DrawRankingNum)
				TxtMinor = TxtMinor & "," & AryMinor(DrawRankingNum)
				TxtDrawRanking = TxtDrawRanking & "," & "99"

				
				'Response.write TxtStudentNumber & " / "
				'Response.write TxtStudentRecordAverage & " / "
				'Response.write TxtCreditSum & " / "
				'Response.write TxtChoiceSemester & " / "
				'Response.write TxtMinor & " / "
				'Response.write "<br>"
				'Response.write "<br>"
			End If
		Next

		'콤마지우기
		ArrTxtStudentNumber = Mid(TxtStudentNumber, 2, Len(TxtStudentNumber))
		ArrTxtxtStudentRecordAverage = Mid(TxtStudentRecordAverage, 2, Len(TxtStudentRecordAverage))
		ArrTxtxtCreditSum = Mid(TxtCreditSum, 2, Len(TxtCreditSum))
		ArrTxtxtChoiceSemester = Mid(TxtChoiceSemester, 2, Len(TxtChoiceSemester))
		ArrTxtxtMinor = Mid(TxtMinor, 2, Len(TxtMinor))
		ArrTxtDrawRanking = Mid(TxtDrawRanking, 2, Len(TxtDrawRanking))

		'Response.write ArrTxtStudentNumber & " / "
		'Response.write ArrTxtxtStudentRecordAverage & " / "
		'Response.write ArrTxtxtCreditSum & " / "
		'Response.write ArrTxtxtChoiceSemester & " / "
		'Response.write ArrTxtxtMinor & " / "
		'Response.write ArrTxtDrawRanking & " / "
		'Response.write "<br>"
		'Response.write "<br>"

		ArrTxtStudentNumber = Split(ArrTxtStudentNumber, ",")
		ArrTxtxtStudentRecordAverage = Split(ArrTxtxtStudentRecordAverage, ",")
		ArrTxtxtCreditSum = Split(ArrTxtxtCreditSum, ",")
		ArrTxtxtChoiceSemester = Split(ArrTxtxtChoiceSemester, ",")
		ArrTxtxtMinor = Split(ArrTxtxtMinor, ",")
		ArrTxtDrawRanking = Split(ArrTxtDrawRanking, ",")
		count2 = ubound(ArrTxtStudentNumber)

		'For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
		'	Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtStudentRecordAverage(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "
		'	Response.write count2 & " / "
		'	Response.write "<br>"
		'	Response.write "<br>"
		'Next

		'// 2순위 교과 성적 평균등급 상위자
		'1칸씩밖에 교체가 안 되므로, 동점자있을 경우 추가로 해야 내림차순이 됨
		For DrawRankingNum2 = 0 To count2
			For StandardNum = 0 To count2 - 1
				For StandardNum2 = StandardNum + 1 To count2
					If ArrTxtxtStudentRecordAverage(StandardNum) < ArrTxtxtStudentRecordAverage(StandardNum2) Then
						StudentRecordAverageTemp = ArrTxtxtStudentRecordAverage(StandardNum)
						ArrTxtxtStudentRecordAverage(StandardNum) = ArrTxtxtStudentRecordAverage(StandardNum2)
						ArrTxtxtStudentRecordAverage(StandardNum2) = StudentRecordAverageTemp

						CreditSumTemp = ArrTxtxtCreditSum(StandardNum)
						ArrTxtxtCreditSum(StandardNum) = ArrTxtxtCreditSum(StandardNum2)
						ArrTxtxtCreditSum(StandardNum2) = CreditSumTemp

						ChoiceSemesterTemp = ArrTxtxtChoiceSemester(StandardNum)
						ArrTxtxtChoiceSemester(StandardNum) = ArrTxtxtChoiceSemester(StandardNum2)
						ArrTxtxtChoiceSemester(StandardNum2) = ChoiceSemesterTemp

						MinorTemp = ArrTxtxtMinor(StandardNum)
						ArrTxtxtMinor(StandardNum) = ArrTxtxtMinor(StandardNum2)
						ArrTxtxtMinor(StandardNum2) = MinorTemp

						StudentNumberTemp = ArrTxtStudentNumber(StandardNum)
						ArrTxtStudentNumber(StandardNum) = ArrTxtStudentNumber(StandardNum2)
						ArrTxtStudentNumber(StandardNum2) = StudentNumberTemp

						DrawRankingTemp = ArrTxtDrawRanking(StandardNum)
						ArrTxtDrawRanking(StandardNum) = ArrTxtDrawRanking(StandardNum2)
						ArrTxtDrawRanking(StandardNum2) = DrawRankingTemp
					ElseIf ArrTxtxtStudentRecordAverage(StandardNum) = ArrTxtxtStudentRecordAverage(StandardNum2) Then
						ArrTxtDrawRanking(StandardNum) = 0
						ArrTxtDrawRanking(StandardNum2) = 0
					End If
				Next			
			Next
		Next

		'For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
		'	Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtStudentRecordAverage(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "
		'	Response.write count2 & " / "
		'	Response.write "<br>"
		'	Response.write "<br>"
		'Next

		'테스트용
		'DrawRanking(0) = 0
		'DrawRanking(2) = 3

		'석차 정하기-동석차는 제외
		For DrawRankingNum3 = 0 To count2		
			If ArrTxtDrawRanking(DrawRankingNum3) <> "0" Then
				'전체 동점자가 2명이었을 때
				'1,2등
				If count = 1 Then
					ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
					DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
				End If

				'전체 동점자가 3명이었을 때
				'석차 1이 있으면 2,3등
				'석차 3이 있으면 1,2등
				'없으면 1,2,3등
				If count = 2 Then
					For DrawRankingNum4 = 0 To count
						If ArrTxtDrawRanking(DrawRankingNum3) = "99" And OldDrawRanking = "" Then
							If DrawRanking(DrawRankingNum4) = 1 Then
								OldDrawRanking = 1
							ElseIf DrawRanking(DrawRankingNum4) = 3 Then
								OldDrawRanking = 3
							Else
								OldDrawRanking = 3							
							End If
						End If
					Next
					If OldDrawRanking = 1 Then
						ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 2
						DrawRanking(DrawRankingNum3+1) = DrawRankingNum3 + 2
					ElseIf OldDrawRanking = 3 Then
						ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
						DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
					End If
				End If

				'전체 동점자가 4명이었을 때
				'석차 1,4등 이면 2,3등
				'석차 1,2등 이면 3,4등
				'석차 3,4등 이면 1,2등
				'석차 1만 있으면 2,3,4등
				'석차 4만 있으면 1,2,3등
				'없으면 1,2,3,4등
				If count = 3 Then
					For DrawRankingNum4 = 0 To count
						If ArrTxtDrawRanking(DrawRankingNum3) = "99" And OldDrawRanking = "" Then

						End If
					Next
				End If

				'Response.write "<br>"
				'Response.write ArrTxtDrawRanking(DrawRankingNum3) & " / "
				'Response.write "<br>"
			End If
		Next

		'For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
		'	Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtStudentRecordAverage(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "	
		'	Response.write DrawRanking(DrawRankingNum) & " / "
		'	Response.write count2 & " / "
		'	Response.write "<br>"
		'	Response.write "<br>"
		'Next
		
		'For DrawRankingNum = 0 To count2
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "
		'	Response.write "<br>"
		'Next

		'2순위 석차 계산 후 동석차가 있는지 체크
		For DrawRankingNum = 0 To count2
			If ArrTxtDrawRanking(DrawRankingNum) = 0 Then
				Draw = True
			Else
				Draw = false
			End If
		Next
	End If

	'테스트용
	Draw = True
	DrawRanking(0) = "1"
	DrawRanking(1) = "0"
	DrawRanking(2) = "0"

	'2순위 석차 계산 후 동석차가 있으면
	If Draw Then
		'변수초기화
		OldDrawRanking = ""

		TxtStudentNumber = ""
		TxtCreditSum = ""
		TxtChoiceSemester = ""
		TxtMinor = ""
		TxtDrawRanking = ""

		For DrawRankingNum = 0 To count
			If DrawRanking(DrawRankingNum) = 0 Then
				'Response.write "수험번호 : " & AryStudentNumber(DrawRankingNum) & "<br>"
				'Response.write "이수합 : " & AryCreditSum(DrawRankingNum) & "<br>"
				'Response.write "반영학기 : " & AryChoiceSemester(DrawRankingNum) & "<br>"
				'Response.write "나이 : " & AryMinor(DrawRankingNum) & "<br>"
				'Response.write "명수 : " & count + 1 & "<br>"
				'Response.write "<br>"

				TxtStudentNumber = TxtStudentNumber & "," & AryStudentNumber(DrawRankingNum)
				TxtStudentRecordAverage = TxtStudentRecordAverage & "," & AryStudentRecordAverage(DrawRankingNum)
				TxtCreditSum = TxtCreditSum & "," & AryCreditSum(DrawRankingNum)
				TxtChoiceSemester = TxtChoiceSemester & "," & AryChoiceSemester(DrawRankingNum)
				TxtMinor = TxtMinor & "," & AryMinor(DrawRankingNum)
				TxtDrawRanking = TxtDrawRanking & "," & "99"

				
				'Response.write TxtStudentNumber & " / "
				'Response.write TxtCreditSum & " / "
				'Response.write TxtChoiceSemester & " / "
				'Response.write TxtMinor & " / "
				'Response.write "<br>"
				'Response.write "<br>"
			End If
		Next		

		'콤마지우기
		ArrTxtStudentNumber = Mid(TxtStudentNumber, 2, Len(TxtStudentNumber))
		ArrTxtxtCreditSum = Mid(TxtCreditSum, 2, Len(TxtCreditSum))
		ArrTxtxtChoiceSemester = Mid(TxtChoiceSemester, 2, Len(TxtChoiceSemester))
		ArrTxtxtMinor = Mid(TxtMinor, 2, Len(TxtMinor))
		ArrTxtDrawRanking = Mid(TxtDrawRanking, 2, Len(TxtDrawRanking))

		'Response.write ArrTxtStudentNumber & " / "
		'Response.write ArrTxtxtCreditSum & " / "
		'Response.write ArrTxtxtChoiceSemester & " / "
		'Response.write ArrTxtxtMinor & " / "
		'Response.write ArrTxtDrawRanking & " / "
		'Response.write "<br>"
		'Response.write "<br>"

		ArrTxtStudentNumber = Split(ArrTxtStudentNumber, ",")
		ArrTxtxtCreditSum = Split(ArrTxtxtCreditSum, ",")
		ArrTxtxtChoiceSemester = Split(ArrTxtxtChoiceSemester, ",")
		ArrTxtxtMinor = Split(ArrTxtxtMinor, ",")
		ArrTxtDrawRanking = Split(ArrTxtDrawRanking, ",")
		count2 = ubound(ArrTxtStudentNumber)

		'For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
		'	Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "
		'	Response.write "count : " & count2 & " / "
		'	Response.write "<br>"
		'	Response.write "<br>"
		'Next

		'//3순위 선택한 학기 이수단위 합이 높은자
		'1칸씩밖에 교체가 안 되므로, 동점자있을 경우 추가로 해야 내림차순이 됨
		For DrawRankingNum2 = 0 To count2
			For StandardNum = 0 To count2 - 1
				For StandardNum2 = StandardNum + 1 To count2
					If ArrTxtxtCreditSum(StandardNum) < ArrTxtxtCreditSum(StandardNum2) Then
						CreditSumTemp = ArrTxtxtCreditSum(StandardNum)
						ArrTxtxtCreditSum(StandardNum) = ArrTxtxtCreditSum(StandardNum2)
						ArrTxtxtCreditSum(StandardNum2) = CreditSumTemp

						ChoiceSemesterTemp = ArrTxtxtChoiceSemester(StandardNum)
						ArrTxtxtChoiceSemester(StandardNum) = ArrTxtxtChoiceSemester(StandardNum2)
						ArrTxtxtChoiceSemester(StandardNum2) = ChoiceSemesterTemp

						MinorTemp = ArrTxtxtMinor(StandardNum)
						ArrTxtxtMinor(StandardNum) = ArrTxtxtMinor(StandardNum2)
						ArrTxtxtMinor(StandardNum2) = MinorTemp

						StudentNumberTemp = ArrTxtStudentNumber(StandardNum)
						ArrTxtStudentNumber(StandardNum) = ArrTxtStudentNumber(StandardNum2)
						ArrTxtStudentNumber(StandardNum2) = StudentNumberTemp

						DrawRankingTemp = ArrTxtDrawRanking(StandardNum)
						ArrTxtDrawRanking(StandardNum) = ArrTxtDrawRanking(StandardNum2)
						ArrTxtDrawRanking(StandardNum2) = DrawRankingTemp
					ElseIf ArrTxtxtCreditSum(StandardNum) = ArrTxtxtCreditSum(StandardNum2) Then
						ArrTxtDrawRanking(StandardNum) = 0
						ArrTxtDrawRanking(StandardNum2) = 0
					End If
				Next			
			Next
		Next

		'For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
		'	Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtStudentRecordAverage(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
		'	Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
		'	Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "
		'	Response.write count2 & " / "
		'	Response.write "<br>"
		'	Response.write "<br>"
		'Next

		'석차 정하기-동석차는 제외
		For DrawRankingNum3 = 0 To count2		
			If ArrTxtDrawRanking(DrawRankingNum3) <> "0" Then
				'전체 동점자가 2명이었을 때
				'1,2등
				If count = 1 Then
					ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
					DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
				End If

				'전체 동점자가 3명이었을 때
				'석차 1이 있으면 2,3등
				'석차 3이 있으면 1,2등
				'없으면 1,2,3등
				If count = 2 Then
					For DrawRankingNum4 = 0 To count
						If ArrTxtDrawRanking(DrawRankingNum3) = "99" And OldDrawRanking = "" Then
							If DrawRanking(DrawRankingNum4) = 1 Then
								OldDrawRanking = 1
							ElseIf DrawRanking(DrawRankingNum4) = 3 Then
								OldDrawRanking = 3
							Else
								OldDrawRanking = 3							
							End If
						End If
					Next
					If OldDrawRanking = 1 Then
						ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 2
						DrawRanking(DrawRankingNum3+1) = DrawRankingNum3 + 2
					ElseIf OldDrawRanking = 3 Then
						ArrTxtDrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
						DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
					End If
				End If

				'전체 동점자가 4명이었을 때
				'석차 1,4등 이면 2,3등
				'석차 1,2등 이면 3,4등
				'석차 3,4등 이면 1,2등
				'석차 1만 있으면 2,3,4등
				'석차 4만 있으면 1,2,3등
				'없으면 1,2,3,4등
				If count = 3 Then
					For DrawRankingNum4 = 0 To count
						If ArrTxtDrawRanking(DrawRankingNum3) = "99" And OldDrawRanking = "" Then

						End If
					Next
				End If

				'Response.write "<br>"
				'Response.write ArrTxtDrawRanking(DrawRankingNum3) & " / "
				'Response.write "<br>"
			End If
		Next

		For DrawRankingNum = 0 To ubound(ArrTxtStudentNumber)
			Response.write ArrTxtStudentNumber(DrawRankingNum) & " / "
			Response.write ArrTxtxtStudentRecordAverage(DrawRankingNum) & " / "
			Response.write ArrTxtxtCreditSum(DrawRankingNum) & " / "
			Response.write ArrTxtxtChoiceSemester(DrawRankingNum) & " / "
			Response.write ArrTxtxtMinor(DrawRankingNum) & " / "
			Response.write ArrTxtDrawRanking(DrawRankingNum) & " / "	
			Response.write DrawRanking(DrawRankingNum) & " / "
			Response.write count2 & " / "
			Response.write "<br>"
			Response.write "<br>"
		Next
		
		For DrawRankingNum = 0 To count
			Response.write DrawRanking(DrawRankingNum) & " / "
			Response.write "<br>"
		Next

		'3순위 석차 계산 후 동석차가 있는지 체크
		For DrawRankingNum = 0 To count2
			If ArrTxtDrawRanking(DrawRankingNum) = 0 Then
				Draw = True
			Else
				Draw = false
			End If
		Next
	End If

End Function

'3. 수시 - 전문대졸이상졸업자전형
'// 1순위 대학 백분위 성적 우수자 - 성적이 같은 자가 동점이므로 생략
'// 2순위 졸업(취득) 학점이 많은 자
Function UnqualifiedStandard_C(StudentNumber, Score, count)

	'Response.write "<br>"

	'// 2순위 졸업(취득) 학점이 많은 자
	'1칸씩밖에 교체가 안 되므로, 동점자있을 경우 추가로 해야 내림차순이 됨
	For DrawRankingNum2 = 0 To count
		For StandardNum = 0 To count - 1
			For StandardNum2 = StandardNum + 1 To count
				If Score(StandardNum) < Score(StandardNum2) Then
					StandardNumTemp = Score(StandardNum)
					Score(StandardNum)	= Score(StandardNum2)
					Score(StandardNum2) = StandardNumTemp

					StudentNumberTemp = StudentNumber(StandardNum)
					StudentNumber(StandardNum) = StudentNumber(StandardNum2)
					StudentNumber(StandardNum2) = StudentNumberTemp

					DrawRankingTemp = DrawRanking(StandardNum)
					DrawRanking(StandardNum) = DrawRanking(StandardNum2)
					DrawRanking(StandardNum2) = DrawRankingTemp
				ElseIf Score(StandardNum) = Score(StandardNum2) Then
					DrawRanking(StandardNum) = 0
					DrawRanking(StandardNum2) = 0
				End If
			Next			
		Next
	Next

	'석차 정하기-동석차는 제외
	For DrawRankingNum3 = 0 To count		
		If DrawRanking(DrawRankingNum3) <> "0" Then
			DrawRanking(DrawRankingNum3) = DrawRankingNum3 + 1
		End If
	
		'입력
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE ChangeScoreTable  "
		SQL = SQL & vbCrLf & "SET	 DrawStanding=?  "
		SQL = SQL & vbCrLf & "		 , UPDT_USID=?, UPDT_DATE=getdate(), UPDT_ADDR=? "
		SQL = SQL & vbCrLf & "WHERE StudentNumber=? "

		'Update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@DrawStanding",			adInteger,		adParamInput,		0,		DrawRanking(DrawRankingNum3)) _

			, Array("@UPDT_USID",				adVarchar,		adParamInput,		20,		SessionUserID) _
			, Array("@UPDT_ADDR",				adVarchar,		adParamInput,		20,		ASP_USER_IP) _
			, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber(DrawRankingNum3)) _
		)

		'objDB.blnDebug = True
		Call objDB.sbExecSQL(SQL, arrParams)
			

		'Response.write DrawRanking(DrawRankingNum3) & " / "
		'Response.write StudentNumber(DrawRankingNum3) & " / "
		'Response.write Score(DrawRankingNum3) & " / "
		'Response.write "<br>"
	Next
	
	'동석차가 있을 경우, True 리턴
	For DrawRankingNum = 0 To count
		If DrawRanking(DrawRankingNum) = 0 Then
			UnqualifiedStandard_Up_Temp = True
		Else
			UnqualifiedStandard_Up_Temp = false
		End If
	Next
	UnqualifiedStandard_C = UnqualifiedStandard_Up_Temp

End Function

'/////////////////////////////////////////////
'// 표준분포표 계산
'/////////////////////////////////////////////

'1. 표준분포표 계산
Function NORMDIST(X, MEAN, STD, CUMULATIVE)
	If CUMULATIVE = 1 Then
		NORMDIST = PHI_2( X, MEAN, STD )
	End If
End Function

'2. 표준분포표 계산
Dim ZTemp
Function PHI_2(Z, MU, SIGMA)
	ScoreDim = "Dim executeTemp : executeTemp = ((" & Z & "-" & MU & ") / " & SIGMA & ")"
	execute(ScoreDim) 
	PHI_2 = PHI_1(executeTemp)
End Function

'3. 표준분포표 계산
Dim ERFTemp, ERFTemp2
Function PHI_1(Z)
	ERFTemp = Sqr(2.0)
	ScoreDim = "Dim executeTemp : executeTemp = " & Z & "/" & ERFTemp
	execute(ScoreDim) 
	ERFTemp2 = ERF(executeTemp)
	ScoreDim = "Dim executeTemp : executeTemp = 0.5 * (1.0 + " & ERFTemp2 & ")"
	execute(ScoreDim) 
	PHI_1 = executeTemp
End Function

'4. 표준분포표 계산
Dim TD
Function ERF(Z)
	TD = 1.0 / (1.0 + 0.5 * ABS(Z))
	ERF = 1 - TD * EXP( -Z * Z   -  1.26551223 + TD * ( 1.00002368 + TD * ( 0.37409196 + TD * ( 0.09678418 + TD * (-0.18628806 + TD * ( 0.27886807 + TD * (-1.13520398 + TD * ( 1.48851587 + TD * (-0.82215223 + TD * ( 0.17087277)	) ) ) )	) )	) ) )
End Function

'/////////////////////////////////////////////
'// 등급 & 점수로 변환
'/////////////////////////////////////////////

'1. 백분율로 등급 만들기
Function PercentageGrade(Score)
	If (0.00			<= Score and 4.00 >= Score) Then PercentageGrade = 1 End if
	If (4.000000000001	<= Score and 11.000	>= Score) Then PercentageGrade = 2 End if
	If (11.000000000001	<= Score and 23.000	>= Score) Then PercentageGrade = 3 End if
	If (23.000000000001	<= Score and 40.000	>= Score) Then PercentageGrade = 4 End if
	If (40.000000000001	<= Score and 60.000	>= Score) Then PercentageGrade = 5 End if
	If (60.000000000001	<= Score and 77.000	>= Score) Then PercentageGrade = 6 End if
	If (77.000000000001	<= Score and 89.000	>= Score) Then PercentageGrade = 7 End if
	If (89.000000000001	<= Score and 96.000	>= Score) Then PercentageGrade = 8 End if
	If (96.000000000001	<= Score and 100.000 >= Score) Then PercentageGrade = 9 End If
End Function

'2. 검정고시 점수로 등급 만들기
Function QualificationGrade(Score)
	If (00.00 <= Score and 61.99 >= Score) Then QualificationGrade = 9 End if
	If (62.00 <= Score and 64.99 >= Score) Then QualificationGrade = 8 End if
	If (65.00 <= Score and 69.99 >= Score) Then QualificationGrade = 7 End if
	If (70.00 <= Score and 76.99 >= Score) Then QualificationGrade = 6 End if
	If (77.00 <= Score and 84.99 >= Score) Then QualificationGrade = 5 End if
	If (85.00 <= Score and 90.99 >= Score) Then QualificationGrade = 4 End if
	If (91.00 <= Score and 95.99 >= Score) Then QualificationGrade = 3 End if
	If (96.00 <= Score and 98.99 >= Score) Then QualificationGrade = 2 End if
	If (99.00 <= Score and 100.0 >= Score) Then QualificationGrade = 1 End If
End Function

'3. 수능 영어등급 점수로 바꾸기
Function CSATGrad(Score)
	If 1 = Score Then CSATGrad = 95 End if
	If 2 = Score Then CSATGrad = 85 End if
	If 3 = Score Then CSATGrad = 75 End if
	If 4 = Score Then CSATGrad = 65 End if
	If 5 = Score Then CSATGrad = 55 End if
	If 6 = Score Then CSATGrad = 45 End if
	If 7 = Score Then CSATGrad = 35 End if
	If 8 = Score Then CSATGrad = 25 End if
	If 9 = Score Then CSATGrad = 10 End If
End Function
%>

			</div>
			<!-- 공식입력란 끝-->	
		</div>		
	</div>
</div>
<!-- #InClude Virtual = "/Common/Bottom.asp" -->