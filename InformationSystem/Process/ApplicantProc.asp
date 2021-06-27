<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim process					: process = fnR("process", "")
Dim ProcessType				: ProcessType = fnR("ProcessType", "")
Dim LogDivision				: LogDivision = "ApplicantProc"

'구분값(체크박스 복수선택-사용안 함)
'Dim StudentNumberHidden		: StudentNumberHidden = fnRF("StudentNumberHidden")
'Dim MyearHidden				: MyearHidden = fnRF("MyearHidden")

'구분값2(개별선택)
Dim StudentNumber			: StudentNumber = fnRF("StudentNumber")
Dim Myear					: Myear = fnRF("Myear")

'가산점, 면접, 실기 점수
Dim ExtraPoint				: ExtraPoint = fnRF("ExtraPoint")
Dim InterviewerPoint		: InterviewerPoint = fnR("Interviewer", "0")
Dim PracticalPoint			: PracticalPoint = fnR("Practical", "0")

'면접, 실기 평가비율
Dim InterviewerRatio		: InterviewerRatio = fnR("InterviewerRatio", "0")
Dim PracticalRatio			: PracticalRatio = fnR("PracticalRatio", "0")

'생기부, 검정, 수능 동의
Dim StudentRecord			: StudentRecord = fnRF("StudentRecord")
Dim Qualification			: Qualification = fnRF("Qualification")
Dim CSAT					: CSAT = fnRF("CSAT")

'필수서류
Dim document1				: document1 = fnR("document1", "0")
Dim document2				: document2 = fnR("document2", "0")
Dim document3				: document3 = fnR("document3", "0")
Dim document4				: document4 = fnR("document4", "0")
Dim document5				: document5 = fnR("document5", "0")
Dim document6				: document6 = fnR("document6", "0")
Dim document7				: document7 = fnR("document7", "0")
Dim document8				: document8 = fnR("document8", "0")
Dim document21				: document21 = fnR("document21", "0")
Dim document22				: document22 = fnR("document22", "0")
Dim document23				: document23 = fnR("document23", "0")
Dim document24				: document24 = fnR("document24", "0")

'필수서류 체크
Dim documentCheck1			
Dim documentCheck2			
Dim documentCheck3			
Dim documentCheck4			
Dim documentCheck5			
Dim documentCheck6			
Dim documentCheck7			
Dim documentCheck8			
Dim documentCheck21			
Dim documentCheck22			
Dim documentCheck23			
Dim documentCheck24	

'자격미달 
Dim DrawStandard
Dim DrawMsg

'필수서류
Dim Document : Document = "N"
Dim DocumentMsg

'입력, 수정
Dim INPT_USID    			: INPT_USID = SessionUserID
Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
Dim UPDT_USID    			: UPDT_USID = SessionUserID
Dim UPDT_ADDR    			: UPDT_ADDR = ASP_USER_IP

'히스토리 카운터
Dim UpdateCnt

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim intNUM

'hidden 값을 하나씩 풀어 배열로 저장(체크박스에 선택된 리스트-사용 안 함)
'Dim StudentNumber			: StudentNumber = Split(StudentNumberHidden, ",")
'Dim Myear					: Myear = Split(MyearHidden, ",")

Dim i, StudnetNumberTemp, MYearTemp
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, strLogMSG2

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'체크박스가 선택된 갯수만큼 반복(사용 안 함)
'For i = 0 To Ubound(StudentNumber)
'	StudnetNumberTemp = StudentNumber(i)
'	MYearTemp = Myear(i)

	'///////////////////////////////////////////////
	'// 필수서류 체크
	'///////////////////////////////////////////////
	SQL = ""
	SQL = SQL & vbCrLf & "Select DocumentaryCheck1, DocumentaryCheck2, DocumentaryCheck3, DocumentaryCheck4, DocumentaryCheck5 "
	SQL = SQL & vbCrLf & "		,DocumentaryCheck6, DocumentaryCheck7, DocumentaryCheck8 "
	SQL = SQL & vbCrLf & "		,DocumentaryCheck21, DocumentaryCheck22, DocumentaryCheck23, DocumentaryCheck24 "
	SQL = SQL & vbCrLf & "from ApplicationTable "
	SQL = SQL & vbCrLf & "where Myear = '" & Myear & "' "
	SQL = SQL & vbCrLf & "And StudentNumber = '" & StudentNumber & "'; "

	'objDB.blnDebug = TRUE
	arrParams = objDB.fnGetArray
	AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)				

	If isArray(AryHash) Then
		documentCheck1				= AryHash(0).Item("DocumentaryCheck1")
		documentCheck2				= AryHash(0).Item("DocumentaryCheck2")
		documentCheck3				= AryHash(0).Item("DocumentaryCheck3")
		documentCheck4				= AryHash(0).Item("DocumentaryCheck4")
		documentCheck5				= AryHash(0).Item("DocumentaryCheck5")
		documentCheck6				= AryHash(0).Item("DocumentaryCheck6")
		documentCheck7				= AryHash(0).Item("DocumentaryCheck7")
		documentCheck8				= AryHash(0).Item("DocumentaryCheck8")
		documentCheck21				= AryHash(0).Item("DocumentaryCheck21")
		documentCheck22				= AryHash(0).Item("DocumentaryCheck22")
		documentCheck23				= AryHash(0).Item("DocumentaryCheck23")
		documentCheck24				= AryHash(0).Item("DocumentaryCheck24")
	End If

	'=============== 자격미달여부 계산 (하드코딩)============================================================
	'====== 하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save, 4. 수능점수 엑셀 Save =====

	'//////////////////////////////////////////////////////////////////////////////////////////////////////
	'//자격미달자 기준별 계산식 (하드코딩) 1 ~ 8번 코드 (Y = 미달자)
	'//하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save, 4. 수능점수 엑셀 Save
	'//***** 1번코드 (수시용)입학원서 등록 시 체크
	'//***** 2번코드 (정시용)입학원서 등록 시 체크, 수능성적 입력 시 체크 -> 자격미달여부 다시 체크
	'//***** 3번 ~ 8번코드 지원자관리에서 관리자가 직접 서류체크 및 점수 입력 -> 자격미달여부 다시 체크
	'//////////////////////////////////////////////////////////////////////////////////////////////////////

	'////////////////////////////////////////////////////////////////////////////////////////////////
	'//1번코드 (공통-수시)국내 고등학교(검정고시포함) 졸업(예정) 학력 소유
	'//2번코드 (공통-정시)국내 고등학교 졸업(예정)자로 수학능력시험 성적이 있는 자
	'//3번코드 (공통-면접)면접 미응시자(실기포함)
	'//4번코드 (농어촌-1유형)농어촌지역 거주 및 고등학교 졸업(예정)자
	'//5번코드 (농어촌-2유형)12년 동안 연속으로 농어촌지역 학교에 재학한 고등학교 졸업(예정)자
	'//6번코드 (기초생활수급자)지원자 명의의 수급자 증명서 발급 가능
	'//7번코드 (차상위) 증명서 중 1개이상 발급 가능
	'//8번코드 (전문대이상졸업자)4년제 대학 2년이상 수료자, 전문대학 졸업자(예정자 지원 불가)
	'////////////////////////////////////////////////////////////////////////////////////////////////

	'1번 코드
	If documentCheck1 = "3" Then
		If document1 = "1" Then
			DrawStandard = "N"
		Else
			DrawStandard = "Y"
			document1 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>외국고교 졸업자(제출서류) :</b> 국내고교 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
			Else
				DrawMsg = DrawMsg & "= <b>외국고교 졸업자(제출서류) :</b> 국내고교 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
			End If
		End If		
	ElseIf documentCheck1 = "4" Then
		If document1 = "1" Then
			DrawStandard = "N"
		Else
			DrawStandard = "C"
			document1 = "4"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "등록된 고교코드가 없습니다. 국내 모든 고등학교를 고교코드에 입력하여 주세요."
			Else
				DrawMsg = DrawMsg & "= 등록된 고교코드가 없습니다. 국내 모든 고등학교를 고교코드에 입력하여 주세요."
			End If
		End If
	ElseIf documentCheck1 = "5" Then
		If StudentRecord = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document1 = "1"
		Else
			If document1 = "1" Then
				DrawStandard = "N"
			Else
				DrawStandard = "Y"
				document1 = "5"
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
				Else
					DrawMsg = DrawMsg & "= <b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
				End If
			End If	
		End If
	ElseIf documentCheck1 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document1 = "1"
	End If

	'2번 코드
	If documentCheck2 = "3" Then
		If StudentRecord = "1" And CSAT = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document2 = "0"
		ElseIf StudentRecord = "1" Then
			DrawStandard = "C"
			document2 = "10"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
			Else
				DrawMsg = DrawMsg & "= <b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
			End If
		ElseIf CSAT = "1" Then
			If document2 = "1" Then
				If DrawStandard <> "Y" And DrawStandard <> "C" Then
					DrawStandard = "N"
				End If
				document2 = "1"
			Else
				DrawStandard = "Y"
				document2 = "3"
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
				Else
					DrawMsg = DrawMsg & "= <b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
				End If
			End If
		Else 
			If document2 = "1" Then
				DrawStandard = "C"
				document2 = "10"
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
				Else
					DrawMsg = DrawMsg & "= <b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
				End If
			Else
				DrawStandard = "Y"
				document2 = "3"
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>학력/수능성적 자격미달(제출서류) :</b> 고등학교 생활기록부, 수능성적 업로드 필요(미동의한 지원자의 경우 수동입력)"
				Else
					DrawMsg = DrawMsg & "= <b>학력/수능성적 자격미달(제출서류) :</b> 고등학교 생활기록부, 수능성적 업로드 필요(미동의한 지원자의 경우 수동입력)"
				End If
			End If
		End If
	ElseIf documentCheck2 = "10" Then
		If CSAT = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document2 = "0"
		Else
			DrawStandard = "C"
			document2 = "10"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
			Else
				DrawMsg = DrawMsg & "= <b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
			End If
		End If
	ElseIf documentCheck2 = "5" Then
		If StudentRecord = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document2 = "0"
		Else
			If document2 = "1" Then
				If DrawStandard <> "Y" And DrawStandard <> "C" Then
					DrawStandard = "N"
				End If
				document2 = "1"
			Else
				DrawStandard = "Y"
				document2 = "3"
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
				Else
					DrawMsg = DrawMsg & "= <b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
				End If
			End If
		End If
	ElseIf documentCheck2 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document2 = "1"
	End IF

	'4번 코드
	If documentCheck4 = "3" Then
		If document4 = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document4 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>농어촌1유형 자격미달(제출서류) :</b> 농어촌전형 추천서, 중/고등학교 생활기록부, 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
			Else
				DrawMsg = DrawMsg & "= <b>농어촌1유형 자격미달(제출서류) :</b> 농어촌전형 추천서, 중/고등학교 생활기록부, 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
			End If
		End If
	ElseIf documentCheck4 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document4 = "1"
	End If

	'5번 코드
	If documentCheck5 = "3" Then
		If document5 = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document5 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>농어촌2유형 자격미달(제출서류) :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
			Else
				DrawMsg = DrawMsg & "= <b>농어촌2유형 자격미달(제출서류) :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
			End If
		End If
	ElseIf documentCheck5 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document5 = "1"
	End If

	'6번 코드
	If documentCheck6 = "3" Then
		If document6 = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document6 = "3"
			If StudentRecord = "1" Then
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>기초수급자 자격미달(제출서류) :</b> 지원자 명의의 수급자 증명서"
				Else
					DrawMsg = DrawMsg & "= <b>기초수급자 자격미달(제출서류) :</b> 지원자 명의의 수급자 증명서"
				End If
			Else
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>기초수급자 자격미달(제출서류) :</b> 고등학교 생활기록부. 또는 검정고시 합격증명서, 성적증명서, 지원자 명의의 수급자 증명서"
				Else
					DrawMsg = DrawMsg & "= <b>기초수급자 자격미달(제출서류) :</b> 고등학교 생활기록부. 또는 검정고시 합격증명서, 성적증명서, 지원자 명의의 수급자 증명서"
				End If
			End IF
		End If
	ElseIf documentCheck6 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document6 = "1"
	End If

	'7번 코드
	If documentCheck7 = "3" Then
		If document7 = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document7 = "3"
			If StudentRecord = "1" Then
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>차상위계층 자격미달(제출서류) :</b> 장애수당, 장애인연금, 자활근로자, 한부모가족, 우선돌봄차상위, 차상위본인부담경감 중 1부"
				Else
					DrawMsg = DrawMsg & "= <b>차상위계층 자격미달(제출서류) :</b> 장애수당, 장애인연금, 자활근로자, 한부모가족, 우선돌봄차상위, 차상위본인부담경감 중 1부"
				End If
			Else
				If LEN(DrawMsg) < 1 Then
					DrawMsg = "<b>차상위계층 자격미달(제출서류) :</b> 고교 생기부 or 검정고시 합격증명서, 성적증명서, (장애수당, 장애인연금, 자활근로자, 한부모가족, 우선돌봄차상위, 차상위본인부담경감 중 1부)"
				Else
					DrawMsg = DrawMsg & "= <b>차상위계층 자격미달(제출서류) :</b> 고교 생기부 or 검정고시 합격증명서, 성적증명서, (장애수당, 장애인연금, 자활근로자, 한부모가족, 우선돌봄차상위, 차상위본인부담경감 중 1부)"
				End If
			End If
		End If
	ElseIf documentCheck7 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document7 = "1"
	End If

	'8번 코드
	If documentCheck8 = "3" Then
		If document8 = "1" Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document8 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>대학졸업자 자격미달(제출서류) :</b> 전적 대학 졸업(수료)증명서 1부. 및  성적증명서 1부"
			Else
				DrawMsg = DrawMsg & "= <b>대학졸업자 자격미달(제출서류) :</b> 전적 대학 졸업(수료)증명서 1부. 및  성적증명서 1부"
			End If
		End If
	ElseIf documentCheck8 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document8 = "1"
	End If

	'3번 코드 (C외에 다른 코드들도 쓰므로 맨 아래서 체크)
	If documentCheck3 = "3" Then
		If InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint > 0 And PracticalPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
				document3 = "1"
			End If
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint = 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "D"
			End If
			document3 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			End If	
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "E"
			End If
			document3 = "4"
				If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			End If
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "F"
			End If
			document3 = "5"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "G"
			End If
			document3 = "6"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			End If
		End If
	ElseIf documentCheck3 = "4" Then
		If InterviewerRatio > 0 And InterviewerPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
				document3 = "1"
			End If
		ElseIf InterviewerRatio > 0 And InterviewerPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "E"
			End If
			document3 = "4"
				If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "G"
			End If
			document3 = "6"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접이 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접이 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			End If
		End If
	ElseIf documentCheck3 = "5" Then
		If PracticalRatio > 0 And PracticalPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
				document3 = "1"
			End If
		ElseIf PracticalRatio > 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "F"
			End If
			document3 = "5"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "G"
			End If
			document3 = "6"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			Else
				DrawMsg = DrawMsg & "= <b>실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			End If
		End If
	ElseIf documentCheck3 = "6" Then
		If InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint > 0 And PracticalPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
				document3 = "1"
			End If
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint = 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "D"
			End If
			document3 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			End If	
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And InterviewerPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "E"
			End If
			document3 = "4"
				If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			End If
		ElseIf InterviewerRatio > 0 And PracticalRatio > 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "F"
			End If
			document3 = "5"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			End If
		ElseIf PracticalRatio > 0 And PracticalPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document3 = "1"
		ElseIf InterviewerRatio > 0 And InterviewerPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "E"
			End If
			document3 = "4"
				If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			End If
		ElseIf InterviewerRatio > 0 And InterviewerPoint > 0 Then
			If DrawStandard <> "Y" And DrawStandard <> "C" Then
				DrawStandard = "N"
			End If
			document3 = "1"	
		ElseIf PracticalRatio > 0 And PracticalPoint = 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "F"
			End If
			document3 = "5"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
			End If
		Else
			If DrawStandard <> "Y" Then
				DrawStandard = "G"
			End If
			document3 = "6"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
			End If
		End If
	ElseIf documentCheck3 = "1" Then
		If DrawStandard <> "Y" And DrawStandard <> "C" Then
			DrawStandard = "N"
		End If
		document3 = "1"
	End If

	'Y, N, C가 아니면 자격미달기준이 없는 것이므로 자격미달 아님(N)
	If DrawStandard <> "Y" And DrawStandard <> "C" And DrawStandard <> "D" And DrawStandard <> "E" And DrawStandard <> "F" And DrawStandard <> "G" Then
		DrawStandard = "N"
	End If
	'=============== 자격미달여부 계산 끝 ===============


	'=============== 필수서류여부 계산 (하드코딩)===============
	'====== 하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc =====

	'//////////////////////////////////////////////////////////////////////////////////////////////////////
	'//필수서류 기준별 계산식 (하드코딩) 1 ~ 5번 코드 (Y = 체크)
	'//하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc
	'//***** 지원자관리에서 관리자가 직접 서류체크 -> 필수서류여부 다시 체크
	'//////////////////////////////////////////////////////////////////////////////////////////////////////

	'////////////////////////////////////////////////////////////////////////////////////////////////
	'//1번코드 (일반/일반고/전문(직업)과정 전형) 생기부 or 검정고시 합격증명서, 성적증명서
	'//			X05010(일반전형), X05041(일반고전형), X05042(전문(직업)과정졸업자)
	'//2번코드 (농어촌전형) 1유형 : 농어촌전형 추천서 1부, 중/고등학교 생활기록부 , 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서
	'//						2유형 : 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본
	'//			X05110(농어촌전형), X05111(농어촌1유형전형), X05112(농어촌2유형전형)
	'//3번코드 (기초생활수급자 및 차상위계층 전형) 
	'//			공통 : 생기부 or 검정고시 합격증명서, 성적증명서
	'//			기초 : 지원자 명의의 수급자 증명서 1부
	'//			차상위 : 장애수당 대상자 확인서, 장애인연금 대상자 확인서, 자활근로자 증명서, 한부모 가족 증명서, 우선돌봄 차상위 확인서, 차상위 본인부담경감 대상자 증명서 중 1부.
	'//			X05120(기초및차상위전형)
	'//4번코드 (전문대졸이상졸업자전형) 전적 대학 졸업(수료)증명서 1부. 및  성적증명서 1부
	'//			X05130(대학졸업자전형)
	'////////////////////////////////////////////////////////////////////////////////////////////////

	'1번 코드
	If documentCheck21 = "3" Then
		If StudentRecord = "1" Then
			'If Document <> "C" Then
				Document = "N"
			'End If
			document21 = "0"			
		Else
			If document21 = "1" Then
				If Document <> "C" Then
					Document = "N"
				End If
			Else
				Document = "C"
				document21 = "3"
				If LEN(DocumentMsg) < 1 Then
					DocumentMsg = "<b>일반전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				Else
					DocumentMsg = DocumentMsg & "= <b>일반전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				End If
			End If
		End If
	ElseIf documentCheck21 = "4" Then
		If StudentRecord = "1" Then
			'If Document <> "C" Then
				Document = "N"
			'End If
			document21 = "0"			
		Else
			If document21 = "1" Then
				If Document <> "C" Then
					Document = "N"
				End If
			Else
				Document = "C"
				document21 = "4"
				If LEN(DocumentMsg) < 1 Then
					DocumentMsg = "<b>일반고전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				Else
					DocumentMsg = DocumentMsg & "= <b>일반고전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				End If
			End If
		End If
	ElseIf documentCheck21 = "5" Then
		If StudentRecord = "1" Then
			'If Document <> "C" Then
				Document = "N"
			'End If
			document21 = "0"			
		Else
			If document21 = "1" Then
				If Document <> "C" Then
					Document = "N"
				End If
			Else
				Document = "C"
				document21 = "5"
				If LEN(DocumentMsg) < 1 Then
					DocumentMsg = "<b>전문(직업)과정전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				Else
					DocumentMsg = DocumentMsg & "= <b>전문(직업)과정전형 :</b> 생기부 또는 검정고시 합격증명서, 성적증명서"
				End If
			End If
		End IF
	ElseIf documentCheck21 = "1" Then
		If Document <> "C" Then
			Document = "N"
		End If
		document21 = "1"
	End If

	'2번 코드
	If documentCheck22 = "3" Then
		If document22 = "1" Then
			If Document <> "C" Then
				Document = "N"
			End If
		Else
			Document = "C"
			document22 = "3"
			If LEN(DocumentMsg) < 1 Then
				DocumentMsg = "<b>농어촌전형 :</b> 유형이 없습니다. 유형을 확인하여 필수서류를 체크하세요."
			Else
				DocumentMsg = DocumentMsg & "= <b>농어촌전형 :</b> 유형이 없습니다. 유형을 확인하여 필수서류를 체크하세요."
			End If
		End If
	ElseIf documentCheck22 = "4" Then
		If document22 = "1" Then
			If Document <> "C" Then
				Document = "N"
			End If
		Else
			Document = "C"
			document22 = "4"
			If LEN(DocumentMsg) < 1 Then
				DocumentMsg = "<b>농어촌전형 1유형 :</b> 농어촌전형 추천서, 중/고등학교 생활기록부 , 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
			Else
				DocumentMsg = DocumentMsg & "= <b>농어촌전형 1유형 :</b> 농어촌전형 추천서, 중/고등학교 생활기록부 , 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
			End If
		End If
	ElseIf documentCheck22 = "5" Then
		If document22 = "1" Then
			If Document <> "C" Then
				Document = "N"
			End If
		Else
			Document = "C"
			document22 = "5"
			If LEN(DocumentMsg) < 1 Then
				DocumentMsg = "<b>농어촌전형 2유형 :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
			Else
				DocumentMsg = DocumentMsg & "= <b>농어촌전형 2유형 :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
			End If
		End If
	ElseIf documentCheck22 = "1" Then
		If Document <> "C" Then
			Document = "N"
		End If
		document22 = "1"
	End If

	'3번 코드
	If documentCheck23 = "3" Then
		If document23 = "1" Then
			If Document <> "C" Then
				Document = "N"
			End If
		Else
			Document = "C"
			document23 = "3"
			If StudentRecord = "1" Then
				If LEN(DocumentMsg) < 1 Then
					DocumentMsg = "<b>기초 및 차상위 전형 :</b> 기초생활수급자 : 지원자 명의의 수급자 증명서 <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
					DocumentMsg = DocumentMsg & "차상위 : 장애수당 대상자 확인서, 장애인연금 대상자 확인서, 자활근로자 증명서, 한부모 가족 증명서, 우선돌봄 차상위 확인서, 차상위 본인부담경감 대상자 증명서 중 1부."
				Else
					DocumentMsg = DocumentMsg & "= <b>기초 및 차상위 전형 :</b> 기초생활수급자 : 지원자 명의의 수급자 증명서 <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
					DocumentMsg = DocumentMsg & "차상위 : 장애수당 대상자 확인서, 장애인연금 대상자 확인서, 자활근로자 증명서, 한부모 가족 증명서, 우선돌봄 차상위 확인서, 차상위 본인부담경감 대상자 증명서 중 1부."
				End If
			Else
				If LEN(DocumentMsg) < 1 Then
					DocumentMsg = "<b>기초 및 차상위 전형 :</b> 공통 : 고등학교 생활기록부. 또는 검정고시 합격증명서, 성적증명서.<br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					DocumentMsg = DocumentMsg & "기초생활수급자 : 지원자 명의의 수급자 증명서 <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
					DocumentMsg = DocumentMsg & "차상위 : 장애수당 대상자 확인서, 장애인연금 대상자 확인서, 자활근로자 증명서, 한부모 가족 증명서, 우선돌봄 차상위 확인서, 차상위 본인부담경감 대상자 증명서 중 1부."
				Else
					DocumentMsg = DocumentMsg & "= <b>기초 및 차상위 전형 :</b> 공통 : 고등학교 생활기록부. 또는 검정고시 합격증명서, 성적증명서.<br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					DocumentMsg = DocumentMsg & "기초생활수급자 : 지원자 명의의 수급자 증명서 <br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "
					DocumentMsg = DocumentMsg & "차상위 : 장애수당 대상자 확인서, 장애인연금 대상자 확인서, 자활근로자 증명서, 한부모 가족 증명서, 우선돌봄 차상위 확인서, 차상위 본인부담경감 대상자 증명서 중 1부."
				End If
			End If
		End If
	ElseIf documentCheck23 = "1" Then
		If Document <> "C" Then
			Document = "N"
		End If
		document23 = "1"
	End If

	'4번 코드
	If documentCheck24 = "3" Then
		If document24 = "1" Then
			If Document <> "C" Then
				Document = "N"
			End If
		Else
			Document = "C"
			document24 = "3"
			If LEN(DocumentMsg) < 1 Then
				DocumentMsg = "<b>전문대졸이상졸업자 전형 :</b> 전적 대학 졸업(수료)증명서. 및  성적증명서 "
			Else
				DocumentMsg = DocumentMsg & "= <b>전문대졸이상졸업자 전형 :</b> 전적 대학 졸업(수료)증명서. 및  성적증명서"
			End If
		End If
	ElseIf documentCheck24 = "1" Then
		If Document <> "C" Then
			Document = "N"
		End If
		document24 = "1"
	End If

	'=============== 필수서류여부 계산 끝 ===============

	Select Case process
		Case "RegApplicant" '지원자
			Call setApplicant()
		Case "RegApplicantAdd" '지원자 서류체크(관리자 외)
			Call setApplicantAdd()
	End Select

	'=============== 지원자 정보 입력 ===============
	Sub setApplicant()

		'On Error Resume Next
		
		'///////////////////////////////////////////////////////////////////
		'// 지원자 별 가산점, 동의여부, 필수서류 저장(위반자 추가해야 함)
		'///////////////////////////////////////////////////////////////////

		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE ApplicationTable SET "
		SQL = SQL & vbCrLf & "		ExtraPoint = ?, StudentRecordAgreement = ?, QualificationAgreement = ?, CSATAgreement = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryCheck1 = ?, DocumentaryCheck2 = ?, DocumentaryCheck3 = ?, DocumentaryCheck4 = ?, DocumentaryCheck5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryCheck6 = ? ,DocumentaryCheck7 = ?, DocumentaryCheck8 = ? "
		SQL = SQL & vbCrLf & "		,DrawStandard = ? ,DrawMsg =?, InterviewerPoint = ?, PracticalPoint = ? "
		SQL = SQL & vbCrLf & "		,Document = ? ,DocumentMsg =?,DocumentaryCheck21 = ?, DocumentaryCheck22 = ?, DocumentaryCheck23 = ?, DocumentaryCheck24 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE MYear = ? "
		SQL = SQL & vbCrLf & " AND StudentNumber = ? "

		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@ExtraPoint",					adInteger,		adParamInput,		0,		ExtraPoint) _
			, Array("@StudentRecordAgreement",		adInteger,		adParamInput,		0,		StudentRecord) _
			, Array("@QualificationAgreement",		adInteger,		adParamInput,		0,		Qualification) _
			, Array("@CSATAgreement",				adInteger,		adParamInput,		0,		CSAT) _

			, Array("@DocumentaryCheck1",			adInteger,		adParamInput,		0,		document1) _
			, Array("@DocumentaryCheck2",			adInteger,		adParamInput,		0,		document2) _
			, Array("@DocumentaryCheck3",			adInteger,		adParamInput,		0,		document3) _
			, Array("@DocumentaryCheck4",			adInteger,		adParamInput,		0,		document4) _
			, Array("@DocumentaryCheck5",			adInteger,		adParamInput,		0,		document5) _
			, Array("@DocumentaryCheck6",			adInteger,		adParamInput,		0,		document6) _
			, Array("@DocumentaryCheck7",			adInteger,		adParamInput,		0,		document7) _
			, Array("@DocumentaryCheck8",			adInteger,		adParamInput,		0,		document8) _

			, Array("@DrawStandard",				adVarchar,		adParamInput,		255,	DrawStandard) _
			, Array("@DrawMsg",						adVarchar,		adParamInput,		5000,	DrawMsg) _
			, Array("@InterviewerPoint",			adInteger,		adParamInput,		0,		InterviewerPoint) _
			, Array("@PracticalPoint",				adInteger,		adParamInput,		0,		PracticalPoint) _

			, Array("@Document",					adVarchar,		adParamInput,		255,	Document) _
			, Array("@DocumentMsg",					adVarchar,		adParamInput,		5000,	DocumentMsg) _
			, Array("@DocumentaryCheck21",			adInteger,		adParamInput,		0,		document21) _
			, Array("@DocumentaryCheck22",			adInteger,		adParamInput,		0,		document22) _
			, Array("@DocumentaryCheck23",			adInteger,		adParamInput,		0,		document23) _
			, Array("@DocumentaryCheck24",			adInteger,		adParamInput,		0,		document24) _
		
			, Array("@UPDT_USID",					adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",					adVarchar,		adParamInput,		20,		UPDT_ADDR) _
			, Array("@MYear",						adVarchar,		adParamInput,		50,		MYear) _
			, Array("@StudentNumber",				adVarchar,		adParamInput,		50,		StudentNumber) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)

		strLogMSG = "지원자관리 > "& MYear &"_"& StudentNumber &" 지원자 정보가 수정되었습니다."
		'복수선택용(사용 안 함)
		'strLogMSG = "지원자관리 > "& MYearTemp &"_"& StudnetNumberTemp &" 지원자 등 " & UpdateCnt &" 건의 지원자 정보가 수정되었습니다."

		InsertType = "Update"
		'UpdateCnt = UpdateCnt + 1
		'=============== 지원자 정보 입력 끝 ===============
	End Sub

	'=============== 지원자 서류체크 입력 ===============
	Sub setApplicantAdd()

		'On Error Resume Next
		
		'///////////////////////////////////////////////////////////////////
		'// 지원자 별 자격미달, 필수서류 저장(위반자 추가해야 함)
		'///////////////////////////////////////////////////////////////////

		'// 수정 ================
		SQL = ""
		SQL = SQL & vbCrLf & "UPDATE ApplicationTable SET "
		SQL = SQL & vbCrLf & "		DocumentaryCheck1 = ?, DocumentaryCheck2 = ?, DocumentaryCheck3 = ?, DocumentaryCheck4 = ?, DocumentaryCheck5 = ?  "
		SQL = SQL & vbCrLf & "		,DocumentaryCheck6 = ? ,DocumentaryCheck7 = ?, DocumentaryCheck8 = ? "
		SQL = SQL & vbCrLf & "		,DrawStandard = ? ,DrawMsg =? "
		SQL = SQL & vbCrLf & "		,Document = ? ,DocumentMsg =?,DocumentaryCheck21 = ?, DocumentaryCheck22 = ?, DocumentaryCheck23 = ?, DocumentaryCheck24 = ? "
		SQL = SQL & vbCrLf & "		,UPDT_USID = ?,UPDT_DATE = getdate(),UPDT_ADDR = ?, InsertTime = getdate() "
		SQL = SQL & vbCrLf & " WHERE MYear = ? "
		SQL = SQL & vbCrLf & " AND StudentNumber = ? "

		'update일 때는 UPDT입력
		arrParams = Array(_
			  Array("@DocumentaryCheck1",			adInteger,		adParamInput,		0,		document1) _
			, Array("@DocumentaryCheck2",			adInteger,		adParamInput,		0,		document2) _
			, Array("@DocumentaryCheck3",			adInteger,		adParamInput,		0,		document3) _
			, Array("@DocumentaryCheck4",			adInteger,		adParamInput,		0,		document4) _
			, Array("@DocumentaryCheck5",			adInteger,		adParamInput,		0,		document5) _
			, Array("@DocumentaryCheck6",			adInteger,		adParamInput,		0,		document6) _
			, Array("@DocumentaryCheck7",			adInteger,		adParamInput,		0,		document7) _
			, Array("@DocumentaryCheck8",			adInteger,		adParamInput,		0,		document8) _

			, Array("@DrawStandard",				adVarchar,		adParamInput,		255,	DrawStandard) _
			, Array("@DrawMsg",						adVarchar,		adParamInput,		5000,	DrawMsg) _

			, Array("@Document",					adVarchar,		adParamInput,		255,	Document) _
			, Array("@DocumentMsg",					adVarchar,		adParamInput,		5000,	DocumentMsg) _

			, Array("@DocumentaryCheck21",			adInteger,		adParamInput,		0,		document21) _
			, Array("@DocumentaryCheck22",			adInteger,		adParamInput,		0,		document22) _
			, Array("@DocumentaryCheck23",			adInteger,		adParamInput,		0,		document23) _
			, Array("@DocumentaryCheck24",			adInteger,		adParamInput,		0,		document24) _
		
			, Array("@UPDT_USID",					adVarchar,		adParamInput,		20,		UPDT_USID) _
			, Array("@UPDT_ADDR",					adVarchar,		adParamInput,		20,		UPDT_ADDR) _

			, Array("@MYear",						adVarchar,		adParamInput,		50,		MYear) _
			, Array("@StudentNumber",				adVarchar,		adParamInput,		50,		StudentNumber) _
		)
		
		'objDB.blnDebug = true
		Call objDB.sbExecSQL(SQL, arrParams)

		strLogMSG = "지원자관리 > "& MYear &"_"& StudentNumber &" 지원자를 "& SessionUserID & "가 서류체크 하였습니다."

		InsertType = "Update"
		'UpdateCnt = UpdateCnt + 1
		'=============== 지원자 서류체크 입력 끝 ===============
	End Sub	
'Next

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "지원자정보 저장 완료"
	'objDB.sbCommitTrans 
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
</Metissoft>