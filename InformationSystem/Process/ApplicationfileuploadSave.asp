<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim LogDivision				: LogDivision = "ApplicationExcelSave"

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim InsertType

Dim IDX							'인덱스																	
Dim MYear                 		'입력한 년도
Dim Division0             		'모집시기
Dim StudentNumber         		'수험번호
Dim StudentNameKor        		'이름(한글)
Dim StudentNameUsa        		'이름(영어)
Dim StudentNameChi        		'이름(한문)
Dim Citizen1              		'주민1
Dim Citizen2              		'주민2
Dim Sex                   		'성별
Dim HighCode              		'고교코드
Dim HighSubject           		'고교학과
Dim HighGraduationYear    		'고교졸업년
Dim HighGraduationDivision		'고교졸업여부
Dim Qualification				'검정고시여부
Dim QualificationAreaCode 		'검정고시합격지구코드
Dim QualificationYear     		'검정고시합격년
Dim Subject               		'학과
Dim Semester              		'생기부 성적 반영 학기
Dim UniversityCode        		'대졸자 출신대학명
Dim AugScore              		'평균점수
Dim PerfectScore          		'만점
Dim Credit                		'이수학점
Dim Division1             		'전형
Dim Division2             		'전형2
Dim Division3             		'전형3
Dim Division4             		'전형4
Dim HighDivision          		'고교(과정)구분
Dim RefundDivision        		'환불방법
Dim RefundAccountHolder   		'환불예금주
Dim RefundBankCode        		'환불은행
Dim RefundAccount         		'환불계좌
Dim Tel1                  		'자택번호
Dim Tel2                  		'핸드폰번호
Dim Tel3                  		'보호자핸드폰번호
Dim Email                 		'이메일
Dim Zipcode               		'우편번호
Dim Address1              		'기본주소
Dim Address2              		'상세주소
Dim StudentNameAgreement  		'내용확인자 성명
Dim StudentAgreement      		'수험생확인동의
Dim StudentRecordAgreement		'학교생활기록부 온라인동의
Dim QualificationAgreement		'검정고시합격성적 온라인동의
Dim CSATAgreement				'수능동의
Dim ReceiptDate					'접수일
Dim ReceiptTime					'접수시간	
Dim SubjectCode					'오산대 모집단위코드 조합 : 학과코드 + 모집시기 + 전형

'입력
Dim INPT_USID, INPT_ADDR

'-------------자격미달여부 관련 변수---------------

'자격미달여부 결과 값(필수서류)
Dim document1 : document1 = "0"				
Dim document2 : document2 = "0"				
Dim document3 : document3 = "0"				
Dim document4 : document4 = "0"				
Dim document5 : document5 = "0"				
Dim document6 : document6 = "0"				
Dim document7 : document7 = "0"				
Dim document8 : document8 = "0"	

'평가 비율
Dim InterviewerRatio : InterviewerRatio = "0"	
Dim PracticalRatio : PracticalRatio = "0"	

'자격미달여부
Dim DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6	

'가산점 여부
Dim ExtraPoint, ExtraPoint1, ExtraPoint2, ExtraPoint3, ExtraPoint4, ExtraPoint5, ExtraPoint6				

'자격미달결과 및 필수서류메세지
Dim DrawStandard, DrawMsg 

'-------------자격미달여부 관련 변수 끝------------

'-------------필수서류 관련 변수------------

'필수서류여부
Dim DocumentaryEvidence1 : DocumentaryEvidence1 = "0"
Dim DocumentaryEvidence2 : DocumentaryEvidence2 = "0"
Dim DocumentaryEvidence3 : DocumentaryEvidence3 = "0"
Dim DocumentaryEvidence4 : DocumentaryEvidence4 = "0"
Dim DocumentaryEvidence5 : DocumentaryEvidence5 = "0"
Dim DocumentaryEvidence6 : DocumentaryEvidence6 = "0"

'-------------필수서류 관련 변수 끝------------

'건수
Dim ApplicationCount			: ApplicationCount = fnRF("ApplicationCount")

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, i, j
Dim arrParams2, AryHash2
Dim count : count = "0"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'=======================================================
'=====입학원서 엑셀등록=================================
'=======================================================

'On Error Resume Next

'///////////////////////////////////////////////////////
'// 입학원서 임시테이블 조회
'///////////////////////////////////////////////////////
SQL = ""
SQL = SQL & vbCrLf & "Select Myear,Division0,Subject,Division1,Division2,Division3,Division4,StudentNumber,StudentNameKor,StudentNameUsa,StudentNameChi,Citizen1,Citizen2,Sex,HighGraduationYear,HighCode,HighSubject,HighGraduationDivision,Qualification,QualificationYear,QualificationAreaCode,Semester,UniversityName,AugScore,PerfectScore,Credit,HighDivision,RefundDivision,RefundAccountHolder,RefundBankCode,RefundAccount,Tel1,Tel2,Tel3,Email,ZipCode,Address1,Address2,StudentNameAgreement,StudentRecordAgreement,QualificationAgreement,CSATAgreement,StudentAgreement,ReceiptDate,ReceiptTime,INPT_USID,INPT_DATE,INPT_ADDR "
SQL = SQL & vbCrLf & "From ##ApplicationTable "

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

For i = 0 To Ubound(AryHash,1)	

	MYear							=	AryHash(i).Item("Myear") 																	'입력한 년도                                             
	Division0             			=	AryHash(i).Item("Division0") 																'모집시기                                                
	StudentNumber         			=	AryHash(i).Item("StudentNumber")															'수험번호                                                
	StudentNameKor        			=	AryHash(i).Item("StudentNameKor")															'이름(한글)                                              
	StudentNameUsa        			=	AryHash(i).Item("StudentNameUsa")															'이름(영어)                                              
	StudentNameChi        			=	AryHash(i).Item("StudentNameChi")															'이름(한문)                                              
	Citizen1              			=	AryHash(i).Item("Citizen1")																	'주민1                                                   
	Citizen2              			=	AryHash(i).Item("Citizen2")																	'주민2                                                   
	Sex                   			=	AryHash(i).Item("Sex") 																		'성별                                                    
	HighCode              			=	AryHash(i).Item("HighCode")																	'고교코드                                                
	HighSubject           			=	AryHash(i).Item("HighSubject")																'고교학과                                                
	HighGraduationYear    			=	AryHash(i).Item("HighGraduationYear")														'고교졸업년                                              
	HighGraduationDivision			=	AryHash(i).Item("HighGraduationDivision")													'고교졸업여부   
	Qualification 					=	AryHash(i).Item("Qualification")															'검정고시여부
	QualificationAreaCode 			=	AryHash(i).Item("QualificationAreaCode")													'검정고시합격지구코드                                    
	QualificationYear     			=	AryHash(i).Item("QualificationYear")														'검정고시합격년                                          
	Subject               			=	AryHash(i).Item("Subject") 																	'학과                                                    
	Semester              			=	AryHash(i).Item("Semester") 																'생기부 성적 반영 학기                                   
	UniversityCode        			=	AryHash(i).Item("UniversityName")															'대졸자 출신대학명                                       
	AugScore              			=	AryHash(i).Item("AugScore")																	'평균점수                                                
	PerfectScore          			=	AryHash(i).Item("PerfectScore")																'만점                                                    
	Credit                			=	AryHash(i).Item("Credit") 																	'이수학점                                                
	Division1             			=	AryHash(i).Item("Division1")																'전형
	Division2             			=	AryHash(i).Item("Division2")																'전형2
	Division3             			=	AryHash(i).Item("Division3")																'전형3
	Division4             			=	AryHash(i).Item("Division4")																'전형4
	HighDivision          			=	AryHash(i).Item("HighDivision")																'고교(과정)구분                                          
	RefundDivision        			=	AryHash(i).Item("RefundDivision")															'환불방법                                                
	RefundAccountHolder   			=	AryHash(i).Item("RefundAccountHolder") 														'환불예금주                                              
	RefundBankCode        			=	AryHash(i).Item("RefundBankCode")															'환불은행                                                
	RefundAccount         			=	AryHash(i).Item("RefundAccount")															'환불계좌                                                
	Tel1                  			=	AryHash(i).Item("Tel1")																		'자택번호                                                
	Tel2                  			=	AryHash(i).Item("Tel2")																		'핸드폰번호                                              
	Tel3                  			=	AryHash(i).Item("Tel3")																		'보호자핸드폰번호                                        
	Email                 			=	AryHash(i).Item("Email")																	'이메일                                                  
	Zipcode               			=	AryHash(i).Item("ZipCode")																	'우편번호                                                
	Address1              			=	AryHash(i).Item("Address1")																	'기본주소                                                
	Address2              			=	AryHash(i).Item("Address2")																	'상세주소                                                
	StudentNameAgreement  			=	AryHash(i).Item("StudentNameAgreement") 													'내용확인자 성명                                         
	StudentAgreement      			=	AryHash(i).Item("StudentAgreement") 														'수험생확인동의                                          
	StudentRecordAgreement			=	AryHash(i).Item("StudentRecordAgreement")													'학교생활기록부 온라인동의                               
	QualificationAgreement			=	AryHash(i).Item("QualificationAgreement")													'검정고시합격성적 온라인동의                             
	CSATAgreement					=	AryHash(i).Item("CSATAgreement")															'수능동의                                                
	ReceiptDate						=	AryHash(i).Item("ReceiptDate")																'접수일                                                  
	ReceiptTime						=	AryHash(i).Item("ReceiptTime")																'접수시간	                                             
	SubjectCode						=	AryHash(i).Item("Subject") & AryHash(0).Item("Division0") & AryHash(0).Item("Division1") 	'오산대 모집단위코드 조합 : 학과코드 + 모집시기 + 전형   
	INPT_USID						=	AryHash(i).Item("INPT_USID")																'입력자
	INPT_ADDR						=	AryHash(i).Item("INPT_ADDR")																'입력IP

	If Not(isnull(ReceiptTime)) Then
		ReceiptTime	= FormatDateTime(ReceiptTime,4)
	End If

	'///////////////////////////////////////////////////////
	'// 평가비율에 따른 자격미달자, 가산점, 필수서류 구분
	'///////////////////////////////////////////////////////
	SQL = ""
	SQL = SQL & vbCrLf & "Select InterviewerRatio, PracticalRatio, DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6 "
	SQL = SQL & vbCrLf & "		 , ExtraPoint1, ExtraPoint2, ExtraPoint3, ExtraPoint4, ExtraPoint5, ExtraPoint6 "
	SQL = SQL & vbCrLf & "		 , DocumentaryEvidence1, DocumentaryEvidence2, DocumentaryEvidence3, DocumentaryEvidence4, DocumentaryEvidence5, DocumentaryEvidence6 "
	SQL = SQL & vbCrLf & "from AppraisalTable "
	SQL = SQL & vbCrLf & "where MYear = ? "
	SQL = SQL & vbCrLf & "And SubjectCode = ?; "

	Call objDB.sbSetArray("@MYear", adVarchar, adParamInput, 50, MYear)
	Call objDB.sbSetArray("@SubjectCode", adVarchar, adParamInput, 50, SubjectCode)

	'objDB.blnDebug = TRUE
	arrParams2 = objDB.fnGetArray
	AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

	If isArray(AryHash2) Then
		DrawStandard1			= AryHash2(0).Item("DrawStandard1")
		DrawStandard2			= AryHash2(0).Item("DrawStandard2")
		DrawStandard3			= AryHash2(0).Item("DrawStandard3")
		DrawStandard4			= AryHash2(0).Item("DrawStandard4")
		DrawStandard5			= AryHash2(0).Item("DrawStandard5")
		DrawStandard6			= AryHash2(0).Item("DrawStandard6")
		ExtraPoint1				= AryHash2(0).Item("ExtraPoint1")
		ExtraPoint2				= AryHash2(0).Item("ExtraPoint2")
		ExtraPoint3				= AryHash2(0).Item("ExtraPoint3")
		ExtraPoint4				= AryHash2(0).Item("ExtraPoint4")
		ExtraPoint5				= AryHash2(0).Item("ExtraPoint5")
		ExtraPoint6				= AryHash2(0).Item("ExtraPoint6")
		InterviewerRatio		= AryHash2(0).Item("InterviewerRatio")  
		PracticalRatio			= AryHash2(0).Item("PracticalRatio")  
		DocumentaryEvidence1	= AryHash2(0).Item("DocumentaryEvidence1")
		DocumentaryEvidence2	= AryHash2(0).Item("DocumentaryEvidence2")
		DocumentaryEvidence3	= AryHash2(0).Item("DocumentaryEvidence3")
		DocumentaryEvidence4	= AryHash2(0).Item("DocumentaryEvidence4")
		DocumentaryEvidence5	= AryHash2(0).Item("DocumentaryEvidence5")
		DocumentaryEvidence6	= AryHash2(0).Item("DocumentaryEvidence6")
	End If

	'///////////////////////////////////////////////
	'// 등록되어 있는 고교(국내)코드 값 구하기
	'///////////////////////////////////////////////
	SQL = ""
	SQL = SQL & vbCrLf & "Select SubCode, SubCodeName, Temp1 "
	SQL = SQL & vbCrLf & "from CodeSub "
	SQL = SQL & vbCrLf & "where MasterCode = 'HighCode'; "

	'objDB.blnDebug = TRUE
	arrParams2 = objDB.fnGetArray
	AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

	'/////////////////////////////////////////////////////////
	'// 검정고시자 여부 체크
	'// 0 고교졸업자 (고교졸업년도와 졸업고교가 있으면)
	'// 1 검정고시자 (고시합격년도와 합격지구가 있으면)
	'// 2 구분 불가  (고교와 고시 데이터가 혼합 or 없으면)
	'/////////////////////////////////////////////////////////
	If Qualification <> "0" And Qualification <> "1" Then
		If Not(Isnull(HighCode)) And Not(Isnull(HighGraduationYear)) Then
			Qualification = "0"
		ElseIf Not(Isnull(QualificationAreaCode)) And Not(Isnull(QualificationYear)) Then
			Qualification = "1"
		Else
			Qualification = "2"
		End If
	End If

	'=============== 자격미달여부 계산 (하드코딩)============================================================
	'====== 하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save, 4. 수능점수 엑셀 Save =====

	'//////////////////////////////////////////////////////////////////////////////////////////////////////
	'//자격미달자 기준별 계산식 (하드코딩) 1 ~ 8번 코드 (Y = 미달자)
	'//하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save, 4. 수능점수 엑셀 Save
	'//***** 1번코드 (수시용)입학원서 등록 시 체크
	'//***** 2번코드 (정시용)입학원서 등록 시 체크, 수능성적 입력 시 체크 -> 자격미달여부 다시 체크
	'//***** 3번 ~ 8번코드 지원자관리에서 관리자가 직접 서류체크 및 점수 입력 -> 자격미달여부 다시 체크
	'//////////////////////////////////////////////////////////////////////////////////////////////////////ㄷ

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
	'1. 수시자 중 고등학교코드에 등록되어 있는 학교를 졸업 or 예정이면 N
	'2. 수시자 중 검정고시 합격이 있으면 N
	'3. 외국고교이거나 고교코드가 등록되어 있지 않으면 증빙자료가 필요하므로 document에 3(구분)저장
	Dim Draw1Temp
	If Division0 = "X03021" Or Division0 = "X03022" Then '수시1, 수시2
		If Division1 = "X05010" Or Division1 = "X05041" Or Division1 = "X05042" Then '일반, 일반고, 전문
			If DrawStandard1 = "1" Or DrawStandard2 = "1" Or DrawStandard3 = "1" Or DrawStandard4 = "1" Or DrawStandard5 = "1" Or DrawStandard6 = "1" Then
				If Qualification = "0" Or Qualification = "1" Then
					If isArray(AryHash) Then
						For j = 0 to ubound(AryHash2,1)
							If AryHash2(j).Item("SubCode") = HighCode And DrawStandard <> "N" Then
								Draw1Temp = "Y"
							Else
								If Draw1Temp <> "Y" Then
									Draw1Temp = "N"
								End If
							End If
						Next
						If Draw1Temp = "Y" Then
							DrawStandard = "N"
						Else
							DrawStandard = "Y"
							document1 = "3"
							If LEN(DrawMsg) < 1 Then
								DrawMsg = "<b>고교미등록/외국고교 졸업자(제출서류) :</b> 국내고교 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
							Else
								DrawMsg = DrawMsg & "= <b>고교미등록/외국고교 졸업자(제출서류) :</b> 국내고교 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
							End If
						End If
					Else
						DrawStandard = "C"
						document1 = "4"
						If LEN(DrawMsg) < 1 Then
							DrawMsg = "등록된 고교코드가 없습니다. 국내 모든 고등학교를 고교코드에 입력하여 주세요."
						Else
							DrawMsg = DrawMsg & "= 등록된 고교코드가 없습니다. 국내 모든 고등학교를 고교코드에 입력하여 주세요."
						End If
					End If		
				Else
					If StudentRecordAgreement <> "1" Then
						DrawStandard = "Y"
						document1 = "5"
						If LEN(DrawMsg) < 1 Then
							DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
						Else
							DrawMsg = DrawMsg & "= <b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부 or 검정고시 합격증명서, 검정고시 성적증명서"
						End If
					End If
				End If	
			End If
		End If
	End if

	'1. 정시공통 무조건 C
	'2. 입학원서 접수 시점에 수능성적 체크가 불가하므로, 차후 수능성적 등록 시 자동체크
	'3. document에 3(구분)저장
	If Division0 = "X03031" Or Division0 = "X03050" Then '정시, 추가
		If Division1 = "X05010" Or Division1 = "X05041" Or Division1 = "X05042" Then '일반, 일반고, 전문
			If DrawStandard1 = "2" Or DrawStandard2 = "2" Or DrawStandard3 = "2" Or DrawStandard4 = "2" Or DrawStandard5 = "2" Or DrawStandard6 = "2" Then
				If StudentRecordAgreement <> "1" And CSATAgreement <> "1" Then
					document2 = "3"	
					DrawStandard = "Y"
					If LEN(DrawMsg) < 1 Then
						DrawMsg = "<b>학력/수능성적 자격미달(제출서류) :</b> 고등학교 생활기록부, 수능성적 업로드 필요"
					Else
						DrawMsg = DrawMsg & "= <b>학력/수능성적 자격미달(제출서류) :</b> 고등학교 생활기록부, 수능성적 업로드 필요"
					End If
				ElseIf CSATAgreement <> "1" Then 
					document2 = "10"	
					DrawStandard = "C"
					If LEN(DrawMsg) < 1 Then
						DrawMsg = "<b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
					Else
						DrawMsg = DrawMsg & "= <b>수능성적 자격미달(제출서류) :</b> 수능성적 업로드 필요"
					End If
				ElseIf StudentRecordAgreement <> "1" Then
					document2 = "5"	
					DrawStandard = "Y"
					If LEN(DrawMsg) < 1 Then
						DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
					Else
						DrawMsg = DrawMsg & "= <b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
					End If
				End IF
			End If
		End If
	End if
	'1. 농어촌전형 무조건 C
	'2. 필수서류가 필요하므로 관리자가 지원자관리에서 필수서류 체크
	'3. document에 3(구분)저장  
	If Division1 = "X05110" Or Division1 = "X05111" Then '농어촌, 농어촌1유형
		If DrawStandard1 = "4" Or DrawStandard2 = "4" Or DrawStandard3 = "4" Or DrawStandard4 = "4" Or DrawStandard5 = "4" Or DrawStandard6 = "4" Then
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
	End If
	'1. 농어촌 무조건 C
	'2. 필수서류가 필요하므로 관리자가 지원자관리에서 필수서류 체크
	'3. document에 3(구분)저장
	If Division1 = "X05110" Or Division1 = "X05112" Then '농어촌, 농어촌2유형
		If DrawStandard1 = "5" Or DrawStandard2 = "5" Or DrawStandard3 = "5" Or DrawStandard4 = "5" Or DrawStandard5 = "5" Or DrawStandard6 = "5" Then
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
	End if
	'1. 기초수급자 무조건 C
	'2. 필수서류가 필요하므로 관리자가 지원자관리에서 필수서류 체크
	'3. document에 3(구분)저장
	If Division1 = "X05120" Then '기초수급자 및 차상위
		If DrawStandard1 = "6" Or DrawStandard2 = "6" Or DrawStandard3 = "6" Or DrawStandard4 = "6" Or DrawStandard5 = "6" Or DrawStandard6 = "6" Then
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document6 = "3"
			If StudentRecordAgreement = "1" Then
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
			End If
		End If
	End If
	'1. 차상위 무조건 C
	'2. 필수서류가 필요하므로 관리자가 지원자관리에서 필수서류 체크
	'3. document에 3(구분)저장
	If Division1 = "X05120" Then '기초수급자 및 차상위
		If DrawStandard1 = "7" Or DrawStandard2 = "7" Or DrawStandard3 = "7" Or DrawStandard4 = "7" Or DrawStandard5 = "7" Or DrawStandard6 = "7" Then
			If DrawStandard <> "Y" Then
				DrawStandard = "C"
			End If
			document7 = "3"
			If StudentRecordAgreement = "1" Then
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
	End if
	'1. 대학졸업자 무조건 C
	'2. 필수서류가 필요하므로 관리자가 지원자관리에서 필수서류 체크
	'3. document에 3(구분)저장
	If Division1 = "X05130" Then '전문대이상 졸업자
		If DrawStandard1 = "8" Or DrawStandard2 = "8" Or DrawStandard3 = "8" Or DrawStandard4 = "8" Or DrawStandard5 = "8" Or DrawStandard6 = "8" Then
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
	End if
	'1. 면접(실기) 무조건 C
	'2. 면접 시스템과 연동하지 않아, 면접 미응시자는 지원자관리에서 관리자가 개별적으로 입력
	'3. document에 3(면접/실기 수동입력필요), 4(면접 수동입력필요), 5(실기 수동입력필요), 6(자격미달여부에는 있으나 평가점수 비율에는 없음) 저장
	If DrawStandard1 = "3" Or DrawStandard2 = "3" Or DrawStandard3 = "3" Or DrawStandard4 = "3" Or DrawStandard5 = "3" Or DrawStandard6 = "3" Then
		If InterviewerRatio > 0 And PracticalRatio > 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "D"
			End If
			document3 = "3"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
			End If	
		ElseIf InterviewerRatio > 0 Then
			If DrawStandard <> "Y" Then
				DrawStandard = "E"
			End If
			document3 = "4"
			If LEN(DrawMsg) < 1 Then
				DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			Else
				DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
			End If
		ElseIf PracticalRatio > 0 Then
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
	End If
	'Y, N, C, D, E, F, G가 아니면 자격미달기준이 없는 것이므로 자격미달 아님(N)
	If DrawStandard <> "Y" And DrawStandard <> "C" And DrawStandard <> "D" And DrawStandard <> "E" And DrawStandard <> "F" And DrawStandard <> "G" Then
		DrawStandard = "N"
	End If



	'////////////////////////////////////////////////////////////////////////////////////////////////
	'//(필요시 추가)
	'//산점 기준별 계산식 (하드코딩)
	'//가산점은 현재 모집단위별 기준 없음
	'////////////////////////////////////////////////////////////////////////////////////////////////



	'=============== 입학원서 입력 ===============

	'// 입력 =================
	SQL = ""
	SQL = SQL & vbCrLf & "INSERT INTO ApplicationTable ( "
	SQL = SQL & vbCrLf & "		MYear, SubjectCode, Division0, StudentNumber, StudentNameKor, StudentNameUsa, StudentNameChi  "
	SQL = SQL & vbCrLf & "		 , Citizen1, Citizen2, Sex, HighCode, HighSubject, HighGraduationYear, HighGraduationDivision  "
	SQL = SQL & vbCrLf & "		 , QualificationAreaCode, QualificationYear  "
	SQL = SQL & vbCrLf & "		 , Subject, Semester, UniversityName, AugScore, PerfectScore, Credit  "
	SQL = SQL & vbCrLf & "		 , Division1, HighDivision, RefundDivision, RefundAccountHolder  "
	SQL = SQL & vbCrLf & "		 , RefundBankCode, RefundAccount, Tel1, Tel2, Tel3, Email  "
	SQL = SQL & vbCrLf & "		 , Zipcode, Address1, Address2, StudentNameAgreement  "
	SQL = SQL & vbCrLf & "		 , StudentAgreement, StudentRecordAgreement, QualificationAgreement, CSATAgreement  "
	SQL = SQL & vbCrLf & "		 , ReceiptDate, ReceiptTime  "
	SQL = SQL & vbCrLf & "		 , DrawStandard, DrawMsg, Qualification  "
	SQL = SQL & vbCrLf & "		 , DocumentaryCheck1, DocumentaryCheck2, DocumentaryCheck3, DocumentaryCheck4, DocumentaryCheck5  "
	SQL = SQL & vbCrLf & "		 , DocumentaryCheck6, DocumentaryCheck7, DocumentaryCheck8  "
	SQL = SQL & vbCrLf & "		 , INPT_USID, INPT_DATE, INPT_ADDR "
	SQL = SQL & vbCrLf & " ) VALUES ( "
	SQL = SQL & vbCrLf & "		?, ?, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, ?, ? "
	SQL = SQL & vbCrLf & "		, ?, getdate(), ?"
	SQL = SQL & vbCrLf & " ) "

	'insert일 때는 INPT입력
	arrParams = Array(_
		  Array("@MYear",					adVarchar,		adParamInput,		4,		MYear) _
		, Array("@SubjectCode",				adVarchar,		adParamInput,		60,		SubjectCode) _
		, Array("@Division0",				adVarchar,		adParamInput,		60,		Division0) _
		, Array("@StudentNumber",			adVarchar,		adParamInput,		10,		StudentNumber) _
		, Array("@StudentNameKor",			adVarchar,		adParamInput,		30,		StudentNameKor) _
		, Array("@StudentNameUsa",			adVarchar,		adParamInput,		30,		StudentNameUsa) _
		, Array("@StudentNameChi",			adVarchar,		adParamInput,		30,		StudentNameChi) _
		, Array("@Citizen1",				adInteger,		adParamInput,		0,		Citizen1) _
		, Array("@Citizen2",				adInteger,		adParamInput,		0,		Citizen2) _
		, Array("@Sex",						adInteger,		adParamInput,		0,		Sex) _
		, Array("@HighCode",				adVarchar,		adParamInput,		10,		HighCode) _
		, Array("@HighSubject",				adVarchar,		adParamInput,		30,		HighSubject) _
		, Array("@HighGraduationYear",		adVarchar,		adParamInput,		4,		HighGraduationYear) _
		, Array("@HighGraduationDivision",	adInteger,		adParamInput,		0,		HighGraduationDivision) _
		, Array("@QualificationAreaCode",	adVarchar,		adParamInput,		30,		QualificationAreaCode) _
		, Array("@QualificationYear",		adVarchar,		adParamInput,		4,		QualificationYear) _
		, Array("@Subject",					adVarchar,		adParamInput,		50,		Subject) _
		, Array("@Semester",				adVarchar,		adParamInput,		20,		Semester) _
		, Array("@UniversityName",			adVarchar,		adParamInput,		60,		UniversityCode) _
		, Array("@AugScore",				adVarchar,		adParamInput,		10,		AugScore) _
		, Array("@PerfectScore",			adVarchar,		adParamInput,		10,		PerfectScore) _
		, Array("@Credit",					adInteger,		adParamInput,		0,		Credit) _
		, Array("@Division1",				adVarchar,		adParamInput,		60,		Division1) _
		, Array("@HighDivision",			adVarchar,		adParamInput,		60,		HighDivision) _
		, Array("@RefundDivision",			adInteger,		adParamInput,		0,		RefundDivision) _
		, Array("@RefundAccountHolder",		adVarchar,		adParamInput,		50,		RefundAccountHolder) _
		, Array("@RefundBankCode",			adVarchar,		adParamInput,		50,		RefundBankCode) _
		, Array("@RefundAccount",			adVarchar,		adParamInput,		50,		RefundAccount) _
		, Array("@Tel1",					adVarchar,		adParamInput,		20,		Tel1) _
		, Array("@Tel2",					adVarchar,		adParamInput,		20,		Tel2) _
		, Array("@Tel3",					adVarchar,		adParamInput,		20,		Tel3) _
		, Array("@Email",					adVarchar,		adParamInput,		60,		Email) _
		, Array("@Zipcode",					adInteger,		adParamInput,		0,		Zipcode) _
		, Array("@Address1",				adVarchar,		adParamInput,		100,	Address1) _
		, Array("@Address2",				adVarchar,		adParamInput,		100,	Address2) _
		, Array("@StudentNameAgreement",	adVarchar,		adParamInput,		30,		StudentNameAgreement) _
		, Array("@StudentAgreement",		adVarchar,		adParamInput,		10,		StudentAgreement) _
		, Array("@StudentRecordAgreement",	adVarchar,		adParamInput,		10,		StudentRecordAgreement) _
		, Array("@QualificationAgreement",	adVarchar,		adParamInput,		10,		QualificationAgreement) _
		, Array("@CSATAgreement",			adVarchar,		adParamInput,		10,		CSATAgreement) _
		, Array("@ReceiptDate",				adVarchar,		adParamInput,		255,	ReceiptDate) _
		, Array("@ReceiptTime",				adVarchar,		adParamInput,		255,	ReceiptTime) _
		, Array("@DrawStandard",			adVarchar,		adParamInput,		255,	DrawStandard) _
		, Array("@DrawMsg",					adVarchar,		adParamInput,		5000,	DrawMsg) _
		, Array("@Qualification",			adVarchar,		adParamInput,		10,		Qualification) _
		, Array("@DocumentaryCheck1",		adInteger,		adParamInput,		0,		document1) _
		, Array("@DocumentaryCheck2",		adInteger,		adParamInput,		0,		document2) _
		, Array("@DocumentaryCheck3",		adInteger,		adParamInput,		0,		document3) _
		, Array("@DocumentaryCheck4",		adInteger,		adParamInput,		0,		document4) _
		, Array("@DocumentaryCheck5",		adInteger,		adParamInput,		0,		document5) _
		, Array("@DocumentaryCheck6",		adInteger,		adParamInput,		0,		document6) _
		, Array("@DocumentaryCheck7",		adInteger,		adParamInput,		0,		document7) _
		, Array("@DocumentaryCheck8",		adInteger,		adParamInput,		0,		document8) _
		, Array("@INPT_USID",				adVarchar,		adParamInput,		20,		INPT_USID) _
		, Array("@INPT_ADDR",				adVarchar,		adParamInput,		20,		INPT_ADDR) _
	)

	'objDB.blnDebug = True
	Call objDB.sbExecSQL(SQL, arrParams)

	'SQL = " SELECT @@IDENTITY; "
	'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
	'IDX = CInt(aryList(0, 0))\

	count = count + 1

Next

'////////////////////////////////////
'// 등록 히스토리 
'////////////////////////////////////
strLogMSG = "입학원서관리 > "& MYear &"학년도_"& count &"건의 입학원서가 엑셀 등록되었습니다."
InsertType = "Insert"

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "입학원서  저장 완료"
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