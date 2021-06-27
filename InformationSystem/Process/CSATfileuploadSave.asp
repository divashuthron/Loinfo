<?xml version="1.0" encoding="utf-8"?>
<Metissoft>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Dim LogDivision				: LogDivision = "CSATExcelSave"

Dim strResult				: strResult = "failure"
Dim returnMSG
Dim InsertType

Dim IDX							'인덱스		
Dim SCHL_YEAR					'년도
Dim COLL_FLAG					'모집구분
Dim EXAM_NUMB					'수험번호

Dim LGFD_EXFG					'언어영역응시구분
Dim LGFD_SDSC					'언어영역표준점수
Dim LGFD_CENT					'언어영역백분위
Dim LGFD_GRAD					'언어영역등급

Dim MTFD_EXFG					'수리영역응시구분
Dim MTFD_EXTP					'수리영역응시유형
Dim MTFD_SDSC					'수리영역표준점수
Dim MTFD_CENT					'수리영역백분위
Dim MTFD_GRAD					'수리영역등급

Dim FLFD_EXFG					'외국어영역응시구분
Dim FLFD_SDSC					'외국어영역표준점수
Dim FLFD_CENT					'외국어영역백분위
Dim FLFD_GRAD					'외국어영역등급

Dim RSFD_EXFG					'탐구영역응시구분
Dim RSFD_FLAG					'탐구영역구분
Dim RSFD_CCCT					'탐구영역선택과목수
Dim RSFD_SBJ1					'탐구영역과목1
Dim RSFD_SCR1					'탐구영역표준점수1
Dim RSFD_CNT1					'탐구영역백분위1
Dim RSFD_GRD1					'탐구영역등급1
Dim RSFD_SBJ2					'탐구영역과목2
Dim RSFD_SCR2					'탐구영역표준점수2
Dim RSFD_CNT2					'탐구영역백분위2
Dim RSFD_GRD2					'탐구영역등급2
Dim RSFD_SBJ3					'탐구영역과목3
Dim RSFD_SCR3					'탐구영역표준점수3
Dim RSFD_CNT3					'탐구영역백분위3
Dim RSFD_GRD3					'탐구영역등급3
Dim RSFD_SBJ4					'탐구영역과목4
Dim RSFD_SCR4					'탐구영역표준점수4
Dim RSFD_CNT4					'탐구영역백분위4
Dim RSFD_GRD4					'탐구영역등급4

Dim SCFL_EXFG					'제2외국어영역응시구분
Dim SCFL_SBJT					'제2외국어영역과목
Dim SCFL_SDSC					'제2외국어표준점수
Dim SCFL_CENT					'제2외국어백분위
Dim SCFL_GRAD					'제2외국어등급
Dim REMK_TEXT					'비고

'입력
Dim INPT_USID, INPT_ADDR

'-------------자격미달여부 관련 변수---------------

'자격미달여부 결과 값(필수서류)
Dim document2 : document2 = "1"				

'자격미달여부
Dim DrawStandard1, DrawStandard2, DrawStandard3, DrawStandard4, DrawStandard5, DrawStandard6					

'자격미달결과 및 필수서류메세지
Dim DrawStandard : DrawStandard = "N"
Dim DrawMsg 

'-------------자격미달여부 관련 변수 끝------------

'-------------필수서류 관련 변수------------

'필수서류여부
Dim documentCheck2			
Dim documentCheck3			
Dim documentCheck4			
Dim documentCheck5			
Dim documentCheck6			
Dim documentCheck7			
Dim documentCheck8	

Dim StudentRecordAgreement

'-------------필수서류 관련 변수 끝------------

'건수
Dim CSATCount			: CSATCount = fnRF("CSATCount")

Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG, i
Dim arrParams2, AryHash2
Dim count : count = "0"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB
'objDB.sbBeginTrans()

'=======================================================
'=====수능성적 엑셀등록=================================
'=======================================================

'On Error Resume Next

'=============== 수능성적 테이블에 임시테이블 데이터 입력 ===============

'// 입력 =================
SQL = ""
SQL = SQL & vbCrLf & "INSERT INTO IPSICSAT (SCHL_YEAR,COLL_FLAG,EXAM_NUMB,LGFD_EXFG,LGFD_SDSC,LGFD_CENT,LGFD_GRAD,MTFD_EXFG,MTFD_EXTP,MTFD_SDSC,MTFD_CENT,MTFD_GRAD,FLFD_EXFG,FLFD_SDSC,FLFD_CENT,FLFD_GRAD,RSFD_EXFG,RSFD_FLAG,RSFD_CCCT,RSFD_SBJ1,RSFD_SCR1,RSFD_CNT1,RSFD_GRD1,RSFD_SBJ2,RSFD_SCR2,RSFD_CNT2,RSFD_GRD2,RSFD_SBJ3,RSFD_SCR3,RSFD_CNT3,RSFD_GRD3,RSFD_SBJ4,RSFD_SCR4,RSFD_CNT4,RSFD_GRD4,SCFL_EXFG,SCFL_SBJT,SCFL_SDSC,SCFL_CENT,SCFL_GRAD,REMK_TEXT,INPT_USID,INPT_DATE,INPT_ADDR) "
SQL = SQL & vbCrLf & "Select SCHL_YEAR,COLL_FLAG,EXAM_NUMB,LGFD_EXFG,LGFD_SDSC,LGFD_CENT,LGFD_GRAD,MTFD_EXFG,MTFD_EXTP,MTFD_SDSC,MTFD_CENT,MTFD_GRAD,FLFD_EXFG,FLFD_SDSC,FLFD_CENT,FLFD_GRAD,RSFD_EXFG,RSFD_FLAG,RSFD_CCCT,RSFD_SBJ1,RSFD_SCR1,RSFD_CNT1,RSFD_GRD1,RSFD_SBJ2,RSFD_SCR2,RSFD_CNT2,RSFD_GRD2,RSFD_SBJ3,RSFD_SCR3,RSFD_CNT3,RSFD_GRD3,RSFD_SBJ4,RSFD_SCR4,RSFD_CNT4,RSFD_GRD4,SCFL_EXFG,SCFL_SBJT,SCFL_SDSC,SCFL_CENT,SCFL_GRAD,REMK_TEXT,INPT_USID,INPT_DATE,INPT_ADDR "
SQL = SQL & vbCrLf & "From ##CSATTable "

'objDB.blnDebug = True
Call objDB.sbExecSQL(SQL, arrParams)

'=============== 수능성적 테이블에 임시테이블 데이터 입력 끝 ===============

'=============== 수능성적 입력에 따른 입학원서(자격미달여부) 수정 ===============

'///////////////////////////////////////////////////////
'// 입학원서 임시테이블 조회
'///////////////////////////////////////////////////////
SQL = ""
SQL = SQL & vbCrLf & "Select SCHL_YEAR,COLL_FLAG,EXAM_NUMB,LGFD_EXFG,LGFD_SDSC,LGFD_CENT,LGFD_GRAD,MTFD_EXFG,MTFD_EXTP,MTFD_SDSC,MTFD_CENT,MTFD_GRAD,FLFD_EXFG,FLFD_SDSC,FLFD_CENT,FLFD_GRAD,RSFD_EXFG,RSFD_FLAG,RSFD_CCCT,RSFD_SBJ1,RSFD_SCR1,RSFD_CNT1,RSFD_GRD1,RSFD_SBJ2,RSFD_SCR2,RSFD_CNT2,RSFD_GRD2,RSFD_SBJ3,RSFD_SCR3,RSFD_CNT3,RSFD_GRD3,RSFD_SBJ4,RSFD_SCR4,RSFD_CNT4,RSFD_GRD4,SCFL_EXFG,SCFL_SBJT,SCFL_SDSC,SCFL_CENT,SCFL_GRAD,REMK_TEXT,INPT_USID,INPT_DATE,INPT_ADDR,InsertTime "
SQL = SQL & vbCrLf & "From ##CSATTable "

'objDB.blnDebug = TRUE
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

For i = 0 To Ubound(AryHash,1)	

	SCHL_YEAR		=	AryHash(i).Item("SCHL_YEAR") 		'년도                       
	COLL_FLAG		=	AryHash(i).Item("COLL_FLAG") 		'모집구분                   
	EXAM_NUMB		=	AryHash(i).Item("EXAM_NUMB")		'수험번호                   
																		                          
	LGFD_EXFG		=	AryHash(i).Item("LGFD_EXFG")		'언어영역응시구분           
	LGFD_SDSC		=	AryHash(i).Item("LGFD_SDSC")		'언어영역표준점수           
	LGFD_CENT		=	AryHash(i).Item("LGFD_CENT")		'언어영역백분위             
	LGFD_GRAD		=	AryHash(i).Item("LGFD_GRAD")		'언어영역등급               
																		                          
	MTFD_EXFG		=	AryHash(i).Item("MTFD_EXFG")		'수리영역응시구분           
	MTFD_EXTP		=	AryHash(i).Item("MTFD_EXTP")		'수리영역응시유형           
	MTFD_SDSC		=	AryHash(i).Item("MTFD_SDSC")		'수리영역표준점수           
	MTFD_CENT		=	AryHash(i).Item("MTFD_CENT")		'수리영역백분위           
	MTFD_GRAD		=	AryHash(i).Item("MTFD_GRAD")		'수리영역등급             
																		                          
	FLFD_EXFG		=	AryHash(i).Item("FLFD_EXFG")		'외국어영역응시구분         
	FLFD_SDSC		=	AryHash(i).Item("FLFD_SDSC") 		'외국어영역표준점수         
	FLFD_CENT		=	AryHash(i).Item("FLFD_CENT") 		'외국어영역백분위           
	FLFD_GRAD		=	AryHash(i).Item("FLFD_GRAD")		'외국어영역등급             
																		                          
	RSFD_EXFG		=	AryHash(i).Item("RSFD_EXFG")		'탐구영역응시구분           
	RSFD_FLAG		=	AryHash(i).Item("RSFD_FLAG") 		'탐구영역구분               
	RSFD_CCCT		=	AryHash(i).Item("RSFD_CCCT")		'탐구영역선택과목수       
	RSFD_SBJ1		=	AryHash(i).Item("RSFD_SBJ1")		'탐구영역과목1            
	RSFD_SCR1		=	AryHash(i).Item("RSFD_SCR1")		'탐구영역표준점수1        
	RSFD_CNT1		=	AryHash(i).Item("RSFD_CNT1")		'탐구영역백분위1          
	RSFD_GRD1		=	AryHash(i).Item("RSFD_GRD1")		'탐구영역등급1              
	RSFD_SBJ2		=	AryHash(i).Item("RSFD_SBJ2")		'탐구영역과목2              
	RSFD_SCR2		=	AryHash(i).Item("RSFD_SCR2") 		'탐구영역표준점수2          
	RSFD_CNT2		=	AryHash(i).Item("RSFD_CNT2")		'탐구영역백분위2            
	RSFD_GRD2		=	AryHash(i).Item("RSFD_GRD2")		'탐구영역등급2              
	RSFD_SBJ3		=	AryHash(i).Item("RSFD_SBJ3")		'탐구영역과목3              
	RSFD_SCR3		=	AryHash(i).Item("RSFD_SCR3")		'탐구영역표준점수3          
	RSFD_CNT3		=	AryHash(i).Item("RSFD_CNT3")		'탐구영역백분위3            
	RSFD_GRD3		=	AryHash(i).Item("RSFD_GRD3")		'탐구영역등급3              
	RSFD_SBJ4		=	AryHash(i).Item("RSFD_SBJ4")		'탐구영역과목4              
	RSFD_SCR4		=	AryHash(i).Item("RSFD_SCR4")		'탐구영역표준점수4          
	RSFD_CNT4		=	AryHash(i).Item("RSFD_CNT4")		'탐구영역백분위4            
	RSFD_GRD4		=	AryHash(i).Item("RSFD_GRD4") 		'탐구영역등급4              
																		                          
	SCFL_EXFG		=	AryHash(i).Item("SCFL_EXFG")		'제2외국어영역응시구분      
	SCFL_SBJT		=	AryHash(i).Item("SCFL_SBJT")		'제2외국어영역과목          
	SCFL_SDSC		=	AryHash(i).Item("SCFL_SDSC")		'제2외국어표준점수          
	SCFL_CENT		=	AryHash(i).Item("SCFL_CENT")		'제2외국어백분위            
	SCFL_GRAD		=	AryHash(i).Item("SCFL_GRAD")		'제2외국어등급              
	REMK_TEXT		=	AryHash(i).Item("REMK_TEXT")		'비고                       

	INPT_USID		=	AryHash(i).Item("INPT_USID")		'입력자
	INPT_ADDR		=	AryHash(i).Item("INPT_ADDR")		'입력IP


	'///////////////////////////////////////////////
	'// 자격미달여부 체크
	'///////////////////////////////////////////////
	SQL = ""
	SQL = SQL & vbCrLf & "Select DocumentaryCheck1, DocumentaryCheck2, DocumentaryCheck3, DocumentaryCheck4, DocumentaryCheck5 "
	SQL = SQL & vbCrLf & "		,DocumentaryCheck6, DocumentaryCheck7, DocumentaryCheck8, StudentRecordAgreement "
	SQL = SQL & vbCrLf & "from ApplicationTable "
	SQL = SQL & vbCrLf & "where Myear = '" & SCHL_YEAR & "' "
	SQL = SQL & vbCrLf & "And StudentNumber = '" & EXAM_NUMB & "'; "

	'objDB.blnDebug = TRUE
	arrParams2 = objDB.fnGetArray
	AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)				

	If isArray(AryHash2) Then
		documentCheck2				= AryHash(0).Item("DocumentaryCheck2")
		documentCheck3				= AryHash(0).Item("DocumentaryCheck3")
		documentCheck4				= AryHash(0).Item("DocumentaryCheck4")
		documentCheck5				= AryHash(0).Item("DocumentaryCheck5")
		documentCheck6				= AryHash(0).Item("DocumentaryCheck6")
		documentCheck7				= AryHash(0).Item("DocumentaryCheck7")
		documentCheck8				= AryHash(0).Item("DocumentaryCheck8")
		StudentRecordAgreement      = AryHash(0).Item("StudentRecordAgreement")
	End If

	'=============== 자격미달여부 계산 (하드코딩)============================================================
	'====== 하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save, 4. 수능점수 엑셀 Save =====

	'//////////////////////////////////////////////////////////////////////////////////////////////////////
	'//자격미달자 기준별 계산식 (하드코딩) 1 ~ 8번 코드 (Y = 미달자)
	'//하드코딩 위치 : 1. 입학원서Proc, 2. 지원자Proc, 3.입학원서 엑셀 Save
	'//***** 2번코드 (정시용)입학원서 등록 시 체크, 수능성적 입력 시 체크 -> 자격미달여부 다시 체크
	'//////////////////////////////////////////////////////////////////////////////////////////////////////

	'////////////////////////////////////////////////////////////////////////////////////////////////
	'//2번코드 (공통-정시)국내 고등학교 졸업(예정)자로 수학능력시험 성적이 있는 자
	'////////////////////////////////////////////////////////////////////////////////////////////////

	'2번 코드(정시 자격미달 체크)
	If documentCheck2 = "3" Or documentCheck2 = "5" Then '학력/수능점수 미달이었을 시
		document2 = "5"	
		DrawStandard = "Y"
		DrawMsg = "<b>학력 자격미달(제출서류) :</b> 고등학교 생활기록부"
	ElseIf documentCheck2 = "10" Then '수능점수 미달이었을 시
		document2 = "1"	
	End If

	'4번 코드
	If documentCheck4 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "C"
		End If
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>농어촌1유형 자격미달(제출서류) :</b> 농어촌전형 추천서, 중/고등학교 생활기록부, 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
		Else
			DrawMsg = DrawMsg & "= <b>농어촌1유형 자격미달(제출서류) :</b> 농어촌전형 추천서, 중/고등학교 생활기록부, 지원자 본인 및 부모 주민등록 초본, 지원자 가족관계증명서"
		End If		
	End If

	'5번 코드
	If documentCheck5 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "C"
		End If
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>농어촌2유형 자격미달(제출서류) :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
		Else
			DrawMsg = DrawMsg & "= <b>농어촌2유형 자격미달(제출서류) :</b> 농어촌전형 추천서 1부, 초/중/고등학교 생활기록부, 지원자 본인 주민등록 초본"
		End If		
	End If

	'6번 코드
	If documentCheck6 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "C"
		End If
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

	'7번 코드
	If documentCheck7 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "C"
		End If
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

	'8번 코드
	If documentCheck8 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "C"
		End If
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>대학졸업자 자격미달(제출서류) :</b> 전적 대학 졸업(수료)증명서 1부. 및  성적증명서 1부"
		Else
			DrawMsg = DrawMsg & "= <b>대학졸업자 자격미달(제출서류) :</b> 전적 대학 졸업(수료)증명서 1부. 및  성적증명서 1부"
		End If		
	End If

	'3번 코드
	If documentCheck3 = "3" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "D"
		End If
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
		Else
			DrawMsg = DrawMsg & "= <b>면접/실기 자격미달 :</b> 면접/실기 점수를 업로드해주세요."
		End If			
	ElseIf documentCheck3 = "4" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "E"
		End If	
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
		Else
			DrawMsg = DrawMsg & "= <b>면접 미응시 자격미달 :</b> 면접 점수를 업로드해주세요."
		End If		
	ElseIf documentCheck3 = "5" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "F"
		End If
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
		Else
			DrawMsg = DrawMsg & "= <b>실기 미응시 자격미달 :</b> 실기 점수를 업로드해주세요."
		End If			
	ElseIf documentCheck3 = "6" Then
		If DrawStandard <> "Y" Then
			DrawStandard = "G"
		End If	
		If LEN(DrawMsg) < 1 Then
			DrawMsg = "<b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
		Else
			DrawMsg = DrawMsg & "= <b>면접/실기가 평가점수비율에 포함되어 있지 않습니다. 관리자에게 문의하세요."
		End If		
	End If

	'=============== 자격미달여부 계산 끝 ===============

	'=============== 입학원서 입력 ===============

	'// 수정 ================
	SQL = ""
	SQL = SQL & vbCrLf & "UPDATE ApplicationTable SET "
	SQL = SQL & vbCrLf & "		DocumentaryCheck2 = ?, DrawStandard = ?, DrawMsg =?  "
	SQL = SQL & vbCrLf & "		,UPDT_USID = ?, UPDT_DATE = getdate(), UPDT_ADDR = ?, InsertTime = getdate() "
	SQL = SQL & vbCrLf & " WHERE MYear = ? "
	SQL = SQL & vbCrLf & " AND StudentNumber = ? "

	'update일 때는 UPDT입력
	arrParams = Array(_
		  Array("@DocumentaryCheck2",			adInteger,		adParamInput,		0,		document2) _
		, Array("@DrawStandard",				adVarchar,		adParamInput,		255,	DrawStandard) _
		, Array("@DrawMsg",						adVarchar,		adParamInput,		5000,	DrawMsg) _	
		, Array("@UPDT_USID",					adVarchar,		adParamInput,		20,		INPT_USID) _
		, Array("@UPDT_ADDR",					adVarchar,		adParamInput,		20,		INPT_ADDR) _

		, Array("@MYear",						adVarchar,		adParamInput,		50,		SCHL_YEAR) _
		, Array("@StudentNumber",				adVarchar,		adParamInput,		50,		EXAM_NUMB) _
	)
	
	'objDB.blnDebug = true
	Call objDB.sbExecSQL(SQL, arrParams)

	InsertType = "Update"

	'SQL = " SELECT @@IDENTITY; "
	'aryList = objDB.fnExecSQLGetRows(SQL, nothing)
	'IDX = CInt(aryList(0, 0))\

	'=============== 입학원서 입력 끝 ===============

	count = count + 1

Next

'=============== 수능성적 입력에 따른 입학원서(자격미달여부) 수정 끝 ===============

'////////////////////////////////////
'// 등록 히스토리 
'////////////////////////////////////
strLogMSG = "지원자관리 > "& SCHL_YEAR &"학년도_"& count &"건의 수능성적이 엑셀 등록되었습니다."
InsertType = "Insert"

'트랜젝션 처리
If Err.Number <> 0 Then 
	strResult = "Error"
	returnMSG = Err.Number&":"&Err.Description
	'objDB.sbRollbackTrans
Else 
	strResult = "Complete"
	returnMSG = "수능성적 엑셀 저장 완료"
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