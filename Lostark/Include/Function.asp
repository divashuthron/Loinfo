<%
'Option Explicit
'Response.Expires = -1
%>
<!--metadata type="typelib" file="c:\program files\common files\system\ado\msado15.dll"-->
<!--#InClude Virtual = "/Lostark/Include/DBHelperClass.asp"-->
<!--#InClude Virtual = "/Lostark/Include/Function.CONFIG.asp" -->
<!--#InClude Virtual = "/Lostark/Include/Function.STRING.asp" -->
<!--#InClude Virtual = "/Lostark/Include/Function.XML.asp" -->
<!--#InClude Virtual = "/Lostark/Include/Function.BaseCode.asp" -->
<!--#InClude Virtual = "/Lostark/Include/Function.Log.asp" -->
<!--#InClude Virtual = "/Lostark/Include/Function.VisitCount.asp" -->
<%
'== 캐쉬 초기화 Proc =========================================
Call sbNoCacheProc()

'============= URL 관련 세팅 =================================
Dim ASP_SELF_DOMAIN			: ASP_SELF_DOMAIN			= "http://rhby.zulus.site/"
Dim ASP_SELF_URL			: ASP_SELF_URL				= Request.ServerVariables("URL")
Dim ASP_SELF_URL_All		: ASP_SELF_URL_All			= Request.ServerVariables("URL") & + "?" + request.ServerVariables("QUERY_STRING")
Dim ASP_SELF_ENCURL			: ASP_SELF_ENCURL			= fnGetURLEncode(ASP_SELF_URL_All)
Dim ASP_SELF_PATH			: ASP_SELF_PATH				= Request.ServerVariables("PATH_INFO")
Dim ASP_BEFORE_URL			: ASP_BEFORE_URL			= Request.ServerVariables("HTTP_REFERER")
Dim ASP_USER_IP				: ASP_USER_IP				= Request.ServerVariables("REMOTE_ADDR")

'============= 로그인 정보 세팅 ==============================
Dim SessionUserID			: SessionUserID				= base64_decode(Session("MemberID"))
'Dim SessionUserID			: SessionUserID				= base64_decode(Request.Cookies("InformationAdmin")("EmpID"))
Dim SessionUserName			: SessionUserName			= Request.Cookies("LoInfo")("NickName")
Dim SessionClientLevel		: SessionClientLevel		= Request.Cookies("LoInfo")("ClientLevel")

'============== 환경설정 정보 세팅 ============================
'Dim SessionMYear			: SessionMYear				= Request.Cookies("LoInfo")("MYear")
'Dim SessionDivision			: SessionDivision			= Request.Cookies("LoInfo")("Division")
'Dim SessionSubject			: SessionSubject			= Request.Cookies("LoInfo")("Subject")
'Dim SessionDivision1		: SessionDivision1			= Request.Cookies("LoInfo")("Division1")
'Dim SessionDivision2		: SessionDivision2			= Request.Cookies("LoInfo")("Division2")
'Dim SessionSchoolName		: SessionSchoolName			= Request.Cookies("LoInfo")("SchoolName")
'Dim SessionSchoolSmsNumber	: SessionSchoolSmsNumber	= Request.Cookies("LoInfo")("SchoolSmsNumber")
'Dim SessionApplyConfirm		: SessionApplyConfirm		= Request.Cookies("LoInfo")("ApplyConfirm")
'Dim SessionApplyPrintConfirm: SessionApplyPrintConfirm	= Request.Cookies("LoInfo")("ApplyPrintConfirm")
'Dim SessionInterviewConfirm	: SessionInterviewConfirm	= Request.Cookies("LoInfo")("InterviewConfirm")

' ##################################################################################
' 공백일때 대체 문자 처리
' ##################################################################################
  function getParameter(m,s)
    if m = "" or isNull(m) then
      getParameter = Trim(s)
    else
      getParameter = Trim(m)
    end if  
  end function

  function getIntParameter(im,s)
    if im = "" or not IsNumeric(im) then
      getIntParameter = Clng(Trim(s))
    else
      getIntParameter = Clng(Trim(im))
    end if  
  end function

'============= 년도 세팅 ====================================
Dim GlobalMYear				:GlobalMYear				= SessionMYear
'Dim GlobalMYear			:GlobalMYear				= getDateYear()
'Dim GlobalMYear			:GlobalMYear				= "2016"

'== Query 실행 후 Error 체크 =================================
Function dbError(ByRef dbConn)
	dbError = dbConn.Errors.Count
End Function

'== Query 실행 후 Error Message 리턴 =========================
Function dbErrorMsg(ByRef dbConn)
	errMsg = ""
	If dbConn.Errors.Count > 0 Then
		For each Error in dbConn.Errors
			errMsg = errMsg & "Error # :" & Error.Number & "\n"
			errMsg = errMsg & "Error Description : " & Error.Description  & "\n\n"
		Next
	End If
	dbErrorMsg = errMsg
End Function

'== 이번년도 구하기 ==========================================
Function getDateYear()
	getDateYear = Year(Now)
End Function

Function getDateMonth()
	getDateMonth = fnGetCipherProc(Month(Now))
End Function

Function getDateDay()
	getDateDay = fnGetCipherProc(Day(Now))
End Function

'== 오늘날짜 구하기 ==========================================
Function getDateNow(delimiter)
	getDateNow = Year(Now) & delimiter & fnGetCipherProc(Month(Now)) & delimiter & fnGetCipherProc(Day(Now))
End Function

'== 오늘날짜에서 지정된 월만큼 더한다 ========================
Function getTodayAddMonth(mm, delimiter)
	toDate = Date
	addDate = DateAdd("m", mm, toDate)
	arrDate = Split(addDate, "-")
	getTodayAddMonth = arrDate(0) & delimiter & fnGetCipherProc(arrDate(1)) & delimiter & fnGetCipherProc(arrDate(2))
End Function

'== 오늘날짜에서 지정된 일만큼 더한다 ========================
Function getTodayAddDay(mm, delimiter)
	toDate = Date
	addDate = DateAdd("d", mm, toDate)
	arrDate = Split(addDate, "-")
	getTodayAddDay = arrDate(0) & delimiter & fnGetCipherProc(arrDate(1)) & delimiter & fnGetCipherProc(arrDate(2))
End Function

'== 요일 구하기 ==========================================
Function getWeekName(DateValue)
	Dim WeekName : WeekName = weekDay(DateValue)
	select case WeekName
		case "1" WeekName = "일"
		case "2" WeekName = "월"
		case "3" WeekName = "화"
		case "4" WeekName = "수"
		case "5" WeekName = "목"
		case "6" WeekName = "금"
		case "7" WeekName = "토"
	end select
	getWeekName = WeekName
End Function

' ##################################################################################
' 쿼리 검색어 제한
' ##################################################################################
Function getQueryFilter( text )
	text = getParameter( text, "" )
	text = Replace(text,"'","")
	text = Replace(text,"--","")
	text = Replace(text,"insert","")
	text = Replace(text,"select","")
	text = Replace(text,"delete","")
	text = Replace(text,"update","")
	text = Replace(text,"or ","")
	getQueryFilter = text
End Function
%>