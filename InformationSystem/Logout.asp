<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
	Dim LogDivision			: LogDivision = "LogOut"
	Dim strLogMSG
	strLogMSG = "로그아웃  > " & SessionUserID & "가 로그아웃 하였습니다."

	'// 로그기록
	Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)


	'Response.Cookies("InterviewAdmin").Domain = "info.metissoft.co.kr"
	'// 세션삭제
	Session.Abandon
	'// 쿠키삭제
	Response.Cookies("InformationAdmin")("EmpID") = ""
	Response.Cookies("InformationAdmin")("EmpName") = ""
	Response.Cookies("InformationAdmin")("ClientLevel") = ""
	Response.Cookies("InformationAdmin")("MYear") = ""
	Response.Cookies("InformationAdmin")("Division") = ""
	'Response.Cookies("InformationAdmin")("Subject") = ""
	'Response.Cookies("InformationAdmin")("Division1") = ""
	'Response.Cookies("InformationAdmin")("Division2") = ""
	'Response.Cookies("InformationAdmin")("SchoolName") = ""
	'Response.Cookies("InformationAdmin")("SchoolSmsNumber") = ""
	'Response.Cookies("InformationAdmin")("ApplyConfirm") = ""
	'Response.Cookies("InformationAdmin")("ApplyPrintConfirm") = ""
	'Response.Cookies("InformationAdmin")("InterviewConfirm") = ""	
%>

<script langauge="javascript">
    document.location.href = "/Login.asp"
</script>