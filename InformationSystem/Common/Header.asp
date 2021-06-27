<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
If IsE(SessionUserID) Then
	Response.Write "<script language='javascript'>"
	Response.Write "location.href='/Login.asp';"
	Response.Write "</script>"
	Response.End
End If
%>

<!DOCTYPE html>
<html>

<head>

    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">


	<title>대학종합정보시스템 - (주)메티소프트</title>



    <!-- Default Style -->
	<link href="/css/bootstrap.min.css" rel="stylesheet">
    <link href="/font-awesome/css/font-awesome.css" rel="stylesheet">

    <!-- Plugins Style-->
	<link href="/css/plugins/iCheck/custom.css" rel="stylesheet">
    <link href="/css/plugins/select2/select2.min.css" rel="stylesheet">
	<link href="/css/plugins/chosen/bootstrap-chosen.css" rel="stylesheet">
	<link href="/css/plugins/dataTables/datatables.min.css" rel="stylesheet">
	<link href="/css/plugins/awesome-bootstrap-checkbox/awesome-bootstrap-checkbox.css" rel="stylesheet"> 
	<link href="/css/plugins/clockpicker/clockpicker.css" rel="stylesheet">
	
	<!--
	<link href="/css/plugins/chosen/bootstrap-chosen.css" rel="stylesheet">
    <link href="/css/plugins/bootstrap-tagsinput/bootstrap-tagsinput.css" rel="stylesheet">
    <link href="/css/plugins/colorpicker/bootstrap-colorpicker.min.css" rel="stylesheet">
    <link href="/css/plugins/cropper/cropper.min.css" rel="stylesheet">
    <link href="/css/plugins/switchery/switchery.css" rel="stylesheet">
    <link href="/css/plugins/jasny/jasny-bootstrap.min.css" rel="stylesheet">
    <link href="/css/plugins/nouslider/jquery.nouislider.css" rel="stylesheet">
    <link href="/css/plugins/datapicker/datepicker3.css" rel="stylesheet">
    <link href="/css/plugins/ionRangeSlider/ion.rangeSlider.css" rel="stylesheet">
    <link href="/css/plugins/ionRangeSlider/ion.rangeSlider.skinFlat.css" rel="stylesheet">
    <link href="/css/plugins/daterangepicker/daterangepicker-bs3.css" rel="stylesheet">
	<link href="/css/plugins/touchspin/jquery.bootstrap-touchspin.min.css" rel="stylesheet">
    <link href="/css/plugins/dualListbox/bootstrap-duallistbox.min.css" rel="stylesheet">
	-->

	<!-- Basic Style -->
	<link href="/css/animate.css" rel="stylesheet">
    <link href="/css/style.css" rel="stylesheet">

	<!-- Your Style -->
	<link href="/css/your_style.css" rel="stylesheet">

	<!-- FAVICONS -->
	<link rel="shortcut icon" href="/img/favicon/favicon.ico" type="image/x-icon">
	<link rel="icon" href="/img/favicon/favicon.ico" type="image/x-icon">

	<!-- GOOGLE FONT -->
	<link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Open+Sans:400italic,700italic,300,400,700">

	<!-- Link to Google CDN's jQuery + jQueryUI; fall back to local -->
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>

	<!-- daum 주소찾기 -->
	<script src="http://dmaps.daum.net/map_js_init/postcode.v2.js"></script>

	<script>
		if (!window.jQuery) {
			document.write('<script src="js/libs/jquery-2.1.1.min.js"><\/script>');
		}
	</script>

	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
	<script>
		if (!window.jQuery.ui) {
			document.write('<script src="js/libs/jquery-ui-1.10.3.min.js"><\/script>');
		}
	</script>
</head>

<body>

	<div id="wrapper">

		<nav class="navbar-default navbar-static-side" role="navigation">
			<div class="sidebar-collapse">
				<ul class="nav metismenu" id="side-menu">
					<li class="nav-header">
						<div class="dropdown profile-element">
								<!-- 메인타이틀과 아이디 시작 -->
								<!--<a data-toggle="dropdown" class="dropdown-toggle" href="/">-->
								<a href="/">
									<span class="clear"> <span class="block m-t-xs"> <strong class="font-bold" style="font-size:25px;">메티스대학교</strong>
									</span> <span class="text-muted text-xs block" style="font-size:20px;"><%= SessionUserName %> <b class="caret"></b></span> </span>
								</a>
								 <!-- 메인타이틀과 아이디 끝 -->
								 <!--아이디 누르면 드롭다운 시작 -->
								 <!--
								<ul class="dropdown-menu animated fadeInRight m-t-xs">
									<li><a href="/Logout.asp">Logout</a></li>
								</ul>
								-->
								 <!--아이디 누르면 드롭다운 끝 -->
						</div>
						<!-- 메뉴축소 버튼 누르면 시작 -->
						<div class="logo-element">
							IN+
						</div>
						<!-- 메뉴축소 버튼 누르면 끝 -->
					</li>
					<!-- 메뉴바 시작 -->
					<!-- 코드관리 -->
					<li class="<% If TopMenuSeq = "1" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-cc"></i> <span class="nav-label">코드관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Code" Then Response. write "active" End If %>"><a href="/CodeList.asp">코드 등록</a></li>
							<li class="<% If LeftMenuCode = "CodeHistory" Then Response. write "active" End If %>"><a href="/CodeHistory.asp">코드 등록 히스토리</a></li>
						</ul>
					</li>
					<!-- 모집단위관리 -->
					<li class="<% If TopMenuSeq = "2" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-sitemap"></i> <span class="nav-label">모집단위관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Subject" Then Response. write "active" End If %>"><a href="/SubjectList.asp">모집단위 등록</a></li>
							<li class="<% If LeftMenuCode = "SubjectHistory" Then Response. write "active" End If %>"><a href="/SubjectHistory.asp">모집단위 히스토리</a></li>
						</ul>
					</li>
					<!-- 평가기준관리 -->
					<li class="<% If TopMenuSeq = "3" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-tasks"></i> <span class="nav-label">평가기준관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Appraisal" Then Response. write "active" End If %>"><a href="/AppraisalList.asp">평가비율 설정</a></li>
							<li class="<% If LeftMenuCode = "StudentRecord" Then Response. write "active" End If %>"><a href="/StudentRecord.asp">생기부 설정</a></li>
							<li class="<% If LeftMenuCode = "CSAT" Then Response. write "active" End If %>"><a href="/CSAT.asp">수능 설정</a></li>
							<li class="<% If LeftMenuCode = "AppraisalHistory" Then Response. write "active" End If %>"><a href="/AppraisalHistory.asp">평가기준 히스토리</a></li>
						</ul>
					</li>
					<!-- 입학원서관리 -->
					<li class="<% If TopMenuSeq = "4" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-address-card-o"></i> <span class="nav-label">입학원서관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Application" Then Response. write "active" End If %>"><a href="/ApplicationList.asp">입학원서 조회</a></li>
							<li class="<% If LeftMenuCode = "ApplicationAdd" Then Response. write "active" End If %>"><a href="/ApplicationAddList.asp">입학원서 수동입력</a></li>
							<li class="<% If LeftMenuCode = "ApplicationHistory" Then Response. write "active" End If %>"><a href="/ApplicationHistory.asp">입학원서 히스토리</a></li>
						</ul>
					</li>
					<!-- 지원자관리 -->
					<li class="<% If TopMenuSeq = "5" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-address-card"></i> <span class="nav-label">지원자관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Applicant" Then Response. write "active" End If %>"><a href="ApplicantList.asp">지원자 조회</a></li>
							<li class="<% If LeftMenuCode = "ApplicantAdd" Then Response. write "active" End If %>"><a href="ApplicantAddList.asp">지원자 서류체크</a></li>
							<li class="<% If LeftMenuCode = "ApplicantHistory" Then Response. write "active" End If %>"><a href="ApplicantHistory.asp">지원자 히스토리</a></li>
						</ul>
					</li>
					<!-- 사정관리 -->
					<li class="<% If TopMenuSeq = "6" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-check"></i> <span class="nav-label">사정관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Assessment" Then Response. write "active" End If %>"><a href="Assessment.asp">사정처리</a></li>
							<li class="<% If LeftMenuCode = "AssessmentHistory" Then Response. write "active" End If %>"><a href="AssessmentHistory.asp">사정 히스토리</a></li>
						</ul>
					</li>
					<!-- 합격자발표관리 -->
					<li class="<% If TopMenuSeq = "7" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-bullhorn"></i> <span class="nav-label">합격자발표관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "SucCandidate" Then Response. write "active" End If %>"><a href="SuccessfulStudent.asp">합격자 조회</a></li>
							<li class="<% If LeftMenuCode = "Bill" Then Response. write "active" End If %>"><a href="BillList.asp">고지서 설정</a></li>
							<li class="<% If LeftMenuCode = "Demands" Then Response. write "active" End If %>"><a href="DemandsList.asp">유의사항 설정</a></li>							
							<li class="<% If LeftMenuCode = "Report" Then Response. write "active" End If %>"><a href="Report.asp">레포트 출력</a></li>
							<li class="<% If LeftMenuCode = "Report2" Then Response. write "active" End If %>"><a href="Report2.asp">레포트 출력2</a></li>
							<li class="<% If LeftMenuCode = "SuccessfulStudentHistory" Then Response. write "active" End If %>"><a href="SuccessfulStudentHistory.asp">합격자 히스토리</a></li>
						</ul>
					</li>
					<!-- 통계관리 -->
					<li class="<% If TopMenuSeq = "8" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-bar-chart-o"></i> <span class="nav-label">통계관리</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "Charts" Then Response. write "active" End If %>"><a href="Charts.asp">사정 현황표</a></li>
							<li class="<% If LeftMenuCode = "DetailsCharts" Then Response. write "active" End If %>"><a href="DetailsCharts.asp">세부 사정 현황표</a></li>
							<li class="<% If LeftMenuCode = "Statistics" Then Response. write "active" End If %>"><a href="Statistics.asp">종합통계</a></li>
						</ul>
					</li>
					<!-- 환경설정 -->
					<li class="<% If TopMenuSeq = "9" Then Response. write "active" End If %>">
						<a href="#"><i class="fa fa-cog"></i> <span class="nav-label">환경설정</span><span class="fa arrow"></span></a>
						<ul class="nav nav-second-level collapse">
							<li class="<% If LeftMenuCode = "DefaultConfig" Then Response. write "active" End If %>"><a href="/DefaultConfig.asp">기본환경설정</a></li>
							<li class="<% If LeftMenuCode = "Employee" Then Response. write "active" End If %>"><a href="/EmployeeList.asp">사용자관리</a></li>
						</ul>
					</li>
					<!-- 메뉴바 끝 -->
				</ul>

			</div>
		</nav>

		<div id="page-wrapper" class="gray-bg">
			<div class="row border-bottom">
				<nav class="navbar navbar-static-top white-bg" role="navigation" style="margin-bottom: 0">
					<div class="navbar-header">
						<a class="navbar-minimalize minimalize-styl-2 btn btn-primary " href="#"><i class="fa fa-bars"></i> </a>
							<div class="navbar-form-custom">
							   <div class="form-control"><%= LeftMenuName %></div>
							</div>
						<!--
						<form role="search" class="navbar-form-custom" method="post" action="#">
							<div class="form-group">
							   <input type="text" placeholder="Search for something..." class="form-control" name="top-search" id="top-search">
							</div>
						</form>
						-->
					</div>
					<ul class="nav navbar-top-links navbar-right">
						<li class="dropdown">
						<!-- 알림 시작 -->
						<%
								'미확인 건수, 반복획수(최대 5번, 5번보다 적으면 배열 미확인 갯수), 메뉴아이콘
								Dim AryHashCnt, ForNum, iconStyle

								Set objDB = New clsDBHelper
								objDB.strConnectionString = strDBConnString
								objDB.sbConnectDB
								
								'미확인 히스토리만 열기
								SQL = ""
								SQL = SQL & vbCrLf & "select a.Division, a.ActivityContent, a.RegDate "
								SQL = SQL & vbCrLf & "from ActivityHistory a left outer join AlarmHistory b "
								SQL = SQL & vbCrLf & "on a.IDX = b.HistoryIDX "
								SQL = SQL & vbCrLf & "where a.MYear = " & SessionMYear
								SQL = SQL & vbCrLf & "and b.IDX is null  " 
								SQL = SQL & vbCrLf & "order by a.IDX desc  "

								'objDB.blnDebug = TRUE
								arrParams = objDB.fnGetArray
								AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)		

								Set objDB = Nothing
								
								'반복획수(최대 5번, 5번보다 적으면 배열 미확인 갯수)
								If isArray(AryHash) Then
									AryHashCnt = ubound(AryHash,1) + 1
									If AryHashCnt > 4 Then
										ForNum = 4
									Else
										ForNum = AryHashCnt -1
									End if
									%>
								<a class="dropdown-toggle count-info" data-toggle="dropdown" href="#">
									<i class="fa fa-bell"></i>  <span class="label label-primary"><%=AryHashCnt%></span>
								</a>
								<ul class="dropdown-menu dropdown-alerts">
								<%								
									'For i = 0 to ubound(AryHash,1)
									For i = 0 to ForNum
										'메뉴아이콘 정하기
										Select Case AryHash(i).Item("Division")
											Case "Code"
												iconStyle = "fa fa-cc"
											Case "Subject", "SubjectExcelSave"
												iconStyle = "fa fa-sitemap fa-fw"
											Case "Appraisal", "StudentRecord", "CSAT", "BasicDataSet"
												iconStyle = "fa fa-tasks"
											Case "Login"
												iconStyle = "fa fa-sign-in"
											Case "LogOut"
												iconStyle = "fa fa-sign-out"
											Case "Application", "ApplicationList", "ApplicationAddList", "ApplicationExcelSave"
												iconStyle = "fa fa-address-card"
											Case "ApplicantList", "ApplicantProc", "StudentRecordAdd", "CSATAdd", "ApplicantAddList", "CSATExcelSave"
												iconStyle = "fa fa-address-card"
											Case "EmployeeList", "EmployeeView", "EmployeeProc", "DefaultConfig"
												iconStyle = "fa fa-cog"
											Case "SuccessfulStudentList", "BillProc", "DemandsProc", "DrawStandingSet"
												iconStyle = "fa fa-bullhorn"
											Case "AssessmentProc"
												iconStyle = "fa fa-check"
										End Select
								%>
									<li>
										<a href="unconfirmedHistory.asp">
											<div>
												<i class="<%=iconStyle%>"></i> <%= AryHash(i).Item("ActivityContent") %>
												<span class="pull-right text-muted small"><%= AryHash(i).Item("RegDate") %></span>
											</div>
										</a>
									</li>
									<li class="divider"></li>

								<%
									Next
								Else
								%>
								<a class="dropdown-toggle count-info" href="#">
									<i class="fa fa-bell"></i>  <span class="label label-primary"></span>
								</a>
								<ul class="dropdown-menu dropdown-alerts">
								<%
								end If						
						%>
									<li>
										<div class="text-center link-block">
											<a href="unconfirmedHistory.asp">
												<strong>모두 보기</strong>
												<i class="fa fa-angle-right"></i>
											</a>
										</div>
									</li>
								</ul>
							</li>
						<!-- 알림 끝 -->

						<!-- 로그아웃 시작 -->
						<li>
							<a href="/Logout.asp">
								<i class="fa fa-sign-out"></i> Log out
							</a>
						</li>
						<!-- 로그아웃 끝 -->
					</ul>

				</nav>
			</div>

            <div class="row wrapper border-bottom white-bg page-heading">
                <div class="col-lg-10">
                    <h2><%= LeftMenuNameDetail %></h2>
                </div>
            </div>

			<!-- 메인 컨텐츠 -->
			<!--<div class="wrapper wrapper-content animated fadeInRight">-->
			<div class="wrapper wrapper-content">		
