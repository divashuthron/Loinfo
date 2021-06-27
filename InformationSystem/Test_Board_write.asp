<%@  codepage="65001" language="VBScript"%>
<!DOCTYPE html>
<html lang="en-us">
<head>
<%
	Dim sessionId		:	sessionId = Session("MemberID")
	Dim sessionName		:	sessionName = Session("MemberName")
%>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<meta http-equiv="Expires" content="-1"> 
	<meta http-equiv="Pragma" content="no-cache"> 
	<meta http-equiv="Cache-Control" content="No-Cache"> 

	<title>테스트 보드-박송림</title>

	<meta name="description" content="">
	<meta name="author" content="">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

	<!-- Basic Styles -->
	<link rel="stylesheet" type="text/css" media="screen" href="/css/bootstrap.min.css">
	<link rel="stylesheet" type="text/css" media="screen" href="/css/animate.css">
	<link rel="stylesheet" type="text/css" media="screen" href="/css/style.css">

	<!-- your_style -->
	<link rel="stylesheet" type="text/css" media="screen" href="/css/your_style.css">



	<!-- Link to Google CDN's jQuery + jQueryUI; fall back to local -->
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
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
	<link rel="stylesheet" href="/Js/plugins/summernote/summernote.css">
	<script type="text/JavaScript" src="/Js/plugins/summernote/summernote.js"></script>
	<script type="text/JavaScript" src="/Js/plugins/summernote/lang/summernote-ko-KR.js"></script>
	<script type="text/JavaScript" src="/js/plugins/summernote/summernote-image-attributes.js"></script>
	<script type="text/JavaScript">
		var SID = "<%=sessionId%>";
		$(function(){
			$('#summernote').summernote({
				height:500, // 높이
				width:960,
				minHeight: null, //최소높이
				maxHeight: null, //최대높이
				focus: true,
				lang : "ko-Kr",
				placeholder:null
			});
			
			if(!SID || SID.length==0 || SID==""){
				alert("로그인이 필요합니다.");
				location.href='Test_Board_login.asp';
			}
		})

		function wirting() {
			$("#content1").val($('#summernote').summernote('code'));
		if(!SID || SID.length==0 || SID==""){
			alert("로그인이 필요합니다.");
			location.href='Test_Board_login.asp';
		}else{
			$('#boardfrm').submit();
		}
	}
	</script>
	<style>
		
		.container{
			min-width: 900px;
			max-width: none !important;
		}
		table, th, td{
			border-right: none !important;
			border-left: none !important;
			text-align: left !important;
		}
		.content-box{
			margin: 0 auto;
			max-width: 970px !important;
		}
	</style>

</head>
	
<body class="gray-bg">
	<div class="container">
		<div class="row content-box">
			<div class="col-xs-12 text-center">
				<h2>글작성</h2>
			</div>
			<form method="POST" action="Process/Test_Board_writProc.asp" id="boardfrm">
			<input type="hidden" name="content1" id="content1" />
				<table class="table table-bordered " style="background-color: white;">
					<tbody>
						<tr>
							<td>
								<div class="col-xs-1">
									<h4 style="line-height: 1.53em; text-align: center;">제목</h4>
								</div>
								<div class="col-xs-11">
									<input type="text" id="title" name="title" class="form-control" placeholder="제목을 입력해 주세요">
								</div>
							</td>
						</tr>
						<tr>
							<td>
								<div class="col-xs-1">
									<h4 style="line-height: 1.53em; text-align: center;">파일첨부</h4>
								</div>
								<div class="col-xs-11">
									<input type="file" value="파일 선택">
								</div>
							</td>
						</tr>
						<tr>
							<td>
								<div style="width: 960px; margin: 0 auto;">
									<textarea id="summernote" name="contents"> </textarea>
								</div>
							</td>
						</tr>
					</tbody>
				</table>
				<div style="float: right;">
						<button type="button" class="btn btn-labeled btn-primary" onclick="wirting()" > <i class="glyphicon glyphicon-check"></i> 작성완료 </button>
					<button type="button" class="btn btn-labeled bg-color-redLight txt-color-white" onclick= "location.href='Test_Board_list.asp'" > <i class="glyphicon glyphicon-random"></i> 돌아가기 </button>
				</div>
			</form>
		</div>
	</div>

	<!--<script src="js/jquery-3.1.1.min.js"></script>-->
	<script src="/js/bootstrap.min.js"></script>

	<!-- JQUERY SELECT2 INPUT -->
	<script src="/js/plugins/select2/select2.min.js"></script>

	<!-- MetisSoft -->
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.cookie.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.form.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/base64.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.simplemodal.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.blockUI.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.url.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.Calendar.1.7.js"></script>
	<script type="text/javascript" src="/Js/MetisSoft/Jquery.plugins/jquery.printThis.js"></script>

	<script type="text/javascript" src="/Js/MetisSoft/common.js"></script>
		
</body>
</html>