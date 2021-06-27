<%@  codepage="65001" language="VBScript" %>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<!DOCTYPE html>
<html lang="en-us">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<meta http-equiv="Expires" content="-1"> 
	<meta http-equiv="Pragma" content="no-cache"> 
	<meta http-equiv="Cache-Control" content="No-Cache"> 

	<title>대학종합정보시스템 - (주)메티소프트</title>


	<meta name="description" content="">
	<meta name="author" content="">
		
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">

	<!-- Basic Styles -->
	<link rel="stylesheet" type="text/css" media="screen" href="/css/bootstrap.min.css">
	<link rel="stylesheet" type="text/css" media="screen" href="/css/font-awesome.min.css">

	<link rel="stylesheet" type="text/css" media="screen" href="/css/animate.css">
	<link rel="stylesheet" type="text/css" media="screen" href="/css/style.css">

	<!-- your_style -->
	<link rel="stylesheet" type="text/css" media="screen" href="/css/your_style.css">

	<style>
		.logo-name {
			font-size: 100px !important;
		}
		.loginscreen {
			width: 400px !important;
		}
		.loginform {
			padding:0 30px 0 30px;
		}
	</style>


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
</head>
	
<body class="gray-bg">


	<script language="javascript">
		$(function () {
			var $UserID = $("#UserID");
			var $Passwd = $("#Passwd");
			
			$("#goLogin").click(function () {
				if (!$.chkInputValue($UserID,		"아이디를 입력해 주시기 바랍니다.")) { return; }
				if (!$.chkInputValue($Passwd,		"비밀번호를 입력해 주시기 바랍니다.")) { return; }
				
				$("#form").submit();
			});
				
			$("#form").ajaxForm({
				dataType: "xml",
				beforeSubmit: function () { },
				success: function (datas, state) {
					var $objList = $(datas).find("List");
					
					if ($objList.find("Result").text() == "true") {
						document.location.href = "/Lostark/Index.asp";
					} else {
						alert("입력하신 아이디 혹은 비밀번호가 일치하지 않습니다.");
						$UserID.focus();
					}
				},
				error: function (reason, e) {
					alert('서버연결에 실패했습니다.');
				}
			});

			$UserID.keyup(function (event) {
				if (event.keyCode == 13) {
					$Passwd.focus();
				}
			});
			
			$Passwd.keyup(function () {
				if (event.keyCode == 13) {
					$("#goLogin").click();
				}
			});

			$(window).load(function () {
				$("#UserID").focus();
			});
		});
	</script>


    <div class="middle-box text-center loginscreen animated fadeInDown">
        <div>
            <div>

                <h1 class="logo-name">MetisSoft</h1>

            </div>
            <h2 class="bold">대학종합정보시스템</h2>
            <p>Perfectly designed and precisely prepared admin theme with over 50 pages with extra new web app views.
                <!--Continually expanded and constantly improved Inspinia Admin Them (IN+)-->
            </p>
            <p>Login in. To see it in action.</p>
			<div class="loginform">
				<form class="m-t" role="form" id="form" action="/Process/LoginProc.asp" method="post">
					<div class="form-group">
						<input type="text" name="UserID" id="UserID" maxlength="20" class="form-control" placeholder="아이디" required="">
					</div>
					<div class="form-group">
						<input type="password" name="Passwd" id="Passwd" maxlength="20" class="form-control" placeholder="비밀번호" required="">
					</div>
					<span id="goLogin" class="btn btn-primary block full-width m-b" style="height:30px">로그인</span>


					<p class="text-muted text-center"><small>※ <span style="color:#e5322b;">비밀번호</span>는 영문 + 숫자 + 특수문자 조합으로 8자리 이상 입력</small></p>
					<!--<a class="btn btn-sm btn-white btn-block" href="register.html">Create an account</a>-->
				</form>
			</div>
            <p class="m-t"> <small>Copyright © 2019 | MetisSoft, Inc.</small> </p>
        </div>
    </div>

	<!-- Mainly scripts -->
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
	<!--<script type="text/javascript" src="/Js/MetisSoft/common.MetisSoft.js"></script>-->

</body>
</html>