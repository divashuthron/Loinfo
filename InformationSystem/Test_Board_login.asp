<%@  codepage="65001" language="VBScript" %>
<!DOCTYPE html>
<html lang="en-us">
<head>
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

	<style>

	</style>
</head>
	
<body class="gray-bg">
	<div class="text-center animated fadeInDown" style="width: 300px; margin: 0 auto;">
		<div >

			<div class="form-group">
				<h1>Test Board Lend</h1>
			</div>

			<div class="form-group">
				<form action="Process/Test_Board_loginProc.asp" method="post" role="form" id="form">
					<div class="form-group">
						<input type="text" name="memberID" placeholder="아이디" class="form-control">
					</div>
					<div class="form-group">
						<input type="password" name="memberPW" placeholder="비밀번호" class="form-control">
					</div>
					<div class="form-group">
						<input type="submit" name="login" value="로그인" class="btn btn-primary block full-width m-b" style="height:30px">
					</div>	
					<div class="form-group">
						<button name="register" type="button" class="btn btn-success block full-width m-b" style="height:30px" onclick="location.href='Test_Board_register.asp'">회원가입</button>
					</div>
				</form>
			</div>
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