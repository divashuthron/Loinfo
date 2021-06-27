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
	<div class="animated fadeInDown" style="width: 300px; margin: 0 auto;">
		<div>

			<div class="form-group text-center">
				<h1>Test Board Lend</h1>
			</div>

			<div class="form-group text-center">
				<h1>회원가입</h1>
			</div>

			<div class="form-group">
				<form action="Process/Test_Board_registerPorc.asp" method="post" role="form" id="form">

					<div class="form-group">
						<input id="Mname" type="text" name="memberName" class="form-control" placeholder="이름">
						<div id="error_name" class="text-danger animated fadeInDown m-l-sm" style="display: none;">필수 정보입니다.</div>
					</div>

					<div class="form-group">
						<input id="Mid" type="text" name="memberID" class="form-control" placeholder="아이디">
						<div id="error_id" class="text-danger animated fadeInDown m-l-sm" style="display: none;">필수 정보입니다.</div>
					</div>

					<div class="form-group">
						<input type="password" id="pw" name="memberPW" class="form-control" placeholder="비밀번호">
					</div>

					<div class="form-group">
						<input type="password" id="pw2" name="memberPWck" class="pw_ck form-control" placeholder="비밀번호확인">
						<div id="error_pwck" class="text-danger animated fadeInDown m-l-sm" style="display: none;">비밀번호가 다릅니다.</div>
					</div>

					<div class="form-group text-center" style="justify-content: space-around; display: flex;">
						<button type="button" name="login" class="btn btn-info" style="height:30px; width: 100px;" onclick="checkForm();">가입완료</button>
						<button name="register" type="button" class="btn btn-warning" style="height:30px; width: 100px;" onclick="location.href='Test_Board_login.asp'">돌아가기</button>
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
	<!--<script type="text/javascript" src="/Js/MetisSoft/common.js"></script>-->

	<script type="text/javascript">
		checkcode = 0;

		// 비어있는지 검사
		$('.form-control').blur(function(){
			var input_box = $(this);
			var input_text = input_box.val();
			var input_name = input_box.attr('name');
			var err_txt="";
			if(!input_text && input_name != 'memberPWck'){
				input_box.next().css("display","block");
				input_box.next().addClass('text-danger');
				input_box.next().text("필수 정보 입니다.");
			}else{
				if(input_name == "memberName"){
					var nameCheck = RegExp(/^[가-힣]{2,6}$/);
					if (nameCheck.test(input_text)) {
						input_box.next().css("display", "none");
					} else {
						input_box.next().text("2~6글자의 한글만 입력 가능합니다.");
						input_box.next().css("display", "block");
					}
				}
				if (input_name == "memberID") {
					userIdCheck = RegExp(/^[A-za-z]+[A-za-z0-9]{5,19}$/g);
					if (userIdCheck.test(input_text)) {
						$.ajax({
							url: "Process/Test_Board_register_IDcheck.asp",
							data: { 'input_text': input_text },
							type: "post",
							dataType: "xml",
							cache: false,
							success: function (xml) {
								datas = $(xml).find('itemlist').find('datas').text();
								if (datas == 1) {
									$("#error_id").text('중복된 아이디 입니다.');
									// input_box.parent().addClass('has-danger');
									input_box.next().css("display", "block");
									input_box.next().removeClass('text-info');
									input_box.next().addClass('text-danger');
								} else {
									$("#error_id").text('사용가능한 아이디 입니다.');
									input_box.next().css("display", "block");
									input_box.next().removeClass('text-danger');
									input_box.next().addClass('text-info');
								}
								return;
							},
							error: function (xhr, ajaxOptions, thrownError) {
								alert("error!!!");
								//alert("statusText : "+xhr.statusText);
								//alert("responseText : "+xhr.responseText);
								alert("thrownError : " + xhr + '\n\n' + ajaxOptions + '\n\n' + thrownError);
								return;
							}
						});
					} else {
						input_box.next().text("아이디는 영문자로 시작하는 5~20자 영문자 또는 숫자이어야 합니다.");
						input_box.next().removeClass('text-info');
						input_box.next().addClass('text-danger');
						input_box.next().css("display","block");
					}
				}
				
			}
		});

		// 비밀번호 확인
		$('.pw_ck').blur(function(){
			var pw1 = $("#pw").val();
			var pw2 = $("#pw2").val();
			if(pw1 != pw2){
				$(this).next().css("display","block");
			}else if(pw1 == pw2){
				$(this).next().css("display","none");
			}
		});

		// 폼전송
		function checkForm(){
			var form = $('#form');
			var frmArr = form.serializeArray();
			
			if(!$('#Mname').val()){checkcode=1}
			if(!$('#Mid').val()){checkcode=1}
			if(!$('#pw').val()){checkcode=1}
			if(!$('#pw2').val()){checkcode=1}
			if(checkcode==0){
				form.submit();
			}else{
				alert("정보를 모두 입력해 주세요.");
			}
		}


	</script>
</body>

</html>