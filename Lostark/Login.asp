<%@  codepage="65001" language="VBScript" %>
<!-- InClude Virtual = "/Lostark/Include/Function.asp" -->
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8"/>
    <meta http-equiv="X-UA-Compatible" content="IE=edge"/>
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no"/>
    <meta name="description" content=""/>
    <meta name="author" content=""/>
    <title>[로아룸] 로스트아크 파티 매칭시스템</title>
    <!-- loader-->
    <link href="assets/css/pace.min.css" rel="stylesheet"/>
    <script src="assets/js/pace.min.js"></script>
    <!--favicon-->
    <link rel="icon" href="assets/images/favicon.ico" type="image/x-icon">
    <!-- Bootstrap core CSS-->
    <link href="assets/css/bootstrap.min.css" rel="stylesheet"/>
    <!-- animate CSS-->
    <link href="assets/css/animate.css" rel="stylesheet" type="text/css"/>
    <!-- Icons CSS-->
    <link href="assets/css/icons.css" rel="stylesheet" type="text/css"/>
    <!-- Custom Style-->
    <link href="assets/css/app-style.css" rel="stylesheet"/>

    <script src="http://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
    <script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
</head>
<body class="bg-theme bg-theme2">
<script language="javascript">
    $(function () {
        var $UserID = $("#UserID");
        var $Passwd = $("#Passwd");
        
        $("#goLogin").click(function () {
            $UserID = $("#UserID");
            $Passwd = $("#Passwd");

            $("#form").submit();
        });

        $("#form").ajaxForm({
            dataType: "xml",
            beforeSubmit: function () { },
            success: function (datas, state) {
                var $objList = $(datas).find("List");
                
                if ($objList.find("Result").text() == "true") {
                    document.location.href = "/Lostark/forms.html";
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

<!-- start loader -->
<div id="pageloader-overlay" class="visible incoming"><div class="loader-wrapper-outer"><div class="loader-wrapper-inner" ><div class="loader"></div></div></div></div>
<!-- end loader -->

<!-- Start wrapper-->
 <div id="wrapper">

 <div class="loader-wrapper"><div class="lds-ring"><div></div><div></div><div></div><div></div></div></div>
	<div class="card card-authentication1 mx-auto my-5">
		<div class="card-body">
		 <div class="card-content p-2">
		 	<div class="text-center">
		 		<img src="assets/images/logo-icon.png" alt="logo icon">
		 	</div>
		  <div class="card-title text-uppercase text-center py-3">[Lostark Room] 로아룸</div>
		    <form id="form" role="form" action="/Lostark/Process/LoginProc.asp" method="post">
			  <div class="form-group">
			  <label for="ID" class="sr-only">ID</label>
			   <div class="position-relative has-icon-right">
				  <input type="text" id="ID" name="ID" class="form-control input-shadow" placeholder="ID를 입력하세요">
				  <div class="form-control-position">
					  <i class="icon-user"></i>
				  </div>
			   </div>
			  </div>
			  <div class="form-group">
			  <label for="Password" class="sr-only">비밀번호</label>
			   <div class="position-relative has-icon-right">
				  <input type="password" id="Password" name="password" class="form-control input-shadow" placeholder="비밀번호를 입력하세요">
				  <div class="form-control-position">
					  <i class="icon-lock"></i>
				  </div>
			   </div>
			  </div>
			<div class="form-row">
			 <div class="form-group col-6">
			   <div class="icheck-material-white">
                <input type="checkbox" id="Save" name="Save" checked="" />
                <label for="user-checkbox">ID 저장</label>
			  </div>
			 </div>
       <!--
			 <div class="form-group col-6 text-right">
			  <a href="reset-password.html">비밀번호 재설정</a>
			 </div>
       -->
			</div>
			 <span id="goLogin" class="btn btn-light btn-block">로그인</span>
			  <!--<div class="text-center mt-3">비밀번호 찾기</div>-->
			  
        <!--
			 <div class="form-row mt-4">
			  <div class="form-group mb-0 col-6">
			   <button type="button" class="btn btn-light btn-block"><i class="fa fa-comment"></i> 카카오톡 로그인</button>
			 </div>
			 <div class="form-group mb-0 col-6 text-right">
			  <button type="button" class="btn btn-light btn-block"><i class="fa fa-google"></i> 구글 로그인</button>
			 </div>
			</div>
      -->
			 
			 </form>
		   </div>
		  </div>
		  <div class="card-footer text-center py-3">
		    <p class="text-warning mb-0">아직 회원이 아니신가요? <a href="">회원가입</a></p>
		  </div>
	     </div>
    
     <!--Start Back To Top Button-->
    <a href="javaScript:void();" class="back-to-top"><i class="fa fa-angle-double-up"></i> </a>
    <!--End Back To Top Button-->
	
	<!--start color switcher-->
  <!--
   <div class="right-sidebar">
    <div class="switcher-icon">
      <i class="zmdi zmdi-settings zmdi-hc-spin"></i>
    </div>
    <div class="right-sidebar-content">

      <p class="mb-0">Gaussion Texture</p>
      <hr>
      
      <ul class="switcher">
        <li id="theme1"></li>
        <li id="theme2"></li>
        <li id="theme3"></li>
        <li id="theme4"></li>
        <li id="theme5"></li>
        <li id="theme6"></li>
      </ul>

      <p class="mb-0">Gradient Background</p>
      <hr>
      
      <ul class="switcher">
        <li id="theme7"></li>
        <li id="theme8"></li>
        <li id="theme9"></li>
        <li id="theme10"></li>
        <li id="theme11"></li>
        <li id="theme12"></li>
		<li id="theme13"></li>
        <li id="theme14"></li>
        <li id="theme15"></li>
      </ul>
      
     </div>
   </div>
   -->
  <!--end color switcher-->
	
	</div><!--wrapper-->
	
  <!-- Bootstrap core JavaScript-->
  <script src="assets/js/jquery.min.js"></script>
  <script src="assets/js/popper.min.js"></script>
  <script src="assets/js/bootstrap.min.js"></script>
	
  <!-- sidebar-menu js -->
  <script src="assets/js/sidebar-menu.js"></script>
  
  <!-- Custom scripts -->
  <script src="assets/js/app-script.js"></script>
  <script src="assets/js/jquery.form.min.js"></script>
  
</body>
</html>