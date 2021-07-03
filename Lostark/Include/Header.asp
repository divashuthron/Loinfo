<!--#include Virtual="/Lostark/include/Function.asp"-->

<%
    Dim HeaderObjDB, HeaderSQL, HeaderAryHash, HeaderStrWhere
    Dim HeaderNickName
    Dim HeaderCNickName, HeaderCDivision, HeaderCJob, HeaderCLevel, HeaderCImgSrc

    '// 대표캐릭터가 없을 경우의 로고, 배경이미지
    Const LogoImgSrc = "https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/logo.png"
    Const BackgroundImgSrc = "https://cdn-lostark.game.onstove.com/uploadfiles/banner/8e7f80c7a9584c92960eff87fd6f6f9c.jpg"

    '// DB 오픈
    Set HeaderObjDB = New clsDBHelper
    HeaderObjDB.strConnectionString = strDBConnString
    HeaderObjDB.sbConnectDB

    '// 서브쿼리
    HeaderStrWhere = HeaderStrWhere & VbCrLf & "    And M.ID = '" & SessionUserID & "'"

    '// SQL
    HeaderSQL = ""
    HeaderSQL = HeaderSQL & VbCrLf & "    Select "
    HeaderSQL = HeaderSQL & VbCrLf & "        M. NickName "
    HeaderSQL = HeaderSQL & VbCrLf & "	      , CL.CharacterNickName "
    HeaderSQL = HeaderSQL & VbCrLf & "	      , CL.CharacterDivision "
    HeaderSQL = HeaderSQL & VbCrLf & "	      , CL.CharacterJob "
    HeaderSQL = HeaderSQL & VbCrLf & "	      , CL.CharacterLevel "
    HeaderSQL = HeaderSQL & VbCrLf & "	      , CL.CharacterImgSrc "
    HeaderSQL = HeaderSQL & VbCrLf & "    From Member M "
    HeaderSQL = HeaderSQL & VbCrLf & "    Left Outer Join CharacterList CL "
    HeaderSQL = HeaderSQL & VbCrLf & "    On M.ID = CL.ID  "
    HeaderSQL = HeaderSQL & VbCrLf & "    Where 1=1 "
    '// 대표 캐릭터만
    HeaderSQL = HeaderSQL & VbCrLf & "    And CL.CharacterIsMain = 'Y' "
    HeaderSQL = HeaderSQL & VbCrLf & HeaderStrWhere

    'HeaderObjDB.blnDebug = true
    'arrParams = HeaderObjDB.fnGetArray
    'aryList = HeaderObjDB.fnExecSQLGetRows(SQL, arrParams)
    HeaderAryHash = HeaderObjDB.fnExecSQLGetHashMap(HeaderSQL, Null)

    If IsArray(HeaderAryHash) Then
        '// 사용자 정보
        HeaderNickName = HeaderAryHash(0).Item("NickName")

        '// 캐릭터 정보
        HeaderCNickName = HeaderAryHash(0).Item("CharacterNickName")
        HeaderCDivision = HeaderAryHash(0).Item("CharacterDivision")
        HeaderCJob = HeaderAryHash(0).Item("CharacterJob")
        HeaderCLevel = HeaderAryHash(0).Item("CharacterLevel")
        HeaderCImgSrc = HeaderAryHash(0).Item("CharacterImgSrc")
    Else
        HeaderCJob = "대표 캐릭터 정보가 없습니다."
    End If

    '// DB 클로즈
    Set HeaderObjDB = Nothing
%>

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
  <link href="/Lostark/assets/css/pace.min.css" rel="stylesheet"/>
  <script src="/Lostark/assets/js/pace.min.js"></script>
  <!--favicon-->
  <link rel="icon" href="/Lostark/assets/images/favicon.ico" type="image/x-icon">
  <!-- simplebar CSS-->
  <link href="/Lostark/assets/plugins/simplebar/css/simplebar.css" rel="stylesheet"/>
  <!-- Bootstrap core CSS-->
  <link href="/Lostark/assets/css/bootstrap.min.css" rel="stylesheet"/>
  <!-- animate CSS-->
  <link href="/Lostark/assets/css/animate.css" rel="stylesheet" type="text/css"/>
  <!-- Icons CSS-->
  <link href="/Lostark/assets/css/icons.css" rel="stylesheet" type="text/css"/>
  <!-- Sidebar CSS-->
  <link href="/Lostark/assets/css/sidebar-menu.css" rel="stylesheet"/>
  <!-- Custom Style-->
  <link href="/Lostark/assets/css/app-style.css" rel="stylesheet"/>
  
  
</head>

<body class="bg-theme bg-theme2">

<!-- start loader -->
   <div id="pageloader-overlay" class="visible incoming"><div class="loader-wrapper-outer"><div class="loader-wrapper-inner" ><div class="loader"></div></div></div></div>
   <!-- end loader -->

<!-- Start wrapper-->
 <div id="wrapper">

 <!--Start sidebar-wrapper-->
   <div id="sidebar-wrapper" data-simplebar="" data-simplebar-auto-hide="true">
     <div class="brand-logo">
      <a href="Index.asp">
       <img src="/Lostark/assets/images/logo-icon.png" class="logo-icon" alt="logo icon">
       <h5 class="logo-text">LoInfo</h5>
     </a>
   </div>
   <ul class="sidebar-menu do-nicescrol">
      <%'// 관리자용 메뉴 %>
      <%If SessionClientLevel = "Developer" Then%>
      <li class="sidebar-header"><i class="zmdi zmdi-device-hub"></i>　관리자용</li>

      <li class="<%If SessionNavMenu = "CodeList" Then Response.Write "active" End If %>">
        <a href="CodeList.asp">
          <i class="zmdi zmdi-archive"></i> <span>코드 관리</span>
        </a>
      </li>

      <li class="<%If SessionNavMenu = "ActivityHistory" Then Response.Write "active" End If %>">
        <a href="ActivityHistory.asp">
          <i class="zmdi zmdi-menu"></i> <span>로그 내역</span>
        </a>
      </li>
      <%End If%>

      <%'// 공통 메뉴 %>
      <li class="sidebar-header"><i class="zmdi zmdi-accounts-list"></i>　정보 관리</li>

      <li class="<%If SessionNavMenu = "Index" Then Response.Write "active" End If %>">
        <a href="Index.asp">
          <i class="zmdi zmdi-account"></i> <span>내 정보</span>
        </a>
      </li>

      <!--
      <li>
        <a href="icons.html">
          <i class="zmdi zmdi-invert-colors"></i> <span>UI Icons</span>
        </a>
      </li>

      <li>
        <a href="forms.html">
          <i class="zmdi zmdi-format-list-bulleted"></i> <span>Forms</span>
        </a>
      </li>

      <li>
        <a href="tables.html">
          <i class="zmdi zmdi-grid"></i> <span>Tables</span>
        </a>
      </li>

      <li>
        <a href="calendar.html">
          <i class="zmdi zmdi-calendar-check"></i> <span>Calendar</span>
          <small class="badge float-right badge-light">New</small>
        </a>
      </li>

      <li>
        <a href="profile.html">
          <i class="zmdi zmdi-face"></i> <span>Profile</span>
        </a>
      </li>

      <li>
        <a href="login.html" target="_blank">
          <i class="zmdi zmdi-lock"></i> <span>Login</span>
        </a>
      </li>

       <li>
        <a href="register.html" target="_blank">
          <i class="zmdi zmdi-account-circle"></i> <span>Registration</span>
        </a>
      </li>
	  
      <li class="sidebar-header">LABELS</li>
      <li><a href="javaScript:void();"><i class="zmdi zmdi-coffee text-danger"></i> <span>Important</span></a></li>
      <li><a href="javaScript:void();"><i class="zmdi zmdi-chart-donut text-success"></i> <span>Warning</span></a></li>
      <li><a href="javaScript:void();"><i class="zmdi zmdi-share text-info"></i> <span>Information</span></a></li>
      -->

    </ul>
   
   </div>
   <!--End sidebar-wrapper-->
  

<!--Start topbar header-->
<header class="topbar-nav">
 <nav class="navbar navbar-expand fixed-top">
  <ul class="navbar-nav mr-auto align-items-center">
    <li class="nav-item">
      <a class="nav-link toggle-menu" href="javascript:void();">
       <i class="icon-menu menu-icon"></i>
     </a>
    </li>
    <li class="nav-item">
      <form id="NavSearchForm" name="NavSearchForm" method="Post" action="Index.asp" class="search-bar">
        <input type="text" name="SearchNickName" class="form-control" placeholder="캐릭터 검색">
         <a href="javascript:void();"><i class="icon-magnifier"></i></a>
      </form>
    </li>
  </ul>
    
  <ul class="navbar-nav align-items-center right-nav-link">
    <%'// 관리자용 메뉴 (로그)%>
    <%If SessionClientLevel = "Developer" Then%>
    <li class="nav-item dropdown-lg">
      <a class="nav-link dropdown-toggle dropdown-toggle-nocaret waves-effect" data-toggle="dropdown" href="javascript:void();">
      <i class="fa fa-envelope-open-o"></i></a>
    </li>
    <li class="nav-item dropdown-lg">
      <a class="nav-link dropdown-toggle dropdown-toggle-nocaret waves-effect" data-toggle="dropdown" href="javascript:void();">
      <i class="fa fa-bell-o"></i></a>
    </li>
    <li class="nav-item language">
      <a class="nav-link dropdown-toggle dropdown-toggle-nocaret waves-effect" data-toggle="dropdown" href="javascript:void();"><i class="fa fa-flag"></i></a>
      <ul class="dropdown-menu dropdown-menu-right">
          <li class="dropdown-item"> <i class="flag-icon flag-icon-gb mr-2"></i> English</li>
          <li class="dropdown-item"> <i class="flag-icon flag-icon-fr mr-2"></i> French</li>
          <li class="dropdown-item"> <i class="flag-icon flag-icon-cn mr-2"></i> Chinese</li>
          <li class="dropdown-item"> <i class="flag-icon flag-icon-de mr-2"></i> German</li>
        </ul>
    </li>
    <%End If%>
    <li class="nav-item">
      <a class="nav-link dropdown-toggle dropdown-toggle-nocaret" data-toggle="dropdown" href="#">
        <span class="user-profile"><img src="<%If HeaderCImgSrc = "" Then Response.Write LogoImgSrc Else Response.Write HeaderCImgSrc End If %>" class="img-circle" alt="user avatar"></span>
      </a>
      <ul class="dropdown-menu dropdown-menu-right">
       <li class="dropdown-item user-details">
        <a href="javaScript:void();">
           <div class="media">
             <div class="avatar"><img class="align-self-start mr-3" src="<%If HeaderCImgSrc = "" Then Response.Write LogoImgSrc Else Response.Write HeaderCImgSrc End If %>" alt="user avatar"></div>
            <div class="media-body">
            <h6 class="mt-2 user-title"><%= HeaderNickName %></h6>
            <p class="user-subtitle"><%= HeaderCNickName %> [<%= HeaderCLevel %>&nbsp;<%= HeaderCJob %>]</p>
            </div>
           </div>
          </a>
        </li>

        <li class="dropdown-divider"></li>
        <a href="Index.asp">
            <li class="dropdown-item"><i class="zmdi zmdi-account mr-2"></i> 내 정보</li>
        </a>

        <li class="dropdown-divider"></li>
        <a href="Config.asp">
            <li class="dropdown-item"><i class="icon-settings mr-2"></i> 환경설정</li>
        </a>

        <li class="dropdown-divider"></li>
        <%' 테스트 기간용%>
        <a href="Login.asp">
        <!-- <a href="/Lostark/Process/Logout.asp"> -->
            <li class="dropdown-item"><i class="icon-power mr-2"></i> 로그아웃</li>
        </a>
      </ul>
    </li>
  </ul>
</nav>
</header>
<!--End topbar header-->
<div class="clearfix"></div>