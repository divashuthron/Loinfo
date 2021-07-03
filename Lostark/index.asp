<!--#include Virtual="/Lostark/include/header.asp"-->

<%
    '// 메뉴 선택 정보
    Dim SessionNavMenu : SessionNavMenu = "Index"

    Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strLogMSG
    Dim CNickName, CDivision, CJob, CLevel, CIsMain, CImgSrc

    '// 검색정보
    Dim NickName : NickName = fnR("SearchNickName", SessionUserName)

    '// DB 오픈
    Set objDB = New clsDBHelper
    objDB.strConnectionString = strDBConnString
    objDB.sbConnectDB

    '// 서브쿼리
    strWhere = strWhere & VbCrLf & "    And M.NickName = '" & NickName & "'"

    '// SQL
    SQL = ""
    SQL = SQL & VbCrLf & "Select "
    SQL = SQL & VbCrLf & "  M.NickName "
    SQL = SQL & VbCrLf & "	, CL.CharacterNickName "
    SQL = SQL & VbCrLf & "	, CL.CharacterDivision "
    SQL = SQL & VbCrLf & "	, CL.CharacterJob "
    SQL = SQL & VbCrLf & "	, CL.CharacterLevel "
    SQL = SQL & VbCrLf & "	, CL.CharacterIsMain "
    SQL = SQL & VbCrLf & "	, CL.CharacterImgSrc "
    SQL = SQL & VbCrLf & "From Member M "
    SQL = SQL & VbCrLf & "	Left Outer Join CharacterList CL "
    SQL = SQL & VbCrLf & "	On M.ID = CL.ID  "
    SQL = SQL & VbCrLf & "Where 1=1 "
    SQL = SQL & VbCrLf & strWhere

    'objDB.blnDebug = true
    'arrParams = objDB.fnGetArray
    'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
    AryHash = objDB.fnExecSQLGetHashMap(SQL, Null)

    If IsArray(AryHash) Then
      '// 사용자 정보
      NickName = AryHash(0).Item("NickName")
      
      '// 캐릭터 정보
      CNickName = AryHash(0).Item("CharacterNickName")
      CDivision = AryHash(0).Item("CharacterDivision")
      CJob = AryHash(0).Item("CharacterJob")
      CLevel = AryHash(0).Item("CharacterLevel")
      CIsMain = AryHash(0).Item("CharacterIsMain")
      CImgSrc = AryHash(0).Item("CharacterImgSrc")
    Else
      NickName = SessionUserName

      Response.Write "<script>alert('일치하는 정보가 없습니다.'); history.back();</script>"
    End If

    '// DB 클로즈
    Set objDB = Nothing
%>

<div class="clearfix"></div>
	
<div class="content-wrapper">
  <div class="container-fluid">
  <!-- Container Start -->

    <div class="row"> <!-- first row -->

      <div class="col-12 col-lg-12"> <!-- first row 12 -->
        <div class="card mt-3">
          <div class="card-header color-orange">
            <kbd><%= NickName %></kbd>
          </div>
        </div>
      </div> <!-- first row 12 end -->

    </div> <!-- first row end -->

    <div class="row"> <!-- second row -->
      <div class="col-lg-4"> <!-- second row 4 -->
        <div class="card profile-card-2">
          <div class="card-header">
            <i class="zmdi zmdi-account-circle"></i>대표 캐릭터
            <%
              'HeaderCNickName  대표캐릭터명
              'HeaderCDivision  대표캐릭터 분류
              'HeaderCJob       대표캐릭터 직업
              'HeaderCLevel     대표캐릭터 레벨
              'HeaderCImgSrc    대표캐릭터 사진
            %>
          </div>
          <div class="card-img-block">
            <%'대표 배경화면%>
            <img class="img-fluid" src="<%= BackgroundImgSrc %>" alt="Card image cap">
          </div>
          <div class="card-body pt-5">
            <img src="<%If HeaderCImgSrc = "" Then Response.Write LogoImgSrc Else Response.Write HeaderCImgSrc End If %>" alt="profile-image" class="profile">
            <h5 class="card-title"><%= HeaderCNickName %></h5>
            <p class="card-text">[<%= HeaderCLevel %>&nbsp;<%= HeaderCJob %>]</p>
            <div class="icon-block">
              <a href="javascript:void();"><i class="fa fa-facebook bg-facebook text-white"></i></a>
              <a href="javascript:void();"> <i class="fa fa-twitter bg-twitter text-white"></i></a>
              <a href="javascript:void();"> <i class="fa fa-google-plus bg-google-plus text-white"></i></a>
          </div>
        </div>

        <%'일정 정보 %>
        <!--
        <div class="card-body border-top border-light">

          <div class="media align-items-center">
            <div><img src="assets/images/timeline/html5.svg" class="skill-img" alt="skill img"></div>
            <div class="media-body text-left ml-3">
              <div class="progress-wrapper">
                <p>HTML5 <span class="float-right">65%</span></p>
                <div class="progress" style="height: 5px;">
                  <div class="progress-bar" style="width:65%"></div>
                </div>
              </div>                   
            </div>
          </div>

          <hr>
          <div class="media align-items-center">
            <div><img src="assets/images/timeline/bootstrap-4.svg" class="skill-img" alt="skill img"></div>
            <div class="media-body text-left ml-3">
              <div class="progress-wrapper">
                <p>Bootstrap 4 <span class="float-right">50%</span></p>
                <div class="progress" style="height: 5px;">
                  <div class="progress-bar" style="width:50%"></div>
                </div>
              </div>                   
            </div>
          </div>

          <hr>
          <div class="media align-items-center">
            <div><img src="assets/images/timeline/angular-icon.svg" class="skill-img" alt="skill img"></div>
            <div class="media-body text-left ml-3">
              <div class="progress-wrapper">
                <p>AngularJS <span class="float-right">70%</span></p>
                <div class="progress" style="height: 5px;">
                  <div class="progress-bar" style="width:70%"></div>
                </div>
              </div>                   
            </div>
          </div>

          <hr>
          <div class="media align-items-center">
            <div><img src="assets/images/timeline/react.svg" class="skill-img" alt="skill img"></div>
            <div class="media-body text-left ml-3">
              <div class="progress-wrapper">
                <p>React JS <span class="float-right">35%</span></p>
                <div class="progress" style="height: 5px;">
                  <div class="progress-bar" style="width:35%"></div>
                </div>
              </div>  
            </div>                 
          </div>
          -->

      </div> <!-- second row 4 end -->
    </div> <!-- second row end -->

  <!-- Container End -->
  </div> 
</div>

<!--#include Virtual="/Lostark/include/Bottom.asp"-->