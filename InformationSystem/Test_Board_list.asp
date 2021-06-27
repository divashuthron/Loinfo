<%@  codepage="65001" language="VBScript" %>
<%OPTION EXPLICIT%>
<!DOCTYPE html>
<head>
<%


Dim sessionId		:	sessionId = Session("MemberID")
Dim Dbcon , Rs, strSql


Set Dbcon = Server.CreateObject("ADODB.Connection") '디비 준비
Set Rs = Server.CreateObject("ADODB.RecordSet") '레코드셋 준비

Dim strConnect

strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect
Dim pageRow, pageGroup, nowPage, startPage, endPage, startCnt, endCnt, totalPage, totalCount, searchType, searchText, strWhere

pageRow = 5
pageGroup = 5

nowPage = Request("nowPage")
searchType = Request("searchType")
searchText = Request("searchText")
Select case (searchType)
	case 1 '제목
	strWhere = "AND title LIKE '%"& searchText &"%' "
	case 2 '내용
	strWhere = "AND content1 LIKE '%"& searchText &"%' "
	case 3 '제목 + 내용
	strWhere = "AND (title LIKE '%"& searchText &"%' OR content1 LIKE '%"& searchText &"%')"
	case 4 '작성자'
	strWhere = "AND LIKE INPT_USID %"& searchText &"% "
	case else
	strWhere = ""
End Select

strSql="select * from Test_board where UseType = 'Y' " & strWhere
Rs.Open strSql, Dbcon, 1, 1
totalCount = Rs.RecordCount

Rs.close

if nowPage="" or isNull(nowPage) Then
 nowPage = 1
end if

'보여질 페이지 수 계산
if (totalCount Mod pageRow) = 0 then
 totalPage = FIX((totalCount / pageRow))
else
 totalPage = FIX((totalCount / pageRow)) +1
end if
startCnt = (nowPage-1) * pageRow +1
endCnt = nowPage * pageRow


strSql=""
strSql= strSql & vbCrLf & "SELECT *"
strSql= strSql & vbCrLf & "FROM(SELECT ROW_NUMBER() OVER(ORDER BY NUM DESC ) ROWNUM, *"
strSql= strSql & vbCrLf & "		FROM("
strSql= strSql & vbCrLf & "			SELECT ROW_NUMBER() OVER(ORDER BY idx) AS NUM, *"
strSql= strSql & vbCrLf & "			FROM Test_Board "
strSql= strSql & vbCrLf & "			WHERE UseType = 'Y' " & strWhere
strSql= strSql & vbCrLf & "			)A)B"
strSql= strSql & vbCrLf & "WHERE ROWNUM between "&startCnt&" and "&endCnt
strSql= strSql & vbCrLf & "ORDER BY idx DESC"
Rs.Open strSql, Dbcon, 1, 1

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

	<style>
		.container{
			min-width: 900px;
			max-width: none !important;
		}
		.table-active:hover{
			cursor:pointer;
		}
		tr td:nth-child(3){
			padding-left:15px !important;
			text-align:left;
		}
	</style>

</head>
	
<body class="gray-bg">
	<div class="container">
		<div style="float: right;">
			<button class="btn btn-danger" onclick="logout()"> 로그아웃 </button>
		</div>
		<div class="row content-box">
			<div class="col-xs-12"style="min-width: 560px;">
				<div>
					<div class="col-xs-6" style="padding-top: 30px;">
						<span>전체 <%=totalCount%>건</span>
					</div>
					<div class="col-xs-6">
						<form class="form-inline" style="float: right;margin: 0; padding-top: 15px;" method="POST" id="searchFrm">
							<select name="searchType" id="searchType" class="form-control input-sm">
								<option value="0">구분</option>
								<option value="1">제목</option>
								<option value="2">내용</option>
								<option value="3">제목+내용</option>
								<option value="4">작성자</option>
							</select>
							<input type="hidden" id="nowPage" name="nowPage" value="">
							<input class="form-control" type="text" placeholder="Search" name="searchText" style="display: inline-block; width: 180px; height: 28px;" value="<%=searchText%>">
							<button class="btn btn-info" type="submit">Search</button>
						</form>
					</div>
				</div>
				<table class="table table-striped table-bordered table-hover" style="margin: 0 auto;">
					<thead>
						<tr>
							<th scope="col" style="width:35px;">No.</th>
							<th scope="col">글제목</th>
							<th scope="col" style="width:100px;">작성자</th>
							<th scope="col" style="width:100px;">작성일</th>
							<th scope="col" style="width:50px;">조회수</th>
						</tr>
					</thead>
					<tbody>
					<% if Rs.BOF or rs.EOF then %>
						<tr class="table-active" style="pointer-events: none">
							<td colspan="5">데이터가 없습니다.</td>
						</tr>
					<%else%>
					<%Do until Rs.EOF%>
						<tr class="table-active">
							<td class="idx" name="123" style="display:none;"><%=Rs("idx")%></td>
							<td><%=Rs("NUM")%></td>
							<td><%=Rs("Title")%></td>
							<td><%=Rs("INPT_USID")%></td>
							<td><%=left(Rs("INPT_DATE"),11)%></td>
							<td><%=Rs("readcnt")%></td>
						</tr>
					<%
					Rs.MoveNext
					loop
					%>
					<%end if%>
					</tbody>
				</table>
				<script type="text/javascript">
					$(function() {
						$(document).ready(function() {
							// 페이징 영역 생성
							$.makePage(<%=nowPage%>,<%=pageGroup%>,<%=totalPage%>,".paging");

							$('.pageNUM').click(function(){
								var pagenum = $(this).text();
								if(pagenum =="다음"){
									pagenum = parseInt($(this).prev().text())+1;
								}else if(pagenum == "이전"){
									pagenum = parseInt($(this).next().text())-1;
								}else{
									$("#nowPage").val(pagenum);
									$("#searchFrm").submit();
								}
							});
						});
					});
				</script>
				<div style="margin:8px 5px; float:right;">
					<button class="btn btn-lg btn-info" onclick="location.href='Test_Board_write.asp'"> 글 쓰 기</button>
				</div>
				<div class="paging pad_r10" style="text-align : center;">
				</div>
			</div>
			
		</div>
	</div>
		
	<script type="text/javascript">
		var SID;
	$(function() {
		SID = "<%=sessionId%>";
		$("#searchType option:eq(<%=searchType%>)").attr("selected","selected");
		if(!SID || SID.length==0 || SID==""){
			alert("로그인이 필요합니다.");
			location.href='Test_Board_login.asp';
		}
	});

	function logout(){
			location.href='Test_Board_login.asp';
		alert('로그아웃!');
		location.href='Test_Board_logout.asp';
	}

	function Writego() {
		if(!SID || SID.length==0 || SID==""){
			alert("로그인이 필요합니다.");
		}else{
			location.href='Test_Board_write.asp';
		}
	}
	
	$('.table-active').click(function(){
		var idx = $(this).children('td.idx').text();
		location.href='Test_Board_content.asp?idx='+idx;
	});
	
	</script>
	
	<!--<script src="js/jquery-3.1.1.min.js"></script>-->
	<script src="/js/bootstrap.min.js"></script>

	<!-- JQUERY SELECT2 INPUT -->
	<script src="/js/plugins/SELECT2/SELECT2.min.js"></script>

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
<%
	Rs.close
    set Rs = nothing
	Dbcon.close
    set Dbcon = nothing
%>