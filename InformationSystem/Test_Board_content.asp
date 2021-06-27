<%@  codepage="65001" language="VBScript" %>
<!DOCTYPE html>
<html lang="en-us">
<head>
<%

Dim sessionId		:	sessionId = Trim(Session("MemberID"))
Dim Dbcon , Rs
Dim idx				:	idx = Request("idx")
Set Dbcon = Server.CreateObject("ADODB.Connection") '디비 준비
Set Rs = Server.CreateObject("ADODB.RecordSet") '레코드셋 준비

Dim strConnect

StrSql=""
StrSql= StrSql & vbCrLf & "UPDATE "
StrSql= StrSql & vbCrLf & "	Test_Board "
StrSql= StrSql & vbCrLf & "SET "
StrSql= StrSql & vbCrLf & "	readcnt = readcnt +1 "
StrSql= StrSql & vbCrLf & "WHERE idx = " & idx

strConnect = "Provider=SQLOLEDB; Data Source=SQLMISS; Initial Catalog=InformationSystem; user ID=InterViewMng; password=east12!@;"
Dbcon.Open strConnect
Dbcon.Execute(StrSql)

StrSql=""
StrSql= StrSql & vbCrLf & "SELECT "
StrSql= StrSql & vbCrLf & "	idx, Title, content1, INPT_USID, INPT_DATE, readcnt"
StrSql= StrSql & vbCrLf & "FROM "
StrSql= StrSql & vbCrLf & "	Test_Board "
StrSql= StrSql & vbCrLf & "WHERE UseType = 'Y'"
StrSql= StrSql & vbCrLf & "AND idx = " & idx
Rs.Open StrSql, Dbcon

Dim ContentType		:	ContentType = Request("ContentType")
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
		}	
		td{
			padding-left: 15px!important;
			padding-right: 15px!important;
		}
		
		
	</style>

</head>

<body class="gray-bg">
	
	
	<div class="container">
	<h1>ㅇㅇㅁㄴㅇ ㅁㄴㅇㅁㄴ ㅇㅁㄴㅇㅁㄴㅇㅁ</h1>
		<%if ContentType = "" then%>
		<div class="row content-box">
			<form id = "EditForm" method="post">
				<div style="display:none;">
					<input name ="IDX" value="<%=idx%>">
					<input id ="EditType" name="ContentType" type="hedden" value="Edit">
				</div>
			</form>
			<div class="col-xs-12 text-center">
				<h2>글 상세보기</h2>
			</div>
			<table class="table table-bordered " style="background-color: white;">
				<tbody>
					<tr >
						<td colspan="2" style="background-color: gray; color: white;">
							<span style="float: right;">조회수 : <%=Rs("readcnt")%> </span>
							<span style="float: right;" class="m-r-lg"><%=left(Rs("INPT_DATE"),19)%> </span>
						</td>
					</tr>
					<tr>
						<td>
							<b>제목 : <%=Rs("Title")%></b> 
							<span style="float: right;"> 작성자 : <b> <%=Rs("INPT_USID")%></b> </span>
							
						</td>
					</tr>
					<tr>
						<td colspan="2">
							<div style="padding: 15px 0px 10px 0px; min-height:500px">
								<%=Rs("content1")%>
							</div>
						</td>
					</tr>
				</tbody>
			</table>
			<div style="float: right;">
				<%if StrComp(sessionId,Replace(Rs("INPT_USID"), Chr(13)&Chr(10),""))= 0 then%>
					<button type="button" class="btn btn-labeled bg-color-redLight txt-color-white" onclick="erase()">
						<i class="glyphicon glyphicon-random"></i> 삭제하기
					</button>
					<button type="button" class="btn btn-labeled bg-color-yellow txt-color-white" onclick="$('#EditForm').submit()" >
						<i class="glyphicon glyphicon-check"></i> 수정하기 
					</button>
				<%end if%>
				<button type="button" class="btn btn-labeled btn-primary" onclick= "location.href='Test_Board_list.asp'">
					<i class="glyphicon glyphicon-home"></i> 돌아가기 
				</button>
			</div>
		</div>
		<%else%>

		<div class="row content-box">
			<form method="POST" action="Process/Test_Board_writProc.asp" id="boardfrm">
			<input type="hidden" name="content1" id="content1" />
			<input type="hidden" name="idx" value="<%=idx%>" />
			<div class="col-xs-12 text-center">
				<h2>글 수정</h2>
			</div>
				<table class="table table-bordered " style="background-color: white;">
					<tbody>
						<tr>
							<td>
								<div class="col-xs-1">
									<h4 style="line-height: 1.53em; text-align: center;">제목</h4>
								</div>
								<div class="col-xs-11">
									<input type="text" id="title" name="title" class="form-control" placeholder="제목을 입력해 주세요" value="<%=Replace(Rs("Title"), Chr(13)&Chr(10),"")%>">
								</div>
							</td>
						</tr>
						<tr>
							<td>
								<div style="width: 960px; margin: 0 auto;" id="summernote" name="contents">
									<%=Replace(Rs("content1"), Chr(13)&Chr(10),"")%>
								</div>
							</td>
						</tr>
					</tbody>
				</table>
				<div style="float: right;">
					<button type="button" class="btn btn-labeled btn-primary"onclick="modifiy();"> <i class="glyphicon glyphicon-check"></i> 수정완료 </button>
					<button type="button" class="btn btn-labeled bg-color-redLight txt-color-white" onclick= "location.href='Test_Board_list.asp'" > <i class="glyphicon glyphicon-random"></i> 취소 </button>
				</div>
			</form>
		</div>
		<%end if%>
	</div>
		
	<script type="text/javascript">
		var SID;
		SID = "<%=sessionId%>";
		$(function() {
			if(!SID || SID.length==0 || SID==""){
				alert("로그인이 필요합니다.");
				location.href='Test_Board_login.asp';
			}
		});

		function erase(){
			if(confirm("정말 삭제하시겠습니까?")){
				location.href='/Process/Test_Board_del.asp?idx=<%=idx%>';
			}else{
				return;
			}
		}

		function modifiy(){
			$("#content1").val($('#summernote').summernote('code'));
			if(!SID || SID.length==0 || SID==""){
				alert("로그인이 필요합니다.");
				location.href='Test_Board_login.asp';
			}else{
				$('#boardfrm').submit();
			}
		}

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
		});
	</script>
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
<%
Rs.close
    set Rs = nothing
    set Dbcon = nothing
%>