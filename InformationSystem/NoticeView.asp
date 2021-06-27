<%@  codepage="65001" language="VBScript" %>
<%
'Dim TopMenuSeq : TopMenuSeq = 7
'Dim LeftMenuCode : LeftMenuCode = "Bill"
Dim LeftMenuName : LeftMenuName = "Home / 공지사항"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "공지사항"
Dim LogDivision				: LogDivision = "NoticeView"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere
Dim i, strMSG, intNUM, intNUM2, strTEMP, strRESULT

Dim StrURL			: StrURL = "/index.asp"
Dim StrViewURL		: StrViewURL = "/NoticeView.asp"

Dim IDX				: IDX	= fnR("IDX", 0)
Dim ContentType		: ContentType = FnR("ContentType", "View")

Dim ProcessType
Dim MYear, Department, DepartmentName, Title, content1, file1, file2, file3, file4, file5, INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "select "
SQL = SQL & vbCrLf & "		MYear, Division, Title, content1, file1, file2, file3, file4, file5, INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR, InsertTime "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Department', Division) AS DivisionName "
SQL = SQL & vbCrLf & "from NoticeTable " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 " 
SQL = SQL & vbCrLf & "	AND IDX = ?; "

Call objDB.sbSetArray("@IDX", adInteger, adParamInput, 0, IDX)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

if IsArray(AryHash) Then
	ProcessType				= "Update"
	MYear					= AryHash(0).Item("MYear")
	Department				= AryHash(0).Item("Division")
	DepartmentName			= AryHash(0).Item("DivisionName")
	Title					= AryHash(0).Item("Title")
	content1				= AryHash(0).Item("content1")
	file1					= AryHash(0).Item("file1")	
	file2					= AryHash(0).Item("file2")	
	file3					= AryHash(0).Item("file3")	
	file4					= AryHash(0).Item("file4")	
	file5					= AryHash(0).Item("file5")	
	InsertTime				= AryHash(0).Item("InsertTime")
Else
	ProcessType = "Insert"
	ContentType = "Add"
End if

Set objDB	= Nothing
%>
<link rel="stylesheet" href="/Js/plugins/summernote/summernote.css">
<script type="text/JavaScript" src="/Js/plugins/summernote/summernote.js"></script>
<script type="text/JavaScript" src="/Js/plugins/summernote/lang/summernote-ko-KR.js"></script>
<script type="text/JavaScript" src="/js/plugins/summernote/summernote-image-attributes.js"></script>
<script type="text/JavaScript" src="/js/jquery.print.js"></script>
<script type="text/javascript">
$(function() {
	// 저장
	$("#btnSave").click(function() {
		$('#InputForm').attr("target","");
		$('#InputForm').attr("action","/Process/NoticeProc.asp");
		$('#InputForm').attr("enctype","application/x-www-form-urlencoded");

		// 폼검사
		if ($.setValidation($("#InputForm"))) {	
			if (!$("#title").val()) {
				alert("공지사항명을 입력해주세요.");
				return;
			}else if(!$("#summernote").summernote('code')){
				alert("내용을 입력해주세요.");
				return;
			}			

			if ($(".filenameClass").length > 5 ) {alert("파일은 최대 5개까지 등록할 수 있습니다.");$("#InputForm").focus();return false;}

			var markupStr = $('#summernote').summernote('code');
			$("#content1").val(markupStr);

			// 저장
			if (confirm("공지사항을 등록 하시겠습니까?")) {
				var objOpt = {"url":"","param":"","dataType":"xml","before":"","success":"$.setNotice(datas)","complete":"","clear":"","reset":""};
				objOpt["url"] = $("#InputForm").attr("action");
				$.Ajax4Form($("#InputForm"), objOpt);
				$("#InputForm").submit();
			}
		}
	});

	// 저장처리결과
	$.setNotice = function(datas) {
		var $objList = $(datas).find("List");	
		var strMSG;
		
		if ($objList.find("Result").text() == "Complete") {
			alert("공지사항 등록이 완료 되었습니다.");

			if ($("#InputForm input[name='IDX']").val() == "0") {
				document.location.href = "/index.asp";
			} else {
				document.location.href = "/index.asp";
				//document.location.reload();
			}
		} else {
			alert("공지사항 등록 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
	
	// 취소
	$("#btnCancel").click(function() {
		$.goURL("<%= StrURL %>");
	});

	// 목록
	$("#btnClose").click(function() {
		$.goURL("<%= StrURL %>");
	});

	// 수정
	$("#btnEdit").click(function() {
		EditForm.submit();
	});

	// 첨부파일 삭제
	$(document).on("click", ".deleteBtn", function() {
		$(this).closest("div").remove();
	});
/*
	// 삭제
	$("#btnDelete").click(function() {
	});

	// 인쇄
	$("#btnPrint").click(function() {
		$("#printarea").print();
	});
*/
});

</script>


<style type="text/css">
	body {margin: 0;padding: 0;}
	* {box-sizing: border-box;-moz-box-sizing: border-box;}
	.all_wrap {overflow:hidden;}
	.page {width: 21cm;min-height: 29.7cm;padding: 1.35555cm;background:#eee;float:left;}
	.subpage p {line-height:10px;}
	.subpage,.subpage .subpage_left,.subpage .subpage_left .left_btit,.subpage .subpage_left .left_member,.subpage .subpage_left .mborder,.subpage .subpage_right,.subpage .subpage_right .cont_warp .tg,.subpage .subpage_right .cont_warp .tg td {outline:#000 solid thin\9;}
	.subpage,.subpage .subpage_left,.subpage .subpage_left .left_btit,.subpage .subpage_left .left_member,.subpage .subpage_left .mborder,.subpage .subpage_right,.subpage .subpage_right .cont_warp .tg,.subpage .subpage_right .cont_warp .tg td {outline-width:0px\9;}
	.subpage {border: 1px solid #000;background:#fff;height: 270mm;overflow:hidden;padding:20px;}
	.subpage .subpage_left {width:48%;float:left;border:1px solid #000;padding:10px;height:985px;}
	.subpage .subpage_left .left_btit{border:1px solid #000;text-align:center;font-weight:bold;height:40px;background:#ececec;padding-top:7px;}
	.subpage .subpage_left .left_stit{font-size:11px;margin:10px 0;}
	.subpage .subpage_left .left_member{border:1px solid #000;padding:5px;}
	.subpage .subpage_left .left_member table tr td{font-size:11px;}
	.subpage .subpage_left .left_member table tr td.table_left{font-weight:bold;}
	.subpage .subpage_left .left_member table tr td.table_left.title1{letter-spacing:7px;}
	.subpage .subpage_left .left_member table tr td.table_left.title2{letter-spacing:19px;}
	.subpage .subpage_left .mborder{border:0.5px solid #000; margin:10px 0;}
	.subpage .subpage_left .mtitle{text-align:center;font-weight:bold;margin-bottom:10px;}
	.subpage .subpage_left .cont_warp{font-size:11px;}
	.subpage .subpage_left .cont_warp ol{padding-left:15px;line-height: 18px;}
	.subpage .subpage_left .cont_warp ol li{font-weight:bold;}
	.subpage .subpage_left .cont_warp ol li .li_subcont{font-weight:normal;}
	.subpage .subpage_left .cont_warp ol li .li_subcont ol {padding-left:17px;}
	.subpage .subpage_left .cont_warp ol li .li_subcont ol li{font-weight:normal;}
	.subpage .subpage_right {width:48%;float:right;border:1px solid #000;padding:10px;height:985px;}
	.subpage .subpage_right .cont_warp{font-size:11px;}
	.subpage .subpage_right .cont_warp ol{padding-left:15px;line-height: 18px;}
	.subpage .subpage_right .cont_warp ol li{font-weight:bold;}
	.subpage .subpage_right .cont_warp ol li .li_subcont{font-weight:normal;}
	.subpage .subpage_right .cont_warp ol li .li_subcont ol {padding-left:17px;}
	.subpage .subpage_right .cont_warp ol li .li_subcont ol li{font-weight:normal;}
	.subpage .subpage_right .cont_warp .tg {margin-top:10px;font-size:11px;border:1px solid #000;}
	.subpage .subpage_right .cont_warp .tg td{background:#fff;border:1px solid #000;}
	.subpage .subpage_right .cont_warp .tg .tg_tit{text-align:center;background:#f7f7f7;}
	.subpage .subpage_right .cont_warp .tg .tg_tit.table_top{padding:7px 0;}
	.subpage .subpage_right .cont_warp .tg .tg_cont{text-align:center;}
	.subpage .subpage_right .cont_warp .bcont{font-weight:normal;margin-top:15px;}
	.subpage .subpage_right .bday{font-weight:bold;text-align:center;font-size:20px;margin-top:60px;}
	.subpage .subpage_right .bschool{font-weight:bold;text-align:center;font-size:30px;margin-top:30px;}

	.editor {float:left;overflow:hidden;}
	.editor .left_wrap {width:300px;float:left;}
	.editor .right_wrap {width:300px;float:right;}

	@page {size: A4;margin: 0;}
	@media print {html, body {width: 210mm;height: 297mm;-webkit-print-color-adjust: exact;}.page {margin: 0;border: initial;width: initial;min-height: initial;box-shadow: initial;background: initial;page-break-after: always;}}

	/* IE10과 IE11 : -ms-high-contrast */
	@media all and (-ms-high-contrast:none){
		.subpage .subpage_left .left_btit {padding-top:12px;}
	}	
</style>
<%
If ContentType = "View" Then
%>
<!-- 뷰어 -->
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			<div class="ibox-content">
				<form id = "EditForm" method="post">
					<div style="display:none;">
						<input name ="IDX" value="<%=IDX%>">
						<input id ="EditType" name="ContentType" type="hedden" value="Edit">
					</div>
				</form>
				<div>		
					<div class="row show-grid">
						<div class="col-md-1 col-xs-2">
							<h5>구분</h5>
						</div>
						<div class="col-md-1 col-xs-2">
							<span style="margin-left:11px;"><%= Myear %></span>
						</div>
						<div class="col-md-1 col-xs-2">
							<span><%= DepartmentName %></span>
						</div>
						<div class="col-md-6 col-xs-2">	</div>
						<div class="col-md-3 col-xs-2" style="margin-left:130px;">
							<span><%= InsertTime %></span>
							<%If SessionClientLevel = "Admin" Or SessionClientLevel = "SchoolAdmin" Then%>
							<span class="btnBasic btnTypeSave" id="btnEdit" style="margin-left:10px;">수정</span>	
							<%End If%>
							<span class="btnBasic btnTypeClose" id="btnClose" style="margin-left:10px;">목록</span>
						</div>
					</div>
					<div class="row show-grid">					
						<div class="col-md-1 col-xs-1">
							<h5>제목</h5>
						</div>
						<div class="col-md-11 col-xs-1">
							<span><%= title %></span>
						</div>						
					</div>	
					
					<div class="row show-grid">
						<div class="col-md-1 col-xs-1">
							<h5>파일첨부</h5>
						</div>
						<div class="col-md-6 col-xs-2">
							<div><a href="/upload/Files/<%=file1%>" target="_Blank"><%=file1%></a></div>
							<div><a href="/upload/Files/<%=file2%>" target="_Blank"><%=file2%></a></div>
							<div><a href="/upload/Files/<%=file3%>" target="_Blank"><%=file3%></a></div>
							<div><a href="/upload/Files/<%=file4%>" target="_Blank"><%=file4%></a></div>
							<div><a href="/upload/Files/<%=file5%>" target="_Blank"><%=file5%></a></div>							
						</div>
					</div>
				</div>
			</div>

			<div class="ibox-content"> 
				<div>
					<table style="margin:auto; width:100%; height:600px;">
						<tr>
							<td>
								<div class="subpage_left" id="subpage_left" style="width:100%;float:left;padding:10px;padding:10px;height:100%;">
									<%=content1%>
								</div>
							</td>
						</tr>
					</table>
				</div>
			</div>
			<!-- 테이블 -->
		</div>
	</div>
</div>
<%
Else
%>
<!-- 등록창 -->
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			<div class="ibox-content">
				<div>
					<form name="InputForm" id="InputForm" method="post" target="FileUploadFrame" ENCTYPE="multipart/form-data" action="fileupload2.asp">
					<div style="display:none;">
						<input type="hidden" name="ProcessType" id="ProcessType" value="<%= ProcessType %>">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="hidden" name="content1" id="content1">
					</div>					

					<div class="row show-grid">
						<div class="col-md-1 col-xs-1">
							<h5>구분</h5>
						</div>
						<div class="col-md-1 col-xs-2">
							<% Call SubCodeSelectBox("Myear", "사용년도", Myear, "사용년도를 입력해주세요.", "", "Myear") %>
						</div>
						<div class="col-md-1 col-xs-2">
							<% Call SubCodeSelectBox("Department", "부서선택", Department, "부서를 입력해주세요.", "", "Department") %>
						</div>
					</div>
					<div class="row show-grid">					
						<div class="col-md-1 col-xs-1">
							<h5>제목</h5>
						</div>
						<div class="col-md-10 col-xs-1">
							<input type="text" id="title" name="title" value="<%=title%>" style="width:800px; margin-top:2px;" class="form-control input-sm" maxlength="100" placeholder="공지사항명을 입력해주세요.">
						</div>						
					</div>	
					
					<div class="row show-grid">
						<div class="col-md-1 col-xs-1">
							<h5>파일첨부</h5>
						</div>
						<div class="col-md-6 col-xs-2">
							<!--엑셀 2010-->
							<!--<input type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" id="file" name="FileUpload" id="FileUpload">-->
							<!--엑셀 2003 ~ 2007 
							<input type="file" style="width:600px;" accept="application/vnd.ms-excel" id="file" name="FileUpload" id="FileUpload">-->

							<input type="file" style="float:left; margin-right:15px;" name="FileUpload" id="FileUpload" onChange="InputForm.submit();">
							<p style="font-size:5px; color:red;">#파일은 최대 5개까지 등록할 수 있습니다.</p>
							<div id="FilesName">
							<%If file1 <> "" Then%>
								<div class='filenameClass' ><span><a href="/upload/Files/<%=file1%>" target="_Blank"><%=file1%></a><input type='hidden' name='filename' value='<%=file1%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>
							<%End If%>
							<%If file2 <> "" Then%>
								<div class='filenameClass' ><span><a href="/upload/Files/<%=file2%>" target="_Blank"><%=file2%></a><input type='hidden' name='filename' value='<%=file2%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>
							<%End If%>
							<%If file3 <> "" Then%>
								<div class='filenameClass' ><span><a href="/upload/Files/<%=file3%>" target="_Blank"><%=file3%></a><input type='hidden' name='filename' value='<%=file3%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>
							<%End If%>
							<%If file4 <> "" Then%>
								<div class='filenameClass' ><span><a href="/upload/Files/<%=file4%>" target="_Blank"><%=file4%></a><input type='hidden' name='filename' value='<%=file4%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>
							<%End If%>
							<%If file5 <> "" Then%>
								<div class='filenameClass' ><span><a href="/upload/Files/<%=file5%>" target="_Blank"><%=file5%></a><input type='hidden' name='filename' value='<%=file5%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>
							<%End If%>
							</div>

							<!-- Contacts -->
							<div style="display:none;"><IFRAME name="FileUploadFrame" id="FileUploadFrame" src="" width="0" height="0" scrolling="no" frameborder="0" marginwidth="0" marginheight="0" style="frame-border:0;"></IFRAME></div>
						</div>
					</div>
				</div>
			</div>

			<div class="ibox-content"> 
				<div>
					<table style="margin:auto; width:100%;">
						<tr>
							<td>
								<div>
									<div class="editor" id="summernote">
										<%=content1%>
									</div>
								</div>											
							</td>
						</tr>
					</table>
				</div>
			</div>
			<div class="row show-grid grid_sub_button">
				<div class="col-md-12">
					<!--<span class="btnBasic btnTypePrint" id="btnPrint">인쇄</span>-->
					<span class="btnBasic btnTypeSave" id="btnSave"><% If ProcessType = "Update" Then %>수정<% Else %>등록<% End If %></span>
					<span class="btnBasic btnTypeCancel" id="btnCancel">취소</span>					
				</div>
			</div>

			</form>
		

			<!-- 테이블 -->
		</div>
	</div>
</div>
<%End IF%>
<script type="text/javascript">

//var option = [['style', ['bold', 'italic', 'underline', 'clear']],['font', ['strikethrough', 'superscript', 'subscript']],['fontsize', ['fontsize']],['color', ['color']],['para', ['ul', 'ol', 'paragraph']],['height', ['height']],['link', ['linkDialogShow', 'unlink']]];

$(document).ready(function() {
	
	$('#summernote').summernote({
		lang: 'ko-KR'
		, height:500
		, popover: {
			image: [
				['custom', ['imageAttributes']],
				['imagesize', ['imageSize100', 'imageSize50', 'imageSize33', 'imageSize25']],
				['float', ['floatLeft', 'floatRight', 'floatNone']],
				['remove', ['removeMedia']]
				]
		}
		, imageAttributes:{
			icon:'<i class="note-icon-pencil"/>',
			removeEmpty:false // true = remove attributes | false = leave empty if present
		}
		, callbacks: {
			onImageUpload : function(files, editor, welEditable) {
				for(var i = files.length - 1; i >= 0; i--) {
					$.SendFile(files[i], this, welEditable);
				}
			}
			/*
			, onChange: function() {
				//$("#subpage_left").html(($('#summernote').summernote('code')));
			}
			*/
		}
	});

	$.SendFile = function (file, editor, welEditable) {
		data = new FormData();
		data.append("callbackfile", file);

		//console.log('image upload:', file, editor, welEditable);
		//console.log(data);

		$.ajax({
			type: "POST",
			url: "/fileupload3.asp",
			enctype: "multipart/form-data",
			data: data,
			dataType: "text",
			cache: false,
			//contentType: 'multipart/form-data',
			contentType: false,
			processData: false,
			success: function(ImgURL) {
				$(editor).summernote('editor.insertImage', $.trim(ImgURL));
			},
			error: function (reason, e) {
				alert("파일 업로드에 실패했습니다. : " + e);
			}
		});
	}
});

/*
function removeTag( str ) {
	return str.replace(/(<([^>]+)>)/gi, "");
}
*/
</script>
<!-- 메인 컨텐츠 -->

<!-- #InClude Virtual = "/Common/Bottom.asp" -->