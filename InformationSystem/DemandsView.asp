<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 7
Dim LeftMenuCode : LeftMenuCode = "Demands"
Dim LeftMenuName : LeftMenuName = "Home / 합격자발표관리 / 유의사항 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "유의사항 설정"
Dim LogDivision				: LogDivision = "DemandsView"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere
Dim i, strMSG, intNUM, intNUM2, strTEMP, strRESULT

Dim StrURL			: StrURL = "/DemandsList.asp"
Dim StrViewURL		: StrViewURL = "/DemandsView.asp"

Dim IDX				: IDX	= fnR("IDX", 0)
Dim conut, conutName

Dim ProcessType
Dim MYear,Division0,Title,State,StateName,option1,option2,option3,option4,option5,option6,option7,option8,option9,option10 
Dim content1,content2,INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	IDX,MYear,Division0,Title,State, (CASE  State "
SQL = SQL & vbCrLf & "											WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "											WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName ,option1,option2,option3,option4,option5,option6,option7,option8,option9,option10 "
SQL = SQL & vbCrLf & "	,content1,content2,INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime "
SQL = SQL & vbCrLf & "FROM DemandsTable "
SQL = SQL & vbCrLf & "WHERE 1 = 1 " 
SQL = SQL & vbCrLf & "	AND IDX = ?; "

Call objDB.sbSetArray("@IDX", adInteger, adParamInput, 0, IDX)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

if IsArray(AryHash) Then
	ProcessType				= "DemandsUpdate"
	MYear					= AryHash(0).Item("MYear")
	Division0				= AryHash(0).Item("Division0")
	Title					= AryHash(0).Item("Title")
	State					= AryHash(0).Item("State")
	StateName				= AryHash(0).Item("StateName")
	option1					= AryHash(0).Item("option1")
	option2					= AryHash(0).Item("option2")
	option3					= AryHash(0).Item("option3")
	If Not(isnull(option3)) Then
		option3					= FormatDateTime(option3,4)
	End If
	option4					= AryHash(0).Item("option4")
	option5					= AryHash(0).Item("option5")
	option6					= AryHash(0).Item("option6")
	option7					= AryHash(0).Item("option7")
	option8					= AryHash(0).Item("option8")
	option9					= AryHash(0).Item("option9")
	option10				= AryHash(0).Item("option10")
	content1				= AryHash(0).Item("content1")
	content2				= AryHash(0).Item("content2")
	InsertTime				= AryHash(0).Item("InsertTime")
Else
	ProcessType = "DemandsInsert"
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
	//update나 add인 경우에는 차수계산을 위해 년도, 모집단위 변경불가
	var ProcessType = '<%=ProcessType%>'

	if (ProcessType != 'DemandsInsert')	{
		$("#InputForm [name='Myear']").attr("disabled", true).trigger("chosen:updated");
		$("#InputForm [name='Division0']").attr("disabled", true).trigger("chosen:updated");
	}

/*	//날짜,시간 받기
	var ReceiptDate = '<%=option2%>'
	var CheckTime = '<%=option3%>'

	$("#ReceiptDate").val(ReceiptDate);
	$("#CheckTime").val(CheckTime);
*/
	// 저장
	$("#btnSave").click(function() {
		// 폼검사
		if ($.setValidation($("#InputForm"))) {
			/*if (!$("#ReceiptDate").val()) {
				alert("일자를 입력해주세요.");
				return;
			}else if(!$("#CheckTime").val()){
				alert("시간을 입력해주세요.");
				return;
			}*/

			var markupStr = $('#summernote').summernote('code');
			var markupStr1 = $('#summernote1').summernote('code');

			$("#content1").val(markupStr);
			$("#content2").val(markupStr1);
			$("#InputForm [name='Myear']").attr("disabled", false).trigger("chosen:updated");
			$("#InputForm [name='Division0']").attr("disabled", false).trigger("chosen:updated");

			// 저장
			if (confirm("유의사항을 저장 하시겠습니까?")) {
				var objOpt = {"url":"","param":"","dataType":"xml","before":"","success":"$.setDemands(datas)","complete":"","clear":"","reset":""};
				objOpt["url"] = $("#InputForm").attr("action");
				$.Ajax4Form($("#InputForm"), objOpt);
				$("#InputForm").submit();
			}
		}
	});

	// 저장처리결과
	$.setDemands = function(datas) {
		var $objList = $(datas).find("List");	
		var strMSG;
		
		if ($objList.find("Result").text() == "Complete") {
			alert("유의사항 저장이 완료 되었습니다.");

			if ($("#InputForm input[name='IDX']").val() == "0") {
				document.location.href = "/DemandsList.asp";
			} else {
				document.location.href = "/DemandsList.asp";
				//document.location.reload();
			}
		} else {
			alert("유의사항 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
	
	// 취소
	$("#btnCancel").click(function() {
		$.goURL("<%= StrURL %>");
	});
/*
	// 삭제
	$("#btnDelete").click(function() {
	});
*/
	// 인쇄
	$("#btnPrint").click(function() {
		$("#printarea").print();
	});


	// 시간표시
	$("#CheckTime").clockpicker({
		placement: "bottom", donetext: "Done"
	});
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
<!-- 메인 컨텐츠 -->
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<div class="ibox-title">
				<h5>변수 설정</h5>
			</div>

			<div class="ibox-content">
				<div>
					<form name="InputForm" id="InputForm" method="post" action="/Process/DemandsProc.asp">
					<div style="display:none;">
						<input type="hidden" name="ProcessType" id="ProcessType" value="<%= ProcessType %>">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="hidden" name="content1" id="content1">
						<input type="hidden" name="content2" id="content2">
					</div>					

					<div class="row show-grid">
						<div class="col-md-1 col-xs-1 grid_sub_title">
							사용년도(*)
						</div>
						<div class="col-md-2 col-xs-2">
							<% Call SubCodeSelectBox("Myear", "사용년도", Myear, "고지서년도를 입력해주세요.", "", "Myear") %>
						</div>
						<div class="col-md-1 col-xs-1 grid_sub_title">
							모집시기(*)
						</div>
						<div class="col-md-2 col-xs-2">
							<% Call SubCodeSelectBox("Division0", "모집시기", Division0, "고지서년도를 입력해주세요.", "", "Division0") %>
						</div>
						<!--
						<div class="col-md-1 col-xs-2 grid_sub_title">
							일자(*)
						</div>
						<div class="col-md-2 col-xs-3 grid_sub_title">
							<div class="input-group viewCalendarBtn" Obj="ReceiptDate" >
								<input type="text" name="ReceiptDate" id="ReceiptDate" class="form-control input-sm" maxlength="10" style="background-color:white;" readonly>
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							시간(*)
						</div>
						<div class="col-md-2 col-xs-3 grid_sub_title">
							<div class="input-group">
								<input type="text" name="CheckTime" id="CheckTime" class="form-control input-sm" maxlength="5" data-autoclose="true" style="background-color:white;" readonly>
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>			
						-->
						<div class="col-md-1 col-xs-2 grid_sub_title">
							사용여부(*)
						</div>
						<div class="col-md-2 col-xs-3">
							<select name="State" style="width:100%;">
								<option value="" <% If State = "" then response.write "selected" end if%>>구분</option>
								<option value="1" <% If State = 1 then response.write "selected" end if%>>사용</option>
								<option value="0" <% If State = 0 then response.write "selected" end if%>>미사용</option>
							</select>
						</div>	
					</div>			
				</div>
			</div>
			<!-- 검색조건 끝-->

			<div class="ibox-content">
				<div class="pad_5">
					<table>
						<tr>
							<td>
								<div class="left_wrap">
									<div class="editor" id="summernote">
										<%=content1%>
									</div>
								</div>											
							</td>
							<td>
								<div class="right_wrap">
									<div class="editor" id="summernote1">
										<%=content2%>
									</div>
								</div>											
							</td>
							<td>
								<div style="width: 625px;min-height: 29.7cm;margin: 0 auto;float:right;" id="printarea" class="col-md-6">
									<div style="border: 1px solid #000;background:#fff;height: 290mm;overflow:hidden;padding:5px;">
										<div class="subpage_left" id="subpage_left" style="width:300px;float:left;border:1px solid #000;padding:10px;padding:10px;height:1084px;">
											<%=content1%>
										</div>
										<div class="subpage_right" id="subpage_right" style="width:300px;float:left;border:1px solid #000;border-left:0;padding:10px;height:1084px;">
											<%=content2%>
										</div>
									</div>
								</div>
							</td>
						</tr>
					</table>
				</div>
			</div>
			<div class="row show-grid grid_sub_button">
				<div class="col-md-12">
					<span class="btnBasic btnTypePrint" id="btnPrint">인쇄</span>
					<span class="btnBasic btnTypeSave" id="btnSave"><% If ProcessType = "DemandsUpdate" Then %>수정<% Else %>등록<% End If %></span>
					<span class="btnBasic btnTypeCancel" id="btnCancel">취소</span>					
				</div>
			</div>

			</form>
			<!-- 테이블 -->
		</div>
	</div>
</div>
<script type="text/javascript">

var $summernote = $('#summernote');
var $summernote1 = $('#summernote1');
var option = [['style', ['bold', 'italic', 'underline', 'clear']],['font', ['strikethrough', 'superscript', 'subscript']],['fontsize', ['fontsize']],['color', ['color']],['para', ['ul', 'ol', 'paragraph']],['height', ['height']],['link', ['linkDialogShow', 'unlink']]];
$(document).ready(function() {
	$summernote.summernote({
		lang: 'ko-KR',
		height:990,
		popover: {
			image: [
				['custom', ['imageAttributes']],
				['imagesize', ['imageSize100', 'imageSize50', 'imageSize33', 'imageSize25']],
				['float', ['floatLeft', 'floatRight', 'floatNone']],
				['remove', ['removeMedia']]
				]
		},
		imageAttributes:{
			icon:'<i class="note-icon-pencil"/>',
			removeEmpty:false // true = remove attributes | false = leave empty if present
		},
		callbacks: {
			onImageUpload : function(files, editor, welEditable) {
				for(var i = files.length - 1; i >= 0; i--) {
					sendFile(files[i], this);
				}
			},
			onChange: function() {
				$("#subpage_left").html(($('#summernote').summernote('code')));
			}
		}
	});

	$summernote1.summernote({
		lang: 'ko-KR',
		height:990,
		popover: {
			image: [
				['custom', ['imageAttributes']],
				['imagesize', ['imageSize100', 'imageSize50', 'imageSize33', 'imageSize25']],
				['float', ['floatLeft', 'floatRight', 'floatNone']],
				['remove', ['removeMedia']]
				]
		},
		imageAttributes:{
			icon:'<i class="note-icon-pencil"/>',
			removeEmpty:false // true = remove attributes | false = leave empty if present
		},
		callbacks: {
			onImageUpload : function(files, editor, welEditable) {
				for(var i = files.length - 1; i >= 0; i--) {
					sendFile(files[i], this);
				}
			},
			onChange: function() {
				$("#subpage_right").html(($('#summernote1').summernote('code')));
			}
		}
	});
});
function removeTag( str ) {
	return str.replace(/(<([^>]+)>)/gi, "");
}
</script>
<!-- 메인 컨텐츠 -->


<!-- #InClude Virtual = "/Common/Bottom.asp" -->