<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 6
Dim LeftMenuCode : LeftMenuCode = "Assessment"
Dim LeftMenuName : LeftMenuName = "Home / 사정관리 / 사정처리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "사정처리"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
'(필요 시) 사정처리에 대한 조건 추가 예정
'(필요 시) 데이터 이관에 대한 조건 추가 예정
'(필요 시) 사정처리할 대한 검색 후 출력된 리스트만 사정처리하는 기능 추가 예정(division 별 사정처리 기능)
%>
<script type="text/javascript">

</script>
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">

			<!-- 내용  -->
			<div class="ibox-content">
				<div class="pad_5">
					<form id="AssessmentForm" method="post">
						<div style="display:none;">
							<input type="text" name="AssessmentMyear" id="AssessmentMyear" value="">
							<input type="text" name="AssessmentDivision0" id="AssessmentDivision0" value="">
						</div>
						<div class="pad_t10"></div>
						<div style="text-align: center;">
							<h2>[사정 처리]버튼을 누르면 석차가 확정되며,</h2>
							<h2>합격자와 불합격자, 점수, 예비순위가 생성됩니다.</h2>
							<h5>*지원자 관리 최종 완료 여부가 완료인 지원자를 대상으로 실시됩니다.</h5>
							<h5>*미완료의 경우 자격 미달로 불합격 처리되며 충원 대상자가 아니므로 석차 및 예비순위에서 제외됩니다.</h5>
							<br>
							<h2>[데이터 이관]버튼을 누르면</h2>
							<h2>사정 처리된 데이터가 충원 서버와 납부 환불 서버에 적용됩니다.</h2>						
							<div class="pad_t10"></div>
							<div class="pad_t10"></div>
							<div class="pad_t10"></div>
							<div class="row show-grid" style="text-align: left;">
								<!--<span class="btnBasic btnTypeSearch">사정 처리</span>
								<span class="btnBasic btnTypePrint">사정 처리</span>
								<span class="btnBasic btnTypeAccept">사정 처리</span>
								<span class="btnBasic btnTypeNew">사정 처리</span>-->							
								<div class="col-md-2 col-xs-2 " style="text-align: center; margin-right:80px;"></div>
								<div class="col-md-1 col-xs-2 " style="text-align: center;">
									년도 *
								</div>
								<div class="col-md-1 col-xs-7">
									<% Call SubCodeSelectBox("MYear", "년도 선택", "", "년도를 선택하여주세요.", "", "MYear") %>
								</div>		
								<div class="col-md-1 col-xs-2 " style="text-align: center;">
									모집시기 *
								</div>
								<div class="col-md-1 col-xs-7" style="center; margin-right:50px;">
									<% Call SubCodeSelectBox("Division0", "모집시기 선택", "", "모집시기를 선택하여주세요.", "", "Division0") %>
								</div>	
								<div class="col-md-1 col-xs-2 ">
									<span id="CheckBtn" class="btnBasic btnTypeExcel">대상 확인</span>					
								</div>
								<div class="col-md-1 col-xs-2 ">
									<span id="AssessmentBtn" class="btnBasic btnTypeSave" style="display:none;">사정 처리</span>
									<!--<span class="btnBasic btnTypeDelete">사정 처리</span>
									<span class="btnBasic btnTypeAdd">사정 처리</span>
									<span class="btnBasic btnTypeEdit">사정 처리</span>
									<span class="btnBasic btnTypeComplete">사정 처리</span>
									<span class="btnBasic btnTypeClose">사정 처리</span>
									<span class="btnBasic btnTypeExcel">사정 처리</span>
									<span class="btnBasic btnTypeCancel">사정 처리</span>
									<span class="btnBasic btnTypePeport">사정 처리</span>-->
								</div>
								<div class="col-md-1 col-xs-2 ">
									<span id="DataTransferBtn" class="btnBasic btnTypeAccept" style="display:none;">데이터 이관</span>
								</div>
							</div>
						</div>
						<div class="pad_t10"></div>
					</form>
				</div>		
			</div>
			<!-- 내용 끝 -->

			<!-- 뷰어 -->
			<div class="ibox-title View" style="display:none;">
				<div class="CountTable"></div>
			</div>	
			<div class="ibox-content View" style="display:none;"> 
				<div class="pad_5" style="overflow:auto; height:387px;">
					<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
						<thead>			                
							<tr>
								<th data-hide="phone">No.</th> 
								<th data-hide="phone">년도</th>           
								<th>모집시기</th>                         
								<th>학과</th>                           
								<th>전형</th>                            
								<th>수험번호</th>                            
								<th>이름</th>                            
								<th>최종결과</th>
							</tr>
						</thead>
						<tbody class='viewTable'></tbody>
					</table>
				</div>				
			</div>
			<!-- 뷰어 끝 -->

			<div class="pad_t10"></div>

		</div>		
	</div>
</div>

<script type="text/javascript">
$(function() {
	// 확인 버튼
	$("#CheckBtn").click(function() {	
		if (!$.chkInputValue($("select[name=MYear]"),			"년도를 선택해 주시기 바랍니다.")) { return; }
		if (!$.chkInputValue($("select[name=Division0]"),		"모집시기를 선택해 주시기 바랍니다.")) { return; }		

		var objOpt = {'url':'','param':'','dataType':'text','before':'','success':'$.setCheck(datas)','complete':'','clear':'','reset':''};
		objOpt["url"] = "/Process/AssessmentCheckProc.asp";
		$.Ajax4Form("#AssessmentForm", objOpt);
		$("#AssessmentForm").submit();
	});	

	// 확인 결과
	$.setCheck = function(datas) {
		var datasStr	= "";
				
		//View 출력
		$(".View").css("display","block");

		//내용과 카운터 나누기
		datasStr = datas.split("count");

		//파일내용 출력
		$(".viewTable").empty().html(datasStr[0]);

		//입학원서 건수 출력			
		$(".CountTable").html("<h5>목록 - 전체 " + datasStr[1] + "건</h5>");	

		//완료인 지원자가 있으면
		if (datasStr[1] != "0"){
			// 확인하면 사정처리, 데이터이관용 년도와 모집시기 넣어주기 
			$("#AssessmentMyear").val($("select[name=MYear]").val());	//년도
			$("#AssessmentDivision0").val($("select[name=Division0]").val());	//모집시기

			// 사정처리버튼 생성
			$("#AssessmentBtn").css("display","block");	
			// 데이터이관버튼 생성
			$("#DataTransferBtn").css("display","block");
		}else{
			// 사정처리버튼 숨기기
			$("#AssessmentBtn").css("display","none");	
			// 데이터이관버튼 숨기기
			$("#DataTransferBtn").css("display","none");
		}
	}

	// 사정 처리 버튼
	$("#AssessmentBtn").click(function() {	
		if (!$.chkInputValue($("#AssessmentMyear"),			"년도를 선택해 주시기 바랍니다.")) { return; }
		if (!$.chkInputValue($("#AssessmentDivision0"),		"모집시기를 선택해 주시기 바랍니다.")) { return; }

		if (confirm("사정처리를 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setAssessment(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/AssessmentProc.asp";
			$.Ajax4Form("#AssessmentForm", objOpt);
			$("#AssessmentForm").submit();
		}
	});

	// 사정 결과
	$.setAssessment = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {			
			alert("사정처리가 완료 되었습니다. 합격자조회 페이지에서 결과를 확인하세요.");
				
		} else {
			alert("사정처리 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 데이터 이관 버튼
	$("#DataTransferBtn").click(function() {
		alert("기능이 활성화되지 않았습니다."); return;

		if (!$.chkInputValue($("#AssessmentMyear"),			"년도를 선택해 주시기 바랍니다.")) { return; }
		if (!$.chkInputValue($("#AssessmentDivision0"),		"모집시기를 선택해 주시기 바랍니다.")) { return; }

		if (confirm("데이터를 충원 서버와 납부 환불 서버로 이관 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setDataTransfer(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/DataTransferProc.asp";
			$.Ajax4Form("#AssessmentForm", objOpt);
			$("#AssessmentForm").submit();
		}
	});

	// 데이터 이관 결과
	$.setDataTransfer = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {
			alert("데이터 이관이 완료 되었습니다. 충원 & 환불 서버에서 결과를 확인하세요.");
				
		} else {
			alert("데이터 이관 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
});
</script>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->