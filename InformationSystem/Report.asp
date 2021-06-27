<%@  codepage="65001" language="VBScript" %>

<%
Dim TopMenuSeq : TopMenuSeq = 7
Dim LeftMenuCode : LeftMenuCode = "Report"
Dim LeftMenuName : LeftMenuName = "Home / 합격자발표관리 / 레포트출력"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "레포트출력"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, strWhere2
Dim i, strMSG, intNUM, strTEMP, strRESULT

Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageSize		: PageSize	= 20
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'쿼리
SQL = ""
SQL = SQL & vbCrLf & " select IDX, SubCode, SubCodeName, Step, Temp1, Temp2 "
SQL = SQL & vbCrLf & " from CodeSub "
SQL = SQL & vbCrLf & " where MasterCode = 'Report' "
SQL = SQL & vbCrLf & " and State = 'Y' "
SQL = SQL & vbCrLf & " order by Step  "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
'AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB = Nothing

if IsArray(aryList) Then
	'// 페이지 계산
	TotalCount = ubound(aryList,2) + 1
	PageCount = int((TotalCount - 1) / PageSize) + 1
	StartNum = (PageNum * PageSize) - PageSize
	EndNum = StartNum + PageSize - 1
	intNUM = TotalCount - (PageNum * PageSize) + PageSize

	If EndNum > TotalCount - 1 Then
		EndNum = TotalCount - 1
	End If
End If

'개인정보가 있는 합격자리스트는 조회도 기록함
'strLogMSG = "레포트 출력  > " & SessionUserID  &"가/이 레포트 리스트를 조회 했습니다."
'Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)
%>

<div class="row">
	<div class="col-lg-6">
		<div class="ibox float-e-margins">
			<form id="ExcelForm">
				<input type="hidden" id="Division0" name="Division0" value="">
				<input type="hidden" id="Division1" name="Division1" value=""> 
				<input type="hidden" id="Subject" name="Subject" value=""> 
			</form>
			<form name="InputForm" id="InputForm" method="post" >

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div style="float:right;">					
					<!--<span class="btnBasic btnExcel btnTypeComplete" file="ReportExcel_R2.asp">test</span>
					<span class="btnBasic btnTypePrint" id="ReportPut" >레포트 출력</span>-->
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					
					<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
					<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
						<colgroup>
							<col width="10%"></col>
							<col width="80%"></col>
							<col width="10%"></col>
							<!--<col width="5%"></col>
							<col width="5%"></col>
							<col width="10%"></col>-->
						</colgroup>
						<thead>			                
							<tr>
								<th data-hide="phone">No.</th>    
								<!--<th data-hide="phone">년도</th> -->
								<th data-hide="phone">레포트명</th>  
								<!--<th data-hide="phone">저장여부</th> 
								<th data-hide="phone">출력여부</th> -->
								<th data-hide="phone">정렬순서</th>
							</tr>
						</thead>
						<tbody>
						<%
							' 0 IDX, 1 SubCode, 2 SubCodeName, 3 Step, 4.Temp1, 5.Temp2

							'If Not IsNull(AryHash) Then
							If isArray(aryList) Then
								'For i = 0 to ubound(AryHash,1)
								For i = StartNum to EndNum
						%>
							<tr class="viewDetail_SetDate_2" style="background-color: <%=BGColor%>;">
								<td><%= intNUM %></td>
								<!--<td><%= aryList(4, i) %></td>-->
								<td><%= aryList(2, i) %></td>
								<!--<td>O</td>
								<td><%= aryList(5, i) %></td>-->
								<td><%= aryList(3, i) %>
									<div class="DataField" style="display:none;">
										<li Columnvalue="<%= Trim(aryList(1, i)) %>"					ColumnName="IDX"></li>
									</div>
								</td>
							</tr>
						<%
									intNUM = intNUM - 1
								Next
							Else
						%>
							<tr>
								<td colspan="12" style="height:50px; vertical-align: middle;">검색된 자료가 없습니다.</td>
							</tr>
						<%
							end if
						%>
						</tbody>
					</table>

					<div class="paging pad_r10">&nbsp;</div>

				</div>
			</div>
			<!-- 테이블 -->

			<div class="pad_t10"></div>

			<!-- 히든 값 -->
			<input type="hidden" name="IDX" id="IDX" value="">

			</form>

		</div>		
	</div>

	<div class="col-lg-6">
		<div class="ibox float-e-margins">
			<!-- 검색조건 -->
			<div class="ibox-title">
				<h5>검색조건</h5>
				<div style="float:right;">
					<span class="btnBasic btnTypeComplete" id="ReportExcel" >파일저장</span>
				</div>
			</div>
			<!-- 기본 -->
			<div class="ibox-content" id="TypeBasic" style="padding : 250px 0px 250px 0px; text-align:center;">
				<h3>레포트를 선택하세요.</h3>
			</div>
			<!-- 일자별 원서접수 경쟁률의 검색조건 -->
			<div class="ibox-content" id="TypeR10" style="padding : 0px 0px 30px 0px; display:none;">
				<div class="row show-grid">
					<div class="col-md-3 col-xs-1 grid_sub_title" style="padding-left:20px;">
						모집시기
					</div>
					<div class="col-md-9 col-xs-2">
						<% Call SubCodeSelectBox("Division0_R10", "모집시기선택", SearchDivision, "", "All", "Division0") %>
					</div>
				</div>
			</div>

		</div>
	</div>
</div>

	



<!-- #InClude Virtual = "/Common/Bottom.asp" -->
<script>
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// tr선택 시 
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
		var idx = $("#IDX").val();
	
		$("#TypeBasic").css("display", "none");
		$("#Type"+idx).css("display", "block");
	});

	// 레포트출력 버튼
	$("#ReportPut").click(function() {
		if (!$.chkInputValue($("#IDX"),		"출력할 레포트를 선택해주세요.")) { return; }
		
		//if (confirm("레포트를 출력 하시겠습니까?")) {
			var idx = $("#IDX").val();

			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setReportPut(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/ReportPut_"+ idx +".asp";
			$.Ajax4Form("#InputForm", objOpt);
			$("#InputForm").submit();
		//}
	});

	// 레포트출력 결과
	$.setReportPut = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {
			//alert("레포트가 출력 되었습니다.");
		
		} else if ($objList.find("Result").text() == "Excel")	{
			alert("해당 레포트는 파일저장만 가능합니다.");	
		} else {
			alert("레포트 출력 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 레포트저장 버튼
	$("#ReportExcel").click(function() {
		if (!$.chkInputValue($("#IDX"),		"저장할 레포트를 선택해주세요.")) { return; }
			var idx = $("#IDX").val();

			$("#Division0").val($("select[name=Division0_"+idx+"]").val());
			$("#Division1").val($("select[name=Division1_"+idx+"]").val());
			$("#Subject").val($("select[name=Subject_"+idx+"]").val());			
			
			var blnChkResult = true;
			var $ParentForm = $("#ExcelForm");		
			var $objInput = $ParentForm.find(".form-control");

			$objInput.each(function()  {
				var alertMSG = $(this).attr("alert");
				
				if (alertMSG != undefined && alertMSG)  {
					//alert(alertMSG);
					blnChkResult = blnChkResult && $.chkInputValue($(this), alertMSG)
					return blnChkResult;
				}
			});

			if (blnChkResult) {
				$ParentForm.attr("method", "post");
				$ParentForm.attr("target", "ExcelFrame");
				$ParentForm.attr("action", "ReportExcels_"+idx+".asp");
				$ParentForm.submit();
			}
	});
});

</script>