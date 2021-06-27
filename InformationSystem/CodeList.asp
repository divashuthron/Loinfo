<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Code"
Dim LeftMenuName : LeftMenuName = "Home / 환경설정 / 코드관리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "코드등록"
%>
<!--#InClude Virtual = "/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 4
Dim PageNum			: PageNum	= fnR("page", 1)
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/CodeList.asp?type=Code"

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

if (not(IsE(SearchText))) then
	strWhere = strWhere & "	And MasterCodeName like '%' + ? + '%' "
	Call objDB.sbSetArray("@MasterCodeName", adVarchar, adParamInput, 255, SearchText)
end If

' MasterCode, MasterCodeName, State, StateName, 		0~3
' RegDate, RegID, EditDate, EditID						4~7

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	MasterCode, MasterCodeName, State "
SQL = SQL & vbCrLf & "	, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID"
SQL = SQL & vbCrLf & "FROM CodeMaster AS A " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & strWhere
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB	= Nothing
%>

<div class="row">
	<!-- 마스터 코드 -->
	<div class="col-lg-5">
		<div class="ibox float-e-margins">
			<!-- 마스터 코드 리스트 -->
			<div class="ibox-title">
				<h5>마스터 코드</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="table-responsive">				
					<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
					<table id="dt_MasterCode" class="table table-striped table-bordered table-hover" width="100%">
						<thead>			                
							<tr>
								<th data-hide="phone">코드</th>
								<th>코드명</th>
								<th data-hide="phone,tablet">상태</th>
							</tr>
						</thead>
						<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If IsArray(AryHash) then
									intNUM = 1
									For i = 0 to ubound(AryHash,1)
									' MasterCode, MasterCodeName, State, StateName, 		0~3
									' RegDate, RegID, EditDate, EditID						4~7														
							%>
								<tr class="viewDetail viewSubCode" MasterCode="<%= AryHash(i).Item("MasterCode") %>" State="<%= AryHash(i).Item("State") %>">
									<td><%= AryHash(i).Item("MasterCode") %><!--<% 'aryList(0, i) %>--></td>
									<td><%= AryHash(i).Item("MasterCodeName") %></td>
									<td><%= AryHash(i).Item("StateName") %></td>
								</tr>
							<%
									Next
								end if
							%>
						</tbody>
					</table>
				</div>
			</div>
			<!-- 마스터 코드 리스트 -->

			<!-- 마스터 코드 입력 DIV -->
			<div id="MasterCodeDIV" style="width:500px; display:none;">
			<!--<div id="MasterCodeDIV" style="width:500px;">-->
				<form name="MasterCodeform" id="MasterCodeform" method="post" action="/Process/CodeProc.asp">
				<input type="hidden" name="process" value="RegMasterCode">
				<input type="hidden" name="ProcessType" id="MasterCodeProcessType" value="Insert">
				<table style="width:465px;">
					<tr><td  bgcolor="#d3d3d3" height="1" colspan="2"></td></tr>
					<tr bgcolor="#FAFAFA">
						<td height="35" class="sb_tit_span" colspan="2" align="left">마스터 코드 관리</td>
					</tr>
					<tr>
						<td width="130" height="35" class="sb_tit" align="left">마스터 코드</td>
						<td class="sb_tit_sub" align="left">
							<input type="text" name="MasterCode" id="MasterCode" maxlength="40" class="form-control input-sm" alert="마스터 코드를 입력하세요.">
						</td>
					</tr>
					<tr><td height="1" colspan="2" bgcolor="#d3d3d3"></td></tr>
					<tr>
						<td height="35" class="sb_tit" align="left">마스터 코드명</td>
						<td class="sb_tit_sub" align="left">
							<input type="text" name="MasterCodeName" id="MasterCodeName" maxlength="255" class="form-control input-sm" alert="마스터 코드명을 입력하세요.">
						</td>
					</tr>
					<tr><td height="1" colspan="2" bgcolor="#d3d3d3"></td></tr>
					<tr>
						<td height="35" class="sb_tit" align="left">상태</td>
						<td class="sb_tit_sub" align="left">
							<select name="MasterCodeState" id="MasterCodeState" class="form-control input-sm" alert="상태를 선택하세요.">
								<option value="">상태선택</option>
								<option value="Y">사용</option>
								<option value="N">미사용</option>
							</select>
						</td>
					</tr>
					<tr><td height="1" colspan="2" bgcolor="#d3d3d3"></td></tr>
				</table>
				</form>
				
				<table style="width:465px;">
					<tr>
						<td align="center" style="padding-top:10px;">
							<span class="btnBasic btnTypeSave" id="RegMasterCode" style="width:80px;">저장</span>
							<span class="btnBasic btnTypeClose SelfCloseDIV" style="width:80px;">취소</span>
						</td>
					</tr>
				</table>
			</div>
			<!-- 마스터 코드 입력 DIV -->
		</div>
	</div>
	<!-- 마스터 코드 -->

	<!-- 상세정보 & 서브코드 -->
	<div class="col-lg-7" style="padding-left:0px;">
		<div class="ibox float-e-margins">
			<!-- 상세정보 -->
			<div class="ibox-title">
				<h5>상세정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div style="border:none;padding-top:0px;">
					<form name="SubCodeform" id="SubCodeform" method="post" action="/Process/CodeProc.asp">
						<div style="display:none;">
							<input type="hidden" name="process" value="RegSubCode">
							<input type="hidden" name="ProcessType" id="SubCodeProcessType" value="Insert">
							<input type="hidden" name="SubMasterCodeState" id="SubMasterCodeState" value="">
							<input type="hidden" name="SubCodePage" id="SubCodePage" value="1" />
						</div>

						<div class="row show-grid">
							<div class="col-md-2 grid_sub_title">
								마스터코드명
							</div>
							<div class="col-md-6">
								<input type="text" name="SubMasterCodeName" id="SubMasterCodeName" class="form-control input-sm" readonly>
								<input type="hidden" name="SubMasterCode" id="SubMasterCode">
							</div>
							<div class="col-md-4">
								<span class="btnBasic btnTypeEdit" id="viewMasterCode2">마스터코드수정</span>
							</div>
						</div>
						<div class="row show-grid">
							<div class="col-md-2 grid_sub_title">
								서브 코드
							</div>
							<div class="col-md-4">
								<input type="text" name="SubCode" id="SubCode" class="form-control input-sm" maxlength="25">
								<input type="hidden" name="SubCodeOld" id="SubCodeOld">
							</div>
							<div class="col-md-2 grid_sub_title2">
								서브 코드명
							</div>
							<div class="col-md-4">
								<input type="text" name="SubCodeName" id="SubCodeName" class="form-control input-sm" maxlength="255">
							</div>
						</div>
						<div class="row show-grid">
							<div class="col-md-2 grid_sub_title">
								기타정보1
							</div>
							<div class="col-md-4">
								<input type="text" name="Temp1" id="Temp1" class="form-control input-sm" maxlength="25">
							</div>
							<div class="col-md-2 grid_sub_title2">
								기타정보2
							</div>
							<div class="col-md-4">
								<input type="text" name="Temp2" id="Temp2" class="form-control input-sm" maxlength="255">
							</div>
						</div>
						<div class="row show-grid">
							<div class="col-md-2 grid_sub_title">
								순번
							</div>
							<div class="col-md-4">
								<input type="text" name="Step" id="Step" class="form-control input-sm" maxlength="25" class="KeyTypeNUM">
							</div>
							<div class="col-md-2 grid_sub_title2">
								상태
							</div>
							<div class="col-md-4">
								<select name="State" id="State" class="form-control input-sm">
									<option value="">상태선택</option>
									<option value="Y">사용</option>
									<option value="N">미사용</option>
								</select>
							</div>
						</div>

						<div class="row show-grid grid_sub_button">
							<div class="col-md-12">
								<span class="btnBasic btnTypeSave" id="RegSubCode">저장</span>
								<span class="btnBasic btnTypeCancel" id="ResetSubCode">취소</span>
							</div>
						</div>

					</form>
				</div>
			</div>
			<!-- 상세정보 -->

			<div class="pad_t10"></div>

			<!-- 서브코드 -->
			<div class="ibox-title">
				<h5>서브 코드</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<div>
						<table id="SubCodeListTable" class="table table-striped table-bordered table-hover" style="margin-bottom:0px;">
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>
									<th data-hide="phone,tablet">코드</th>
									<th>코드명</th>
									<th data-hide="phone,tablet">순번</th>
									<th data-hide="phone,tablet">상태</th>
								</tr>
							</thead>
							<tbody>
								<tr>
									<td align="center" colspan="5">등록된 데이터가 없습니다.</td>
								</tr>
							</tbody>
						</table>
					</div>

					<div id="SubCodeListPaging" class="paging" style="display:none;"></div>
				</div>
			</div>
			<!-- 서브코드 -->
		
		</div>
	</div>
	<!-- 상세정보 & 서브코드 -->
</div>


<script type="text/javascript">

	$(function() {
		// 코드관리 -> 마스터코드 테이블에만 사용
		// 테이블 중간에 코드등록 버튼 추가
		var responsiveHelper_dt_MasterCode = undefined;
		var breakpointDefinition2 = {
			tablet : 1024,
			phone : 480
		};

		$("#dt_MasterCode").dataTable({
			// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
			//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
			//"sDom": "<'dt-toolbar'<'col-sm-6 col-xs-12 hidden-xs'f>r>" +
			"sDom": "<'dt-toolbar'<'col-xs-7 col-sm-8'f><'col-xs-5 col-sm-4'<'RegCodeDIV'>r>"+

				"t"+
				//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-6 col-sm-6'p>>",
				//"<'dt-toolbar-footer'<'col-xs-6 col-sm-6'i><'col-xs-12 col-sm-6'p>>"
				"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
			, "oLanguage": {
				"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
			}
			, "preDrawCallback" : function() {
				// Initialize the responsive datatables helper once.
				if (!responsiveHelper_dt_MasterCode) {
					responsiveHelper_dt_MasterCode = new ResponsiveDatatablesHelper($('#dt_MasterCode'), breakpointDefinition2);
				}
			}
			, "rowCallback" : function(nRow) {
				responsiveHelper_dt_MasterCode.createExpandIcon(nRow);
			}
			, "drawCallback" : function(oSettings) {
				responsiveHelper_dt_MasterCode.respond();
			}
			, "pageLength": 15
			, "autoWidth" : true
			, "bSort": false				// 정렬
			, "ordering": false				// 정렬
			//, "bSortClasses": false		// 정렬
			// ,"sScrollY": "200px"			// 스크롤
			//, "scrollCollapse": true		// 스크롤뷰 자동
			//, "bPaginate": false			// 페이징
			//,"paging": false				// 페이징
			//, "info": false				// 상단 인포
		});
		// 코드 등록 버튼 이벤트 추가
		$("div.RegCodeDIV").html('<div class="text-right"><span class="btnBasicTemp" id="viewMasterCode">코드등록</span></div>');
		// 검색버튼 Style Width 값 변경
		$("input[type='search']").css("width", "170px");
		// 버튼 모양 변경
		$(".btnBasicTemp").each(function () {
			var strHtml = "";
			var strTitle = $(this).html();
			var strURL = $(this).attr("url");
			strHtml = "<a class=\"btn btn-labeled btn-danger\"> <span class=\"btn-label\"><i class=\"glyphicon glyphicon-plus-sign\"></i></span>" + strTitle + "</a>";
			$(this).html(strHtml);
			if (strURL != undefined && strURL) { $(this).click(function() { $.goURL(strURL); }); }
		}).css("cursor", "pointer");

		// 마스터 코드 저장
		$("#RegMasterCode").click(function() {
			//if (!$.chkInputValue($("#MasterCode"),			"마스터 코드를 입력해 주시기 바랍니다.")) { return; }
			//if (!$.chkInputValue($("#MasterCodeName"),		"마스터 코드명을 입력해 주시기 바랍니다.")) { return; }
			//if (!$.chkInputValue($("#MasterCodeState"),		"상태를 선택해 주시기 바랍니다.")) { return; }

			if ($.setValidation($("#MasterCodeform"))) {
				if (confirm("마스터 코드를 저장 하시겠습니까?")) {
					$.Ajax4FormSubmit($("#MasterCodeform"), "마스터 코드 저장이 완료되었습니다.");
				}
			}
		});

		// 마스터 코드 등록 오픈
		$("#viewMasterCode").click(function() {
			$("#MasterCode").val("").prop("readonly", false);
			$("#MasterCodeName").val("");
			$("#MasterCodeState").val("").trigger("change");
			$("#MasterCodeProcessType").val("Insert");
			$.openMadal($("#MasterCodeDIV"), "2");
		});
		
		// 마스터 코드 수정 오픈
		$("#viewMasterCode2").click(function() {
			if (!$.chkInputValue($("#SubMasterCode"),		"마스터 코드를 선택해 주시기 바랍니다.")) { return; }
			
			$("#MasterCode").val($("#SubMasterCode").val()).prop("readonly", true);
			$("#MasterCodeName").val($("#SubMasterCodeName").val());
			$("#MasterCodeState").val($("#SubMasterCodeState").val()).trigger("change");
			$("#MasterCodeProcessType").val("Update");
			$.openMadal($("#MasterCodeDIV"), "2");
		});
		
		// 내용 상세보기 값 입력 2 (DataField input 사용)
		$(document).on("click", "tr.viewSubCode", function(){
		//$(document).delegate("tr.viewSubCode", "click", function() {
			// 마스터 코드 정보 입력
			$("#SubMasterCode").val($(this).attr("MasterCode"));
			$("#SubMasterCodeName").val($(this).children("td").eq(1).text());
			$("#SubMasterCodeState").val($(this).attr("State"));
			
			// 서브 코드 정보 초기화 & 리스트 호출
			$.ReSetSubCodeForm();
			$.getSubCodeList();
		});

		// 서브 코드 저장
		$("#RegSubCode").click(function() {
			if (!$.chkInputValue($("#SubMasterCode"),		"마스터 코드를 선택해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($("#SubCode"),				"서브 코드를 입력해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($("#SubCodeName"),			"서브 코드명을 입력해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($("#Step"),				"순번을 입력해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($("#State"),				"상태를 선택해 주시기 바랍니다.")) { return; }
			
			if (confirm("서브 코드를 저장 하시겠습니까?")) {
				var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setSubCode(datas)','complete':'','clear':'','reset':''};
				objOpt["url"] = "/Process/CodeProc.asp";
				$.Ajax4Form("#SubCodeform", objOpt);
				$("#SubCodeform").submit();
			}
		});
		
		// 서브 코트 입력 취소
		$("#ResetSubCode").click(function() {
			if (confirm("서브 코드 입력을 취소 하시겠습니까?")) {
				$.ReSetSubCodeForm();
			}
		});

		// SubCode 리스트 XML 받기
		$.getSubCodeList = function () {
			var aryData = { "process": "getSubCodeList","MasterCode": $("#SubMasterCode").val(), "page": $("#SubCodePage").val() }
			$.Ajax4Get("/Process/CodeProc.asp", aryData, "$.setSubCodeList(datas)", "xml", "", "", false);
		}
		
		// SubCode 리스트 설정
		$.setSubCodeList = function (datas) {
			var $objList = $(datas).find("Lists");
			var $objPageInfo = $objList.find("PageInfo");
			var $objItem = $objList.find("item");
			var strHTML = "";
			
			if ($objList.find("Result").text() == "Complete") {
				// SubCode 리스트 설정
				strHTML += "<thead>";
				strHTML += "	<tr>";
				strHTML += "		<th data-hide=\"phone\">No.</th>";
				strHTML += "		<th data-hide=\"phone,tablet\">코드</th>";
				strHTML += "		<th>코드명</th>";
				strHTML += "		<th data-hide=\"phone,tablet\">순번</th>";
				strHTML += "		<th data-hide=\"phone,tablet\">상태</th>";
				strHTML += "	</tr>";
				strHTML += "</thead>";
				
				if ($objItem.length != 0) {
					strHTML += "<tbody>";
					$objItem.each(function(i) {
						// SubCode, SubCodeName, Step, Temp1, Temp2, 
						// Temp3, Temp4, TempEtc, UseYN, State, StateName
						strHTML += "	<tr class=\"viewDetail3 viewSubCodeInfo\" "
						strHTML += "				SubCode=\""+ $(this).find("SubCode").text() +"\" "
						strHTML += "				SubCodeName=\""+ $(this).find("SubCodeName").text() +"\" "
						strHTML += "				Step=\""+ $(this).find("Step").text() +"\" "
						strHTML += "				Temp1=\""+ $(this).find("Temp1").text() +"\" "
						strHTML += "				Temp2=\""+ $(this).find("Temp2").text() +"\" "
						strHTML += "				State=\""+ $(this).find("State").text() +"\" "
						strHTML += "	>";
						strHTML += "		<td align=\"center\">"+ $(this).find("Num").text() +"</td>";
						strHTML += "		<td align=\"center\">"+ $(this).find("SubCode").text() +"</td>";
						strHTML += "		<td class=\"pd20\">"+ $(this).find("SubCodeName").text() +"</a>";
						strHTML += "		</td>";
						strHTML += "		<td align=\"center\">"+ $(this).find("Step").text() +"</td>";
						strHTML += "		<td align=\"center\">"+ $(this).find("StateName").text() +"</td>";
						strHTML += "	</tr>";
					});
					strHTML += "</tbody>";
				} else {
					strHTML += "<tbody>";
					strHTML += "	<tr>";
					strHTML += "		<td align=\"center\" colspan=\"5\">등록된 데이터가 없습니다.</td>";
					strHTML += "	</tr>";
					strHTML += "</tbody>";
				}
				
				// SubCode 리스트 삽입
				$("#SubCodeListTable").empty().append(strHTML);
				// 서브 코드 상세보기
				$(".viewSubCodeInfo").click(function() {
					$("#SubCode").val($(this).attr("SubCode"));
					$("#SubCodeOld").val($(this).attr("SubCode"));
					$("#SubCodeName").val($(this).attr("SubCodeName"));
					$("#Temp1").val($(this).attr("Temp1"));
					$("#Temp2").val($(this).attr("Temp2"));
					$("#Step").val($(this).attr("Step"));
					$("#State").val($(this).attr("State")).trigger("change");
					$("#SubCodeProcessType").val("Update");
				}).css("cursor", "pointer");

				$(".viewDetail3").click(function() {
					// 색상 입력
					//$(this).siblings().css({ backgroundColor:""});	// 형제 노드 초기화
					$(this).parent().find("tr").css({ backgroundColor:""});
					$(this).css({ backgroundColor:"#F5F5DC"});		// 선택 노트 색생 변경
				});

				// 페이지 생성
				$.makePageBlock($objPageInfo.find("PageNum").text(), $objPageInfo.find("PageBlock").text(), $objPageInfo.find("PageCount").text(), "#SubCodeListPaging");
				
				// 페이지 번호 클릭 시 SubCode 리스트 재호출
				$(".pageNUM").click(function () {
					if ($(this).attr("value") != undefined) {
						$("#SubCodePage").val($(this).attr("value"));
						$.getSubCodeList();
					}
				});

				// 데이터가 없으면 페이징 보이지 않게 처리
				if ($objItem.length != 0) {
					$("#SubCodeListPaging").show()
				} else {
					$("#SubCodeListPaging").hide()
				}
			} else {
				alert("서브 코드 추출 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
				return;
			}
		}
		
		// 서브 코드 저장 결과
		$.setSubCode = function(datas) {
			var $objList	= $(datas).find("List");	
			var strMSG;
			
			if ($objList.find("Result").text() == "Complete") {
				alert("서브 코드 저장이 완료 되었습니다.");
				
				// 서브 코드 정보 초기화 & 리스트 호출
				$.ReSetSubCodeForm();
				$.getSubCodeList();
			} else {
				alert("서브 코드 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
				return;
			}
		}

		// 서브 코드 폼 리셋
		$.ReSetSubCodeForm = function () {
			$("#SubCode").val("");
			$("#SubCodeOld").val("");
			$("#SubCodeName").val("");
			$("#Temp1").val("");
			$("#Temp2").val("");
			$("#Step").val("");
			$("#State").val("").trigger("change");
			$("#SubCodePage").val("1");
			$("#SubCodeProcessType").val("Insert");
		}
	})

</script>


<!--#InClude Virtual = "/Common/Bottom.asp" -->