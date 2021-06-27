<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 3
Dim LeftMenuCode : LeftMenuCode = "Appraisal"
Dim LeftMenuName : LeftMenuName = "Home / 평가기준관리 / 평가비율 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "평가비율 설정"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, AryHash2
Dim i, strMSG, intNUM, strTEMP, strRESULT

'검색 조건
'Dim SearchMYear		: SearchMYear = fnR("SearchMYear", SessionMYear)
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", "")
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", "")
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", "")
Dim SearchDivision2	: SearchDivision2 = fnR("SearchDivision2", "")
Dim SearchDivision3	: SearchDivision3 = fnR("SearchDivision3", "")
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/AppraisalList.asp"

'페이지설정(사이즈는 검색)
Dim PageSize		: PageSize = getIntParameter(FnR("PageSize", 5), 5)
Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'서브쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And a.Division0 = ? "
	Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 50, SearchDivision)
end If

if not(IsE(SearchSubject)) And SearchSubject <> "All" then
	strWhere = strWhere & " And a.Subject = ? "
	Call objDB.sbSetArray("@Subject", adVarchar, adParamInput, 50, SearchSubject)
end If

if not(IsE(SearchDivision1)) And SearchDivision1 <> "All" then
	strWhere = strWhere & " And a.Division1 = ? "
	Call objDB.sbSetArray("@Division1", adVarchar, adParamInput, 50, SearchDivision1)
end If

if not(IsE(SearchDivision2)) And SearchDivision2 <> "All" then
	strWhere = strWhere & " And a.Division2 = ? "
	Call objDB.sbSetArray("@Division2", adVarchar, adParamInput, 50, SearchDivision2)
end If

if not(IsE(SearchDivision3)) And SearchDivision3 <> "All" then
	strWhere = strWhere & " And a.Division3 = ? "
	Call objDB.sbSetArray("@Division3", adVarchar, adParamInput, 50, SearchDivision3)
end If

'평가비율 리스트 쿼리
SQL = ""
SQL = SQL & vbCrLf & " SELECT "
SQL = SQL & vbCrLf & "	b.IDX,a.MYear,a.SubjectCode"
SQL = SQL & vbCrLf & "	, a.Division0, a.Subject, a.Division1, a.Division2, a.Division3 "
SQL = SQL & vbCrLf & "	, a.Quorum,a.QuorumFix"
SQL = SQL & vbCrLf & "	, a.RF1,a.RF2,a.RF3,a.RF4,a.RF5,a.RF6,a.RF7,a.RF8,a.RF9,a.RF10,a.RF11"
SQL = SQL & vbCrLf & "	, b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio"
SQL = SQL & vbCrLf & "	, b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5, b.DrawStandard6"
SQL = SQL & vbCrLf & "	, b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5, b.UnqualifiedStandard6"
SQL = SQL & vbCrLf & "	, b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5, b.ExtraPoint6"
SQL = SQL & vbCrLf & "	, b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5, b.Scholarship6"
SQL = SQL & vbCrLf & "	, b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5"
SQL = SQL & vbCrLf & "	, b.DocumentaryEvidence6, b.DocumentaryEvidence7, b.DocumentaryEvidence8, b.DocumentaryEvidence9, b.DocumentaryEvidence10"
SQL = SQL & vbCrLf & "	, b.INPT_USID, b.INPT_DATE, b.INPT_ADDR, b.UPDT_USID, b.UPDT_DATE, b.UPDT_ADDR, b.InsertTime"

SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', a.Division0) AS Division0Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', a.Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division3', a.Division3) AS Division3Name "

SQL = SQL & vbCrLf & "FROM SubjectTable AS a "
SQL = SQL & vbCrLf & "	left outer join AppraisalTable AS b"
SQL = SQL & vbCrLf & "		on a.SubjectCode = b.SubjectCode"
SQL = SQL & vbCrLf & "WHERE 1 = 1 " & strWhere
SQL = SQL & vbCrLf & "ORDER BY a.IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB = Nothing

if IsArray(AryHash) Then
	'// 페이지 계산
	TotalCount = ubound(AryHash,1) + 1
	PageCount = int((TotalCount - 1) / PageSize) + 1
	StartNum = (PageNum * PageSize) - PageSize
	EndNum = StartNum + PageSize - 1
	intNUM = TotalCount - (PageNum * PageSize) + PageSize

	If EndNum > TotalCount - 1 Then
		EndNum = TotalCount - 1
	End If
End If
%>

<script type="text/javascript">
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// 목록 클릭 시 업데이트 프로세스로 변경
	//$(document).delegate("tr.viewDetail_SetDate_2", "click", function() {
	//	$("#InputForm input[name='ProcessType']").val("Update");
	//});

	// 신규
	$("#btnNew").click(function () {
		if (confirm("입력되어 있던 내용이 초기화 됩니다.\n신규로 입력 하시겠습니까?")) {
			$.FormReset($("#InputForm"));
		}
	});

	// 저장
	$("#btnSave").click(function () { 
		if ($.setValidation($("#InputForm"))) {
			if (!$("input[name=chk]:checked").val()) {
				alert("선택된 모집단위가 없습니다.");
				return;
			}

			//비율 합계가 100이 아니면 경고
			var StudentRecordRatio
			var InterviewerRatio
			var PracticalRatio
			var CSATRatio
			var TotalRatio

			if (!$("select[name=StudentRecordRatio]").val()) {
				StudentRecordRatio = 0;
			}else {
				StudentRecordRatio = $("select[name=StudentRecordRatio]").val();
			}
			if (!$("select[name=InterviewerRatio]").val()) {
				InterviewerRatio = 0;
			}else {
				InterviewerRatio = $("select[name=InterviewerRatio]").val();
			}
			if (!$("select[name=PracticalRatio]").val()) {
				PracticalRatio = 0;
			}else {
				PracticalRatio = $("select[name=PracticalRatio]").val();
			}
			if (!$("select[name=CSATRatio]").val()) {
				CSATRatio = 0;
			}else {
				CSATRatio = $("select[name=CSATRatio]").val();
			}

			TotalRatio = parseInt(StudentRecordRatio) + parseInt(InterviewerRatio) + parseInt(PracticalRatio) + parseInt(CSATRatio)

			if (TotalRatio != 100) {
				alert("비율 합이 " + TotalRatio + "% 입니다. 100%로 설정하셔야 합니다.");
				return;
			}	

			if (confirm("입력하신 내용을 저장 하시겠습니까?")) {
				/////////////////////////////////////////////////////////////////////////////////////////
				//선택된 체크박스의 값(모집코드)을 가져와서 SubjectCodeHidden변수에 넣어준 후 submit
				/////////////////////////////////////////////////////////////////////////////////////////
				var checkBoxArr = [];
				var checkBoxArr2 = [];
				$("input[name=chk]:checked").each(function(i){					
					checkBoxArr.push($(this).val());
					checkBoxArr2.push($(this).attr('Myear'));	
				});
				$("#SubjectCodehidden").val(checkBoxArr);
				$("#MyearHidden").val(checkBoxArr2);
				$.Ajax4FormSubmit($("#InputForm"), "입력하신 정보 저장이 완료되었습니다.");
			}
		}
	});

	/*
	// 삭제
	$("#btnDelete").click(function () {
		if (!$.chkInputValue($("#InputForm input[name='IDX']"),	"삭제할 항목을 선택해 주세요.")) { return; }

		if (confirm("선택된 항목을 삭제 하시겠습니까?")) {
			$("#InputForm input[name='ProcessType']").val("Delete");
			$.Ajax4FormSubmit($("#InputForm"), "선택된 항목이 삭제되었습니다.");
		}
	});
	*/

	// 전체선택 체크박스
    $("#checkall").click(function(){
        if($("#checkall").prop("checked")){
            $("input[name=chk]").prop("checked",true);
        }else{
            $("input[name=chk]").prop("checked",false);
        }
    });

	// tr선택 시 체크
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
		var $Checkbox = $(this).find("input[type='Checkbox']")
		if ($Checkbox.is(":checked")) {
			$Checkbox.prop("checked", false); 
		} else {
			$Checkbox.prop("checked", true); 
		}
	});

	// 체크박스 선택 시 체크 
	$(document).on("click", "input.CheckboxCheck", function(){
		var $Checkbox = $(this)
		if ($Checkbox.is(":checked")) {
			$Checkbox.prop("checked", false); 
		} else {
			$Checkbox.prop("checked", true); 
		}
	});	

	// 기본 데이터 설정(모달) 오픈
	$("#BasicDataSet").click(function() {
		$("#BasicDataSetProcessType").val("Update");
		$.openMadal($("#BasicDataSetModal"), "2");
	});

	// 기본 데이터 저장
	$("#RegBasicDataSet").click(function() {
		if (!$.chkInputValue($("select[name=BasicDataBtn]"),		"버튼번호를 선택해 주시기 바랍니다.")) { return; }
		
		if (confirm("기본데이터를 저장 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setBasicData(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/BasicDataProc.asp";
			$.Ajax4Form("#BasicDataSetForm", objOpt);
			$("#BasicDataSetForm").submit();
		}
	});

	// 기본 데이터 저장 결과
	$.setBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {
			alert("기본 데이터가 저장 되었습니다.");
				
		} else {
			alert("기본 데이터가 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 기본 데이터 넣기
	$.PutBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
			
		if ($objList.find("Result").text() == "Complete") {
			 
			 $("select[name=StudentRecordRatio]").val($objList.find("StudentRecordRatio").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=InterviewerRatio]").val($objList.find("InterviewerRatio").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=PracticalRatio]").val($objList.find("PracticalRatio").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=CSATRatio]").val($objList.find("CSATRatio").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=DrawStandard1]").val($objList.find("DrawStandard1").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DrawStandard2]").val($objList.find("DrawStandard2").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DrawStandard3]").val($objList.find("DrawStandard3").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DrawStandard4]").val($objList.find("DrawStandard4").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DrawStandard5]").val($objList.find("DrawStandard5").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DrawStandard6]").val($objList.find("DrawStandard6").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=UnqualifiedStandard1]").val($objList.find("UnqualifiedStandard1").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=UnqualifiedStandard2]").val($objList.find("UnqualifiedStandard2").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=UnqualifiedStandard3]").val($objList.find("UnqualifiedStandard3").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=UnqualifiedStandard4]").val($objList.find("UnqualifiedStandard4").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=UnqualifiedStandard5]").val($objList.find("UnqualifiedStandard5").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=UnqualifiedStandard6]").val($objList.find("UnqualifiedStandard6").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=ExtraPoint1]").val($objList.find("ExtraPoint1").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=ExtraPoint2]").val($objList.find("ExtraPoint2").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=ExtraPoint3]").val($objList.find("ExtraPoint3").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=ExtraPoint4]").val($objList.find("ExtraPoint4").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=ExtraPoint5]").val($objList.find("ExtraPoint5").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=ExtraPoint6]").val($objList.find("ExtraPoint6").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=Scholarship1]").val($objList.find("Scholarship1").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Scholarship2]").val($objList.find("Scholarship2").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Scholarship3]").val($objList.find("Scholarship3").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Scholarship4]").val($objList.find("Scholarship4").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Scholarship5]").val($objList.find("Scholarship5").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=Scholarship6]").val($objList.find("Scholarship6").text()).prop("selected", true).trigger("chosen:updated");

			 $("select[name=DocumentaryEvidence1]").val($objList.find("DocumentaryEvidence1").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence2]").val($objList.find("DocumentaryEvidence2").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence3]").val($objList.find("DocumentaryEvidence3").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence4]").val($objList.find("DocumentaryEvidence4").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence5]").val($objList.find("DocumentaryEvidence5").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence6]").val($objList.find("DocumentaryEvidence6").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence7]").val($objList.find("DocumentaryEvidence7").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence8]").val($objList.find("DocumentaryEvidence8").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence9]").val($objList.find("DocumentaryEvidence9").text()).prop("selected", true).trigger("chosen:updated");
			 $("select[name=DocumentaryEvidence10]").val($objList.find("DocumentaryEvidence10").text()).prop("selected", true).trigger("chosen:updated");
				
		} else {
			alert("기본 데이터 넣기 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
});

//기본 데이터 가져오기
function BasicDataBtn(num) {
	if (confirm(num + "번 기본데이터를 넣으시겠습니까?")) {
		$("#BasicDataBtnNum").val(num);

		var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.PutBasicData(datas)','complete':'','clear':'','reset':''};
		objOpt["url"] = "/Process/BasicDataSelect.asp";
		$.Ajax4Form("#BasicDataBtnFrom", objOpt);
		$("#BasicDataBtnFrom").submit();
	}
}
</script>
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 검색조건 -->
			<div class="ibox-title">
				<h5>검색정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div>
					<form id="SearchForm" method="get">
					<input type="hidden" name="Page" value="<%= PageNum %>">
						<div class="row show-grid">
							<div class="col-md-1 col-xs-1 grid_sub_title">
								모집시기
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchDivision", "모집시기선택", SearchDivision, "", "All", "Division0") %>
							</div>
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								학과
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchSubject", "학과명선택", SearchSubject, "", "All", "Subject") %>
							</div>
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								구분1
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchDivision1", "구분1선택", SearchDivision1, "", "All", "Division1") %>
							</div>	
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								구분2
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchDivision2", "구분2선택", SearchDivision2, "", "All", "Division2") %>
							</div>
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								구분3
							</div>
							<div class="col-md-2 col-xs-2">
								<% Call SubCodeSelectBox("SearchDivision3", "구분3선택", SearchDivision3, "", "All", "Division3") %>
							</div>
						</div>
						<div class="pad_t10 pad_r10 text-right">
							<span class="btnBasic btnSubmit">조회</span>
						</div>	
					<!--</form>-->
				</div>
			</div>
			<!-- 검색조건 끝 -->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div class="ibox-tools">
					<!--<a class="collapse-link">-->
						<!--<form id="PageSizeForm" method="get">-->
							<div class="col-md-1 col-xs-2" style="float:right;">
								<select name = "PageSize" onChange="SearchForm.submit();">
									<option value="5" <% If PageSize = 5 then response.write "selected" end if%>>5개씩 보기</option>
									<option value="15" <% If PageSize = 15 then response.write "selected" end if%>>15개씩 보기</option>
									<option value="30" <% If PageSize = 30 then response.write "selected" end if%>>30개씩 보기</option>
									<option value="50" <% If PageSize = 50 then response.write "selected" end if%>>50개씩 보기</option>
									<option value="100" <% If PageSize = 100 then response.write "selected" end if%>>100개씩 보기</option>
									<option value="200" <% If PageSize = 200 then response.write "selected" end if%>>200개씩 보기</option>
								</select>
							</div>
						</form>
						<!--<i class="fa fa-chevron-up"></i>-->
					<!--</a>-->
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<thead>			                
								<tr>
									<!-- 체크박스 추가-평가비율 등록 시, 체크박스가 선택 된 모든 모집단위 없데이트 -->
									<th colspan="1" style="padding: 7px 0px 6px 0px; text-align: center;">
										<input type="checkbox" id="checkall"/>
									</th>
									<!--<th data-hide="phone">No.</th>-->
									<th data-hide="phone">년도</th>
									<th data-hide="phone">모집코드</th>
									<th>모집시기</th>
									<th>학과명</th>
									<th>구분1</th>
									<th>구분2</th>
									<th>구분3</th>
									<th data-hide="phone,tablet">최초입력자</th>
									<th data-hide="phone,tablet">최초등록일</th>
									<th data-hide="phone">최종수정자</th>
									<th data-hide="phone">최종수정일</th>
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									'For i = 0 to ubound(AryHash,1)
									For i = StartNum to EndNum
							%>
								<tr class="viewDetail_SetDate_2" IDX="<%= AryHash(i).Item("IDX") %>">
									<td colspan="1" style="background-color: <%=BGColor%>; padding: 8px 0px 0px 0px; text-align: center;">
										<input class="CheckboxCheck" type="Checkbox" name="chk" ID="Checkbox<%=i%>" value="<%=AryHash(i).Item("SubjectCode")%>" Myear="<%= AryHash(i).Item("MYear") %>">
									</td>
									<!--<td><%= intNUM %></td>-->
									<td><%= AryHash(i).Item("MYear") %></td>
									<td><%= AryHash(i).Item("SubjectCode") %></td>
									<td><%= AryHash(i).Item("Division0Name") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("Division2Name") %></td>
									<td><%= AryHash(i).Item("Division3Name") %></td>
									<td><%= AryHash(i).Item("INPT_USID") %></td>
									<td><%= Left(AryHash(i).Item("INPT_DATE"),10) %></td>
									<td><%= AryHash(i).Item("UPDT_USID") %></td>
									<td>
										<%= Left(AryHash(i).Item("UPDT_DATE"),10) %>
										<div class="DataField" style="display:none;">
											<!--<li Columnvalue="Update"												ColumnName="ProcessType"></li>-->
											<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>"						ColumnName="IDX"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>"						ColumnName="MYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectCode")) %>"				ColumnName="SubjectCode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division0")) %>"					ColumnName="Division0"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Subject")) %>"					ColumnName="Subject"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1")) %>"					ColumnName="Division1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division2")) %>"					ColumnName="Division2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division3 ")) %>"				ColumnName="Division3 "></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordRatio")) %>"		ColumnName="StudentRecordRatio"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("InterviewerRatio")) %>"			ColumnName="InterviewerRatio"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("PracticalRatio")) %>"			ColumnName="PracticalRatio"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("CSATRatio")) %>"					ColumnName="CSATRatio"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard1")) %>"				ColumnName="DrawStandard1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard2")) %>"				ColumnName="DrawStandard2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard3")) %>"				ColumnName="DrawStandard3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard4")) %>"				ColumnName="DrawStandard4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard5")) %>"				ColumnName="DrawStandard5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard6")) %>"				ColumnName="DrawStandard6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard1")) %>"		ColumnName="UnqualifiedStandard1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard2")) %>"		ColumnName="UnqualifiedStandard2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard3")) %>"		ColumnName="UnqualifiedStandard3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard4")) %>"		ColumnName="UnqualifiedStandard4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard5")) %>"		ColumnName="UnqualifiedStandard5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard6")) %>"		ColumnName="UnqualifiedStandard6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint1")) %>"				ColumnName="ExtraPoint1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint2")) %>"				ColumnName="ExtraPoint2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint3")) %>"				ColumnName="ExtraPoint3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint4")) %>"				ColumnName="ExtraPoint4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint5")) %>"				ColumnName="ExtraPoint5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint6")) %>"				ColumnName="ExtraPoint6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship1")) %>"				ColumnName="Scholarship1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship2")) %>"				ColumnName="Scholarship2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship3")) %>"				ColumnName="Scholarship3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship4")) %>"				ColumnName="Scholarship4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship5")) %>"				ColumnName="Scholarship5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship6")) %>"				ColumnName="Scholarship6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence1")) %>"		ColumnName="DocumentaryEvidence1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence2")) %>"		ColumnName="DocumentaryEvidence2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence3")) %>"		ColumnName="DocumentaryEvidence3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence4")) %>"		ColumnName="DocumentaryEvidence4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence5")) %>"		ColumnName="DocumentaryEvidence5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence6")) %>"		ColumnName="DocumentaryEvidence6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence7")) %>"		ColumnName="DocumentaryEvidence7"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence8")) %>"		ColumnName="DocumentaryEvidence8"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence9")) %>"		ColumnName="DocumentaryEvidence9"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence10")) %>"		ColumnName="DocumentaryEvidence10"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division0Name")) %>"				ColumnName="Division0Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectName")) %>"				ColumnName="SubjectName"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1Name")) %>"				ColumnName="Division1Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division2Name")) %>"				ColumnName="Division2Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division3Name")) %>"				ColumnName="Division3Name"></li>
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
					</form>

					<div class="paging pad_r10">&nbsp;</div>
				</div>
				
				
			</div>
			<!-- 테이블 끝 -->

			<div class="pad_t10"></div>

			<!-- 상세보기 -->
			<div class="ibox-title">
				<h5>상세정보</h5>
				<div style="float:right;">
					<span class="btnBasic btnTypeEdit" id="BasicDataSet">기본 데이터 설정</span>
					<span class="btnBasic btnTypeSave" id="BasicData1" onclick="BasicDataBtn(1)">1. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData2" onclick="BasicDataBtn(2)">2. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData3" onclick="BasicDataBtn(3)">3. 기본 데이터 넣기</span>
					<form id="BasicDataBtnFrom" method="post" action="/Process/BasicDataSelect.asp">
						<div style="display:none;">
							<input type="text" name="process" id="process" value="RegAppraisalBasicDataSet">
							<input type="text" name="BasicDataBtnNum" id="BasicDataBtnNum" value="">
						</div>
					</form>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/AppraisalProc.asp">
					<div style="display:none;">
						<input type="text" name="process" id="process" value="RegAppraisal">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<!-- 1개씩 insert or update 할 때 사용.
						<input type="text" name="ProcessType" id="ProcessType" <% If IsEmpty(IDX) Or IDX = "" Then %> value="Insert" <% Else %> value="Update" <% End IF%>>
						-->
						<input type="text" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="text" name="SubjectCodehidden" id="SubjectCodehidden" value="">
						<input type="text" name="MyearHidden" id="MyearHidden" value="">
					</div>

					<!-- 선택한 학과 기본정보 -->
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							년도 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="MYear" class="form-control input-sm" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							모집시기 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division0Name" class="form-control input-sm" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							학과 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="SubjectName" class="form-control input-sm" readonly>
						</div>
					</div>
					<div class="row show-grid">

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							전형 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division1Name" class="form-control input-sm" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							구분2 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division2Name" class="form-control input-sm" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							구분3 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division3Name" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">&nbsp;</div>

					<!-- 평가비율 -->
					<div class="row show-grid">
						<div class="col-md-1">
							생기부비율 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("StudentRecordRatio", "생기부비율 선택", StudentRecordRatio, "", "", "StudentRecordRatio") %>
						</div>
						<div class="col-md-1">
							면접비율 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("InterviewerRatio", "면접비율 선택", InterviewerRatio, "", "", "InterviewerRatio") %>
						</div>

						<div class="col-md-1">
							실기비율
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("PracticalRatio", "실기비율 선택", PracticalRatio, "", "", "PracticalRatio") %>
						</div>

						<div class="col-md-1">
							수능비율 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("CSATRatio", "수능비율 선택", CSATRatio, "", "", "CSATRatio") %>
						</div>
					</div>

					<!-- 자격미달기준 -->
					<div class="row show-grid">
						<div class="col-md-1">
							자격미달기준1
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard1", "자격미달기준 선택", DrawStandard1, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준2
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard2", "자격미달기준 선택", DrawStandard2, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준3
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard3", "자격미달기준 선택", DrawStandard3, "", "", "DrawStandard") %>
						</div>
					<!--</div>
					<div class="row show-grid">
						<div class="col-md-1">
							자격미달기준4
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard4", "자격미달기준 선택", DrawStandard4, "", "", "DrawStandard") %>
						</div>
						<div class="col-md-1">
							자격미달기준5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard5", "자격미달기준 선택", DrawStandard5, "", "", "DrawStandard") %>
						</div>
						<div class="col-md-1">
							자격미달기준6
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard6", "자격미달기준 선택", DrawStandard6, "", "", "DrawStandard") %>
						</div>
					</div>-->

					<!-- 동석차기준 -->
					<!--<div class="row show-grid">-->
						<div class="col-md-1">
							동석차기준
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("UnqualifiedStandard1", "동석차기준 선택", UnqualifiedStandard1, "", "", "UnqualifiedStandard") %>
						</div>

						<!--<div class="col-md-1">
							동석차기준2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard2", "동석차기준 선택", UnqualifiedStandard2, "", "", "UnqualifiedStandard") %>
						</div>

						<div class="col-md-1">
							동석차기준3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard3", "동석차기준 선택", UnqualifiedStandard3, "", "", "UnqualifiedStandard") %>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-1">
							동석차기준4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard4", "동석차기준 선택", UnqualifiedStandard4, "", "", "UnqualifiedStandard") %>
						</div>
						<div class="col-md-1">
							동석차기준5 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard5", "동석차기준 선택", UnqualifiedStandard5, "", "", "UnqualifiedStandard") %>
						</div>
						<div class="col-md-1">
							동석차기준6 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard6", "동석차기준 선택", UnqualifiedStandard6, "", "", "UnqualifiedStandard") %>
						</div>-->
					</div>

					<!-- 가산점 -->
					<!--<div class="row show-grid">
						<div class="col-md-1">
							가산점1 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint1", "가산점 선택", ExtraPoint1, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint2", "가산점 선택", ExtraPoint2, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint3", "가산점 선택", ExtraPoint3, "", "", "ExtraPoint") %>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-1">
							가산점4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint4", "가산점 선택", ExtraPoint4, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점5 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint5", "가산점 선택", ExtraPoint5, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점6 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint6", "가산점 선택", ExtraPoint6, "", "", "ExtraPoint") %>
						</div>
					</div>-->

					<!-- 장학 -
					<div class="row show-grid">
						<div class="col-md-1">
							장학1 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship1", "", Scholarship1, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학2 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship2", "", Scholarship2, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학3 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship3", "", Scholarship3, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학4 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship4", "", Scholarship4, "", "", "Scholarship") %>
						</div>
					</div>->

					<!-- 필수서류 -->
					<!--<div class="row show-grid">
						<div class="col-md-1">
							필수서류1 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence1", "필수서류 선택", DocumentaryEvidence1, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence2", "필수서류 선택", DocumentaryEvidence2, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence3", "필수서류 선택", DocumentaryEvidence3, "", "", "DocumentaryEvidence") %>
						</div>
					</div>
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							필수서류4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence4", "필수서류 선택", DocumentaryEvidence4, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence5", "필수서류 선택", DocumentaryEvidence5, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류6
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence6", "필수서류 선택", DocumentaryEvidence6, "", "", "DocumentaryEvidence") %>
						</div>
					</div>-->

					<br>

					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">초기화</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>

					<br>
					<br>

				</form>
			</div>
			<!-- 상세보기 끝 -->

			<!-- 기본 데이터 설정 -->
			<div id="BasicDataSetModal" style="width:100%; margin:5px; display:none;">
				<form name="BasicDataSetForm" id="BasicDataSetForm" method="post" action="/Process/BasicDataProc.asp">
				<input type="hidden" name="BasicDataSetprocess" value="RegAppraisalBasicDataSet">
				<input type="hidden" name="BasicDataSetProcessType" id="BasicDataSetProcessType" value="Insert">
				<div class="ibox-content">		
					<!-- 버튼 번호 -->
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							버튼번호
						</div>
						<div class="col-md-2" style="text-align:left;">
							<% Call SubCodeSelectBox("BasicDataBtn", "버튼번호 선택", "", "", "", "BasicDataBtn") %>
						</div>
					</div>

					<!-- 평가비율 -->
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							생기부비율
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("StudentRecordRatio", "생기부비율 선택", StudentRecordRatio, "", "", "StudentRecordRatio") %>
						</div>
						<div class="col-md-1">
							면접비율 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("InterviewerRatio", "면접비율 선택", InterviewerRatio, "", "", "InterviewerRatio") %>
						</div>

						<div class="col-md-1">
							실기비율
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("PracticalRatio", "실기비율 선택", PracticalRatio, "", "", "PracticalRatio") %>
						</div>

						<div class="col-md-1">
							수능비율 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("CSATRatio", "수능비율 선택", CSATRatio, "", "", "CSATRatio") %>
						</div>
					</div>

					<!-- 자격미달기준 -->
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							자격미달기준1
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard1", "자격미달기준 선택", DrawStandard1, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준2
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard2", "자격미달기준 선택", DrawStandard2, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준3
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("DrawStandard3", "자격미달기준 선택", DrawStandard3, "", "", "DrawStandard") %>
						</div>
					<!--</div>
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							자격미달기준4
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard4", "자격미달기준 선택", DrawStandard4, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard5", "자격미달기준 선택", DrawStandard5, "", "", "DrawStandard") %>
						</div>

						<div class="col-md-1">
							자격미달기준6
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DrawStandard6", "자격미달기준 선택", DrawStandard6, "", "", "DrawStandard") %>
						</div>
					</div>-->

					<!-- 동석차기준 -->
					<!--<div class="row show-grid" style="text-align:left;">-->
						<div class="col-md-1">
							동석차기준1 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("UnqualifiedStandard1", "동석차기준 선택", UnqualifiedStandard1, "", "", "UnqualifiedStandard") %>
						</div>
					<!--
						<div class="col-md-1">
							동석차기준2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard2", "동석차기준 선택", UnqualifiedStandard2, "", "", "UnqualifiedStandard") %>
						</div>

						<div class="col-md-1">
							동석차기준3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard3", "동석차기준 선택", UnqualifiedStandard3, "", "", "UnqualifiedStandard") %>
						</div>
					</div>
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							동석차기준4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard4", "동석차기준 선택", UnqualifiedStandard4, "", "", "UnqualifiedStandard") %>
						</div>

						<div class="col-md-1">
							동석차기준5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard5", "동석차기준 선택", UnqualifiedStandard5, "", "", "UnqualifiedStandard") %>
						</div>

						<div class="col-md-1">
							동석차기준6 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("UnqualifiedStandard6", "동석차기준 선택", UnqualifiedStandard6, "", "", "UnqualifiedStandard") %>
						</div>-->
					</div>

					<!-- 가산점 -->
					<!--
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							가산점1 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint1", "", ExtraPoint1, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint2", "", ExtraPoint2, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint3", "", ExtraPoint3, "", "", "ExtraPoint") %>
						</div>
					</div>
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							가산점4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint4", "", ExtraPoint4, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint5", "", ExtraPoint5, "", "", "ExtraPoint") %>
						</div>

						<div class="col-md-1">
							가산점6
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("ExtraPoint6", "", ExtraPoint6, "", "", "ExtraPoint") %>
						</div>
					</div>
					-->

					<!-- 장학 
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							장학1 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship1", "", Scholarship1, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학2 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship2", "", Scholarship2, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학3 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship3", "", Scholarship3, "", "", "Scholarship") %>
						</div>

						<div class="col-md-1">
							장학4 
						</div>
						<div class="col-md-2">
							<% Call SubCodeSelectBox("Scholarship4", "", Scholarship4, "", "", "Scholarship") %>
						</div>
					</div>-->

					<!-- 필수서류 -->
					<!--<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							필수서류1 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence1", "필수서류 선택", DocumentaryEvidence1, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류2 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence2", "필수서류 선택", DocumentaryEvidence2, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류3 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence3", "필수서류 선택", DocumentaryEvidence3, "", "", "DocumentaryEvidence") %>
						</div>
					</div>
					<div class="row show-grid" style="text-align:left;">
						<div class="col-md-1">
							필수서류4 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence4", "필수서류 선택", DocumentaryEvidence4, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류5
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence5", "필수서류 선택", DocumentaryEvidence5, "", "", "DocumentaryEvidence") %>
						</div>

						<div class="col-md-1">
							필수서류6 
						</div>
						<div class="col-md-3">
							<% Call SubCodeSelectBox("DocumentaryEvidence6", "필수서류 선택", DocumentaryEvidence6, "", "", "DocumentaryEvidence") %>
						</div>
					</div>-->
				</form>
				
				<br>
				<div class="row show-grid grid_sub_button" >					
					<div class="col-md-12" >
						<span class="btnBasic btnTypeSave" id="RegBasicDataSet" style="width:80px;">저장</span>
						<span class="btnBasic btnTypeClose SelfCloseDIV" style="width:80px;">취소</span>
					</div>
				</div>
				</div>
			</div>
			<!-- 기본 데이터 설정 -->

		</div>		
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->