<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 3
Dim LeftMenuCode : LeftMenuCode = "CSAT"
Dim LeftMenuName : LeftMenuName = "Home / 평가기준관리 / 수능 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "수능 환산 기준 설정"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

'검색 조건
'Dim SearchMYear		: SearchMYear = fnR("SearchMYear", SessionMYear)
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", "")
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", "")
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", "")
Dim SearchDivision2	: SearchDivision2 = fnR("SearchDivision2", "")
Dim SearchDivision3	: SearchDivision3 = fnR("SearchDivision3", "")
Dim PageSize		: PageSize = getIntParameter(FnR("PageSize", 5), 5)
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/AppraisalList.asp"

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
end if

'쿼리
SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	b.IDX,a.MYear,a.SubjectCode"
SQL = SQL & vbCrLf & "	, a.Division0, a.Subject, a.Division1, a.Division2, a.Division3 "
SQL = SQL & vbCrLf & "	, b.Formula1, b.Formula2, b.Formula3, b.Formula4, b.Formula5 "
SQL = SQL & vbCrLf & "	, b.INPT_USID, b.INPT_DATE, b.INPT_ADDR, b.UPDT_USID, b.UPDT_DATE, b.UPDT_ADDR, b.InsertTime "

SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', a.Division0) AS Division0Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', a.Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division3', a.Division3) AS Division3Name "

SQL = SQL & vbCrLf & "FROM SubjectTable AS a "
SQL = SQL & vbCrLf & "	left outer join csat AS b"
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

			if (confirm("입력하신 내용을 저장 하시겠습니까?")) {
				/////////////////////////////////////////////////////////////////////////////////////////
				//선택된 체크박스의 값(모집코드)을 가져와서 SubjectCodeHidden변수에 넣어준 후 submit
				/////////////////////////////////////////////////////////////////////////////////////////
				var checkBoxArr = [];
				$("input[name=chk]:checked").each(function(i){
					checkBoxArr.push($(this).val());
				});
				$("#SubjectCodehidden").val(checkBoxArr);
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
});
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
					<!--</form>--
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
							<colgroup><col width="3%"><col width="5%"></col><col width="5%"></col><col width="10%"></col><col width="13%"></col>
								      <col width="13%"></col><col width="13%"></col><col width="13%"></col><col width="5%"></col><col width="7%"></col>
									  <col width="5%"></col><col width="7%"></col></col></colgroup>
							<thead>			                
								<tr>
									<!-- 체크박스 추가-평가비율 등록 시, 체크박스가 선택 된 모든 모집단위 없데이트 -->
									<th colspan="1" style="padding: 7px 0px 6px 0px; text-align: center;">
										<input type="checkbox" id="checkall"/>
									</th>
									<!--<th data-hide="phone">No.</th>-->
									<th data-hide="phone">년도</th>
									<!--<th data-hide="phone">모집코드</th>-->
									<th>모집시기</th>
									<th>학과명</th>
									<th>구분1</th>
									<!--<th>구분2</th>
									<th>구분3</th>-->
									<th>공식1</th>
									<th>공식2</th>
									<th>공식3</th>
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
										<input class="CheckboxCheck" type="Checkbox" name="chk" ID="Checkbox<%=i%>" value="<%=AryHash(i).Item("SubjectCode")%>">
									</td>
									<!--<td><%= intNUM %></td>-->
									<td><%= AryHash(i).Item("MYear") %></td>
									<!--<td><%= AryHash(i).Item("SubjectCode") %></td>-->
									<td><%= AryHash(i).Item("Division0Name") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<!--<td><%= AryHash(i).Item("Division2Name") %></td>
									<td><%= AryHash(i).Item("Division3Name") %></td>-->

									<td><%= AryHash(i).Item("Formula1") %></td>
									<td><%= AryHash(i).Item("Formula2") %></td>
									<td><%= AryHash(i).Item("Formula3") %></td>

									<td><%= AryHash(i).Item("INPT_USID") %></td>
									<td><%= Left(AryHash(i).Item("INPT_DATE"),10) %></td>
									<td><%= AryHash(i).Item("UPDT_USID") %></td>
									<td>
										<%= Left(AryHash(i).Item("UPDT_DATE"),10) %>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>"						ColumnName="IDX"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>"						ColumnName="MYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectCode")) %>"				ColumnName="SubjectCode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division0Name")) %>"				ColumnName="Division0Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectName")) %>"				ColumnName="SubjectName"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1Name")) %>"				ColumnName="Division1Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division2Name")) %>"				ColumnName="Division2Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division3Name")) %>"				ColumnName="Division3Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Formula1")) %>"					ColumnName="Formula1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Formula2")) %>"					ColumnName="Formula2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Formula3")) %>"					ColumnName="Formula3"></li>
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

			<!-- 변수정의 표 -->
			<div class="ibox-title">				
				<h5>변수 정의 </h5><br><br>
				<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
					<colgroup><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col></colgroup>
					<thead>	<tr><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th><th>L</th><th>M</th><th>U</th><th>V</th></tr></thead>
					<tbody>	<tr><td>국어</td><td>영어</td><td>수학</td><td>과학</td><td>한국사</td><td>사회</td><td>선택과목1</td><td>선택과목2</td><td>공식1의 값</td><td>공식2의 값</td></tr></tbody>
				</table>			
				<div class="ibox-tools">
					<a class="collapse-link">
						<!--<i class="fa fa-chevron-up"></i>-->
					</a>
				</div>
			</div>
			<!-- 변수정의 표 끝 -->

			<div class="pad_t10"></div>

			<!-- 상세보기 -->
			<div class="ibox-title">
				<h5>상세보기</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<!--<i class="fa fa-chevron-up"></i>-->
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/CSATProc.asp">
					<div style="display:none;">
						<input type="text" name="process" id="process" value="RegAppraisal">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<!-- 1개씩 insert or update 할 때 사용.
						<input type="text" name="ProcessType" id="ProcessType" <% If IsEmpty(IDX) Or IDX = "" Then %> value="Insert" <% Else %> value="Update" <% End IF%>>
						-->
						<input type="text" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="text" name="SubjectCodehidden" id="SubjectCodehidden" value="">
					</div>

					<!-- 선택한 학과 기본정보 -->
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							년도 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="MYear" class="form-control input-sm" value="" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							모집시기 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division0Name" class="form-control input-sm" value="" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							학과 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="SubjectName" class="form-control input-sm" value="" readonly>
						</div>
					</div>
					<div class="row show-grid">

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							전형 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division1Name" class="form-control input-sm" value="" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							구분2 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division2Name" class="form-control input-sm" value="" readonly>
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							구분3 
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Division3" class="form-control input-sm" value="" readonly>
						</div>
					</div>
				</div>

				<div class="ibox-content">
					<!-- 수능점수 환산 공식 -->
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
							수능점수 환산 공식 
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Formula1" class="form-control input-sm" value="">
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Formula2" class="form-control input-sm" value="">
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="Formula3" class="form-control input-sm" value="">
						</div>
					</div>

					<br>

					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">초기화</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>

				</form>
			</div>
			<!-- 상세보기 끝 -->

			<!-- 테이블 -->
			<div class="pad_t10"></div>
		</div>		
	</div>
</div>
<!-- #InClude Virtual = "/Common/Bottom.asp" -->