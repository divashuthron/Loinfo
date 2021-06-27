<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 4
Dim LeftMenuCode : LeftMenuCode = "Application"
Dim LeftMenuName : LeftMenuName = "Home / 입학원서관리 / 입학원서조회"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "입학원서조회"
Dim LogDivision	: LogDivision = "ApplicationList"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

'검색조건
Dim SearchMYear		: SearchMYear = fnR("SearchMYear", "")
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", "")
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", "")
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", "")
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/ApplicationList.asp"

'페이지설정(사이즈는 검색)
Dim PageSize		: PageSize = getIntParameter(FnR("PageSize", 5), 5)
Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'DBOpen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'검색조건 쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And Division0 = ? "
	Call objDB.sbSetArray("@Division", adVarchar, adParamInput, 50, SearchDivision)
end If

if not(IsE(SearchSubject)) And SearchSubject <> "All" then
	strWhere = strWhere & " And Subject = ? "
	Call objDB.sbSetArray("@Subject", adVarchar, adParamInput, 50, SearchSubject)
end if

if (not(IsE(SearchText))) then
	if SearchType = "1" then
		SearchPart = "StudentNumber"
	elseif SearchType = "2" then
		 SearchPart = "StudentNameKor"
	Else
		SearchType = "1"
		SearchPart = "StudentNumber"
	end if
	
	strWhere = strWhere & " And "& SearchPart &" like '%' + ? + '%' "
	Call objDB.sbSetArray("@SearchText", adVarchar, adParamInput, 255, SearchText)
end If

'쿼리
SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & " IDX, MYear, Division0, StudentNumber, StudentNameKor, StudentNameUsa, StudentNameChi "
SQL = SQL & vbCrLf & " , Citizen1, Citizen2, Sex, HighCode, HighSubject, HighGraduationYear, HighGraduationDivision "
SQL = SQL & vbCrLf & " , QualificationAreaCode, QualificationYear "
SQL = SQL & vbCrLf & " , Subject, Semester, UniversityName, AugScore, PerfectScore, Credit "
SQL = SQL & vbCrLf & " , Division1, HighDivision, RefundDivision, RefundAccountHolder "
SQL = SQL & vbCrLf & " , RefundBankCode, RefundAccount, Tel1, Tel2, Tel3, Email "
SQL = SQL & vbCrLf & " , Zipcode, Address1, Address2, StudentNameAgreement, CSATAgreement "
SQL = SQL & vbCrLf & " , StudentAgreement, StudentRecordAgreement, QualificationAgreement "
SQL = SQL & vbCrLf & " , INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR, InsertTime "
SQL = SQL & vbCrLf & " , ReceiptDate, ReceiptTime "

SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', Division0) AS DivisionName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', Division1) AS Division1Name "
'SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('HignSchoolDivision', HighDivision) AS HighDivisionName "

SQL = SQL & vbCrLf & "FROM ApplicationTable " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & strWhere
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'개인정보가 있는 입학원서는 조회도 기록함
strLogMSG = "입학원서관리  > " & SessionUserID  &"가/이 입학원서 리스트를 조회 했습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

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
//주소검색
function openDaumPostcode() {
	new daum.Postcode({
	  oncomplete: function(data) {
		var fullAddr = '';
		var extraAddr = '';

		if (data.userSelectedType === 'R') {
		  fullAddr = data.roadAddress;
		} else {
		  fullAddr = data.jibunAddress;
		}

		if(data.userSelectedType === 'R'){
		  if(data.bname !== ''){
			extraAddr += data.bname;
		  }
		  if(data.buildingName !== ''){
			extraAddr += (extraAddr !== '' ? ', ' + data.buildingName : data.buildingName);
		  }
		  fullAddr += (extraAddr !== '' ? ' ('+ extraAddr +')' : '');
		}

		document.getElementById('Zipcode').value = data.zonecode; //5자리 새우편번호 사용
		document.getElementById('Address1').value = fullAddr;
		$("#ZipcodeChk").text("");
		$("#Address1Chk").text("");
		document.getElementById('Address2').focus();
	  }
	}).open();
}

$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	//입학원서 기본값(년도, 모집시기(환경설정에 설정해놓은 값)) 넣기
	var SessionMYear = '<%=SessionMYear%>'
	var SessionDivision0 = '<%=SessionDivision%>'
	$("#InputForm [name=MYear]").val(SessionMYear).prop("selected", true).trigger("chosen:updated");
	$("#InputForm [name=Division0]").val(SessionDivision0).prop("selected", true).trigger("chosen:updated");

	// 목록 클릭 시 업데이트 프로세스로 변경
	$(document).delegate("tr.viewDetail_SetDate_2", "click", function() {
		$("#InputForm input[name='ProcessType']").val("Update");		
		$("#InputForm [name='StudentNumber']").prop("readonly", true).trigger("chosen:updated");
		//년도변경 안 되게 하려면 주석 풀기(목록,신규,저장)
		//$("#InputForm [name='MYear']").attr("disabled", true).trigger("chosen:updated");
	});

	// 신규
	$("#btnNew").click(function () {
		if (confirm("입력되어 있던 내용이 초기화 됩니다.\n신규로 입력 하시겠습니까?")) {
			$.FormReset($("#InputForm"));
			$("#InputForm [name='StudentNumber']").prop("readonly", false).trigger("chosen:updated");
			//$("#InputForm [name='MYear']").prop("disabled", false).trigger("chosen:updated");
			$("#InputForm [name=MYear]").val(SessionMYear).prop("selected", true).trigger("chosen:updated");
			$("#InputForm [name=Division0]").val(SessionDivision0).prop("selected", true).trigger("chosen:updated");
		}
	});

	// 저장
	$("#btnSave").click(function () { 
		if ($.setValidation($("#InputForm"))) {
			//$("#InputForm [name='MYear']").attr("disabled", false).trigger("chosen:updated");

			if (confirm("입력하신 내용을 저장 하시겠습니까?")) {
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

	// 시간표시
	$("#CheckTime").clockpicker({
		placement: "top", donetext: "Done"
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
								학생
							</div>
							<div class="col-md-2 col-xs-2 grid_sub_title">
								<select name="searchType" id="searchType" class="form-control input-sm">
									<option value="">구분</option>
									<option value="1" <%= setSelected(searchType, "1") %>>수험번호</option>
									<option value="2" <%= setSelected(searchType, "2") %>>이름(한글)</option>
								</select>
							</div>
							<div class="col-md-2 col-xs-2">
								<input type="text" name="searchText" id="searchText" value="<%= SearchText %>" class="form-control input-sm"/>
							</div>
						</div>
						<div class="pad_t10 pad_r10 text-right">							
							<span class="btnBasic btnSubmit">입학원서 조회</span>
						</div>
					<!--</form>-->
				</div>
			</div>
			<!-- 검색조건 끝 -->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div style="float:right;">
					<!-- 게시물 갯수 선택 -->
					<!--<a class="collapse-link">-->
					<!--<form id="PageSizeForm" method="get">-->
						<a href="/Download/입학원서샘플.xlsx"><span class="btnBasic btnTypeComplete">입학원서 엑셀샘플</span></a>
						<span class="btnBasic btnTypeExcel" id="btnExcel" onClick="window.open('./ApplicationUpload.asp','ApplicationUpload','toolbar=no,menubar=no,scrollbars=no,resizable=no,width=1200 height=615'); return false;">엑셀로 등록</span>
						<span class="btnBasic btnTypeAccept" onclick="alert('중간테이블과 연결되어 있지 않습니다.');return false;">입학원서 가져오기</span>
						<select name = "PageSize" style="margin-left:10px;" onChange="SearchForm.submit();">
							<option value="5" <% If PageSize = 5 then response.write "selected" end if%>>5개씩 보기</option>
							<option value="15" <% If PageSize = 15 then response.write "selected" end if%>>15개씩 보기</option>
							<option value="30" <% If PageSize = 30 then response.write "selected" end if%>>30개씩 보기</option>
							<option value="50" <% If PageSize = 50 then response.write "selected" end if%>>50개씩 보기</option>
							<option value="100" <% If PageSize = 100 then response.write "selected" end if%>>100개씩 보기</option>
							<option value="200" <% If PageSize = 200 then response.write "selected" end if%>>200개씩 보기</option>
						</select>
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
							<colgroup>
								<col width="3%"></col>
								<col width="4%"></col>
								<col width="5%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
								<col width="8%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
								<col width="5%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
								<col width="3%"></col>
							</colgroup>
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>
									<th data-hide="phone">년도</th>
									<th>모집시기</th>
									<th>학과명</th>
									<th>전형</th>
									<th data-hide="phone,tablet">수험번호</th>
									<th data-hide="phone">이름</th>
									<th data-hide="phone">Tel1</th>
									<th data-hide="phone">Tel2</th>
									<th data-hide="phone">Tel3</th>									
									<th data-hide="phone">생기부동의</th>
									<th data-hide="phone">검정동의</th>
									<th data-hide="phone">내용확인자</th>
									<th data-hide="phone">수험생동의</th>
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									'For i = 0 to ubound(AryHash,1)
									For i = StartNum to EndNum
										'  IDX, MYear, Division0, StudentNumber, StudentNameKor, StudentNameUsa, StudentNameChi 
										' , Citizen1, Citizen2, Sex, HighCode, HighSubject, HighGraduationYear, HighGraduationDivision
										' , QualificationAreaCode, QualificationYear 
										' , Subject, Semester, UniversityName, AugScore, PerfectScore, Credit 
										' , Division1, HighDivision, RefundDivision, RefundAccountHolder 
										' , RefundBankCode, RefundAccount, Tel1, Tel2, Tel3, Email 
										' , Zipcode, Address1, Address2, StudentNameAgreement 
										' , StudentAgreement, StudentRecordAgreement, QualificationAgreement 
										' , INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR, InsertTime 
										' , DivisionName, SubjectName, Division1Name, HighDivisionName
							%>
								<tr class="viewDetail_SetDate_2">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("MYear") %></td>
									<td><%= AryHash(i).Item("DivisionName") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("StudentNumber") %></td>
									<td><%= AryHash(i).Item("StudentNameKor") %></td>
									<td><%= AryHash(i).Item("Tel1") %></td>
									<td><%= AryHash(i).Item("Tel2") %></td>
									<td><%= AryHash(i).Item("Tel3") %></td>
									<td><% if AryHash(i).Item("StudentRecordAgreement") = "1" Then%> Y <%Else%> N <%End If %></td>
									<td><% if AryHash(i).Item("QualificationAgreement") = "1" Then%> Y <%Else%> N <%End If %></td>
									<td><% If AryHash(i).Item("StudentNameAgreement") = "1" Then%> Y <%Else%> N <%End If %></td>
									<td>
										<% if AryHash(i).Item("StudentAgreement") = "1" Then%> Y <%Else%> N <%End If %>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>"								ColumnName="IDX"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>"								ColumnName="MYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>"								ColumnName="MYearHidden"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division0")) %>"							ColumnName="Division0"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNumber")) %>"						ColumnName="StudentNumber"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNameKor")) %>"					ColumnName="StudentNameKor"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNameUsa")) %>"					ColumnName="StudentNameUsa"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNameChi")) %>"					ColumnName="StudentNameChi"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Citizen1")) %>"							ColumnName="Citizen1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Citizen2")) %>"							ColumnName="Citizen2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Sex")) %>"								ColumnName="Sex"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighCode")) %>"							ColumnName="HighCode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighCode")) %>"							ColumnName="HighCodeTemp"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighSubject")) %>"						ColumnName="HighSubject"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighGraduationYear")) %>"				ColumnName="HighGraduationYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighGraduationDivision")) %>"			ColumnName="HighGraduationDivision"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationAreaCode")) %>"				ColumnName="QualificationAreaCode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationAreaCode")) %>"				ColumnName="QualificationAreaCodeTemp"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationYear")) %>"					ColumnName="QualificationYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Subject")) %>"							ColumnName="Subject"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Subject")) %>"							ColumnName="SubjectTemp"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Semester")) %>"							ColumnName="Semester"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UniversityName")) %>"					ColumnName="UniversityName"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("AugScore")) %>"							ColumnName="AugScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("PerfectScore")) %>"						ColumnName="PerfectScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Credit")) %>"							ColumnName="Credit"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1")) %>"							ColumnName="Division1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighDivision")) %>"						ColumnName="HighSchoolDivision"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("RefundDivision")) %>"					ColumnName="RefundDivision"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("RefundAccountHolder")) %>"				ColumnName="RefundAccountHolder"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("RefundBankCode")) %>"					ColumnName="RefundBankCode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("RefundAccount")) %>"						ColumnName="RefundAccount"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel1")) %>"								ColumnName="Tel1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel2")) %>"								ColumnName="Tel2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel3")) %>"								ColumnName="Tel3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Email")) %>"								ColumnName="Email"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Zipcode")) %>"							ColumnName="Zipcode"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Address1")) %>"							ColumnName="Address1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Address2")) %>"							ColumnName="Address2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNameAgreement")) %>"				ColumnName="StudentNameAgreement"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentAgreement")) %>"					ColumnName="StudentAgreement"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordAgreement")) %>"			ColumnName="StudentRecordAgreement"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationAgreement")) %>"			ColumnName="QualificationAgreement"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("CSATAgreement")) %>"						ColumnName="CSATAgreement"></li>
											<li Columnvalue="<% if Not isnull(AryHash(i).Item("ReceiptDate")) Then %><%= Trim(FormatDateTime(AryHash(i).Item("ReceiptDate"),2)) %><%End If%>"		ColumnName="ReceiptDate"></li>
											<li Columnvalue="<% if Not isnull(AryHash(i).Item("ReceiptTime")) Then %><%= Trim(FormatDateTime(AryHash(i).Item("ReceiptTime"),4)) %><%End If%>"		ColumnName="CheckTime"></li>
										</div>
									</td>
								</tr>
							<%
										intNUM = intNUM - 1
									Next
								Else
							%>
								<tr>
									<td colspan="14" style="height:50px; vertical-align: middle;">검색된 자료가 없습니다.</td>
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
				<div class="ibox-tools">
					<a class="collapse-link">
						<!--<i class="fa fa-chevron-up"></i>-->
					</a>
				</div>
			</div>
			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/ApplicationProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegApplication">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="hidden" name="MYearHidden" id="MYearHidden" value="">
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							년도 *
						</div>
						<div class="col-md-1 col-xs-7">
							<% Call SubCodeSelectBox("MYear", "년도 선택", "", "년도를 선택하여주세요.", "", "MYear") %>
						</div>		
						<div class="col-md-1 col-xs-2 grid_sub_title">
							모집시기 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division0", "모집시기 선택", "", "모집시기를 선택하여주세요.", "", "Division0") %>
						</div>	
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							수험번호 *
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="text" name="StudentNumber" class="form-control input-sm" maxlength="7" alert="수험번호를 입력하세요.">
						</div>						
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							성명(한글) *
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="text" id="StudentNameKor" name="StudentNameKor" class="form-control input-sm" maxlength="15" alert="성명(한글)을 입력하세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							성명(영문)
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="text" name="StudentNameUsa" class="form-control input-sm" maxlength="15">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							성명(한문)
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="text" name="StudentNameChi" class="form-control input-sm" maxlength="15">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							주민번호
						</div>
						<div class="col-md-1 col-xs-7">
							<input type="text" name="Citizen1" class="form-control input-sm KeyTypeNUM" maxlength="6">
						</div>
						-
						<div class="col-md-1 col-xs-7">
							<input type="text" name="Citizen2" class="form-control input-sm KeyTypeNUM" maxlength="7">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학력(고교)
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<% Call SubCodeSelectBox("HighGraduationYear", "년도 선택", "", "", "", "HighGraduationYear") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							고교선택
						</div>
						<div class="col-md-2 col-xs-7">
							<%' Call SubCodeSelectBox("HighCode", "고교 선택", "", "", "", "HighCode") %>
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<% Call SubCodeSelectBox("HighGraduationDivision", "졸업여부", "", "", "", "HighGraduationDivision") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							출신고교코드코드
						</div>
						<div class="col-md-1 col-xs-7">
							<input type="text" name="HighCodeTemp" value="<%=HighCodeTemp%>" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							검정고시
						</div>
						<div class="col-md-1 col-xs-7">
							<% Call SubCodeSelectBox("QualificationYear", "년도 선택", "", "", "", "QualificationYear") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							합격지구
						</div>
						<div class="col-md-2 col-xs-3">
							<% Call SubCodeSelectBox("QualificationAreaCode", "합격지구 선택", "", "", "", "QualificationAreaCode") %>
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title"></div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							합격지구코드
						</div>
						<div class="col-md-1 col-xs-7 ">
							<input type="text" name="QualificationAreaCodeTemp" value="<%=QualificationAreaCodeTemp%>" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학교생활기록부 온라인 동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="StudentRecordAgreement" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title">
							검정고시 합격성적 온라인 제공동의
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="QualificationAgreement" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수능 동의 
						</div>
						<div class="col-md-1 col-xs-3 grid_sub_title">
							<input type="checkbox" name="CSATAgreement" class="form-control input-sm" value="1">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							지망학과(전공) *
						</div>
						<div class="col-md-2 col-xs-7 grid_sub_title">
							<% Call SubCodeSelectBox("Subject", "학과 선택", "", "지망학과를 선택해주세요.", "", "Subject") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							학과코드
						</div>
						<div class="col-md-1 col-xs-7">
							<input type="text" name="SubjectTemp" id="<%=SubjectTemp%>" class="form-control input-sm" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학교생활기록부 반영 학기 선택(택1)
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							1학년1학기
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="Semester" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							1학년2학기
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="Semester" class="form-control input-sm" value="2">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							2학년1학기
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="Semester" class="form-control input-sm" value="3">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							2학년2학기
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="Semester" class="form-control input-sm" value="4">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							3학년1학기
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="Semester" class="form-control input-sm" value="5">
						</div>						
					</div>
					<div class="row show-grid">					
						<div class="col-md-2 col-xs-2 grid_sub_title">
							출신대학(전문대학)명
						</div>
						<div class="col-md-2 col-xs-8 grid_sub_title">
							<% 'Call SubCodeSelectBox("UniversityCode", "대학 선택", "", "", "", "UniversityCode") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							평균점수 / 만점
						</div>
						<div class="col-md-1 col-xs-8">
							<input type="text" name="AugScore" class="form-control input-sm KeyTypeNUM" maxlength="4">
						</div>
						/
						<div class="col-md-1 col-xs-8 grid_sub_title">
							<% Call SubCodeSelectBox("PerfectScore", "만점 선택", "", "", "", "PerfectScore") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							이수학점
						</div>
						<div class="col-md-1 col-xs-8">
							<input type="text" name="Credit" class="form-control input-sm KeyTypeNUM" maxlength="3">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							전형구분 *
						</div>
						<div class="col-md-2 col-xs-3 grid_sub_title">
							<% Call SubCodeSelectBox("Division1", "전형 선택", "", "모집전형을 선택해주세요.", "", "Division1") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							고교(과정)구분
						</div>
						<div class="col-md-2 col-xs-3">
							<% Call SubCodeSelectBox("HighSchoolDivision", "고교(과정) 선택", "", "", "", "HighSchoolDivision") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							전형료 비례환불 선택방법
						</div>
						<div class="col-md-2 col-xs-3">
							<% Call SubCodeSelectBox("RefundDivision", "방법 선택", "", "", "", "RefundDivision") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							예금주
						</div>
						<div class="col-md-1 col-xs-3">
							<input type="text" name="RefundAccountHolder" class="form-control input-sm" maxlength="15">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							환불은행
						</div>
						<div class="col-md-1 col-xs-3">
							<% Call SubCodeSelectBox("RefundBankCode", "은행선택", "", "", "", "RefundBankCode") %>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							환불계좌
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="RefundAccount" class="form-control input-sm KeyTypeNUM" maxlength="13">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							자택전화
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel1" class="form-control input-sm KeyTypeNUM" maxlength="25" >
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							휴대폰번호
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel2" class="form-control input-sm KeyTypeNUM" maxlength="25" >
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title2">
							이메일
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Email" class="form-control input-sm" onblur="javascript:ck_email(this);" maxlength="50" >
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							보호자 휴대폰번호
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel3" class="form-control input-sm KeyTypeNUM" maxlength="25" >
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							우편번호
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" id="Zipcode" name="Zipcode" class="form-control input-sm" maxlength="6" readonly>
						</div>
						<div class="col-md-2 col-xs-1">
							<span class="btnBasic" onclick="openDaumPostcode();">주소검색</span>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							기본주소
						</div>
						<div class="col-md-3 col-xs-7">
							<input type="text" name="Address1" id="Address1" class="form-control input-sm" readonly>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							상세주소
						</div>
						<div class="col-md-3 col-xs-7">
							<input type="text" name="Address2" id="Address2" class="form-control input-sm" > 
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							지원자 확인 
						</div>
						<div class="col-md-2 col-xs-3 grid_sub_title">
							<input type="text" name="StudentNameAgreement" class="form-control input-sm" maxlength="15" placeholder="성명을 입력해주세요." >
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수험생 확인동의 
						</div>
						<div class="col-md-1 col-xs-3 grid_sub_title">
							<input type="checkbox" name="StudentAgreement" class="form-control input-sm" value="1">
						</div>						
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							원서접수일자
						</div>
						<div class="col-md-2 col-xs-3 grid_sub_title">
							<div class="input-group viewCalendarBtn" Obj="ReceiptDate" >
								<input type="text" name="ReceiptDate" id="ReceiptDate" class="form-control input-sm" maxlength="10" style="background-color:white;" readonly>
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							원서접수시간
						</div>
						<div class="col-md-2 col-xs-3">
							<div class="input-group">
								<input type="text" name="CheckTime" id="CheckTime" class="form-control input-sm" maxlength="5" data-autoclose="true" style="background-color:white;" readonly>
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
					</div>
					( * 는 필수 입력값입니다.)
					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">신 규</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>

				</form>
			</div>
			<!-- 상세보기 끝 -->

		</div>		
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->