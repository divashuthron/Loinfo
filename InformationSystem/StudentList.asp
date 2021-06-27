<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Student"
Dim LeftMenuName : LeftMenuName = "Home / 입학원서관리 / 입학원서조회"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "입학원서조회"
Dim LogDivision	: LogDivision = "StudentList"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

Dim SearchMYear		: SearchMYear = fnR("SearchMYear", SessionMYear)
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", SessionDivision)
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", SessionSubject)
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", SessionDivision1)

Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/StudentList.asp"

'DBOpen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'검색조건 쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And Division = ? "
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
		 SearchPart = "StudentName"
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
SQL = SQL & vbCrLf & "	IDX, MYear, Division, Subject, Division1, Division2 "
SQL = SQL & vbCrLf & "	, StudentNumber, StudentName, InterviewNumber, TSize, HighSchool "
SQL = SQL & vbCrLf & "	, Birthday, Sex, Tel1, Tel2, Tel3, EnglishPoint "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division', Division) AS DivisionName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', Division1) AS Division1Name "
'SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, '' AS Division2Name "
SQL = SQL & vbCrLf & "	, State, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID "
SQL = SQL & vbCrLf & "FROM StudentTable AS A " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & strWhere
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'개인정보가 있는 리스트는 조회 히스토리(기록)
strLogMSG = "입학원서관리  > 입학원서 리스트가 조회 되었습니다."
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
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// 목록 클릭 시 업데이트 프로세스로 변경
	$(document).delegate("tr.viewDetail_SetDate_2", "click", function() {
		$("#InputForm input[name='ProcessType']").val("Update");
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
								<% Call SubCodeSelectBox("SearchDivision", "모집시기선택", SearchDivision, "", "All", "Division") %>
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
									<option value="2" <%= setSelected(searchType, "2") %>>이름</option>
								</select>
							</div>
							<div class="col-md-2 col-xs-2">
								<input type="text" name="searchText" id="searchText" value="<%= SearchText %>" class="form-control input-sm"/>
							</div>
						</div>
						<div class="pad_t10 pad_r10 text-right">
							<span class="btnBasic btnSubmit">지원자 조회</span>
						</div>
					</form>
				</div>
			</div>
			<!-- 검색조건 -->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>
									<th data-hide="phone">년도</th>
									<th>모집시기</th>
									<th>학과명</th>
									<th>전형</th>
									<th data-hide="phone,tablet">수험번호</th>
									<th data-hide="phone">면접번호</th>
									<th data-hide="phone,tablet">이름</th>
									<th data-hide="phone">성별</th>
									<th data-hide="phone">TSize</th>
									<th data-hide="phone">고등학교</th>
									<th data-hide="phone">생년월일</th>
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									'For i = 0 to ubound(AryHash,1)
									For i = StartNum to EndNum
										'IDX, MYear, Division, Subject, Division1, Division2
										'StudentNumber, StudentName, InterviewNumber, TSize, HighSchool
										'Birthday, Sex, Tel1, Tel2, Tel3
										'DivisionName, SubjectName, Division1Name, Division2Name
										'State, StateName, RegDate, RegID, EditDate, EditID
							%>
								<tr class="viewDetail_SetDate_2" StudentNumber="<%= AryHash(i).Item("StudentNumber") %>">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("MYear") %></td>
									<td><%= AryHash(i).Item("DivisionName") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("StudentNumber") %></td>
									<td><%= AryHash(i).Item("InterviewNumber") %></td>
									<td><%= AryHash(i).Item("StudentName") %></td>
									<td><%= AryHash(i).Item("Sex") %></td>
									<td><%= AryHash(i).Item("TSize") %></td>
									<td><%= AryHash(i).Item("HighSchool") %></td>
									<td>
										<%= AryHash(i).Item("Birthday") %>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>"				ColumnName="IDX"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>"				ColumnName="MYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division")) %>"			ColumnName="Division"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Subject")) %>"			ColumnName="Subject"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1")) %>"			ColumnName="Division1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division2")) %>"			ColumnName="Division2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentNumber")) %>"		ColumnName="StudentNumber"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentName")) %>"		ColumnName="StudentName"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("InterviewNumber")) %>"	ColumnName="InterviewNumber"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("TSize")) %>"				ColumnName="TSize"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("HighSchool")) %>"		ColumnName="HighSchool"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Birthday")) %>"			ColumnName="Birthday"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Sex")) %>"				ColumnName="Sex"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel1")) %>"				ColumnName="Tel1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel2")) %>"				ColumnName="Tel2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Tel3")) %>"				ColumnName="Tel3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("EnglishPoint")) %>"		ColumnName="EnglishPoint"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("State")) %>"				ColumnName="State"></li>
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
			<!-- 테이블 -->

			<div class="pad_t10"></div>

			<!-- 상세보기 -->
			<div class="ibox-title">
				<h5>상세정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/StudentProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegStudnet">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							년도 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("MYear", "년도선택", "", "년도를 선택해 주세요.", "", "MYear") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							모집시기 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division", "모집시기선택", "", "모집시기를 선택해 주세요.", "", "Division") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학과 *
						</div>
						<div class="col-md-3 col-xs-7">
							<% Call SubCodeSelectBox("Subject", "학과명선택", "", "학과명을 선택해 주세요.", "", "Subject") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							전형 *
						</div>
						<div class="col-md-3 col-xs-7">
							<% Call SubCodeSelectBox("Division1", "전형선택", "", "전형을 선택해 주세요.", "", "Division1") %>
						</div>
					</div>


					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							수험번호 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="StudentNumber" class="form-control input-sm" maxlength="50" <% If Not(IsE(StudentNumber)) Then Response.write "readonly" End If %>>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title2">
							이름 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="StudentName" class="form-control input-sm" maxlength="25" alert="이름을 입력해 주세요.">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접번호
						</div>
						<div class="col-md-3 col-xs-7">
							<input type="text" name="InterviewNumber" class="form-control input-sm" maxlength="50" alert="">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							TSize *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("TSize", "TSize선택", "", "TSize를 선택해 주세요.", "", "T-Size") %>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							출신고등학교
						</div>
						<div class="col-md-6 col-xs-8">
							<input type="text" name="HighSchool" class="form-control input-sm" maxlength="50" alert="출신고등학교를 입력해 주세요.">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							영어점수 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="EnglishPoint" class="form-control input-sm" maxlength="25" alert="영어점수를 입력해 주세요.">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							생년월일 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Birthday" class="form-control input-sm KeyTypeNUM" maxlength="8" alert="생년월일을 입력해 주세요.">
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title2">
							성별 *
						</div>
						<div class="col-md-2 col-xs-3">
							<% Call SubCodeSelectBox("Sex", "성별선택", "", "성별을 선택해 주세요.", "", "Sex") %>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							전화번호 1 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel1" class="form-control input-sm" maxlength="25" alert="전화번호를 입력해 주세요.">
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title2">
							전화번호 2 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel2" class="form-control input-sm" maxlength="25" alert="전화번호를 입력해 주세요.">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							전화번호 3 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Tel3" class="form-control input-sm" maxlength="25" alert="전화번호를 입력해 주세요.">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							상태 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="State" class="form-control input-sm" alert="상태를 선택하세요.">
								<option value="">상태선택</option>
								<option value="Y" <%= setSelected(State, "Y") %>>사용</option>
								<option value="N" <%= setSelected(State, "N") %>>미사용</option>
							</select>
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
			<!-- 상세보기 -->

		</div>		
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->