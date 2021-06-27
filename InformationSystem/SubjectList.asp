<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 2
Dim LeftMenuCode : LeftMenuCode = "Subject"
Dim LeftMenuName : LeftMenuName = "Home / 모집단위관리 / 모집단위등록"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "모집단위등록"
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
Dim StrURL			: StrURL = "/SubjectList.asp"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'서브쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And Division0 = ? "
	Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 50, SearchDivision)
end If

if not(IsE(SearchSubject)) And SearchSubject <> "All" then
	strWhere = strWhere & " And Subject = ? "
	Call objDB.sbSetArray("@Subject", adVarchar, adParamInput, 50, SearchSubject)
end If

if not(IsE(SearchDivision1)) And SearchDivision1 <> "All" then
	strWhere = strWhere & " And Division1 = ? "
	Call objDB.sbSetArray("@Division1", adVarchar, adParamInput, 50, SearchDivision1)
end If

if not(IsE(SearchDivision2)) And SearchDivision2 <> "All" then
	strWhere = strWhere & " And Division2 = ? "
	Call objDB.sbSetArray("@Division2", adVarchar, adParamInput, 50, SearchDivision2)
end If

if not(IsE(SearchDivision3)) And SearchDivision3 <> "All" then
	strWhere = strWhere & " And Division3 = ? "
	Call objDB.sbSetArray("@Division3", adVarchar, adParamInput, 50, SearchDivision3)
end if

'쿼리
SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "IDX,MYear,SubjectCode"

SQL = SQL & vbCrLf & " ,dbo.getSubCodeName('Division0', Division0) AS Division0Name "
SQL = SQL & vbCrLf & " ,dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & " ,dbo.getSubCodeName('Division1', Division1) AS Division1Name "
SQL = SQL & vbCrLf & " ,dbo.getSubCodeName('Division2', Division2) AS Division2Name "
SQL = SQL & vbCrLf & " ,dbo.getSubCodeName('Division3', Division3) AS Division3Name "

SQL = SQL & vbCrLf & ", Quorum,QuorumFix"
SQL = SQL & vbCrLf & ", RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11"
SQL = SQL & vbCrLf & ", INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime"
SQL = SQL & vbCrLf & "FROM SubjectTable " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 " & strWhere
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

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
					</form>
				</div>
			</div>
			<!-- 검색조건 끝-->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div style="float:right;">
					<a href="/Download/모집단위샘플.xlsx"><span class="btnBasic btnTypeComplete">모집단위 엑셀샘플</span></a>
					<span class="btnBasic btnTypeExcel" id="btnExcel" onClick="window.open('./SubjectUpload.asp','SubjectUpload','toolbar=no,menubar=no,scrollbars=no,resizable=no,width=1200 height=615'); return false;">엑셀로 등록</span>
					<span class="btnBasic btnTypeAdd" id="btnREG">학과 등록</span> &nbsp;&nbsp;&nbsp;&nbsp;
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
									<th data-hide="phone">모집코드</th>       
									<th>모집시기</th>                         
									<th>학과명</th>                           
									<th>구분1</th>                            
									<th>구분2</th>                            
									<th>구분3</th>                            
									<th data-hide="phone,tablet">입학정원</th>
									<th data-hide="phone,tablet">모집인원</th>
									<th data-hide="phone">등록금</th>         
									<th data-hide="phone">최종등록일</th>							
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									'For i = 0 to ubound(AryHash,1)
									For i = StartNum to EndNum
							%>
								<tr class="viewDetail" IDX="<%= AryHash(i).Item("IDX") %>">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("MYear") %></td>
									<td><%= AryHash(i).Item("SubjectCode") %></td>
									<td><%= AryHash(i).Item("Division0Name") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("Division2Name") %></td>
									<td><%= AryHash(i).Item("Division3Name") %></td>
									<td><%= AryHash(i).Item("QuorumFix") %></td>
									<td><%= AryHash(i).Item("Quorum") %></td>
									<td class="RF11<%=i%>"><%= AryHash(i).Item("RF11") %></td>
									<td><%= Left(AryHash(i).Item("InsertTime"),10) %></td>
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

	// 신규
	$("#btnREG").click(function() {
		$.goURL("/SubjectView.asp");
	});
	
	// 상세보기
	$(document).on("click", "tr.viewDetail", function(){
		var IDX = $(this).attr("IDX")
		$.goURL("/SubjectView.asp?IDX="+ IDX);
	});

	//=== 콤마 처리 ====================
	$(document).ready(function(){
		var StartNum = '<%=StartNum%>'
		var EndNum = '<%=EndNum%>'
		for(StartNum ; EndNum ; StartNum++){
			$('.RF11'+StartNum).html($.commaSplit($('.RF11'+StartNum).html()));
		}
	});
});

</script>