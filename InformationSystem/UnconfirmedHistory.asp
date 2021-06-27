<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "UnconfirmedHistory"
Dim LeftMenuName : LeftMenuName = "Home / 미확인 히스토리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "미확인 히스토리"
Dim LogDivision	: LogDivision = "UnconfirmedHistory"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT, IDX

Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'검색 조건
'Dim SearchMYear		: SearchMYear = fnR("SearchMYear", SessionMYear)
Dim SearchRegId				: SearchRegId = fnR("RegId", "")
Dim SearchActivityContent	: SearchActivityContent = fnR("ActivityContent", "")

Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/UnconfirmedHistory.asp"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'서브쿼리
if not(IsE(SearchRegId)) And SearchRegId <> "All" then
	strWhere = strWhere & " And a.RegId like '%' + ? + '%' "
	Call objDB.sbSetArray("@RegId", adVarchar, adParamInput, 50, SearchRegId)
end If

if not(IsE(SearchActivityContent)) And SearchActivityContent <> "All" then
	strWhere = strWhere & " And a.ActivityContent like '%' + ? + '%' "
	Call objDB.sbSetArray("@ActivityContent", adVarchar, adParamInput, 50, SearchActivityContent)
end If

'쿼리
SQL = ""
SQL = SQL & vbCrLf & "select "
SQL = SQL & vbCrLf & "		a.IDX, a.MYear, a.Division, a.ActivityContent, a.RegID, a.RegDate "
SQL = SQL & vbCrLf & "from ActivityHistory a left outer join AlarmHistory b  " 
SQL = SQL & vbCrLf & "on a.IDX = b.HistoryIDX  "
SQL = SQL & vbCrLf & "where a.MYear = " & SessionMYear
SQL = SQL & vbCrLf & "and b.IDX is null  " & strWhere
SQL = SQL & vbCrLf & "order by a.IDX desc  "

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
								입력자	
							</div>
							<div class="col-md-2 col-xs-2">
								<input type="text" name="SearchRegId" class="form-control input-sm">
							</div>
							<div class="col-md-1 col-xs-1 grid_sub_title2">
								내용
							</div>
							<div class="col-md-4 col-xs-2">
								<input type="text" name="SearchActivityContent" class="form-control input-sm">
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
				<div class="ibox-tools">
					<a class="collapse-link">
						<!--<i class="fa fa-chevron-up"></i>-->
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
									<th data-hide="phone">내용</th>       
									<th>입력자</th>                         
									<th>입력시간</th>                           							
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(aryList) Then
									'For i = 0 to ubound(AryHash,1)
									'For i = 0 To UBound(aryList, 2)
									For i = StartNum to EndNum
										'IDX, MYear, Division, ActivityContent, RegID, RegDate
							%>
								<tr class="viewDetail" IDX="<%= aryList(0, i) %>">
									<td><%= intNUM %></td>
									<td><%= aryList(1, i) %></td>
									<td><%= aryList(3, i) %></td>
									<td><%= aryList(4, i) %></td>
									<td><%= aryList(5, i) %></td>
								</tr>
							<%
									intNUM = intNUM - 1
									'// 히스토리 읽음 기록
									Call AlarmHistory(aryList(0, i), aryList(1, i), SessionUserID)
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
});
</script>