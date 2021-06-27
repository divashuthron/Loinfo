<%@  codepage="65001" language="VBScript" %>
<%
'Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Home"
Dim LeftMenuName : LeftMenuName = "Home"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "Home"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
'DB
'RW(SessionUserName)
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere
Dim i, strMSG, intNUM, intNUM2, strTEMP, strRESULT, IDX

Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageSize		: PageSize	= 5
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'전체 히스토리 쿼리
SQL = ""
SQL = SQL & vbCrLf & "select "
SQL = SQL & vbCrLf & "		IDX, MYear, Division, ActivityContent, RegID, RegDate "
SQL = SQL & vbCrLf & "from ActivityHistory " 
SQL = SQL & vbCrLf & "where MYear = " & SessionMYear
SQL = SQL & vbCrLf & "order by IDX desc  "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

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

'공지사항 쿼리
SQL = ""
SQL = SQL & vbCrLf & "select "
SQL = SQL & vbCrLf & "		IDX, MYear, Division, Title, content1, file1, INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR"
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Department', Division) AS DivisionName "
SQL = SQL & vbCrLf & "from NoticeTable " 
SQL = SQL & vbCrLf & "where MYear = " & SessionMYear
SQL = SQL & vbCrLf & "order by IDX desc  "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB = Nothing
%>
<script type="text/javascript">
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// 신규
	$("#NoticeSet").click(function() {
		$.goURL("/NoticeView.asp");
	});

	// 상세보기
	$(document).on("click", "tr.viewDetail", function(){
		var IDX = $(this).attr("IDX")
		$.goURL("/NoticeView.asp?IDX="+ IDX);
	});
});
</script>
<form id="SearchForm" method="get">
	<input type="hidden" name="Page" value="<%= PageNum %>">
</form>
<div class="row">
	<!--상좌 달성률 -->
	<div class="col-lg-6">
		<div class="ibox float-e-margins">
			<div class="ibox-title">
				<h5>달성률</h5>

				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>
			<div class="ibox-content">
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
				<div >&nbsp;</div>
			</div>
		</div>
	</div>

	<!--상우 전체 히스토리 -->
	<div class="col-lg-6">
		<div class="ibox float-e-margins">
			<div class="ibox-title">
				<h5>전체 히스토리</h5>

				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>
			<div class="ibox-content">
				<div class="pad_5">
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
							If isArray(AryHash) Then
								'For i = 0 to ubound(AryHash,1)
								'For i = 0 To UBound(aryList, 2)
								For i = StartNum to EndNum
									'IDX, MYear, Division, ActivityContent, RegID, RegDate
						%>
							<tr IDX="<%= AryHash(i).Item("IDX") %>">
								<td><%= intNUM %></td>
								<td><%= AryHash(i).Item("MYear") %></td>
								<td><%= AryHash(i).Item("ActivityContent") %></td>
								<td><%= AryHash(i).Item("RegID") %></td>
								<td><%= AryHash(i).Item("RegDate") %></td>
							</tr>
						<%
								intNUM = intNUM - 1
								'// 히스토리 읽음 기록
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
		</div>
	</div>

	<!--하 공지사항 -->
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			<div class="ibox-title">
				<h5>공지사항</h5>
				
				<div style="float:right;">
					<a>
						<span class="btnBasic btnTypeAdd" id="NoticeSet" style="margin-right:5px;">공지사항 등록</span>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_horizontal" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<colgroup>
								<col width="7%"></col>
								<col width="7%"></col>
								<col width="7%"></col>
								<col width=""></col>
								<col width="7%"></col>
								<col width="15%"></col>
								<col width="20%"></col>
							</colgroup>
							<thead>			                
								<tr>
									<th>No.</th>            
									<th>년도</th> 
									<th>구분</th>  
									<th>공지사항명</th>        
									<th>첨부파일여부</th> 
									<th>최초입력자</th>                            
									<th>최초입력시간</th>                            						
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash2) Then
									intNUM2 = ubound(AryHash2,1) + 1
									For i = 0 to ubound(AryHash2,1)
									'For i = StartNum to EndNum
										'IDX, MYear, Division, Title, content1, file1, INPT_USID, INPT_DATE, INPT_ADDR, UPDT_USID, UPDT_DATE, UPDT_ADDR
							%>
								<tr class="viewDetail" IDX="<%= AryHash2(i).Item("IDX") %>">
									<td><%= intNUM2 %></td>
									<td><%= AryHash2(i).Item("MYear") %></td>
									<td><%= AryHash2(i).Item("DivisionName") %></td>
									<td><%= AryHash2(i).Item("Title") %></td>
									<td><% If Not(isnull(AryHash2(i).Item("file1"))) Then %><i class="fa fa-paperclip"></i> <% End If%></td>
									<td><%= AryHash2(i).Item("INPT_USID") %></td>
									<td><%= AryHash2(i).Item("INPT_DATE") %></td>
								</tr>
							<%
										intNUM2 = intNUM2 - 1
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
				</div>
				
			</div>
		</div>
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->