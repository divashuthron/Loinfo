<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 7
Dim LeftMenuCode : LeftMenuCode = "Demands"
Dim LeftMenuName : LeftMenuName = "Home / 합격자발표관리 / 유의사항 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "유의사항 조회"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 4
Dim PageNum			: PageNum	= fnR("page", 1)
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/DemandsList.asp"

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	IDX,MYear,Division0,dbo.getSubCodeName('Division0', Division0) AS DivisionName,Title,State "
SQL = SQL & vbCrLf & "	, (CASE  State "
SQL = SQL & vbCrLf & "			WHEN '1' THEN '사용' "
SQL = SQL & vbCrLf & "			WHEN '0' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName ,option1,option2,option3,option4,option5,option6,option7,option8,option9,option10 "
SQL = SQL & vbCrLf & "	,content1,content2,INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime "
SQL = SQL & vbCrLf & "FROM DemandsTable "
SQL = SQL & vbCrLf & "WHERE 1 = 1 " 
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB	= Nothing
%>

<script type="text/javascript">
$(function() {
	// 신규
	$("#btnREG").click(function() {
		$.goURL("/DemandsView.asp");
	});
	
	// 상세보기
	$(document).on("click", "tr.viewDetail", function(){
		var IDX = $(this).attr("IDX")
		$.goURL("/DemandsView.asp?IDX="+ IDX);
	});
});
</script>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>유의사항목록</h5>				
				<div style="float:right;">
					<span class="btnBasic btnTypeAdd" id="btnREG">유의사항 등록</span>
				</div>
			</div>
			

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_Search" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>            
									<th data-hide="phone">년도</th>    
									<th>모집시기</th>
									<th data-hide="phone">유의사항명</th>   
									<th>상태</th>                           
									<th>최초입력자</th>                            
									<th>최초입력시간</th>                            
									<th>수정자</th>                            
									<th data-hide="phone,tablet">수정입력시간</th>							
								</tr>
							</thead>
							<tbody>
							<%
								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									intNUM = ubound(AryHash,1) + 1
									For i = 0 to ubound(AryHash,1)
									'For i = StartNum to EndNum
							%>
								<tr class="viewDetail" IDX="<%= AryHash(i).Item("IDX") %>">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("MYear") %></td>
									<td><%= AryHash(i).Item("DivisionName") %></td>
									<td><%= AryHash(i).Item("Title") %></td>									
									<td><%= AryHash(i).Item("StateName") %></td>
									<td><%= AryHash(i).Item("INPT_USID") %></td>
									<td><%= AryHash(i).Item("INPT_DATE") %></td>
									<td><%= AryHash(i).Item("UPDT_USID") %></td>
									<td><%= AryHash(i).Item("UPDT_DATE") %></td>
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