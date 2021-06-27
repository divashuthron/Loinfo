<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Subject"
Dim LeftMenuName : LeftMenuName = "Home / 모집단위관리 / 모집단위등록"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "모집단위관리"
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
Dim StrURL			: StrURL = "/SubjectList.asp"

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB


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

Set objDB	= Nothing
%>

<script type="text/javascript">
$(function() {
	// 신규
	$("#btnREG").click(function() {
		$.goURL("/SubjectView.asp");
	});
	
	// 상세보기
	$(document).on("click", "tr.viewDetail", function(){
		var IDX = $(this).attr("IDX")
		$.goURL("/SubjectView.asp?IDX="+ IDX);
	});
});
</script>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>학과 목록</h5>
				<div class="ibox-tools">
					<span class="btnBasic btnTypeAdd" id="btnREG">학과 등록</span>
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="table-responsive">
					<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
					<table id="dt_basic_Search" class="table table-striped table-bordered table-hover" width="100%">
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
								<th data-hide="phone">등록일</th>
							</tr>
						</thead>
						<tbody>
						<%
							'If Not IsNull(AryHash) Then
							Dim IntNum2

							If isArray(AryHash) Then
								intNUM = ubound(AryHash,1) + 1
								IntNum2 = Ubound(AryHash,1) + 1
								For i = 0 to ubound(AryHash,1)
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
							end if
						%>
						</tbody>
					</table>

				</div>
			</div>
			<!-- 테이블 -->
		</div>		
	</div>
</div>


<!-- #InClude Virtual = "/Common/Bottom.asp" -->
<script>
//=== 콤마 처리 ====================
$(document).ready(function(){
	var intNum = '<%=intNum2%>'
	for(var i = 0 ; intNum ; i++){
		$('.RF11'+i).html($.commaSplit($('.RF11'+i).html()));
	}
});
</script>