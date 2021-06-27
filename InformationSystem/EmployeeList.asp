<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 9
Dim LeftMenuCode : LeftMenuCode = "Employee"
Dim LeftMenuName : LeftMenuName = "Home / 환경설정 / 사용자관리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "사용자관리"
Dim LogDivision	: LogDivision = "EmployeeList"
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
Dim StrURL			: StrURL = "/EmployeeList.asp?type=Employee"

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

if not(IsE(SearchState)) then
	strWhere = strWhere & " And State = ? "
	Call objDB.sbSetArray("@State", adVarchar, adParamInput, 1, SearchState)
end if

if (not(IsE(SearchText))) then
	if SearchType = "1" then
		SearchPart = "EmpID"
	elseif SearchType = "2" then
		 SearchPart = "EmpName"
	end if
	
	strWhere = strWhere & " And "& SearchPart &" like '%' + ? + '%' "
	Call objDB.sbSetArray("@SearchText", adVarchar, adParamInput, 255, SearchText)
end If

' IDX, EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, PhoneNumber
' Email, JoinDate, OutDate, EmpInfo, State
' StateName, RegDate, RegID, EditDate, EditID, ClientLevelName

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	IDX, EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, PhoneNumber "
SQL = SQL & vbCrLf & "	, Email, JoinDate, OutDate, EmpInfo, State "
SQL = SQL & vbCrLf & "	, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID "
'SQL = SQL & vbCrLf & "	, (Select SubCodeName from CodeSub where MasterCode = 'SchoolCode' and SubCode = A.ClientCode) AS ClientCodeName "
'SQL = SQL & vbCrLf & "	, (Select SubCodeName from CodeSub where MasterCode = 'UserGrade' and SubCode = A.ClientLevel) AS ClientLevelName "
'SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('SchoolCode', A.ClientCode) AS ClientCodeName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('UserGrade', A.ClientLevel) AS ClientLevelName "
SQL = SQL & vbCrLf & "FROM Employee AS A " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 " & strWhere
SQL = SQL & vbCrLf & "ORDER BY IDX DESC;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

strLogMSG = "사용자관리  > " & SessionUserID & "이/가 사용자 관리 정보 리스트를 조회 했습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

Set objDB	= Nothing
%>

<script type="text/javascript">
	$(function() {
		// 신규
		$("#btnREG").click(function () {
			$.goURL("/EmployeeView.asp");
		});

		// 상세보기
		$(document).on("click", "tr.viewDetail", function(){
		//$(document).delegate("tr.viewDetail", "click", function() {
			var EmpID = $(this).attr("EmpID")
			$.goURL("/EmployeeView.asp?EmpID="+ EmpID);
		});
	})
</script>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>사용자 목록</h5>
				<div class="ibox-tools">
					<span class="btnBasic btnTypeAdd" id="btnREG">사용자 등록</span>
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
								<!--<th data-hide="phone,tablet">소속</th>-->
								<th>아이디</th>
								<th data-hide="phone,tablet">이름</th>
								<th data-hide="phone,tablet">권한</th>
								<th data-hide="phone,tablet">학과</th>
								<th data-hide="phone,tablet">상태</th>
							</tr>
						</thead>
						<tbody>
						<%
							If Not IsNull(AryHash) Then
								intNUM = ubound(AryHash,1) + 1
								For i = 0 to ubound(AryHash,1)
						%>
							<tr class="viewDetail" EmpID="<%= AryHash(i).Item("EmpID") %>">
								<td><%= intNUM %></td>
								<!--<td><%= AryHash(i).Item("ClientCodeName") %></td>-->
								<td><%= AryHash(i).Item("EmpID") %></td>
								<td><%= AryHash(i).Item("EmpName") %></td>
								<td><%= AryHash(i).Item("ClientLevelName") %></td>
								<td><%= AryHash(i).Item("SubjectName") %></td>
								<td>
									<%= AryHash(i).Item("StateName") %>
									<div class="DataField" style="display:none;">
										<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>"			ColumnName="IDX"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("EmpID")) %>"			ColumnName="EmpID"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ClientLevel")) %>"	ColumnName="ClientLevel"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("EmpPWD")) %>"		ColumnName="EmpPWD"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("EmpName")) %>"		ColumnName="EmpName"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("PhoneNumber")) %>"	ColumnName="PhoneNumber"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Email")) %>"			ColumnName="Email"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("EmpInfo")) %>"		ColumnName="EmpInfo"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("State")) %>"			ColumnName="State"></li>
									</div>
								</td>
							</tr>
						<%
									intNUM = intNUM - 1
								Next
							end if
						%>
						</tbody>
						<!--
						<tfoot>
							<tr>
								<th data-hide="phone">No.</th>
								<th>아이디</th>
								<th data-hide="phone,tablet">이름</th>
								<th data-hide="phone,tablet">권한</th>
								<th data-hide="phone,tablet">학과</th>
								<th data-hide="phone,tablet">상태</th>
							</tr>
						</tfoot>
						-->
					</table>
				</div>
			</div>
			<!-- 테이블 -->
		</div>		
	</div>
</div>


<!-- #InClude Virtual = "/Common/Bottom.asp" -->