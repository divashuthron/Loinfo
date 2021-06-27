<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 7
Dim LeftMenuCode : LeftMenuCode = "SucCandidate"
Dim LeftMenuName : LeftMenuName = "Home / 합격자발표관리 / 합격자조회"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "합격자조회"
Dim LogDivision	: LogDivision = "SuccessfulStudentList"

'히스토리
Dim strLogMSG
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

Dim BGColor '동석차가 가려지지 않은 합격자에 대한 색깔표시

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
SQL = SQL & vbCrLf & " select	a.Myear, b.Division0, b.Subject, b.Division1 "
SQL = SQL & vbCrLf & " 			,dbo.getSubCodeName('Division0', a.Division0) AS Division0Name "
SQL = SQL & vbCrLf & "			,dbo.getSubCodeName('Subject', a.Subject) AS SubjectName "
SQL = SQL & vbCrLf & "			,dbo.getSubCodeName('Division1', a.Division1) AS Division1Name "
SQL = SQL & vbCrLf & "			,dbo.getSubCodeName('Division2', a.Division2) AS Division2Name "
SQL = SQL & vbCrLf & "			,dbo.getSubCodeName('Division3', a.Division3) AS Division3Name "
SQL = SQL & vbCrLf & "			,dbo.getSubCodeTemp1('ExtraPoint', b.ExtraPoint) AS ExtraPointScore "
SQL = SQL & vbCrLf & "			,b.StudentNumber, b.StudentNameKor, b.SubjectCode, a.totScore "
SQL = SQL & vbCrLf & "			,a.Standing, a.DrawStanding, a.Result, a.BackupStanding "
SQL = SQL & vbCrLf & "			,a.StudentRecordScore, a.InterviewerScore, a.QualificationScore "
SQL = SQL & vbCrLf & "			,a.UniversityScore, a.CSATScore, a.StudentRecordAverage, a.CreditSum "
SQL = SQL & vbCrLf & "			,a.ChoiceSemester, a.KorLanScore, a.EnglishScore, a.MathematicsScore "
SQL = SQL & vbCrLf & "			,a.UniversityCredit, a.Minor "
SQL = SQL & vbCrLf & " from ChangeScoreTable a join ApplicationTable b "
SQL = SQL & vbCrLf & " on a.StudentNumber = b.StudentNumber "
SQL = SQL & vbCrLf & " and a.Myear = b.Myear "
SQL = SQL & vbCrLf & " where 1 = 1 " & strWhere
SQL = SQL & vbCrLf & " order by a.Division0 desc, a.Subject, a.Division1, a.Standing "

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

'개인정보가 있는 합격자리스트는 조회도 기록함
strLogMSG = "합격자 조회  > " & SessionUserID  &"가/이 합격자 리스트를 조회 했습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)
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
							<!--
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
							-->
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
					<span class="btnBasic btnTypeEdit" id="DrawStandingSet">동석차순위 수정하기</span>
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<!--<table id="dt_basic" class="table table-striped table-bordered table-hover" width="100%">-->
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<colgroup>
								<col width="3%"></col>
								<col width="5%"></col>
								<col width="10%"></col>
								<col width="15%"></col>
								<col width="20%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
								<col width="6%"></col>
							</colgroup>
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>            
									<th data-hide="phone">년도</th>  
									<th data-hide="phone">시기</th>
									<th data-hide="phone">학과</th>  
									<th data-hide="phone">전형</th>  
									<th data-hide="phone">수험번호</th>  
									<th data-hide="phone">이름</th>
									<th data-hide="phone">총점수</th>  
									<th data-hide="phone">석차</th>
									<th data-hide="phone">동석차순위</th>
									<th data-hide="phone">결과</th>
								</tr>
							</thead>
							<tbody>
							<%
								' a.Myear
								' ,dbo.getSubCodeName('Division0', a.Division0) AS Division0Name
								' ,dbo.getSubCodeName('Subject', a.Subject) AS SubjectName
								' ,dbo.getSubCodeName('Division1', a.Division1) AS Division1Name
								' ,dbo.getSubCodeName('Division2', a.Division2) AS Division2Name
								' ,dbo.getSubCodeName('Division3', a.Division3) AS Division3Name
								' ,dbo.getSubCodeTemp1('ExtraPoint', b.ExtraPoint) AS ExtraPointScore
								' ,b.StudentNumber, b.StudentNameKor, b.SubjectCode, a.totScore, a.Standing, a.DrawStanding, a.Result, a.BackupStanding
								' a.StudentRecordScore, a.InterviewerScore, a.QualificationScore, a.UniversityScore, a.CSATScore, 
								' a.StudentRecordAverage, a.CreditSum, a.ChoiceSemester, a.KorLanScore, a.EnglishScore, a.MathematicsScore, a.UniversityCredit, a.Minor

								'If Not IsNull(AryHash) Then
								If isArray(AryHash) Then
									'For i = 0 to ubound(AryHash,1)
									For i = StartNum to EndNum

									If AryHash(i).Item("DrawStanding") = 0 Then
										BGColor = "#74D0F1"
									Else
										BGColor = ""
									End If
							%>
								<tr class="viewDetail_SetDate_2" style="background-color: <%=BGColor%>;" IDX="<%= AryHash(i).Item("IDX") %>">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("Myear") %></td>
									<td style="display:none;"><%= AryHash(i).Item("Division0") %></td>
									<td style="display:none;"><%= AryHash(i).Item("Subject") %></td>
									<td style="display:none;"><%= AryHash(i).Item("Division1") %></td>
									<td><%= AryHash(i).Item("Division0Name") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("StudentNumber") %></td>
									<td><%= AryHash(i).Item("StudentNameKor") %></td>
									<td><%= AryHash(i).Item("totScore") %></td>
									<%If AryHash(i).Item("Result") = "합격" Then%>
										<td><%= AryHash(i).Item("Standing") %></td>
									<%Else%>
										<td><%= "예비 " & AryHash(i).Item("BackupStanding") %></td>
									<%End If%>
									<td><%= AryHash(i).Item("DrawStanding") %></td>
									<td>
										<%= AryHash(i).Item("Result") %>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("Myear")) %>"					ColumnName="MYear"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division0Name")) %>"			ColumnName="Division0Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectName")) %>"			ColumnName="SubjectName"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division1Name")) %>"			ColumnName="Division1Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division2Name")) %>"			ColumnName="Division2Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Division3Name")) %>"			ColumnName="Division3Name"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("totScore")) %>"				ColumnName="totScore"></li>

											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordScore")) %>"	ColumnName="StudentRecordScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("InterviewerScore")) %>"		ColumnName="InterviewerScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("QualificationScore")) %>"	ColumnName="QualificationScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UniversityScore")) %>"		ColumnName="UniversityScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("CSATScore")) %>"				ColumnName="CSATScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPointScore")) %>"		ColumnName="ExtraPointScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordAverage")) %>"	ColumnName="StudentRecordAverage"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("CreditSum")) %>"				ColumnName="CreditSum"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("ChoiceSemester ")) %>"		ColumnName="ChoiceSemester "></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("KorLanScore")) %>"			ColumnName="KorLanScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("EnglishScore")) %>"			ColumnName="EnglishScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("MathematicsScore")) %>"		ColumnName="MathematicsScore"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("UniversityCredit")) %>"		ColumnName="UniversityCredit"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("Minor")) %>"					ColumnName="Minor"></li>
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
				<div style="float:right;">
				</div>
			</div>
			<form name="InputForm" id="InputForm" method="post" action="/Process/AppraisalProc.asp">
			<div class="ibox-content" >				
				<!-- 선택한 합격자 기본정보 -->
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						년도 
					</div>
					<div class="col-md-2 col-xs-7" >
						<input type="text" name="MYear" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						모집시기 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="Division0Name" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						학과 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="SubjectName" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						전형 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="Division1Name" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						구분2 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="Division2Name" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						구분3 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="Division3Name" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<!-- 선택한 합격자 기본정보 끝 -->
			</div>

			<div class="ibox-title">
				<h5>환산점수 / 가산점</h5>
				<div style="float:right;">
				</div>
			</div>

			<div class="ibox-content" >
				<!-- 환산점수 / 가산점 -->
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						생기부환산점수 
					</div>
					<div class="col-md-2 col-xs-7" >
						<input type="text" name="StudentRecordScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						면접점수
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="InterviewerScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						검정고시환산점수 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="QualificationScore" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						대학백분율점수 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="UniversityScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						수능환산성적
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="CSATScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						가산점
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="ExtraPointScore" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<!-- 환산점수 / 가산점 끝 -->
			</div>

			<div class="ibox-title">
				<h5>동석차 기준 </h5>
				<div style="float:right;">
				</div>
			</div>

			<div class="ibox-content" >
				<!-- 동석차 기준 / 총점 -->
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						교과 성적 평균등급 
					</div>
					<div class="col-md-2 col-xs-7" >
						<input type="text" name="StudentRecordAverage" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						이수단위 합
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="CreditSum" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						반영학기 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="ChoiceSemester" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						국어영역 성적 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="KorLanScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						영어영역 성적
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="EnglishScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						수학영역 성적
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="MathematicsScore" class="form-control input-sm InputBGcolor" readonly>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						대학이수학점 
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="UniversityCredit" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						면접점수
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="InterviewerScore" class="form-control input-sm InputBGcolor" readonly>
					</div>

					<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						나이
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="Minor" class="form-control input-sm InputBGcolor" readonly>
					</div>					

					<!--<div class="col-md-2 col-xs-2 grid_sub_title2 text-center">
						총점(환산+가산점)
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="totScore" class="form-control input-sm InputBGcolor" readonly>
					</div>-->
				</div>
				<!-- 동석차 기준 / 총점 끝 -->
				<div class="row show-grid">&nbsp;</div>
			</div>
			</form>
			<!-- 상세보기 끝 -->

			<!-- 동석차 수정하기 -->
			<div id="DrawStandingSetModal" style="margin:5px; display:none;">
				<form name="DrawStandingSetForm" id="DrawStandingSetForm" method="post" action="/Process/DrawStandingProc.asp">
					<div style="display:none;">
						<input type="text" name="DrawMyear" id="DrawMyear" value="">
						<input type="text" name="DrawDivision0" id="DrawDivision0" value="">
						<input type="text" name="DrawSubject" id="DrawSubject" value="">
						<input type="text" name="DrawDivision1" id="DrawDivision1" value="">
					</div>
					<div class="ibox-content">							
						<div class="row show-grid" style="text-align:left;">
							<div class="col-md-3">
								수험번호
							</div>
							<div class="col-md-3">
								<input type="text" id="DrawStudentNumber" name="DrawStudentNumber" class="form-control input-sm InputBGcolor" readonly>
							</div>
							<div class="col-md-3">
								동석차등수
							</div>
							<div class="col-md-3">
								<select id="DrawStanding" name="DrawStanding" class="form-control input-sm">
									<option value="">등수</option>
									<option value="1">1</option>
									<option value="2">2</option>
									<option value="3">3</option>
									<option value="4">4</option>
									<option value="5">5</option>
									<option value="6">6</option>
									<option value="7">7</option>
									<option value="8">8</option>
									<option value="9">9</option>
									<option value="10">10</option>
								</select>
							</div>
						</div>
						<div class="row show-grid" style="text-align:left;">
							*동석차 순위를 수정하면, 해당 모집단위의 석차가 다시 계산됩니다.
						</div>
					<div>
				</form>
				
				<br>
				<div class="row show-grid grid_sub_button" >					
					<div class="col-md-12" >
						<span class="btnBasic btnTypeSave" id="RegDrawStandingSet" style="width:80px;">저장</span>
						<span class="btnBasic btnTypeClose SelfCloseDIV" style="width:80px;">취소</span>
					</div>
				</div>
			</div>
			<!-- 동석차 수정하기 -->
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

	// 기본 데이터 설정(모달) 오픈
	$("#DrawStandingSet").click(function() {
		$.openMadal($("#DrawStandingSetModal"), "2");
	});

	// tr선택 시 '동석차 수정하기' 모달에 기본데이터 넣기
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){		
		$("#DrawMyear").val($(this).find("td").eq(1).html());	//년도
		$("#DrawDivision0").val($(this).find("td").eq(2).html());	//모집시기
		$("#DrawSubject").val($(this).find("td").eq(3).html());	//학과
		$("#DrawDivision1").val($(this).find("td").eq(4).html());	//전형
		$("#DrawStudentNumber").val($(this).find("td").eq(8).html());	//수험번호
	});

	// 동석차 수정하기 저장
	$("#RegDrawStandingSet").click(function() {
		if (!$.chkInputValue($("input[name=DrawStudentNumber]"),	"지원자를 선택하고 수정창을 열면 수험번호가 자동 입력됩니다.")) { return; }
		if (!$.chkInputValue($("select[name=DrawStanding]"),		"등수를 선택해 주시기 바랍니다.")) { return; }
		
		if (confirm("동석차 순위를 수정 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setDrawStanding(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/DrawStandingProc.asp";
			$.Ajax4Form("#DrawStandingSetForm", objOpt);
			$("#DrawStandingSetForm").submit();
		}
	});

	// 동석차 수정하기 저장 결과
	$.setDrawStanding = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;	
			
		if ($objList.find("Result").text() == "Complete") {
			window.location.reload();	
			alert("동석차 순위 수정과 해당 모집단위의 석차가 다시 계산되었습니다.");			
		} else {
			alert("동석차 순위 수정 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
});

</script>