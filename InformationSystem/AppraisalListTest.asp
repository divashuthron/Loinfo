<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Appraisal"
Dim LeftMenuName : LeftMenuName = "Home / 평가기준관리 / 평가비율 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "평가기준관리"
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
Dim StrURL			: StrURL = "/AppraisalList.asp"

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	b.IDX,a.MYear,a.SubjectCode"
SQL = SQL & vbCrLf & "	, a.Division0, a.Subject, a.Division1, a.Division2, a.Division3 "
SQL = SQL & vbCrLf & "	, a.Quorum,a.QuorumFix"
SQL = SQL & vbCrLf & "	, a.RF1,a.RF2,a.RF3,a.RF4,a.RF5,a.RF6,a.RF7,a.RF8,a.RF9,a.RF10,a.RF11"
SQL = SQL & vbCrLf & "	, b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio"
SQL = SQL & vbCrLf & "	, b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5"
SQL = SQL & vbCrLf & "	, b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5"
SQL = SQL & vbCrLf & "	, b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5"
SQL = SQL & vbCrLf & "	, b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5"
SQL = SQL & vbCrLf & "	, b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5"
SQL = SQL & vbCrLf & "	, b.DocumentaryEvidence6, b.DocumentaryEvidence7, b.DocumentaryEvidence8, b.DocumentaryEvidence9, b.DocumentaryEvidence10"
SQL = SQL & vbCrLf & "	, b.INPT_USID, b.INPT_DATE, b.INPT_ADDR, b.UPDT_USID, b.UPDT_DATE, b.UPDT_ADDR, b.InsertTime"
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', a.Division0) AS Division0Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', a.Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division3', a.Division3) AS Division3Name "
SQL = SQL & vbCrLf & "FROM SubjectTable AS a "
SQL = SQL & vbCrLf & "	left outer join AppraisalTable AS b"
SQL = SQL & vbCrLf & "		on a.SubjectCode = b.SubjectCode"
SQL = SQL & vbCrLf & "WHERE 1 = 1 " & strWhere
SQL = SQL & vbCrLf & "ORDER BY a.IDX DESC;"

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
	
	// 목록 클릭 시 업데이트 프로세스로 변경
	//$(document).delegate("tr.viewDetail_SetDate_2", "click", function() {
	//	$("#InputForm input[name='ProcessType']").val("Update");
	//});
});
</script>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>학과 목록</h5>
				<div class="ibox-tools">
					<!--나중에 저장버튼으로 사용 <span class="btnBasic btnTypeAdd" id="btnREG">학과 등록</span>-->
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
								<th data-hide="phone,tablet">최초입력자</th>
								<th data-hide="phone,tablet">최초등록일</th>
								<th data-hide="phone">최종수정자</th>
								<th data-hide="phone">최종수정일</th>
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
							<tr class="viewDetail_SetDate_2" SubjectCode="<%= AryHash(i).Item("SubjectCode") %>">
								<td><%= intNUM %></td>
								<td><%= AryHash(i).Item("MYear") %></td>
								<td><%= AryHash(i).Item("SubjectCode") %></td>
								<td><%= AryHash(i).Item("Division0Name") %></td>
								<td><%= AryHash(i).Item("SubjectName") %></td>
								<td><%= AryHash(i).Item("Division1Name") %></td>
								<td><%= AryHash(i).Item("Division2Name") %></td>
								<td><%= AryHash(i).Item("Division3Name") %></td>
								<td><%= AryHash(i).Item("INPT_USID") %></td>
								<td><%= Left(AryHash(i).Item("INPT_DATE"),10) %></td>
								<td><%= AryHash(i).Item("UPDT_USID") %></td>
								<td>
									<%= Left(AryHash(i).Item("UPDT_DATE"),10) %>

									<div class="DataField" style="display:none;">
										<!--<li Columnvalue="Update"										ColumnName="ProcessType"></li>-->
										<li Columnvalue="<%= Trim(AryHash(i).Item("IDX")) %>						ColumnName="IDX"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("MYear")) %>						ColumnName="MYear"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("SubjectCode")) %>				ColumnName="SubjectCode"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Division0")) %>					ColumnName="Division0"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Subject")) %>					ColumnName="Subject"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Division1")) %>					ColumnName="Division1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Division2")) %>					ColumnName="Division2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Division3 ")) %>					ColumnName="Division3 "></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Quorum")) %>						ColumnName="Quorum"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("QuorumFix")) %>					ColumnName="QuorumFix"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF1")) %>						ColumnName="RF1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF2")) %>						ColumnName="RF2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF3")) %>						ColumnName="RF3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF4")) %>						ColumnName="RF4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF5")) %>						ColumnName="RF5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF6")) %>						ColumnName="RF6"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF7")) %>						ColumnName="RF7"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF8")) %>						ColumnName="RF8"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF9")) %>						ColumnName="RF9"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF10")) %>						ColumnName="RF10"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("RF11")) %>						ColumnName="RF11"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("StudentRecordRatio")) %>			ColumnName="StudentRecordRatio"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("InterviewerRatio")) %>			ColumnName="InterviewerRatio"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("PracticalRatio")) %>				ColumnName="PracticalRatio"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("CSATRatio")) %>					ColumnName="CSATRatio"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard1")) %>				ColumnName="DrawStandard1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard2")) %>				ColumnName="DrawStandard2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard3")) %>				ColumnName="DrawStandard3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard4")) %>				ColumnName="DrawStandard4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DrawStandard5")) %>				ColumnName="DrawStandard5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard1")) %>		ColumnName="UnqualifiedStandard1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard2")) %>		ColumnName="UnqualifiedStandard2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard3")) %>		ColumnName="UnqualifiedStandard3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard4")) %>		ColumnName="UnqualifiedStandard4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UnqualifiedStandard5")) %>		ColumnName="UnqualifiedStandard5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint1")) %>				ColumnName="ExtraPoint1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint2")) %>				ColumnName="ExtraPoint2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint3")) %>				ColumnName="ExtraPoint3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint4")) %>				ColumnName="ExtraPoint4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("ExtraPoint5")) %>				ColumnName="ExtraPoint5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship1")) %>				ColumnName="Scholarship1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship2")) %>				ColumnName="Scholarship2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship3")) %>				ColumnName="Scholarship3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship4")) %>				ColumnName="Scholarship4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("Scholarship5")) %>				ColumnName="Scholarship5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence1")) %>		ColumnName="DocumentaryEvidence1"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence2")) %>		ColumnName="DocumentaryEvidence2"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence3")) %>		ColumnName="DocumentaryEvidence3"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence4")) %>		ColumnName="DocumentaryEvidence4"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence5")) %>		ColumnName="DocumentaryEvidence5"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence6")) %>		ColumnName="DocumentaryEvidence6"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence7")) %>		ColumnName="DocumentaryEvidence7"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence8")) %>		ColumnName="DocumentaryEvidence8"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence9")) %>		ColumnName="DocumentaryEvidence9"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryEvidence10")) %>		ColumnName="DocumentaryEvidence10"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("INPT_USID")) %>					ColumnName="INPT_USID"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("INPT_DATE")) %>					ColumnName="INPT_DATE"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("INPT_ADDR")) %>					ColumnName="INPT_ADDR"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UPDT_USID")) %>					ColumnName="UPDT_USID"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UPDT_DATE")) %>					ColumnName="UPDT_DATE"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("UPDT_ADDR")) %>					ColumnName="UPDT_ADDR"></li>
										<li Columnvalue="<%= Trim(AryHash(i).Item("InsertTime")) %>					ColumnName="InsertTime"></li>

									</div>

								</td>
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