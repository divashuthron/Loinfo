<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 3
Dim LeftMenuCode : LeftMenuCode = "StudentRecord"
Dim LeftMenuName : LeftMenuName = "Home / 평가기준관리 / 생기부 설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "생기부 환산 기준 설정"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere

Dim StrURL			: StrURL = "/StudentRecord.asp"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'공식 가져오기 쿼리(공식3번까지만 사용, Formula4,5번은 예비)
SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "		 MYear, FormulaNum, FormulaName, Formula1, Formula2, Formula3, Formula4, Formula5, INPT_USID, INPT_DATE, UPDT_USID, UPDT_DATE "
SQL = SQL & vbCrLf & "FROM StudentRecord "
SQL = SQL & vbCrLf & "WHERE 1 = 1 " & strWhere
SQL = SQL & vbCrLf & "ORDER BY FormulaNum;"

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB = Nothing

%>

<script type="text/javascript">
$(function() {
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
			<div class="pad_t10"></div>

			<!-- 변수정의 표 -->
			<div class="ibox-title">				
				<h3>변수 정의 </h3>
				<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
					<colgroup><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="9%"></col><col width="10%"></col><col width="9%"></col><col width="9%"></col><col width="9"></col></colgroup>
					<thead>	<tr><th>A</th><th>B</th><th>C</th><th>D</th><th>E</th><th>F</th><th>G</th><th>H</th><th>I</th><th>J</th><th>K</th></tr></thead>
					<tbody>	<td>선택학기 평균등급</td><td>검정고시 총등급</td><td>검정고시 과목수</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tbody>
					<thead>	<tr><th>L</th><th>M</th><th>N</th><th>O</th><th>P</th><th>Q</th><th>R</th><th>S</th><th>U</th><th>V</th><th>Z</th></tr></thead>
					<tbody>	<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td>기타공식 값</td><td>최종값</td></tbody>
				</table>				
				<div class="ibox-tools">
					<a class="collapse-link">
						<!--<i class="fa fa-chevron-up"></i>-->
					</a>
				</div>
			</div>
			<!-- 변수정의 표 끝 -->

			<!-- 공식입력란 -->			
			<div class="ibox-title">
				<form name="InputForm" id="InputForm" method="post" action="/Process/StudentRecordProc.asp">
					<div style="display:none;">
						<input type="text" name="process" id="process" value="RegStudentRecord">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">						
					</div>
					<h3>공식 설정 </h3>
				<%
					'If Not IsNull(AryHash) Then
					If isArray(AryHash) Then
						For i = 0 to ubound(AryHash,1)						
				%>
					<input type="hidden" name="FormulaNum" id="FormulaNum" value="<%=AryHash(i).Item("FormulaNum")%>">
					<input type="hidden" name="FormulaName" id="FormulaName" value="<%=AryHash(i).Item("FormulaName")%>">
					<input type="hidden" name="INPT_USID" id="INPT_USID" value="<%=AryHash(i).Item("INPT_USID")%>">
					<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
						<colgroup><col width="20%"></col><col width="10%"></col><col width="10%"></col><col width="10%"></col><col width="10%"></col><col width="15%"></col><col width="10%"></col><col width="15"></col></colgroup>
						<thead>
						<tr>
							<% If AryHash(i).Item("FormulaNum") = "1" Then %>						
									<th>구분</th><th>수시</th><th>정시</th><th>기타공식</th><th>최초입력자</th><th>최초입력시간</th><th>수정자</th><th>수정입력시간</th>
							<% End If %>
						</tr>
						</thead>
						<tbody>
						<tr>
							<td><h4><%= AryHash(i).Item("FormulaName") %></h4></td>
							<td><input type="text" name="Formula1" class="form-control input-sm" value="<%= AryHash(i).Item("Formula1") %>" placeholder="공식1을 입력하세요."></td>
							<td><input type="text" name="Formula2" class="form-control input-sm" value="<%= AryHash(i).Item("Formula2") %>" placeholder="공식2를 입력하세요."></td>
							<td><input type="text" name="Formula3" class="form-control input-sm" value="<%= AryHash(i).Item("Formula3") %>" placeholder="공식3을 입력하세요."></td>
							<td><%= AryHash(i).Item("INPT_USID") %></td>
							<td><%= AryHash(i).Item("INPT_DATE") %></td>
							<td><%= AryHash(i).Item("UPDT_USID") %></td>
							<td><%= AryHash(i).Item("UPDT_DATE") %></td>
						</tr>
						</tbody>
						<%
						Next
					end if
					%>
					</table>
					<br>
					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">초기화</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>
				</form>
			</div>
			<!-- 공식입력란 끝-->	
			<!-- 테이블 -->
			<div class="pad_t10"></div>
		</div>		
	</div>
</div>
<!-- #InClude Virtual = "/Common/Bottom.asp" -->