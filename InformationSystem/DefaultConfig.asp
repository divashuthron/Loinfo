<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 9
Dim LeftMenuCode : LeftMenuCode = "DefaultConfig"
Dim LeftMenuName : LeftMenuName = "Home / 환경설정 / 기본환경설정"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "기본환경설정"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere
Dim i, strMSG, intNUM, strTEMP, strRESULT

Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 10
Dim PageNum			: PageNum	= fnR("page", 1)
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/DefaultConfig.asp"

Dim	ProcessType
Dim Idx, MYear, Division, Subject, Division1, Division2
Dim SchoolName, SchoolAddress, SchoolSmsNumber, SchoolTelNumber
Dim ApplyConfirm, ApplyPrintConfirm, InterviewConfirm
Dim State, StateName, RegDate, RegID
Dim BillConfirm, DemandsConfirm, ApplicationAddConfirm, ApplicantAddConfirm	


Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT TOP 1"
SQL = SQL & vbCrLf & "	Idx, MYear, Division, Subject, Division1, Division2 "
SQL = SQL & vbCrLf & "	, SchoolName, SchoolAddress, SchoolSmsNumber, SchoolTelNumber "
SQL = SQL & vbCrLf & "	, ApplyConfirm, ApplyPrintConfirm, InterviewConfirm "
SQL = SQL & vbCrLf & "	, BillConfirm, DemandsConfirm, ApplicationAddConfirm, ApplicantAddConfirm	"
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', Division) AS DivisionName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', Division1) AS Division1Name "
'SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division2', Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	, '' AS Division2Name "
SQL = SQL & vbCrLf & "	, State, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID "
SQL = SQL & vbCrLf & "FROM ConfigTable AS A " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & "	AND State = 'Y' "
SQL = SQL & vbCrLf & "ORDER BY IDX DESC; "

'objDB.blnDebug = true
'arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, Nothing)

If Not IsNull(AryHash) then
	ProcessType				= "DefaultConfigInsert"
	Idx						= AryHash(0).Item("Idx")
	MYear					= AryHash(0).Item("MYear")
	Division				= AryHash(0).Item("Division")
	Subject					= AryHash(0).Item("Subject")
	Division1				= AryHash(0).Item("Division1")
	Division2				= AryHash(0).Item("Division2")
	SchoolName				= AryHash(0).Item("SchoolName")
	SchoolAddress			= AryHash(0).Item("SchoolAddress")
	SchoolSmsNumber			= AryHash(0).Item("SchoolSmsNumber")
	SchoolTelNumber			= AryHash(0).Item("SchoolTelNumber")
	ApplyConfirm			= AryHash(0).Item("ApplyConfirm")
	ApplyPrintConfirm		= AryHash(0).Item("ApplyPrintConfirm")
	InterviewConfirm		= AryHash(0).Item("InterviewConfirm")
	BillConfirm				= AryHash(0).Item("BillConfirm")
	DemandsConfirm			= AryHash(0).Item("DemandsConfirm")
	ApplicationAddConfirm	= AryHash(0).Item("ApplicationAddConfirm")
	ApplicantAddConfirm		= AryHash(0).Item("ApplicantAddConfirm")
	State					= AryHash(0).Item("State")
	StateName				= AryHash(0).Item("StateName")
	RegDate					= AryHash(0).Item("RegDate")
	RegID					= AryHash(0).Item("RegID")
End if

Set objDB	= Nothing
%>

<script type="text/javascript">
	$(function() {
		// 저장
		$("#btnSave").click(function() {
			if ($.setValidation($("#InputForm"))) {
				if (confirm("지원자 정보를 저장 하시겠습니까?")) {
					$.Ajax4FormSubmit($("#InputForm"), "기본환경설정 저장이 완료되었습니다.");
				}
			}
		});
	});
</script>


<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
		<form id="InputForm" method="post" action="/Process/DefaultConfigProc.asp">
			<div style="display:none;">
				<input type="hidden" name="process" id="process" value="RegDefaultConfig">
				<input type="hidden" name="ProcessType" id="ProcessType" value="<%= ProcessType %>">
				<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
			</div>
			

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>기본정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						년도 *
					</div>
					<div class="col-md-2 col-xs-7">
						<% Call SubCodeSelectBox("MYear", "년도선택", MYear, "년도를 선택해 주세요.", "", "MYear") %>
					</div>
				</div>

				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						모집시기 *
					</div>
					<div class="col-md-2 col-xs-7">
						<% Call SubCodeSelectBox("Division0", "모집시기선택", Division, "모집시기를 선택해 주세요.", "", "Division0") %>
					</div>
				</div>
				<!--
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						학과 *
					</div>
					<div class="col-md-3 col-xs-7">
						<% Call SubCodeSelectBox("Subject", "학과명선택", Subject, "", "", "Subject") %>
					</div>
				</div>

				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						전형 *
					</div>
					<div class="col-md-3 col-xs-7">
						<% Call SubCodeSelectBox("Division1", "전형선택", Division1, "", "", "Division1") %>
					</div>
				</div>
				-->
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						대학명
					</div>
					<div class="col-md-3 col-xs-7">
						<input type="text" name="SchoolName" id="SchoolName" class="form-control input-sm" value="<%= SchoolName %>" maxlength="25" alert="대학명을 입력해 주세요.">
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						대학주소
					</div>
					<div class="col-md-8 col-xs-9">
						<input type="text" name="SchoolAddress" id="SchoolAddress" class="form-control input-sm" value="<%= SchoolAddress %>" maxlength="100">
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						SMS 회신번호
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="SchoolSmsNumber" id="SchoolSmsNumber" class="form-control input-sm" value="<%= SchoolSmsNumber %>" maxlength="50">
					</div>
					<div class="col-md-2 col-xs-2 grid_sub_title2">
						문의전화번호
					</div>
					<div class="col-md-2 col-xs-7">
						<input type="text" name="SchoolTelNumber" id="SchoolTelNumber" class="form-control input-sm" value="<%= SchoolTelNumber %>" maxlength="50">
					</div>
				</div>
			</div>
			<!-- 테이블 -->

			
			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>제어정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>
			
			<div class="ibox-content">
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						고지서 사용여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="BillConfirm" id="BillConfirm" class="form-control input-sm" alert="고지서 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(BillConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(BillConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						유의사항 사용여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="DemandsConfirm" id="DemandsConfirm" class="form-control input-sm" alert="유의사항 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(DemandsConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(DemandsConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						입학원서 수동입력 페이지 사용여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="ApplicationAddConfirm" id="ApplicationAddConfirm" class="form-control input-sm" alert="유의사항 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(ApplicationAddConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(ApplicationAddConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						지원자 서류체크 페이지 사용여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="ApplicantAddConfirm" id="ApplicantAddConfirm" class="form-control input-sm" alert="유의사항 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(ApplicantAddConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(ApplicantAddConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>

				<!--
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						예약신청가능여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="ApplyConfirm" id="ApplyConfirm" class="form-control input-sm" alert="예약신청 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(ApplyConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(ApplyConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						수험표출력가능여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="ApplyPrintConfirm" id="ApplyPrintConfirm" class="form-control input-sm" alert="수험표출력 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(ApplyPrintConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(ApplyPrintConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>-->
				<div class="row show-grid">
					<div class="col-md-2 col-xs-2 grid_sub_title">
						면접평가가능여부
					</div>
					<div class="col-md-2 col-xs-7">
						<select name="InterviewConfirm" id="InterviewConfirm" class="form-control input-sm" alert="면접평가 가능 여부를 선택하세요.">
							<option value="">가능여부</option>
							<option value="Y" <%= setSelected(InterviewConfirm, "Y") %>>사용</option>
							<option value="N" <%= setSelected(InterviewConfirm, "N") %>>미사용</option>
						</select>
					</div>
				</div>
			</div>
			<!-- 테이블 -->

		</form>
		</div>

		<!-- 버튼 -->
		<div class="row show-grid grid_sub_button">
			<div class="col-md-12">
				<div style="padding:0 20px 5px 0;">
				
					<span class="btnBasic btnTypeSave" id="btnSave">저장</span>
					<span class="btnBasic btnTypeCancel" id="btnCancel">취소</span>
				</div>
			</div>
		</div>
		<!-- 버튼 -->

	</div>
</div>



<!-- #InClude Virtual = "/Common/Bottom.asp" -->