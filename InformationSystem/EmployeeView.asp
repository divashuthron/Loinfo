<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 9
Dim LeftMenuCode : LeftMenuCode = "Employee"
Dim LeftMenuName : LeftMenuName = "Home / 환경설정 / 사용자관리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "사용자관리"
Dim LogDivision	: LogDivision = "EmployeeView"
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
Dim StrURL			: StrURL = "/EmployeeList.asp?type=Employee"
Dim StrViewURL		: StrViewURL = "/EmployeeView.asp?type=Employee"

Dim EmpID			: EmpID	= fnR("EmpID", "")
Dim	ProcessType, ClientCode, ClientLevel, EmpPWD, EmpName, PhoneNumber, Email, JoinDate, OutDate, EmpInfo
Dim State, StateName, RegDate, RegID, EditDate, EditID

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	EmpID, ClientCode, ClientLevel, EmpPWD, EmpName, ISNULL(PhoneNumber, '') as PhoneNumber "
SQL = SQL & vbCrLf & "	, ISNULL(Email, '') as Email, ISNULL(JoinDate, '') as JoinDate, ISNULL(OutDate, '') as OutDate, ISNULL(EmpInfo, '') as EmpInfo "
SQL = SQL & vbCrLf & "	, State, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID "
SQL = SQL & vbCrLf & "From Employee AS A "
SQL = SQL & vbCrLf & "Where 1 = 1 "
SQL = SQL & vbCrLf & "	AND EmpID = ?; "

Call objDB.sbSetArray("@EmpID", adVarchar, adParamInput, 25, EmpID)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

If Not IsNull(AryHash) then
	ProcessType		= "EmployeeUpdate"
	EmpID			= AryHash(0).Item("EmpID")
	ClientCode		= AryHash(0).Item("ClientCode")
	ClientLevel		= AryHash(0).Item("ClientLevel")
	EmpPWD			= AryHash(0).Item("EmpPWD")
	EmpName			= AryHash(0).Item("EmpName")
	PhoneNumber		= AryHash(0).Item("PhoneNumber")
	Email			= AryHash(0).Item("Email")
	JoinDate		= AryHash(0).Item("JoinDate")
	OutDate			= AryHash(0).Item("OutDate")
	EmpInfo			= AryHash(0).Item("EmpInfo")
	State			= AryHash(0).Item("State")
	StateName		= AryHash(0).Item("StateName")
	RegDate			= AryHash(0).Item("RegDate")
	RegID			= AryHash(0).Item("RegID")
	EditDate		= AryHash(0).Item("EditDate")
	EditID			= AryHash(0).Item("EditID")

	strLogMSG = "사용자관리  > " & SessionUserID & "이/가 사용자 "& EmpID &" 정보가 조회 했습니다."
	Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)
Else
	ProcessType = "EmployeeInsert"
End If

Set objDB	= Nothing
%>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>사용자 기본정보</h5>
				<div class="ibox-tools">
					<a class="collapse-link">
						<i class="fa fa-chevron-up"></i>
					</a>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/EmployeeProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegEmployee">
						<input type="hidden" name="ProcessType" id="ProcessType" value="<%= ProcessType %>">
						<input type="hidden" name="ClientLevel_OLD" id="ClientLevel_OLD" value="<%= ClientLevel %>">
					</div>
					<!--
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							학교
						</div>
						<div class="col-md-4">
							<% Call SubCodeSelectBox("ClientCode", "학교선택", ClientCode, "학교를 선택해 주세요.", "", "SchoolCode") %>
						</div>
					</div>
					-->
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							사용자 권한 *
						</div>
						<div class="col-md-4">
							<% Call SubCodeSelectBox("ClientLevel", "사용자권한선택", ClientLevel, "사용자 권한을 선택해 주세요.", "", "UserGrade") %>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							사용자 ID *
						</div>
	<% If Not IsNull(AryHash) Then %>
						<div class="col-md-4">
							<input type="text" name="EmpID" id="EmpID" class="form-control input-sm" value="<%= EmpID %>" readonly>
						</div>
						<div class="col-md-4" style="padding-top:10px;">
							※ <span style="color:#e5322b;">사용자 아이디</span>는 변경할 수 없습니다.
						</div>
	<% Else %>
						<div class="col-md-4">
							<input type="text" name="EmpID" id="EmpID" class="form-control input-sm CheckUserId" value="<%= EmpID %>" maxlength="25">
							<input type="hidden" name="CheckID" id="CheckID" value="N">
						</div>
						<div class="col-md-4">
							<span class="btnBasic" id="CheckUserId" style="width:142px;">아이디 중복 확인</span>
						</div>
	<% End If %>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							비밀번호 *
						</div>
						<div class="col-md-4">
							<input type="password" name="EmpPWD" id="EmpPWD" class="form-control input-sm" value="<%= EmpPWD %>" maxlength="25">
						</div>
						<div class="col-md-6" style="padding-top:10px;">
							※ <span style="color:#e5322b;">비밀번호</span>는 영문 + 숫자 + 특수문자 조합으로 8자리 이상 입력해 주세요.
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							사용자 성명 *
						</div>
						<div class="col-md-4">
							<input type="text" name="EmpName" id="EmpName" class="form-control input-sm" value="<%= EmpName %>" maxlength="25">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							전화번호
						</div>
						<div class="col-md-4">
							<input type="text" name="PhoneNumber" id="PhoneNumber" class="form-control input-sm" value="<%= PhoneNumber %>" maxlength="20">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							이메일
						</div>
						<div class="col-md-10">
							<input type="text" name="Email" id="Email" class="form-control input-sm" value="<%= Email %>" maxlength="255">
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							기타
						</div>
						<div class="col-md-10">
							<textarea name="EmpInfo" id="EmpInfo" class="form-control input-sm" maxlength="500" style="height:100px;"><%= EmpInfo %></textarea>
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-2 grid_sub_title">
							상태 *
						</div>
						<div class="col-md-2">
							<select name="State" id="State" class="form-control input-sm" alert="상태를 선택하세요.">
								<option value="">상태선택</option>
								<option value="Y" <%= setSelected(State, "Y") %>>사용</option>
								<option value="N" <%= setSelected(State, "N") %>>미사용</option>
							</select>
						</div>
					</div>
					<!-- 버튼 -->
					( * 는 필수 입력값입니다.)
					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeSave" id="btnSave">저장</span>
							<span class="btnBasic btnTypeCancel" id="btnCancel">취소</span>
						</div>
					</div>
					<!-- 버튼 -->
				</form>
			</div>
			<!-- 테이블 -->
		</div>
	</div>
</div>



<!-- 메인 컨텐츠 -->

<script type="text/javascript">
	$(function() {
		// 아이디 중복 확인
		$("#CheckUserId").click(function() {
			var $EmpID = $("#EmpID");
			var $CheckID = $("#CheckID");
			
			if(!$EmpID.val()){
				alert("사용자 아이디를 입력 하신 후 중복 확인을 해주세요.");
				$CheckID.val("N");
				$EmpID.focus();
				return;
			}
			/*
			if(not_ck_kor($EmpID)) {
				alert("사용자 아이디는 한글을 포함할 수 없습니다.");
				$CheckID.val("N");
				$EmpID.focus();
				return;
			}
			*/
			if(ck_blank($EmpID)) {
				alert("사용자 아이디는 공백을 포함할 수 없습니다.");
				$CheckID.val("N");
				$EmpID.focus();
				return;
			}
			/*
			if(validatenum2($EmpID)) {
				alert("사용자 아이디는 한글/특수문자/공백을 포함할 수 없습니다.");
				$CheckID.val("N");
				$EmpID.focus();
				return;
			}
			*/
			/*
			if($EmpID.val().length < 5){
				alert("사용자 아이디는 5자 이상 입력해 주세요.");
				$CheckID.val("N");
				$EmpID.focus();
				return;
			}
			*/
			
			var aryData = {"process":"CheckID", "EmpID":$EmpID.val()}
			$.Ajax4Get("/Process/EmployeeProc.asp", aryData, "$.SearchID(datas)", "xml", "","", false);
		});

		// 아이디 중복 검사 결과
		$.SearchID = function(datas) {
			var $objList	= $(datas).find("List");	
			var strMSG;
			var $EmpID = $("#EmpID");
			var $EmpPWD = $("#EmpPWD");
			var $CheckID = $("#CheckID");
			
			if ($objList.find("Result").text() == "true") {
				alert("이미 등록되어 있는 아이디 입니다.\n다시 입력해 주시기 바랍니다.");
				$EmpID.focus();
				$EmpID.val("");
				$CheckID.val("N");
			} else {
				alert("사용 가능한 아이디 입니다.");
				$EmpPWD.focus();
				$CheckID.val("Y");
			}
		}
		
		// 아이디 재입력 방지
		$(".CheckUserId").keyup(function() {
			if ($("#CheckID").val() == "Y") { $("#CheckID").val("N"); }
		});

		// 저장
		$("#btnSave").click(function() {
			//var $ClientCode = $("#ClientCode");
			var $ClientLevel = $("form[id='InputForm'] [name='ClientLevel']");
			var $EmpID = $("#EmpID");
			var $CheckID = $("#CheckID");
			var $EmpPWD = $("#EmpPWD");
			var $EmpName = $("#EmpName");
			var $PhoneNumber = $("#PhoneNumber");
			var $Email = $("#Email");
			var $State = $("#State");
			var MenuSeq = "";

			//if (!$.chkInputValue($ClientCode,		"학교를 선택해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($ClientLevel,		"사용자 권한을 선택해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($EmpID,			"사용자 아이디를 입력해 주시기 바랍니다.")) { return; }
			if ($CheckID.val() == "N"){ 			alert("아이디 중복 확인을 해주세요."); $EmpID.focus(); return; }
			if (!$.chkInputValue($EmpPWD,			"비밀번호를 입력해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($EmpName,			"사용자 성명을 입력해 주시기 바랍니다.")) { return; }
			//if (!$.chkInputValue($PhoneNumber,	"전화번호를 입력해 주시기 바랍니다.")) { return; }
			//if (!$.chkInputValue($Email,			"이메일을 입력해 주시기 바랍니다.")) { return; }
			if (!$.chkInputValue($State,			"상태를 선택해 주시기 바랍니다.")) { return; }
			
			//if ($.setValidation($("#InputForm"))) {
				if (confirm("사용자 정보를 저장 하시겠습니까?")) {
					$.Ajax4FormSubmit($("#InputForm"), "사용자 정보 저장이 완료되었습니다.", "/EmployeeList.asp");
				}
			//}
		});
		
		// 취소
		$("#btnCancel").click(function() {
			$.goURL("<%= StrURL %>");
		});

		// 삭제
		$("#btnDelete").click(function() {
		});
	});
</script>


<!-- #InClude Virtual = "/Common/Bottom.asp" -->