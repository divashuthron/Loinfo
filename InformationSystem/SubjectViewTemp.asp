<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 1
Dim LeftMenuCode : LeftMenuCode = "Subject"
Dim LeftMenuName : LeftMenuName = "Home / 면접예약관리 / 학과관리"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "모집단위관리"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere
Dim i, strMSG, intNUM, intNUM2, strTEMP, strRESULT

Dim PageSize		: PageSize	= 15
Dim PageBlock		: PageBlock	= 10
Dim PageNum			: PageNum	= fnR("page", 1)
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/SubjectList.asp"
Dim StrViewURL		: StrViewURL = "/SubjectView.asp"

Dim IDX				: IDX	= fnR("IDX", 0)
Dim ProcessType
Dim MYear, Division, Subject, Division1, Division2, Degree
Dim ApplyTitle, ApplyGroupCount, ApplyTotalNumber, ApplyTotalNumberTemp
Dim ApplyStartDate, ApplyStartTime, ApplyEndDate, ApplyEndTime
Dim ApplyPrintStartDate, ApplyPrintStartTime, ApplyPrintEndDate, ApplyPrintEndTime
Dim InterviewDate, InterviewStartDate, InterviewEndDate, InterviewDays, InterviewEtc, TShirtCheck, StandByRoom, InterviewMemberCnt
Dim StateName, State, Regdate, RegID, EditDate, EditID

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "
SQL = SQL & vbCrLf & "	Idx, MYear, Division, Subject, Division1, Division2, Degree "
SQL = SQL & vbCrLf & "	, ApplyTitle, ApplyGroupCount, ApplyTotalNumber "
SQL = SQL & vbCrLf & "	, ApplyStartDate, ApplyStartTime, ApplyEndDate, ApplyEndTime "
SQL = SQL & vbCrLf & "	, ApplyPrintStartDate, ApplyPrintStartTime, ApplyPrintEndDate, ApplyPrintEndTime "
SQL = SQL & vbCrLf & "	, InterviewStartDate, InterviewEndDate, InterviewDays, InterviewEtc, TShirtCheck, StandByRoom, InterviewMemberCnt "
SQL = SQL & vbCrLf & "	, State, (CASE  State "
SQL = SQL & vbCrLf & "		WHEN 'Y' THEN '사용' "
SQL = SQL & vbCrLf & "		WHEN 'N' THEN '미사용' "
SQL = SQL & vbCrLf & "	END) AS StateName "
SQL = SQL & vbCrLf & "	, RegDate, RegID, EditDate, EditID "
SQL = SQL & vbCrLf & "From SubjectTable "
SQL = SQL & vbCrLf & "Where 1 = 1 "
SQL = SQL & vbCrLf & "	AND IDX = ?; "

Call objDB.sbSetArray("@IDX", adInteger, adParamInput, 0, IDX)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

If Not IsNull(AryHash) then
	ProcessType				= "SubjectUpdate"
	MYear					= AryHash(0).Item("MYear")
	Division				= AryHash(0).Item("Division")
	Subject					= AryHash(0).Item("Subject")
	Division1				= AryHash(0).Item("Division1")
	Division2				= AryHash(0).Item("Division2")
	Degree					= AryHash(0).Item("Degree")
	ApplyTitle				= AryHash(0).Item("ApplyTitle")
	ApplyGroupCount			= AryHash(0).Item("ApplyGroupCount")
	ApplyTotalNumber		= AryHash(0).Item("ApplyTotalNumber")
	ApplyStartDate			= AryHash(0).Item("ApplyStartDate")
	ApplyStartTime			= AryHash(0).Item("ApplyStartTime")
	ApplyEndDate			= AryHash(0).Item("ApplyEndDate")
	ApplyEndTime			= AryHash(0).Item("ApplyEndTime")
	ApplyPrintStartDate		= AryHash(0).Item("ApplyPrintStartDate")
	ApplyPrintStartTime		= AryHash(0).Item("ApplyPrintStartTime")
	ApplyPrintEndDate		= AryHash(0).Item("ApplyPrintEndDate")
	ApplyPrintEndTime		= AryHash(0).Item("ApplyPrintEndTime")
	InterviewStartDate		= AryHash(0).Item("InterviewStartDate")
	InterviewEndDate		= AryHash(0).Item("InterviewEndDate")
	InterviewDays			= AryHash(0).Item("InterviewDays")
	InterviewEtc			= AryHash(0).Item("InterviewEtc")
	TShirtCheck				= AryHash(0).Item("TShirtCheck")
	StandByRoom				= AryHash(0).Item("StandByRoom")
	InterviewMemberCnt		= AryHash(0).Item("InterviewMemberCnt")
	State					= AryHash(0).Item("State")
	StateName				= AryHash(0).Item("StateName")
	RegDate					= AryHash(0).Item("RegDate")
	RegID					= AryHash(0).Item("RegID")
	EditDate				= AryHash(0).Item("EditDate")
	EditID					= AryHash(0).Item("EditID")

	'// 평가조 가져오기
	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "	Idx, SubjectIDX, MYear, Division, Subject, Division1, Division2, Degree "
	SQL = SQL & vbCrLf & "	, GroupName, RoomName "
	SQL = SQL & vbCrLf & "From SubjectGroup "
	SQL = SQL & vbCrLf & "Where 1 = 1 "
	SQL = SQL & vbCrLf & "	AND SubjectIDX = ? "
	'SQL = SQL & vbCrLf & "	AND MYear = ? "
	'SQL = SQL & vbCrLf & "	AND Division = ? "
	'SQL = SQL & vbCrLf & "	AND Subject = ? "
	'SQL = SQL & vbCrLf & "	AND Division1 = ? "
	'SQL = SQL & vbCrLf & "	AND Division2 = ? "
	'SQL = SQL & vbCrLf & "	AND Degree = ? "


	Call objDB.sbSetArray("@SubjectIDX",	adInteger, adParamInput, 0,		IDX)
	'Call objDB.sbSetArray("@MYear",		adVarchar, adParamInput, 4,		MYear)
	'Call objDB.sbSetArray("@Division",		adVarchar, adParamInput, 50,	Division)
	'Call objDB.sbSetArray("@Subject",		adVarchar, adParamInput, 50,	Subject)
	'Call objDB.sbSetArray("@Division1",	adVarchar, adParamInput, 50,	Division1)
	'Call objDB.sbSetArray("@Division2",	adVarchar, adParamInput, 50,	Division2)
	'Call objDB.sbSetArray("@Degree",		adInteger, adParamInput, 0,		Division2)

	'objDB.blnDebug = true
	'arrParams = objDB.fnGetArray
	'AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams)

	'// 면접교시 가져오기
	SQL = ""
	SQL = SQL & vbCrLf & "SELECT "
	SQL = SQL & vbCrLf & "	Idx, SubjectIDX, MYear, Division, Subject, Division1, Division2, Degree "
	SQL = SQL & vbCrLf & "	, InterviewDate, TimeCode, CheckTime, InterviewStartTime, InterviewEndTime, Quorum "
	SQL = SQL & vbCrLf & "From SubjectTime "
	SQL = SQL & vbCrLf & "Where 1 = 1 "
	SQL = SQL & vbCrLf & "	AND SubjectIDX = ? "
	
	Call objDB.sbSetArray("@SubjectIDX",	adInteger, adParamInput, 0,		IDX)

	'objDB.blnDebug = true
	'arrParams = objDB.fnGetArray
	'AryHash3 = objDB.fnExecSQLGetHashMap(SQL, arrParams)
Else
	ProcessType = "SubjectInsert"
End if

Set objDB	= Nothing

'response.write "SessionUserID: " & SessionUserID & "<br>"
'response.End
%>


<script type="text/javascript">
$(function() {
	// 저장
	$("#btnSave").click(function() {
		// 폼검사
		if ($.setValidation($("#InputForm"))) {
			// 대상인원 & 편성인원 비교
			if ($("#ApplyTotalNumberTempDIV").text() != $("#ApplyTotalNumberDIV").text() ) {
				alert("편성인원과 대상인원의 차이가 발생합니다.\n이상이 없다면 계속 진행해 주세요.");
			}
			// 저장
			if (confirm("학과 정보를 저장 하시겠습니까?")) {
				// 입력 / 수정 분기처리 필요해서 Ajax4FormSubmit 사용 안함 -> Ajax4Form 사용
				//$.Ajax4FormSubmit($("#InputForm"), "학과 정보 저장이 완료되었습니다.", "/SubjectView.asp?IDX="+$("#IDX").val());
				var objOpt = {"url":"","param":"","dataType":"xml","before":"","success":"$.setSubject(datas)","complete":"","clear":"","reset":""};
				objOpt["url"] = $("#InputForm").attr("action");
				$.Ajax4Form($("#InputForm"), objOpt);
				$("#InputForm").submit();
			}
		}
	});

	// 저장처리결과
	$.setSubject = function(datas) {
		var $objList = $(datas).find("List");	
		var strMSG;
		
		if ($objList.find("Result").text() == "Complete") {
			alert("학과 정보 저장이 완료 되었습니다.");

			if ($("#InputForm input[name='IDX']").val() == "0") {
				document.location.href = "/SubjectList.asp";
			} else {
				document.location.reload();
			}
		} else {
			alert("학과 정보 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}
	
	// 취소
	$("#btnCancel").click(function() {
		$.goURL("<%= StrURL %>");
	});

	// 삭제
	$("#btnDelete").click(function() {
	});

	// 시간표시
	$("#ApplyStartTime, #ApplyEndTime, #ApplyPrintStartTime, #ApplyPrintEndTime").clockpicker({
		placement: "top", donetext: "Done"
	});
	
	
	// 평가조입력
	$("#InsertSubjectGroup").click(function() {
		var $GroupName = $("#GroupName");
		var $RoomName = $("#RoomName");

		if (!$.chkInputValue($GroupName,	"조이름을 입력해 주세요.")) { return; }
		if (!$.chkInputValue($RoomName,		"고사장을 입력해 주세요.")) { return; }

		// 샘플 복사 후 추가
		$(".SubjectGroupSample").eq(0).clone().appendTo("#SubjectGroupDIV");
		// 복사된 평가조에 값 입력
		$(".SubjectGroupSample:last input[name='GroupName']").val($GroupName.val());
		$(".SubjectGroupSample:last input[name='RoomName']").val($RoomName.val());
		// 입력된 평가조 Show
		$(".SubjectGroupSample:last").removeClass("displayNone");
		// 초기화
		$GroupName.val("").focus();
		$RoomName.val("");
	});

	$(document).on("click", ".GroupDelete", function(){
		$(this).parent().parent().remove();
	});

	
	// 면접교시보기
	$("#btnSubjectTimeEdit").click(function() {
		var $SubjectTimeDIV = $(".displayNoneDIV");
		if ($SubjectTimeDIV.css("display") == "none") {
			$SubjectTimeDIV.show();
		} else {
			$SubjectTimeDIV.hide();
		}
	});

	// 날짜 변경 백드라운드 색상 변경 위한 전역 변수 선언
	var G_InterviewDateTemp = "", G_InterviewDateTemp2 = "", G_LineColorsNum = 0;
	// 면접교시입력
	$("#InsertSubjectTime").click(function() {
		var $InterviewDate = $("#InterviewDate");
		var $TimeCode = $("#TimeCode");
		var $CheckTime = $("#CheckTime");
		var $InterviewStartTime = $("#InterviewStartTime");
		var $InterviewEndTime = $("#InterviewEndTime");
		var $Quorum = $("#Quorum");

		if (!$.chkInputValue($InterviewDate,		"면접일자를 입력해 주세요.")) { return; }
		if (!$.chkInputValue($TimeCode,				"교시를 입력해 주세요.")) { return; }
		if (!$.chkInputValue($CheckTime,			"점검시간을 입력해 주세요.")) { return; }
		if (!$.chkInputValue($InterviewStartTime,	"면접시작시간 입력해 주세요.")) { return; }
		if (!$.chkInputValue($InterviewEndTime,		"면접종료시간 입력해 주세요.")) { return; }
		if (!$.chkInputValue($Quorum,				"배정인원 입력해 주세요.")) { return; }

		// 면접교시 추가
		$.SubjectTimeHtml(
			$InterviewDate.val(), $TimeCode.val(), $CheckTime.val()
			, $InterviewStartTime.val(), $InterviewEndTime.val(), $Quorum.val()
		);

		// 날짜 변경 시 백드라운드 색상 변경
		G_InterviewDateTemp = $InterviewDate.val();
		G_LineColorsNum = (G_InterviewDateTemp == G_InterviewDateTemp2) ? G_LineColorsNum : G_LineColorsNum + 1;
		$(".SubjectTimeSample:last > div").addClass("LineColors_"+ String(G_LineColorsNum));
		G_InterviewDateTemp2 = G_InterviewDateTemp;

		// 면접교시 편성인원 계산 실행
		$.SubjectTimeApplyNumberCount();

		// 초기화
		//$InterviewDate.val("").focus();
		$TimeCode.val("").focus();
		$CheckTime.val("");
		$InterviewStartTime.val("");
		$InterviewEndTime.val("");
		$Quorum.val("");
	});

	// 면접교시 추가
	$.SubjectTimeHtml = function(InterviewDate, TimeCode, CheckTime, InterviewStartTime, InterviewEndTime, Quorum) {
		// 샘플 복사 후 추가
		$(".SubjectTimeSample").eq(0).clone().appendTo("#SubjectTimeDIV");
		// 복사된 면접교시에 값 입력
		$(".SubjectTimeSample:last input[name='InterviewDate']").val(InterviewDate).attr("alert", "면접일자를 입력해 주세요.");
		$(".SubjectTimeSample:last input[name='TimeCode']").val(TimeCode).attr("alert", "교시를 입력해 주세요.");
		$(".SubjectTimeSample:last input[name='CheckTime']").val(CheckTime).attr("alert", "점검시간을 입력해 주세요.");
		$(".SubjectTimeSample:last input[name='InterviewStartTime']").val(InterviewStartTime).attr("alert", "면접시작시간을 입력해 주세요.");
		$(".SubjectTimeSample:last input[name='InterviewEndTime']").val(InterviewEndTime).attr("alert", "면접종료시간을 입력해 주세요.");
		$(".SubjectTimeSample:last input[name='Quorum']").val(Quorum).attr("alert", "배정인원을 입력해 주세요.");
		// 입력된 면접교시 Show
		$(".SubjectTimeSample:last").removeClass("displayNone");
	}

	// 면접교시 삭제
	$(document).on("click", ".TimeDelete", function(){
		$(this).parent().parent().remove();

		// 면접교시 편성인원 계산 실행
		$.SubjectTimeApplyNumberCount();
	});

	// 면접교시 편성인원 계산 (면접교시 입력 / 입력된 배정인원 수정 시 실행 됨)
	$.SubjectTimeApplyNumberCount = function() {
		var ApplyNumberCount = 0;
		$("input[name='Quorum']").each(function() {
			ApplyNumberCount += ($(this).val() == "") ? 0 : parseInt($(this).val());
		});
		// 면접교시 관리 편성인원 수정
		$("#ApplyTotalNumberTempDIV").html(String(ApplyNumberCount));
	}
	
	// 면접교시 샘플 가져오기
	$("#btnSubjectTimeSample").click(function() {
		var $InterviewStartDate = $("#InterviewStartDate");
		var $InterviewEndDate = $("#InterviewEndDate");

		if (!$.chkInputValue($InterviewStartDate,	"면접평가 시작일을 입력해 주세요.")) { return; }
		if (!$.chkInputValue($InterviewEndDate,		"면접평가 종료일을 입력해 주세요.")) { return; }

		if (confirm("면접교시 샘플을 가져옵니다.\n샘플을 가져오면 기존에 입력되어 있던 교시 정보가 초기화 됩니다.\n샘플을 가져오시겠습니까?")) {
			$(".displayNoneDIV").show();
			
			// 면접교시샘플 가져오기
			var aryData = {
				"process":"getSubjectTimeSample"
				, "InterviewStartDate":$InterviewStartDate.val()
				, "InterviewEndDate":$InterviewEndDate.val(), }
			$.Ajax4Get("Process/SubjectProc.asp", aryData, "$.setSubjectTimeSample(datas)", "xml", "","", false);
		}
	});

	// 면접교시샘플 가져오기 처리결과
	$.setSubjectTimeSample = function(datas) {
		var $objInterviewtDate = $(datas).find("InterviewtDate");
		var $objList = $(datas).find("List");
		var InterviewDate = "";
		var intNUM = 0;
		var strMSG;

		if ($objInterviewtDate.length != 0 && $objList.length != 0) {
			// 면접평가기간 만큼 LOOP돌기
			$objInterviewtDate.each(function(j) {
				// 면접평가 날짜 설정
				InterviewDate = $(this).text();
				intNUM += j;
				// 면접시간 만큼 LOOP 돌기
				$objList.each(function(i) {
					// 면접교시 추가
					$.SubjectTimeHtml(
						InterviewDate, $(this).find("TimeCode").text(), $(this).find("CheckTime").text()
						, $(this).find("InterviewStartTime").text(),$(this).find("InterviewEndTime").text()
					);

					// 날짜 변경 시 백드라운드 색상 변경
					$(".SubjectTimeSample:last > div").addClass("LineColors_"+ String(G_LineColorsNum + intNUM + 1));
				});
			});

			// 샘플 시간 가져온 후 전역 변수 입수
			G_LineColorsNum = intNUM;
			G_InterviewDateTemp2 = $("#InterviewEndDate").val();
		}
	}

	// 면접교시샘플 가져오기 처리결과 (자바스크립트로 날짜계산) -> 계산된 날짜가 이상하게 나와서 사용안함
	/*
	$.setSubjectTimeSample_old = function(datas) {
		var $objList = $(datas).find("List");	
		var strMSG;

		// 면접평가기간 가져오기
		var aryDate1 = $("#InterviewStartDate").val().split('-');
		var aryDate2 = $("#InterviewEndDate").val().split('-');

		// 날짜 계산
		var InterviewDate = new Date();
		var InterviewStartDate = new Date(aryDate1[0], aryDate1[1], aryDate1[2]);
		var InterviewEndDate = new Date(aryDate2[0], aryDate2[1], aryDate2[2]);
		var InterviewDateDiff = (InterviewEndDate - InterviewStartDate) / (24 * 60 * 60 * 1000);
		var InterviewDateTemp, InterviewDateTemp2, LineColors;

		if ($objList.length != 0) {
			// 면접평가기간 만큼 LOOP돌기
			for (var intNUM = 0; intNUM <= InterviewDateDiff; intNUM++) {
				$objList.each(function(i) {
					// 면접 시작일부터 날짜 증가
					InterviewDate.setDate(InterviewStartDate.getDate() + intNUM);
					//alert(InterviewDate);
					InterviewDateTemp = InterviewDate.getFullYear() +"-"+ $.right("00" + String(InterviewDate.getMonth() + 1), 2) +"-"+ $.right("00" + String(InterviewDate.getDate()), 2);
					
					// 면접교시 추가
					$.SubjectTimeHtml(
						InterviewDateTemp, $(this).find("TimeCode").text(), $(this).find("CheckTime").text()
						, $(this).find("InterviewStartTime").text(),$(this).find("InterviewEndTime").text() 
					);

					// 날짜 변경 시 백드라운드 색상 변경
					$(".SubjectTimeSample:last > div").addClass("LineColors_"+ String(G_LineColorsNum + intNUM + 1));
				});
			}

			// 샘플 시간 가져온 후 전역 변수 입수
			G_LineColorsNum = intNUM;
			G_InterviewDateTemp2 = $("#InterviewEndDate").val();
		}
	}
	*/
    
	// 예약가능인원 입력 시 면접교시 관리 대상인원 수정
	$(document).on("blur", "#ApplyTotalNumber", function(){
		$("#ApplyTotalNumberDIV").html($(this).val());
	});

	// 이미 입력된 면접교시 정보의 배정인원 변경 시 편성인원 수정
	$(document).on("blur", "input[name='Quorum']", function(){
		// 면접교시 편성인원 계산 실행
		$.SubjectTimeApplyNumberCount();
	});

	// 면접교시 입력 도우미
	$(document).on("keyup", ".InputSubjectTime", function(event){
		// 현재 eq(순서) 가져오기
		var IndexNum = $(this).index(".InputSubjectTime");

		// 엔터 입력 다음 이동
		if (event.keyCode == 13) {
			// 한깐씩 focus 이동
			$(".InputSubjectTime").eq(IndexNum + 1).focus();
			// 배정인원 엔터 입력 시 입력버튼 클릭
			if (IndexNum == 4) {
				$("#InsertSubjectTime").click();
			}
		}

		// maxlength 다음 이동
		if ($(this).val().length == parseInt($(this).attr("maxlength"))) {
			// Input maxlength 시 한깐씩 focus 이동
			$(".InputSubjectTime").eq(IndexNum + 1).focus();
		}
	});

	// 면접교시 시간 입력 시 : 넣기
	$(document).on("keyup", ".keyupTime", function(event){
		if($(this).val().length == 2) {
			$(this).val($(this).val() + ":");
		}
    });
});
</script>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
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
				<form name="InputForm" id="InputForm" method="post" action="/Process/SubjectProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegSubject">
						<input type="hidden" name="ProcessType" id="ProcessType" value="<%= ProcessType %>">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
					</div>

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
							<% Call SubCodeSelectBox("Division", "모집시기선택", Division, "모집시기를 선택해 주세요.", "", "Division") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학과 *
						</div>
						<div class="col-md-3 col-xs-7">
							<% Call SubCodeSelectBox("Subject", "학과명선택", Subject, "학과명을 선택해 주세요.", "", "Subject") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							평가차수 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="Degree" id="Degree" class="form-control input-sm" alert="평가차수를 입력해 주세요.">
								<option value="">평가차수선택</option>
								<% For intNUM = 1 To 10 %>
								<option value="<%= intNUM %>" <%= setSelected(Degree, intNUM) %>><%= intNUM %>차</option>
								<% Next %>
							</select>
						</div>
					</div>

					<!--
					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							평가명 *
						</div>
						<div class="col-md-10 col-xs-10">
							<input type="text" name="ApplyTitle" id="ApplyTitle" class="form-control input-sm" value="<%= ApplyTitle %>" maxlength="100" alert="평가명을 입력해 주세요.">
						</div>
					</div>
					-->

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							예약가능인원 *
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="ApplyTotalNumber" id="ApplyTotalNumber" class="form-control input-sm KeyTypeNUM" value="<%= ApplyTotalNumber %>" maxlength="5" alert="예약가능인원을 입력해 주세요.">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접예약기간 *
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="ApplyStartDate">
								<input type="text" name="ApplyStartDate" id="ApplyStartDate" class="form-control input-sm" value="<%=ApplyStartDate%>" maxlength="10" placeholder="면접예약 시작일" alert="면접예약 시작일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="ApplyStartTime" id="ApplyStartTime" class="form-control input-sm" value="<%=ApplyStartTime%>" maxlength="5" placeholder="면접예약 시작시간" data-autoclose="true" alert="면접예약 시작시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="ApplyEndDate">
								<input type="text" name="ApplyEndDate" id="ApplyEndDate" class="form-control input-sm"  value="<%=ApplyEndDate%>" maxlength="10" placeholder="면접예약 종료일" alert="면접예약 종료일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="ApplyEndTime" id="ApplyEndTime" class="form-control input-sm" value="<%=ApplyEndTime%>" maxlength="5" placeholder="면접예약 종료시간" data-autoclose="true" alert="면접예약 종료시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
						<!--
						<div class="col-md-8 col-xs-10 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="ApplyStartDate">
								<input type="text" name="ApplyStartDate" id="ApplyStartDate" class="form-control input-sm" value="<%=ApplyStartDate%>" maxlength="10" placeholder="면접예약 시작일" alert="면접예약 시작일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
							&nbsp;
							<div class="input-group">
								<input type="text" name="ApplyStartTime" id="ApplyStartTime" class="form-control input-sm" value="<%=ApplyStartTime%>" maxlength="5" placeholder="면접예약 시작시간" data-autoclose="true" alert="면접예약 시작시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
							&nbsp;~&nbsp;
							<div class="input-group viewCalendarBtn" Obj="ApplyEndDate">
								<input type="text" name="ApplyEndDate" id="ApplyEndDate" class="form-control input-sm"  value="<%=ApplyEndDate%>" maxlength="10" placeholder="면접예약 종료일" alert="면접예약 종료일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
							&nbsp;
							<div class="input-group">
								<input type="text" name="ApplyEndTime" id="ApplyEndTime" class="form-control input-sm" value="<%=ApplyEndTime%>" maxlength="5" placeholder="면접예약 종료시간" data-autoclose="true" alert="면접예약 종료시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
						-->
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							평가조관리 *
						</div>
						<div class="col-md-8 col-xs-10" id="SubjectGroupDIV">
							<div>
								<div class="col-xs-5 text-center grid_border" style="background-color:#EAEAEA;"><input type="text" id="GroupName" class="form-control input-sm text-center" maxlength="20" placeholder="조이름"></div>
								<div class="col-xs-5 text-center grid_border grid_border_NoneLeft" style="background-color:#EAEAEA;"><input type="text" id="RoomName" class="form-control input-sm text-center" maxlength="20" placeholder="고사장"></div>
								<div class="col-xs-2 text-center grid_border grid_border_NoneLeft" style="background-color:#EAEAEA;"><span class="btnBasic btnTypeAdd" id="InsertSubjectGroup">입력</span></div>
							</div>

							<div class="SubjectGroupSample displayNone">
								<div class="col-xs-5 grid_borderSubject1"><input type="text" name="GroupName" class="form-control input-sm text-center" value="" maxlength="20" placeholder="조이름"></div>
								<div class="col-xs-5 grid_borderSubject2"><input type="text" name="RoomName" class="form-control input-sm text-center" value="" maxlength="20" placeholder="고사장"></div>
								<div class="col-xs-2 grid_borderSubject2"><span class="btnBasic btnTypeDelete GroupDelete">삭제</span></div>
							</div>
						<%
							'If Not IsNull(AryHash2) Then
							If isArray(AryHash2) Then
								For intNUM = 0 to ubound(AryHash2,1)
						%>
							<div>
								<div class="col-xs-5 grid_borderSubject1"><input type="text" name="GroupName" class="form-control input-sm text-center" value="<%= AryHash2(intNUM).Item("GroupName") %>" maxlength="20" placeholder="조이름"></div>
								<div class="col-xs-5 grid_borderSubject2"><input type="text" name="RoomName" class="form-control input-sm text-center" value="<%= AryHash2(intNUM).Item("RoomName") %>" maxlength="20" placeholder="고사장"></div>
								<div class="col-xs-2 grid_borderSubject2"><span class="btnBasic btnTypeDelete GroupDelete">삭제</span></div>
							</div>
						<%
								Next
							end if
						%>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							수험표출력기간 *
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="ApplyPrintStartDate">
								<input type="text" name="ApplyPrintStartDate" id="ApplyPrintStartDate" class="form-control input-sm" value="<%=ApplyPrintStartDate%>" maxlength="10" placeholder="수험표출력 시작일" alert="수험표출력 시작일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="ApplyPrintStartTime" id="ApplyPrintStartTime" class="form-control input-sm" value="<%=ApplyPrintStartTime%>" maxlength="5" placeholder="출력 시작시간" data-autoclose="true" alert="수험표출력 시작시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="ApplyPrintEndDate">
								<input type="text" name="ApplyPrintEndDate" id="ApplyPrintEndDate" class="form-control input-sm"  value="<%=ApplyPrintEndDate%>" maxlength="10" placeholder="수험표출력 종료일" alert="수험표출력 종료일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="ApplyPrintEndTime" id="ApplyPrintEndTime" class="form-control input-sm" value="<%=ApplyPrintEndTime%>" maxlength="5" placeholder="출력 종료시간" data-autoclose="true" alert="수험표출력 종료시간을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접평가기간 *
						</div>

						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="InterviewStartDate">
								<input type="text" name="InterviewStartDate" id="InterviewStartDate" class="form-control input-sm" value="<%=InterviewStartDate%>" maxlength="10" placeholder="면접평가 시작일" alert="면접평가 시작일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="InterviewStartTimeTemp" id="InterviewStartTimeTemp" class="form-control input-sm" value="00:05" maxlength="5" placeholder="면접평가 시작시간" style="background-color: #FFFFFF;" readonly>
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group viewCalendarBtn" Obj="InterviewEndDate">
								<input type="text" name="InterviewEndDate" id="InterviewEndDate" class="form-control input-sm"  value="<%=InterviewEndDate%>" maxlength="10" placeholder="면접평가 종료일" alert="면접평가 종료일을 입력해 주세요.">
								<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
							</div>
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_inLine">
							<div class="input-group">
								<input type="text" name="InterviewEndTimeTemp" id="InterviewEndTimeTemp" class="form-control input-sm" value="23:55" maxlength="5" placeholder="면접평가 종료시간" style="background-color: #FFFFFF;" readonly>
								<span class="input-group-addon"><i class="fa fa-clock-o"></i></span>
							</div>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접평가일수 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="InterviewDays" id="InterviewDays" class="form-control input-sm" alert="면접평가일수를 입력해 주세요.">
								<option value="">면접평가일수선택</option>
								<% For intNUM = 1 To 10 %>
								<option value="<%= intNUM %>" <%= setSelected(InterviewDays, intNUM) %>><%= intNUM %>일</option>
								<% Next %>
							</select>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접교시관리 *
						</div>
						<div class="col-md-10" >
							<div>
								<span class="btnBasic btnTypeEdit" id="btnSubjectTimeEdit">면접교시 관리</span>
								<span class="btnBasic btnTypeSearch" id="btnSubjectTimeSample">면접교시 샘플 가져오기</span>
							</div>
							<div class="pad_t5 displayNone displayNoneDIV" id="SubjectTimeDIV">
								<div>
									<!--<div class="col-xs-2 text-center grid_border"><input type="text" id="InterviewDate" class="form-control input-sm text-center" maxlength="10" placeholder="면접일자"></div>-->
									<div class="col-xs-2 text-center grid_border LineColors_12">
										<div class="input-group viewCalendarBtn" Obj="InterviewDate">
											<input type="text" id="InterviewDate" class="form-control input-sm text-center" value="<%=InterviewStartDate%>" maxlength="10" placeholder="면접일자">
											<span class="input-group-addon"><i class="fa fa-calendar"></i></span>
										</div>
									</div>
									<div class="col-xs-1 text-center grid_border grid_border_NoneLeft LineColors_12"><input type="text" id="TimeCode" class="form-control input-sm text-center KeyTypeNUM InputSubjectTime" maxlength="2" placeholder="교시"></div>
									<div class="col-xs-2 text-center grid_border grid_border_NoneLeft LineColors_12"><input type="text" id="CheckTime" class="form-control input-sm text-center KeyTypeNUM keyupTime InputSubjectTime" maxlength="5" placeholder="점검시간"></div>
									<div class="col-xs-2 text-center grid_border grid_border_NoneLeft LineColors_12"><input type="text" id="InterviewStartTime" class="form-control input-sm text-center KeyTypeNUM keyupTime InputSubjectTime" maxlength="5" placeholder="면접시작시간"></div>
									<div class="col-xs-2 text-center grid_border grid_border_NoneLeft LineColors_12"><input type="text" id="InterviewEndTime" class="form-control input-sm text-center KeyTypeNUM keyupTime InputSubjectTime" maxlength="5" placeholder="면접종료시간"></div>
									<div class="col-xs-2 text-center grid_border grid_border_NoneLeft LineColors_12"><input type="text" id="Quorum" class="form-control input-sm text-center KeyTypeNUM InputSubjectTime" maxlength="5" placeholder="배정인원"></div>
									<div class="col-xs-1 text-center grid_border grid_border_NoneLeft LineColors_12"><span class="btnBasic btnTypeAdd NoneIcon" id="InsertSubjectTime">입력</span></div>
								</div>

								<div class="SubjectTimeSample displayNone">
									<div class="col-xs-2 grid_borderSubject1"><input type="text" name="InterviewDate" class="form-control input-sm text-center" value="" maxlength="10" placeholder="면접일자"></div>
									<div class="col-xs-1 grid_borderSubject2"><input type="text" name="TimeCode" class="form-control input-sm text-center" value="" maxlength="2" placeholder="교시"></div>
									<div class="col-xs-2 grid_borderSubject2"><input type="text" name="CheckTime" class="form-control input-sm text-center" value="" maxlength="5" placeholder="점검시간"></div>
									<div class="col-xs-2 grid_borderSubject2"><input type="text" name="InterviewStartTime" class="form-control input-sm text-center" value="" maxlength="5" placeholder="면접시작시간"></div>
									<div class="col-xs-2 grid_borderSubject2"><input type="text" name="InterviewEndTime" class="form-control input-sm text-center" value="" maxlength="5" placeholder="면접종료시간"></div>
									<div class="col-xs-2 grid_borderSubject2"><input type="text" name="Quorum" class="form-control input-sm text-center" value="" maxlength="5" placeholder="배정인원"></div>
									<div class="col-xs-1 grid_borderSubject2"><span class="btnBasic btnTypeDelete TimeDelete NoneIcon">삭제</span></div>
								</div>
							<%
								'If Not IsNull(AryHash3) Then
								If isArray(AryHash3) Then
									intNUM2 = 0
									For intNUM = 0 to ubound(AryHash3,1)
										
										'// 면접교시 날짜별 배경색 지정
										If InterviewDate = AryHash3(intNUM).Item("InterviewDate") Then
											intNUM2 = intNUM2
										Else
											intNUM2 = intNUM2 + 1
										End If
										InterviewDate = AryHash3(intNUM).Item("InterviewDate")
							%>
								<div>
									<div class="col-xs-2 grid_borderSubject1 LineColors_<%= intNUM2 %>"><input type="text" name="InterviewDate" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("InterviewDate") %>" maxlength="10" placeholder="면접일자"></div>
									<div class="col-xs-1 grid_borderSubject2 LineColors_<%= intNUM2 %>"><input type="text" name="TimeCode" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("TimeCode") %>" maxlength="2" placeholder="교시"></div>
									<div class="col-xs-2 grid_borderSubject2 LineColors_<%= intNUM2 %>"><input type="text" name="CheckTime" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("CheckTime") %>" maxlength="5" placeholder="점검시간"></div>
									<div class="col-xs-2 grid_borderSubject2 LineColors_<%= intNUM2 %>"><input type="text" name="InterviewStartTime" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("InterviewStartTime") %>" maxlength="5" placeholder="면접시작시간"></div>
									<div class="col-xs-2 grid_borderSubject2 LineColors_<%= intNUM2 %>"><input type="text" name="InterviewEndTime" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("InterviewEndTime") %>" maxlength="5" placeholder="면접종료시간"></div>
									<div class="col-xs-2 grid_borderSubject2 LineColors_<%= intNUM2 %>"><input type="text" name="Quorum" class="form-control input-sm text-center" value="<%= AryHash3(intNUM).Item("Quorum") %>" maxlength="5" placeholder="배정인원"></div>
									<div class="col-xs-1 grid_borderSubject2 LineColors_<%= intNUM2 %>"><span class="btnBasic btnTypeDelete TimeDelete NoneIcon">삭제</span></div>
								</div>
							<%
										'// 편성인원 집계
										ApplyTotalNumberTemp = ApplyTotalNumberTemp + AryHash3(intNUM).Item("Quorum")
									Next
								end if
							%>
							</div>
							<div class="displayNone displayNoneDIV">
								<div class="col-xs-12" style="padding:1px 0 0 0;"></div>
								<div>
									<div class="col-xs-9 text-center grid_border">편성인원</div>
									<div class="col-xs-3 text-center grid_border grid_border_NoneLeft" id="ApplyTotalNumberTempDIV"><% if IsE(ApplyTotalNumberTemp) Then Response.write "0" Else Response.write ApplyTotalNumberTemp End If %></div>
								</div>
								<div class="col-xs-12" style="padding:1px 0 0 0;"></div>
								<div>
									<div class="col-xs-9 text-center grid_border">대상인원</div>
									<div class="col-xs-3 text-center grid_border grid_border_NoneLeft" id="ApplyTotalNumberDIV"><% if IsE(ApplyTotalNumber) Then Response.write "0" Else Response.write ApplyTotalNumber End If %></div>
								</div>
							</div>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							T셔츠 배부여부 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="TShirtCheck" id="TShirtCheck" class="form-control input-sm" alert="T셔츠 배부여부를 선택하세요.">
								<option value="">상태선택</option>
								<option value="Y" <%= setSelected(TShirtCheck, "Y") %>>배부</option>
								<option value="N" <%= setSelected(TShirtCheck, "N") %>>미배부</option>
							</select>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접대기장소
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("StandByRoom", "면접대기장소선택", StandByRoom, "", "", "StandByRoom") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							면접평가진행인원 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="InterviewMemberCnt" id="InterviewMemberCnt" class="form-control input-sm" alert="상태를 선택하세요.">
								<option value="">상태선택</option>
								<% For intNUM = 1 To 10 %>
								<option value="<%= intNUM %>" <%= setSelected(InterviewMemberCnt, intNUM) %>><%= intNUM %>명</option>
								<% Next %>
							</select>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							기타
						</div>
						<div class="col-md-8 col-md-10">
							<input type="text" name="InterviewEtc" id="InterviewEtc" class="form-control input-sm " value="<%= InterviewEtc %>" maxlength="100">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							상태 *
						</div>
						<div class="col-md-2 col-xs-7">
							<select name="State" id="State" class="form-control input-sm" alert="상태를 선택하세요.">
								<option value="">상태선택</option>
								<option value="Y" <%= setSelected(State, "Y") %>>사용</option>
								<option value="N" <%= setSelected(State, "N") %>>미사용</option>
							</select>
						</div>
					</div>

					( * 는 필수 입력값입니다.)
					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeSave" id="btnSave">저장</span>
							<span class="btnBasic btnTypeCancel" id="btnCancel">취소</span>
						</div>
					</div>

				</form>
			</div>
			<!-- 테이블 -->
		</div>		
	</div>
</div>


<!-- #InClude Virtual = "/Common/Bottom.asp" -->