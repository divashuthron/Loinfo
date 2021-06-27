<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 2
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

'모집단위 변수
Dim MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix
Dim Division0Name, SubjectName, Division1Name, Division2Name, Division3Name
'등록금 변수
Dim RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11
'기타 변수(예비)
Dim Etc1,Etc2,Etc3,Etc4,Etc5,Etc6,Etc7,Etc8,Etc9,Etc10
'입력, 수정 변수
Dim INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "SELECT "

SQL = SQL & vbCrLf & "	 dbo.getSubCodeName('Division0', Division0) AS Division0Name "
SQL = SQL & vbCrLf & "	,dbo.getSubCodeName('Subject', Subject) AS SubjectName "
SQL = SQL & vbCrLf & "	,dbo.getSubCodeName('Division1', Division1) AS Division1Name "
SQL = SQL & vbCrLf & "	,dbo.getSubCodeName('Division2', Division2) AS Division2Name "
SQL = SQL & vbCrLf & "	,dbo.getSubCodeName('Division3', Division3) AS Division3Name "

SQL = SQL & vbCrLf & "	,IDX,MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix"
SQL = SQL & vbCrLf & "	,RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11"
SQL = SQL & vbCrLf & "	,INPT_USID,INPT_DATE,INPT_ADDR,UPDT_USID,UPDT_DATE,UPDT_ADDR,InsertTime"
SQL = SQL & vbCrLf & "FROM SubjectTable " 
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & "	AND IDX = ?; "

Call objDB.sbSetArray("@IDX", adInteger, adParamInput, 0, IDX)

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'내용이 있으면 update 및 내용 가져오기
If Not IsNull(AryHash) then
	ProcessType				= "SubjectUpdate"
	MYear					= AryHash(0).Item("MYear")
	SubjectCode				= AryHash(0).Item("SubjectCode")
	Division0  				= AryHash(0).Item("Division0")
	Subject    				= AryHash(0).Item("Subject")
	Division1  				= AryHash(0).Item("Division1")
	Division2  				= AryHash(0).Item("Division2")
	Division3  				= AryHash(0).Item("Division3")
	Quorum     				= AryHash(0).Item("Quorum")
	QuorumFix  				= AryHash(0).Item("QuorumFix")
	RF1        				= AryHash(0).Item("RF1")
	RF2        				= AryHash(0).Item("RF2")
	RF3        				= AryHash(0).Item("RF3")
	RF4        				= AryHash(0).Item("RF4")
	RF5        				= AryHash(0).Item("RF5")
	RF6        				= AryHash(0).Item("RF6")
	RF7        				= AryHash(0).Item("RF7")
	RF8        				= AryHash(0).Item("RF8")
	RF9        				= AryHash(0).Item("RF9")
	RF10       				= AryHash(0).Item("RF10")
	RF11       				= AryHash(0).Item("RF11")
	Etc1       				= AryHash(0).Item("Etc1")
	Etc2       				= AryHash(0).Item("Etc2")
	Etc3       				= AryHash(0).Item("Etc3")
	Etc4       				= AryHash(0).Item("Etc4")
	Etc5       				= AryHash(0).Item("Etc5")
	Etc6       				= AryHash(0).Item("Etc6")
	Etc7       				= AryHash(0).Item("Etc7")
	Etc8       				= AryHash(0).Item("Etc8")
	Etc9       				= AryHash(0).Item("Etc9")
	Etc10      				= AryHash(0).Item("Etc10")
	INPT_USID  				= AryHash(0).Item("INPT_USID")
	INPT_DATE  				= AryHash(0).Item("INPT_DATE")
	INPT_ADDR  				= AryHash(0).Item("INPT_ADDR")
	UPDT_USID  				= AryHash(0).Item("UPDT_USID")
	UPDT_DATE  				= AryHash(0).Item("UPDT_DATE")
	UPDT_ADDR  				= AryHash(0).Item("UPDT_ADDR")
	InsertTime 				= AryHash(0).Item("InsertTime")
	Division0Name			= AryHash(0).Item("Division0Name")
	SubjectName				= AryHash(0).Item("SubjectName")
	Division1Name			= AryHash(0).Item("Division1Name")
	Division2Name			= AryHash(0).Item("Division2Name")
	Division3Name			= AryHash(0).Item("Division3Name")
Else
	ProcessType = "SubjectInsert"
End if

Set objDB	= Nothing

'response.write "SessionUserID: " & SessionDivision & "<br>"
'response.End
%>


<script type="text/javascript">
$(function() {
	// 저장
	$("#btnSave").click(function() {
		// 폼검사
		if ($.setValidation($("#InputForm"))) {
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
	//$("#btnDelete").click(function() {
	//});
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
						<input type="hidden" name="Division0Name" id="Division0Name" value="<%=Division0Name%>">
						<input type="hidden" name="SubjectName" id="SubjectName" value="<%=SubjectName%>">
						<input type="hidden" name="Division1Name" id="Division1Name" value="<%=Division1Name%>">
						<input type="hidden" name="Division2Name" id="Division2Name" value="<%=Division2Name%>">
						<input type="hidden" name="Division3Name" id="Division3Name" value="<%=Division3Name%>">
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
							모집코드
						</div>
						<div class="col-md-2 col-xs-7">
							<input type="text" name="SubjectCode" class="form-control input-sm KeyTypeNUM" maxlength="10" value="<%=SubjectCode%>" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							모집시기 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division0", "모집시기선택", Division0, "모집시기를 선택해 주세요.", "", "Division0") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							학과명 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Subject", "학과명선택", Subject, "학과명을 선택해 주세요.", "", "Subject") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							구분1 *
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division1", "구분1", Division1, "구분1을 선택해 주세요.", "", "Division1") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							구분2 
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division2", "구분2", Division2, "", "", "Division2") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							구분3 
						</div>
						<div class="col-md-2 col-xs-7">
							<% Call SubCodeSelectBox("Division3", "구분3", Division3, "", "", "Division3") %>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							입학정원 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="QuorumFix" class="form-control input-sm KeyTypeNUM" value="<%=QuorumFix%>" maxlength="3" alert="입학정원을 입력해 주세요.">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							모집인원 *
						</div>
						<div class="col-md-2 col-xs-3">
							<input type="text" name="Quorum" class="form-control input-sm KeyTypeNUM" value="<%=Quorum%>" maxlength="3" alert="모집인원을 입력해 주세요.">
						</div>
					</div>

					<div class="row show-grid" style="text-align:center;">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							입학금
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수업료
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							등록금소계
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							학생회비
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							OT비
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							잡부금소계
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							감면액
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							기납입액
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							실납입액
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							예치금
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title">
							총계
						</div>
					</div>
					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF1" name="RF1" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF1%>" maxlength="8" alert="입학금을 입력해 주세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF2" name="RF2" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF2%>" maxlength="8" alert="수업료를 입력해 주세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF3" name="RF3" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF3%>" maxlength="8" readonly>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF4" name="RF4" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF4%>" maxlength="8" alert="학생회비를 입력해 주세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF7" name="RF7" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF7%>" maxlength="8" alert="OT비를 입력해 주세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF10" name="RF10" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF10%>" maxlength="8" readonly>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF8" name="RF8" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF8%>" maxlength="8" alert="감면액을 입력해 주세요.">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF6" name="RF6" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF6%>" maxlength="8" readonly>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF5" name="RF5" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF5%>" maxlength="8" readonly>
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							<input type="text" id="RF9" name="RF9" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF9%>" maxlength="8" alert="예치금을 입력해 주세요.">
						</div>
						<div class="col-md-2 col-xs-2 grid_sub_title">
							<input type="text" id="RF11" name="RF11" class="form-control input-sm KeyTypeNUM input-money" value="<%=RF11%>" maxlength="8" readonly>
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-2 col-xs-2 grid_sub_title">
							입력시각
						</div>
						<div class="col-md-2 col-xs-7">
							<%=InsertTime%>
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
<script>
//=== 콤마 처리 ====================
$(document).ready(function(){
	for(var i = 1 ; 11 ; i++){
		$('#RF'+i).val($.commaSplit($('#RF'+i).val()));
	}
});
</script>