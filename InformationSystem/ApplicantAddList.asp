<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 5
Dim LeftMenuCode : LeftMenuCode = "ApplicantAdd"
Dim LeftMenuName : LeftMenuName = "Home / 지원자관리 / 지원자 서류체크"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "지원자 서류체크"
Dim LogDivision	: LogDivision = "ApplicantAddList"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, objDB2, SQL, arrParams, arrParams2, aryList, AryHash, strWhere, AryHash2
Dim i, strMSG, intNUM, strTEMP, strRESULT

'검색조건
Dim SearchMYear		: SearchMYear = fnR("SearchMYear", "")
Dim SearchDivision	: SearchDivision = fnR("SearchDivision", "")
Dim SearchSubject	: SearchSubject = fnR("SearchSubject", "")
Dim SearchDivision1	: SearchDivision1 = fnR("SearchDivision1", "")
Dim SearchType		: SearchType = fnR("searchType", "")
Dim SearchText		: SearchText = fnR("searchText", "")
Dim SearchState		: SearchState = fnR("SearchState", "")
Dim StrURL			: StrURL = "/ApplicantAddList.asp"
Dim SearchPart
Dim SearchPart2
Dim SearchPart3
Dim SearchPart4
Dim SearchPart5
Dim SearchPart6

'페이지설정(사이즈는 검색)
Dim PageSize		: PageSize = getIntParameter(FnR("PageSize", 5), 5)
Dim PageNum			: PageNum	= fnR("Page", 1)
Dim PageBlock		: PageBlock	= 10
Dim TotalCount		: TotalCount = 0
Dim PageCount		: PageCount = 0
Dim StartNum		: StartNum = 0
Dim EndNum			: EndNum = 0

'DBOpen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

'생기부/검정/수능 데이터, 최종 값
Dim AryHash3, AryHash4, AryHash5, AryHash6, AryHash7, AryHash8
Dim CSATCheck,QualCheck, StuRecCheck
Dim StudentRecordDataStr, QualificationDataStr, CSATDataStr, ReslutStr

'모집시기 검색조건 쿼리
if not(IsE(SearchDivision)) And SearchDivision <> "All" then
	strWhere = strWhere & " And a.Division0 = ? "
	Call objDB.sbSetArray("@Division0", adVarchar, adParamInput, 50, SearchDivision)
end If

'학과 검색조건 쿼리
if not(IsE(SearchSubject)) And SearchSubject <> "All" then
	strWhere = strWhere & " And a.Subject = ? "
	Call objDB.sbSetArray("@Subject", adVarchar, adParamInput, 50, SearchSubject)
end if

'학생 검생조건 쿼리
if (not(IsE(SearchText))) then
	if SearchType = "1" then
		SearchPart = "StudentNumber"
	elseif SearchType = "2" then
		 SearchPart = "StudentNameKor"
	Else
		SearchType = "1"
		SearchPart = "StudentNumber"
	end if
	
	strWhere = strWhere & " And a."& SearchPart &" like '%' + ? + '%' "
	Call objDB.sbSetArray("@SearchText", adVarchar, adParamInput, 255, SearchText)
end If

'리스트 쿼리
SQL = ""
SQL = SQL & vbCrLf & "SELECT   "
SQL = SQL & vbCrLf & " e.StudentNumber as QualCheck, d.StudentNumber as CSATCheck, c.StudentNumber as StuRecCheck  "
SQL = SQL & vbCrLf & " , a.IDX, a.Myear, a.Division0, a.StudentNumber, a.StudentNameKor, a.StudentNameUsa, a.StudentNameChi  "
SQL = SQL & vbCrLf & " , a.Citizen1, a.Citizen2, a.Sex, a.HighCode, a.HighSubject, a.HighGraduationYear, a.HighGraduationDivision  "
SQL = SQL & vbCrLf & " , a.QualificationAreaCode, a.QualificationYear, a.Subject, a.Semester  "
SQL = SQL & vbCrLf & " , a.UniversityName, a.AugScore, a.PerfectScore, a.Credit  "
SQL = SQL & vbCrLf & " , a.Division1, a.Division2, a.Division3, a.Division4, a.HighDivision  "
SQL = SQL & vbCrLf & " , a.RefundDivision, a.RefundAccountHolder, a.RefundBankCode, a.RefundAccount  "
SQL = SQL & vbCrLf & " , a.DrawStandard, a.DrawMsg, a.ExtraPoint, a.InterviewerPoint, a.PracticalPoint, a.Tel1, a.Tel2, a.Tel3, a.Email, a.ZipCode, a.Address1, a.Address2  "
SQL = SQL & vbCrLf & " , a.StudentNameAgreement, a.StudentAgreement, a.StudentRecordAgreement, a.QualificationAgreement  "
SQL = SQL & vbCrLf & " , a.CSATAgreement, a.PersonalCollectionAgreement, a.UniquelyAgreement, a.PersonalTrustAgreement, a.PersonalofferAgreement  "
SQL = SQL & vbCrLf & " , a.ReceiptDate, a.ReceiptTime  "
SQL = SQL & vbCrLf & " , a.DocumentaryCheck1, a.DocumentaryCheck2, a.DocumentaryCheck3, a.DocumentaryCheck4 "
SQL = SQL & vbCrLf & " , a.DocumentaryCheck5, a.DocumentaryCheck6, a.DocumentaryCheck7, a.DocumentaryCheck8 "
SQL = SQL & vbCrLf & " , a.Document, a.DocumentMsg, a.Qualification, a.StudentRecordData, a.QualificationData, a.CSATData, a.Reslut "
SQL = SQL & vbCrLf & " , a.DocumentaryCheck21, a.DocumentaryCheck22, a.DocumentaryCheck23, a.DocumentaryCheck24 "
SQL = SQL & vbCrLf & " , a.INPT_USID, a.INPT_DATE, a.INPT_ADDR, a.UPDT_USID, a.UPDT_DATE, a.UPDT_ADDR, a.InsertTime  "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division0', a.Division0) AS DivisionName  "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName  "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name  "
SQL = SQL & vbCrLf & "	, dbo.getSubCodeName('HignSchoolDivision', a.HighDivision) AS HighDivisionName  "
SQL = SQL & vbCrLf & " , b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio "
SQL = SQL & vbCrLf & " , b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5 "
SQL = SQL & vbCrLf & " , b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5 "
SQL = SQL & vbCrLf & " , b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5 "
SQL = SQL & vbCrLf & " , b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5 "
SQL = SQL & vbCrLf & " , b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5 "
SQL = SQL & vbCrLf & "FROM ApplicationTable AS a "
SQL = SQL & vbCrLf & "	left outer join AppraisalTable AS b "
SQL = SQL & vbCrLf & "		on (a.Myear = b.Myear and a.SubjectCode = b.SubjectCode) "
SQL = SQL & vbCrLf & "	left outer join (	select b.StudentRecordRatio, a.StudentNumber, c.EXAM_NUMB as StdentRecord1, d.EXAM_NUMB as StdentRecord2, e.EXAM_NUMB as StdentRecord3, f.EXAM_NUMB as StdentRecord4  "
SQL = SQL & vbCrLf & "						FROM ApplicationTable AS a   "
SQL = SQL & vbCrLf & "							left outer join AppraisalTable AS b   "
SQL = SQL & vbCrLf & "								on (a.Myear = b.Myear and a.SubjectCode = b.SubjectCode)   "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSI215 b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB)) c  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = c.EXAM_NUMB  "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSI213 b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB) group by b.EXAM_NUMB) d  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = d.EXAM_NUMB  "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSI217 b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB) group by b.EXAM_NUMB) e  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = e.EXAM_NUMB  "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSI237 b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB) group by b.EXAM_NUMB) f  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = f.EXAM_NUMB  "
SQL = SQL & vbCrLf & "						WHERE 1 = 1    "
SQL = SQL & vbCrLf & "						and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0')  "
SQL = SQL & vbCrLf & "						and (c.EXAM_NUMB is not null and d.EXAM_NUMB is not null and e.EXAM_NUMB is not null and f.EXAM_NUMB is not null)) AS c   "
SQL = SQL & vbCrLf & "		on (a.StudentNumber = c.StudentNumber)   "
SQL = SQL & vbCrLf & "	left outer join (	select b.CSATRatio, a.StudentNumber, c.EXAM_NUMB as StdentRecord  "
SQL = SQL & vbCrLf & "						FROM ApplicationTable AS a   "
SQL = SQL & vbCrLf & "							left outer join AppraisalTable AS b   "
SQL = SQL & vbCrLf & "								on (a.Myear = b.Myear and a.SubjectCode = b.SubjectCode)   "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSICSAT b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB)) c  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = c.EXAM_NUMB  "
SQL = SQL & vbCrLf & "						WHERE 1 = 1    "
SQL = SQL & vbCrLf & "						and (b.CSATRatio is not null Or b.CSATRatio <> '' Or b.CSATRatio <> '0')  "
SQL = SQL & vbCrLf & "						and c.EXAM_NUMB is not null ) AS d  "
SQL = SQL & vbCrLf & "		on (a.StudentNumber = d.StudentNumber)   "
SQL = SQL & vbCrLf & "	left outer join (	select b.StudentRecordRatio,a.Qualification, a.StudentNumber, c.EXAM_NUMB as StdentRecord  "
SQL = SQL & vbCrLf & "						FROM ApplicationTable AS a   "
SQL = SQL & vbCrLf & "							left outer join AppraisalTable AS b   "
SQL = SQL & vbCrLf & "								on (a.Myear = b.Myear and a.SubjectCode = b.SubjectCode)   "
SQL = SQL & vbCrLf & "							full outer join  ( select b.EXAM_NUMB from ApplicationTable a join IPSI212 b on (a.Myear = b.SCHL_YEAR and a.StudentNumber = b.EXAM_NUMB) group by b.EXAM_NUMB) c  "
SQL = SQL & vbCrLf & "								on a.StudentNumber = c.EXAM_NUMB  "
SQL = SQL & vbCrLf & "						WHERE 1 = 1    "
SQL = SQL & vbCrLf & "						and (b.StudentRecordRatio is not null Or b.StudentRecordRatio <> '' Or b.StudentRecordRatio <> '0')  "
SQL = SQL & vbCrLf & "						and c.EXAM_NUMB is not null   "
SQL = SQL & vbCrLf & "						and a.Qualification = '1') AS e  "
SQL = SQL & vbCrLf & "		on (a.StudentNumber = e.StudentNumber)   "
SQL = SQL & vbCrLf & "WHERE 1 = 1  "
SQL = SQL & vbCrLf & strWhere
SQL = SQL & vbCrLf & "ORDER BY a.IDX DESC; "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
'aryList = objDB.fnExecSQLGetRows(SQL, arrParams)
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'개인정보가 있는 입학원서는 조회도 기록함
strLogMSG = "지원자관리  > " & SessionUserID  &"가/이 지원자 리스트를 조회 했습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

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

'---------환경설정 - 지원자 서류체크 페이지 사용여부 체크---------
'DBOpen2
Set objDB2 = New clsDBHelper
objDB2.strConnectionString = strDBConnString
objDB2.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & "Select ApplicantAddConfirm "
SQL = SQL & vbCrLf & " from ConfigTable "
SQL = SQL & vbCrLf & "WHERE 1 = 1 "
SQL = SQL & vbCrLf & "	AND State = 'Y' "
SQL = SQL & vbCrLf & "ORDER BY IDX DESC; "

'objDB2.blnDebug = TRUE
arrParams2 = objDB2.fnGetArray
AryHash2 = objDB2.fnExecSQLGetHashMap(SQL, arrParams2)

Set objDB2 = Nothing
%>


<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">

		<% 
		If AryHash2(0).Item("ApplicantAddConfirm") = "N" Then
		%>
			<div class="ibox-content" style="padding:350px;">
				<div style="text-align: center;">
					<h2>지원자 서류체크 페이지 사용이 중단되었습니다. 관리자에게 문의하세요.</h2>
				</div>
			</div>
		<%
		Else
		%>
			
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
								학생
							</div>
							<div class="col-md-2 col-xs-2 grid_sub_title">
								<select name="searchType" id="searchType" class="form-control input-sm">
									<option value="">구분</option>
									<option value="1" <%= setSelected(searchType, "1") %>>수험번호</option>
									<option value="2" <%= setSelected(searchType, "2") %>>이름(한글)</option>
								</select>
							</div>
							<div class="col-md-2 col-xs-2">
								<input type="text" name="searchText" id="searchText" value="<%= SearchText %>" class="form-control input-sm"/>
							</div>
						</div>
						<div class="pad_t10 pad_r10 text-right">
							<span class="btnBasic btnSubmit">지원자 조회</span>
						</div>
					<!--</form>-->
				</div>
			</div>
			<!-- 검색조건 끝 -->

			<div class="pad_t10"></div>

			<!-- 테이블 -->
			<div class="ibox-title">
				<h5>목록 - 전체 <%= TotalCount %>건</h5>
				<div class="ibox-tools">
					<!-- 게시물 갯수 선택 -->
						<div class="col-md-1 col-xs-2" style="float:right;">
							<select name = "PageSize" onChange="SearchForm.submit();">
								<option value="5" <% If PageSize = 5 then response.write "selected" end if%>>5개씩 보기</option>
								<option value="15" <% If PageSize = 15 then response.write "selected" end if%>>15개씩 보기</option>
								<option value="30" <% If PageSize = 30 then response.write "selected" end if%>>30개씩 보기</option>
								<option value="50" <% If PageSize = 50 then response.write "selected" end if%>>50개씩 보기</option>
								<option value="100" <% If PageSize = 100 then response.write "selected" end if%>>100개씩 보기</option>
								<option value="200" <% If PageSize = 200 then response.write "selected" end if%>>200개씩 보기</option>
							</select>
						</div>
					</form>
				</div>
			</div>

			<div class="ibox-content">
				<div class="pad_5">
					<form id="ListForm" method="post">
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<colgroup>
								<col width="5%"></col>
								<col width="5%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
								<col width="20%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
								<col width="10%"></col>
							</colgroup>
							<thead>			                
								<tr>
									<th data-hide="phone">No.</th>
									<th data-hide="phone">년도</th>
									<th>시기</th>
									<th>학과</th>
									<th>전형</th>
									<th data-hide="phone,tablet">수험번호</th>
									<th data-hide="phone">이름</th>
									<th>자격미달여부</th>
								</tr>
							</thead>
							<tbody>
							<%
								Dim StudentRecordStr, QualificationStr, CSATStr, ManualStr
								Dim DrawTemp, DocumentTemp								

								If isArray(AryHash) Then								
									For i = StartNum to EndNum
									' e.StudentNumber as QualCheck, d.StudentNumber as CSATCheck, c.StudentNumber as StuRecCheck 
									' , a.IDX, a.Myear, a.Division0, a.StudentNumber, a.StudentNameKor, a.StudentNameUsa, a.StudentNameChi  
									' , a.Citizen1, a.Citizen2, a.Sex, a.HighCode, a.HighSubject, a.HighGraduationYear, a.HighGraduationDivision  
									' , a.QualificationAreaCode, a.QualificationYear, a.Subject, a.Semester  
									' , a.UniversityName, a.AugScore, a.PerfectScore, a.Credit  
									' , a.Division1, a.Division2, a.Division3, a.Division4, a.HighDivision  
									' , a.RefundDivision, a.RefundAccountHolder, a.RefundBankCode, a.RefundAccount  
									' , a.DrawStandard, a.DrawMsg ,a.ExtraPoint, a.InterviewerPoint, a.PracticalPoint, a.Tel1, a.Tel2, a.Tel3, a.Email, a.ZipCode, a.Address1, a.Address2  
									' , a.StudentNameAgreement, a.StudentAgreement, a.StudentRecordAgreement, a.QualificationAgreement  
									' , a.CSATAgreement, a.PersonalCollectionAgreement, a.UniquelyAgreement, a.PersonalTrustAgreement, a.PersonalofferAgreement 
									' , a.ReceiptDate, a.ReceiptTime  
									' , a.DocumentaryCheck1, a.DocumentaryCheck2, a.DocumentaryCheck3, a.DocumentaryCheck4 "
									' , a.DocumentaryCheck5, a.DocumentaryCheck6, a.DocumentaryCheck7, a.DocumentaryCheck8 "
									' , a.INPT_USID, a.INPT_DATE, a.INPT_ADDR, a.UPDT_USID, a.UPDT_DATE, a.UPDT_ADDR, a.InsertTime  
									'	, dbo.getSubCodeName('Division0', a.Division0) AS DivisionName  
									'	, dbo.getSubCodeName('Subject', a.Subject) AS SubjectName  
									'	, dbo.getSubCodeName('Division1', a.Division1) AS Division1Name  
									'	, dbo.getSubCodeName('HignSchoolDivision', a.HighDivision) AS HighDivisionName  
									' , b.StudentRecordRatio, b.InterviewerRatio, b.PracticalRatio, b.CSATRatio 
									' , b.DrawStandard1, b.DrawStandard2, b.DrawStandard3, b.DrawStandard4, b.DrawStandard5 
									' , b.UnqualifiedStandard1, b.UnqualifiedStandard2, b.UnqualifiedStandard3, b.UnqualifiedStandard4, b.UnqualifiedStandard5 
									' , b.ExtraPoint1, b.ExtraPoint2, b.ExtraPoint3, b.ExtraPoint4, b.ExtraPoint5 
									' , b.Scholarship1, b.Scholarship2, b.Scholarship3, b.Scholarship4, b.Scholarship5 
									' , b.DocumentaryEvidence1, b.DocumentaryEvidence2, b.DocumentaryEvidence3, b.DocumentaryEvidence4, b.DocumentaryEvidence5 
									
									'생기부 동의 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)									
									If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
										StudentRecordStr = "-"
									Else
										if AryHash(i).Item("StudentRecordAgreement") = "1" Then
											StudentRecordStr = "Y"
										Else
											StudentRecordStr = "N"
										End If
									End If

									'검정고시 동의 체크(지원한 모집단위(모집시기+학과+전형)에 생기부 비율이 있으면)
									If isnull(AryHash(i).Item("StudentRecordRatio")) Or AryHash(i).Item("StudentRecordRatio") = "" Or AryHash(i).Item("StudentRecordRatio") = "0" Then
										QualificationStr = "-"
									Else
										'검정고시 자면
										If AryHash(i).Item("Qualification") = "1" Then
											if AryHash(i).Item("QualificationAgreement") = "1" Then
												QualificationStr = "Y"
											Else
												QualificationStr = "N"
											End If
										Else
											QualificationStr = "-"
										End If	
									End If

									'수능 동의 체크(지원한 모집단위(모집시기+학과+전형)에 수능 비율이 있으면)
									If isnull(AryHash(i).Item("CSATRatio")) Or AryHash(i).Item("CSATRatio") = "" Or AryHash(i).Item("CSATRatio") = "0" Then
										CSATStr = "-"
									Else
										if AryHash(i).Item("CSATAgreement") = "1" Then
											CSATStr = "Y"
										Else
											CSATStr = "N"
										End If
									End If													

									'DrawMsg의 =를 <br>로 변경
									If Not isnull(AryHash(i).Item("DrawMsg")) Then
										DrawTemp = replace(AryHash(i).Item("DrawMsg"), "=", "<br>") 
									Else
										DrawTemp = "<b>목록이 없습니다.</b>"
									End If
							%>
								<tr class="viewDetail_SetDate_2">
									<td><%= intNUM %></td>
									<td><%= AryHash(i).Item("Myear")  %></td>
									<td><%= AryHash(i).Item("DivisionName") %></td>
									<td><%= AryHash(i).Item("SubjectName") %></td>
									<td><%= AryHash(i).Item("Division1Name") %></td>
									<td><%= AryHash(i).Item("StudentNumber") %></td>
									<td><%= AryHash(i).Item("StudentNameKor") %></td>
									<td>
										<%= AryHash(i).Item("DrawStandard") %>
										<div class="DataField" style="display:none;">
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck1")) %>"					ColumnName="document1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck2")) %>"					ColumnName="document2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck3")) %>"					ColumnName="document3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck4")) %>"					ColumnName="document4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck5")) %>"					ColumnName="document5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck6")) %>"					ColumnName="document6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck7")) %>"					ColumnName="document7"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck8")) %>"					ColumnName="document8"></li>
											<li Columnvalue="<%= DrawTemp %>"													ColumnName="DrawMsg"></li>

											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck21")) %>"				ColumnName="document21"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck22")) %>"				ColumnName="document22"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck23")) %>"				ColumnName="document23"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck24")) %>"				ColumnName="document24"></li>
											<li Columnvalue="<%= DocumentTemp %>"												ColumnName="DocumentMsg"></li>

											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck1")) %>"					ColumnName="Check1"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck2")) %>"					ColumnName="Check2"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck3")) %>"					ColumnName="Check3"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck4")) %>"					ColumnName="Check4"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck5")) %>"					ColumnName="Check5"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck6")) %>"					ColumnName="Check6"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck7")) %>"					ColumnName="Check7"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck8")) %>"					ColumnName="Check8"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck21")) %>"				ColumnName="Check21"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck22")) %>"				ColumnName="Check22"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck23")) %>"				ColumnName="Check23"></li>
											<li Columnvalue="<%= Trim(AryHash(i).Item("DocumentaryCheck24")) %>"				ColumnName="Check24"></li>
											<li Columnvalue="<%= AryHash(i).Item("Myear") %>"									ColumnName="Myear"></li>
											<li Columnvalue="<%= AryHash(i).Item("StudentNumber") %>"							ColumnName="StudentNumber"></li>

											<li Columnvalue="<%= StudentRecordStr %>"											ColumnName="StudentRecordStr"></li>
											<li Columnvalue="<%= QualificationStr %>"											ColumnName="QualificationStr"></li>
											<li Columnvalue="<%= CSATStr %>"													ColumnName="CSATStr"></li>

										</div>
									</td>
								</tr>
							<%
										intNUM = intNUM - 1
									Next
								Else
							%>
								<tr>
									<td colspan="19" style="height:50px; vertical-align: middle;">검색된 자료가 없습니다.</td>
								</tr>
							<%
								end If
								
							Set objDB = Nothing
							%>
							</tbody>
						</table>
					</form>

					<div class="paging pad_r10">&nbsp;</div>
				</div>
				
				
			</div>
			<!-- 테이블 끝 -->			

			<div class="pad_t10"></div>

			<!-- 상세보기 -->
			<div class="ibox-title">
				<h5>상세정보</h5>
				<div style="float:right;">
					<span class="btnBasic btnTypeEdit" id="BasicDataSet">기본 데이터 설정</span>
					<span class="btnBasic btnTypeSave" id="BasicData1" onclick="BasicDataBtn(1)">1. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData2" onclick="BasicDataBtn(2)">2. 기본 데이터 넣기</span>
					<span class="btnBasic btnTypeSave" id="BasicData3" onclick="BasicDataBtn(3)">3. 기본 데이터 넣기</span>
					<form id="BasicDataBtnFrom" method="post" action="/Process/BasicDataSelect.asp">
						<div style="display:none;">
							<input type="text" name="process" id="process" value="RegApplicantAddBasicDataSet">
							<input type="text" name="BasicDataBtnNum" id="BasicDataBtnNum" value="">
						</div>
					</form>
				</div>
			</div>

			<div class="ibox-content">
				<form name="InputForm" id="InputForm" method="post" action="/Process/ApplicantProc.asp">
					<div style="display:none;">
						<input type="hidden" name="process" id="process" value="RegApplicantAdd">
						<input type="text" name="ProcessType" id="ProcessType" value="Insert">
						<input type="hidden" name="IDX" id="IDX" value="<%=IDX%>">
						<input type="text" name="StudentNumberHidden" id="StudentNumberHidden" value="">
						<input type="text" name="MyearHidden" id="MyearHidden" value="">
						<input type="text" name="Check1" id="Check1" value="">
						<input type="text" name="Check2" id="Check2" value="">
						<input type="text" name="Check3" id="Check3" value="">
						<input type="text" name="Check4" id="Check4" value="">
						<input type="text" name="Check5" id="Check5" value="">
						<input type="text" name="Check6" id="Check6" value="">
						<input type="text" name="Check7" id="Check7" value="">
						<input type="text" name="Check8" id="Check8" value="">		
						<input type="text" name="Check21" id="Check21" value="">
						<input type="text" name="Check22" id="Check22" value="">
						<input type="text" name="Check23" id="Check23" value="">
						<input type="text" name="Check24" id="Check24" value="">	
						<input type="text" name="Myear" id="Myear" value="">
						<input type="text" name="StudentNumber" id="StudentNumber" value="">	
						<input type="text" name="StudentRecordStr" id="StudentRecordStr" value="">
						<input type="text" name="QualificationStr" id="QualificationStr" value="">
						<input type="text" name="CSATStr" id="CSATStr" value="">
					</div>
					
					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 제출서류 &nbsp;&nbsp;&nbsp;&nbsp;
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document1" id="document1" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							정시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document2" id="document2" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							전문대이상
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document8" id="document8" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌1유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document4" id="document4" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌2유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document5" id="document5" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							기초수급자
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document6" id="document6" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							차상위계층
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document7" id="document7" class="form-control input-sm" value="1">
						</div>
					</div>

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 목록
						</div>
						<div class="col-md-11 col-xs-7 grid_sub_title">
							<input type="hidden" name="DrawMsg" id="DrawMsg" class="form-control input-sm">
							<div name="DrawMsgUser" id="DrawMsgUser"></div>
						</div>
					</div>		

					<div class="row show-grid grid_sub_button">
						<div class="col-md-12">
							<span class="btnBasic btnTypeNew" id="btnNew">신 규</span>
							<span class="btnBasic btnTypeSave" id="btnSave">저 장</span>
							<!--<span class="btnBasic btnTypeDelete" id="btnDelete">삭 제</span>-->
						</div>
					</div>

				</form>
			</div>
			<!-- 상세보기 끝 -->

			<!-- 기본 데이터 설정 -->
			<div id="BasicDataSetModal" style="width:100%; margin:5px; display:none;">
				<form name="BasicDataSetForm" id="BasicDataSetForm" method="post" action="/Process/BasicDataProc.asp">
				<input type="hidden" name="BasicDataSetprocess" value="RegApplicantAddBasicDataSet">
				<input type="hidden" name="BasicDataSetProcessType" id="BasicDataSetProcessType" value="Insert">
				<div class="ibox-content">		
					<!-- 버튼 번호 -->
					<div class="row show-grid " style="text-align:left;">
						<div class="col-md-1 grid_sub_title">
							버튼번호
						</div>
						<div class="col-md-2" style="text-align:left;">
							<% Call SubCodeSelectBox("BasicDataBtn", "버튼번호 선택", "", "", "", "BasicDataBtn") %>
						</div>
					</div>					

					<div class="row show-grid">
						<div class="col-md-1 col-xs-2 grid_sub_title">
							자격미달 제출서류 &nbsp;&nbsp;&nbsp;&nbsp;
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							수시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document1" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							정시서류
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document2" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							전문대이상
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document8" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌1유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document4" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							농어촌2유형
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document5" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							기초수급자
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document6" class="form-control input-sm" value="1">
						</div>
						<div class="col-md-1 col-xs-2 grid_sub_title">
							차상위계층
						</div>
						<div class="col-md-1 col-xs-7 grid_sub_title">
							<input type="checkbox" name="document7" class="form-control input-sm" value="1">
						</div>
					</div>		

				</form>

				<br>
				<div class="row show-grid grid_sub_button" >					
					<div class="col-md-12" >
						<span class="btnBasic btnTypeSave" id="RegBasicDataSet" style="width:80px;">저장</span>
						<span class="btnBasic btnTypeClose SelfCloseDIV" style="width:80px;">취소</span>
					</div>
				</div>
				</div>
			</div>
			<!-- 기본 데이터 설정 -->

		<%
		End If
		%>
		</div>		
	</div>
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->
<script type="text/javascript">
$(function() {
	$(document).ready(function() {
		// 페이징 영역 생성
		$.makePage(<%= PageNum %>, <%= PageBlock %>, <%= PageCount %>, ".paging");
	});

	// tr선택 시 
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
		//동의해야 하는 항목만 선택 가능하도록 설정
		if ($("#StudentRecordStr").val() === '-') {
			$("#StudentRecord").prop("disabled", true);
		}else{
			$("#StudentRecord").prop("disabled", false);
		}
		if ($("#QualificationStr").val() === '-') {
			$("#Qualification").prop("disabled", true);
		}else{
			$("#Qualification").prop("disabled", false);
		}
		if ($("#CSATStr").val() === '-') {
			$("#CSAT").prop("disabled", true);
		}else{
			$("#CSAT").prop("disabled", false);
		}
		//서류체크를 해야하는 자격미달 체크박스만 선택 가능하도록 설정
		for (var i=1; i<25; i++) {
			if ($("#Check" + i).val() == "0" || $("#Check" + i).val() == "1" || $("#Check" + i).val() == "10"  ) {
				$("#document" + i ).prop("disabled", true);
			}else if($("#Check" + i).val() != "0") {
				$("#document" + i ).prop("disabled", false);
			}
		}	
		//서류체크를 해야하는 자격미달목록 출력
		$("#DrawMsgUser").html($("#DrawMsg").val());
	});

	// 기본 데이터 설정(모달) 오픈
	$("#BasicDataSet").click(function() {
		$("#BasicDataSetProcessType").val("Update");
		$.openMadal($("#BasicDataSetModal"), "2");
	});

	// 기본 데이터 저장
	$("#RegBasicDataSet").click(function() {
		if (!$.chkInputValue($("select[name=BasicDataBtn]"),		"버튼번호를 선택해 주시기 바랍니다.")) { return; }
		
		if (confirm("기본데이터를 저장 하시겠습니까?")) {
			var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.setBasicData(datas)','complete':'','clear':'','reset':''};
			objOpt["url"] = "/Process/BasicDataProc.asp";
			$.Ajax4Form("#BasicDataSetForm", objOpt);
			$("#BasicDataSetForm").submit();
		}
	});

	// 기본 데이터 저장 결과
	$.setBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
		var strMSG;
			
		if ($objList.find("Result").text() == "Complete") {
			alert("기본 데이터가 저장 되었습니다.");
				
		} else {
			alert("기본 데이터가 저장 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 기본 데이터 넣기
	$.PutBasicData = function(datas) {
		var $objList	= $(datas).find("List");	
			
		if ($objList.find("Result").text() == "Complete") {
			 for (var j=1; j<25; j++) {
				 if ($objList.find("document" + j).text() == "1") {
			 		 $("input[name=document" + j + "]").prop("checked", true).trigger("chosen:updated");
				 }
			 }
		} else {
			alert("기본 데이터 넣기 중 오류가 발생하였습니다.\errorCode : "+ $objList.find("ReturnMSG").text());
			return;
		}
	}

	// 신규
	$("#btnNew").click(function () {
		if (confirm("입력되어 있던 내용이 초기화 됩니다.\n신규로 입력 하시겠습니까?")) {
			$.FormReset($("#InputForm"));
		}
	});

	// 저장
	$("#btnSave").click(function () { 
		if ($.setValidation($("#InputForm"))) {
			if (!$("#DrawMsgUser").html()){
				alert("선택된 지원자가 없습니다.");
				return;
			}

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
//기본 데이터 가져오기
function BasicDataBtn(num) {
	if (confirm(num + "번 기본데이터를 넣으시겠습니까?")) {
		$("#BasicDataBtnNum").val(num);

		var objOpt = {'url':'','param':'','dataType':'xml','before':'','success':'$.PutBasicData(datas)','complete':'','clear':'','reset':''};
		objOpt["url"] = "/Process/BasicDataSelect.asp";
		$.Ajax4Form("#BasicDataBtnFrom", objOpt);
		$("#BasicDataBtnFrom").submit();
	}
}
</script>