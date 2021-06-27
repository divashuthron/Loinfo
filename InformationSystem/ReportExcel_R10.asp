<%@  codepage="65001" language="VBScript" %>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
Response.Buffer = False

Dim LogDivision				: LogDivision = "ReoprtExcel_R10"

Dim objDB, SQL, SQL2, arrParams, aryList, aryList2, AryHash, strWhere, AryHashInterviewItem
Dim i, strMSG, intNUM, intListNUM, strTEMP, strRESULT, count

Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""

SQL = SQL & VbCrLf & "select dbo.getSubCodeName('Subject', a.Subject) AS 학과명 "

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '31'),0) as 일반전형면접위주모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '31' and ReceiptDate = '2020-09-23'),0)  as 일반전형면접위주지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '31' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '31'),0),0,1),2)) as 일반전형면접위주경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '16'),0) as 일반고졸업자모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '16' and ReceiptDate = '2020-09-23'),0)  as 일반고졸업자지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '16' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '16'),0),0,1),2)) as 일반고졸업자경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '18'),0) as 특성화고졸업자모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '18' and ReceiptDate = '2020-09-23'),0)  as 특성화고졸업자지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '18' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '18'),0),0,1),2)) as 특성화고졸업자경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '55'),0) as 특기자모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '55' and ReceiptDate = '2020-09-23'),0)  as 특기자지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '55' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '55'),0),0,1),2)) as 특기자경쟁률"
      
SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('31','29','16','18','55')),0) as 정원내소계모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('31','29','16','18','55') and ReceiptDate = '2020-09-23'),0)  as 정원내소계지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('31','29','16','18','55') and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('31','29','16','18','55')),0),0,1),2)) as 정원내소계경쟁률"
SQL = SQL & VbCrLf & "		, isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '08'),0) as 농어촌모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '08' and ReceiptDate = '2020-09-23'),0)  as 농어촌지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '08' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '08'),0),0,1),2)) as 농어촌경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '13'),0) as 기초차상위모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '13' and ReceiptDate = '2020-09-23'),0)  as 기초차상위지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '13' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '13'),0),0,1),2)) as 기초차상위경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '05'),0) as 전문대졸모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '05' and ReceiptDate = '2020-09-23'),0)  as 전문대졸지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '05' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '05'),0),0,1),2)) as 전문대졸경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '07'),0) as 재외국민모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '07' and ReceiptDate = '2020-09-23'),0)  as 재외국민지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '07' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '07'),0),0,1),2)) as 재외국민경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '27'),0) as 전과정해외모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '27' and ReceiptDate = '2020-09-23'),0)  as 전과정해외지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '27' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '27'),0),0,1),2)) as 전과정해외경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '28'),0) as 북한이탈모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '28' and ReceiptDate = '2020-09-23'),0)  as 북한이탈지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '28' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '28'),0),0,1),2)) as 북한이탈경쟁률"

SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '56'),0) as 특성화고졸모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '56' and ReceiptDate = '2020-09-23'),0)  as 특성화고졸지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '56' and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 = '56'),0),0,1),2)) as 특성화고졸경쟁률"

      
SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('08','13','05','07','27','28','56')),0) as 정원외소계모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('08','13','05','07','27','28','56') and ReceiptDate = '2020-09-23'),0)  as 정원외소계지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('08','13','05','07','27','28','56') and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 and Division1 in ('08','13','05','07','27','28','56')),0),0,1),2)) as 정원외소계경쟁률"

      
SQL = SQL & VbCrLf & "      , isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 ),0) as 합계모집인원"
SQL = SQL & VbCrLf & "      , isnull((select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and ReceiptDate = '2020-09-23'),0)  as 합계지원자"
SQL = SQL & VbCrLf & "      , convert(numeric(5, 2), round(isnull(convert(numeric(5, 2), (select count(StudentNumber) from ApplicationTable where Subject = a.Subject and Division0 = a.Division0 and ReceiptDate = '2020-09-23')),0) / replace(isnull((select Sum(Quorum) from SubjectTable where Subject = a.Subject and Division0 = a.Division0 ),0),0,1),2)) as 합계경쟁률"

SQL = SQL & VbCrLf & "		from SubjectTable as a"
SQL = SQL & VbCrLf & "		where 1=1"
SQL = SQL & VbCrLf & "		and a.Division0 = '6'"
SQL = SQL & VbCrLf & "		group by a.Division0, a.Subject"

'///////////////////////////////////////////////////////
'// 대행사별 집계
'///////////////////////////////////////////////////////
SQL2 = ""

SQL2 = SQL2 & VbCrLf & "select 대행사, isnull([2020-09-23],0),isnull([2020-09-24],0),isnull([2020-09-25],0),isnull([2020-09-26],0),isnull([2020-09-27],0) "
SQL2 = SQL2 & VbCrLf & "		,isnull([2020-09-28],0),isnull([2020-09-29],0),isnull([2020-09-30],0),isnull([2020-10-01],0),isnull([2020-10-02],0) "
SQL2 = SQL2 & VbCrLf & "		,isnull([2020-10-03],0),isnull([2020-10-04],0),isnull([2020-10-05],0),isnull([2020-10-06],0),isnull([2020-10-07],0) "
SQL2 = SQL2 & VbCrLf & "		,isnull([2020-10-08],0),isnull([2020-10-09],0),isnull([2020-10-10],0),isnull([2020-10-11],0),isnull([2020-10-12],0) "
SQL2 = SQL2 & VbCrLf & "		,isnull([2020-10-13],0),isnull([2020-10-14],0) "
SQL2 = SQL2 & VbCrLf & "			from ( "
SQL2 = SQL2 & VbCrLf & "				select (case a.InCompary when '6' then '진학사' when '7' then '유웨이' End)  as 대행사 "
SQL2 = SQL2 & VbCrLf & "						, CONVERT(CHAR(10),a.ReceiptDate,23) as 일자 "
SQL2 = SQL2 & VbCrLf & "						,count(a.IDX) as 지원인원 "
SQL2 = SQL2 & VbCrLf & "						from ApplicationTable as a "
SQL2 = SQL2 & VbCrLf & "						group by a.InCompary, CONVERT(CHAR(10),a.ReceiptDate,23) "
SQL2 = SQL2 & VbCrLf & "						) as e "
SQL2 = SQL2 & VbCrLf & "						pivot( "
SQL2 = SQL2 & VbCrLf & "							sum(e.지원인원) for e.일자 in ([2020-09-23],[2020-09-24],[2020-09-25],[2020-09-26],[2020-09-27],[2020-09-28],[2020-09-29],[2020-09-30] "
SQL2 = SQL2 & VbCrLf & "							,[2020-10-01],[2020-10-02],[2020-10-03],[2020-10-04],[2020-10-05],[2020-10-06],[2020-10-07],[2020-10-08] "
SQL2 = SQL2 & VbCrLf & "							,[2020-10-09],[2020-10-10],[2020-10-11],[2020-10-12],[2020-10-13],[2020-10-14]) "
SQL2 = SQL2 & VbCrLf & "			) p "

'objDB.blnDebug = true
'arrParams = objDB.fnGetArray
aryList = objDB.fnExecSQLGetRows(SQL, nothing)
aryList2 = objDB.fnExecSQLGetRows(SQL2, nothing)
'AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

'Response.End

strLogMSG = "입학원서 수동입력 > 수시 1차 원서접수 현황 엑셀파일을 저장하였습니다."
Call ActivityHistory(strLogMSG, LogDivision, SessionUserID)

Set objDB	= Nothing

Dim Filename : Filename = Server.URLEncode(SessionMYear & "학년도_배화여대_수시1차_경쟁률_" & getDateNow(""))&".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition","attachment; filename=" & Filename
%>
<!-- 메인 컨텐츠 -->
<div id="content" class="row">

	<!-- 테이블 -->
	<div class="row">
		<div>
			<div>
				<table border="1">
					<thead>
						<tr style="text-align:center; mso-number-format:'\@';">
							<td rowspan="3">학과명</td>
							<td colspan="12">정원내 모집</td>
							<td rowspan="2" colspan="3">정원내 소계</td>
							<td colspan="21">정원외 모집</td>
							<td rowspan="2" colspan="3">정원외 소계</td>
							<td rowspan="2" colspan="3">합계</td>
						</tr>
						<tr style="text-align:center; mso-number-format:'\@';">
							<td colspan="3">일반전형(면접위주)</td>
							<td colspan="3">일반고전형</td>
							<td colspan="3">특성화고전형</td>
							<td colspan="3">특기자전형</td>
							<td colspan="3">농어촌전형</td>
							<td colspan="3">기초생활수급자및차상위</td>
							<td colspan="3">특성화고졸재직자</td>
							<td colspan="3">전문대학이상졸업자</td>
							<td colspan="3">재외국민및외국인</td>
							<td colspan="3">전 교육과정해외이수자</td>
							<td colspan="3">북한이탈주민</td>
						</tr>
						<tr style="text-align:center; mso-number-format:'\@';">
							<%
							For i = 0 To 13
							%>
							<td>모집인원</td>
							<td>지원자</td>
							<td>경쟁률</td>
							<%
							Next
							%>
						</tr>
					</thead>
					<tbody>
					<%
						if IsArray(aryList) Then
							For i = 0 to ubound(aryList,2)


					%>
						<tr style="text-align:center; mso-number-format:'\@';">
							<td><%= aryList(0, i)%></td>
							<td><%= aryList(1, i)%></td>
							<td><%= aryList(2, i)%></td>
							<td><%= aryList(3, i)%></td>
							<td><%= aryList(4, i)%></td>
							<td><%= aryList(5, i)%></td>
							<td><%= aryList(6, i)%></td>
							<td><%= aryList(7, i)%></td>
							<td><%= aryList(8, i)%></td>
							<td><%= aryList(9, i)%></td>
							<td><%= aryList(10, i)%></td>
							<td><%= aryList(11, i)%></td>
							<td><%= aryList(12, i)%></td>
							<td><%= aryList(13, i)%></td>
							<td><%= aryList(14, i)%></td>
							<td><%= aryList(15, i)%></td>
							<td><%= aryList(16, i)%></td>
							<td><%= aryList(17, i)%></td>
							<td><%= aryList(18, i)%></td>
							<td><%= aryList(19, i)%></td>
							<td><%= aryList(20, i)%></td>
							<td><%= aryList(21, i)%></td>
							<td><%= aryList(22, i)%></td>
							<td><%= aryList(23, i)%></td>
							<td><%= aryList(24, i)%></td>
							<td><%= aryList(25, i)%></td>
							<td><%= aryList(26, i)%></td>
							<td><%= aryList(27, i)%></td>
							<td><%= aryList(28, i)%></td>
							<td><%= aryList(29, i)%></td>
							<td><%= aryList(30, i)%></td>
							<td><%= aryList(31, i)%></td>
							<td><%= aryList(32, i)%></td>
							<td><%= aryList(33, i)%></td>
							<td><%= aryList(34, i)%></td>
							<td><%= aryList(35, i)%></td>
							<td><%= aryList(36, i)%></td>
							<td><%= aryList(37, i)%></td>
							<td><%= aryList(38, i)%></td>
							<td><%= aryList(39, i)%></td>
							<td><%= aryList(40, i)%></td>
							<td><%= aryList(41, i)%></td>
							<td><%= aryList(42, i)%></td>
						</tr>
					<%
							Next
						End If
					%>
						<tr>
							<!-- 구분선 -->
						</tr>
						<tr style="text-align:center; mso-number-format:'mm\/dd';">
							<td>구분</td>
							<td>2020-09-23</td>
							<td>2020-09-24</td>
							<td>2020-09-25</td>
							<td>2020-09-26</td>
							<td>2020-09-27</td>
							<td>2020-09-28</td>
							<td>2020-09-29</td>
							<td>2020-09-30</td>
							<td>2020-10-01</td>
							<td>2020-10-02</td>
							<td>2020-10-03</td>
							<td>2020-10-04</td>
							<td>2020-10-05</td>
							<td>2020-10-06</td>
							<td>2020-10-07</td>
							<td>2020-10-08</td>
							<td>2020-10-09</td>
							<td>2020-10-10</td>
							<td>2020-10-11</td>
							<td>2020-10-12</td>
							<td>2020-10-13</td>
							<td>2020-10-14</td>
						</tr>
					<%
						if IsArray(aryList2) Then
							For i = 0 to ubound(aryList2,2)
					%>
						<tr style="text-align:center; mso-number-format:'\@';">
							<td><%= aryList2(0, i)%></td>
							<td><%= aryList2(1, i)%></td>
							<td><%= aryList2(2, i)%></td>
							<td><%= aryList2(3, i)%></td>
							<td><%= aryList2(4, i)%></td>
							<td><%= aryList2(5, i)%></td>
							<td><%= aryList2(6, i)%></td>
							<td><%= aryList2(7, i)%></td>
							<td><%= aryList2(8, i)%></td>
							<td><%= aryList2(9, i)%></td>
							<td><%= aryList2(10, i)%></td>
							<td><%= aryList2(11, i)%></td>
							<td><%= aryList2(12, i)%></td>
							<td><%= aryList2(13, i)%></td>
							<td><%= aryList2(14, i)%></td>
							<td><%= aryList2(15, i)%></td>
							<td><%= aryList2(16, i)%></td>
							<td><%= aryList2(17, i)%></td>
							<td><%= aryList2(18, i)%></td>
							<td><%= aryList2(19, i)%></td>
							<td><%= aryList2(20, i)%></td>
							<td><%= aryList2(21, i)%></td>
							<td><%= aryList2(22, i)%></td>
						</tr>
					<%
							Next
						End If
					%>
					</tbody>
				</table>
			</div>
		</div>
	</div>
	<!-- 테이블 -->

</div>
<!-- 메인 컨텐츠 -->
