<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 8
Dim LeftMenuCode : LeftMenuCode = "DetailsCharts"
Dim LeftMenuName : LeftMenuName = "Home / 통계관리 / 세부사정현황표"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "세부사정현황표"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%

'데이터 변수
Dim objDB, SQL, arrParams, aryList, AryHash, AryHash2, strWhere

'현황 항목 변수(정원내)
Dim Myear				'년도
Dim Subject				'학과
Dim SubjectCount		'학과전형수
Dim SubjectTemp			'이전학과
Dim Division1			'전형
Dim Quorum2				'가용
Dim RegistRecord		'등록완료인원 (등록완료 - 환불신청자)  : 환불신청하였으나 환불완료하지 않은 인원도 적용
Dim BankRecord			'납부인원 (납부인원 - 등록완료자) : 납부하였으나 등록완료 처리 하지 않은 인원
Dim RefundRecord		'환불신청인원 : 환불신청 하였으나 환불완료처리 하지 않은 인원
Dim Shortfall			'미달인원
Dim AchievementRate		'달성률
Dim Resource			'자원(정원내만 METIS의 자원을 보여준다. 정원외는 데이터 없음)

'현황 항목 변수(정원외)
Dim MyearOut			'년도
Dim SubjectOut			'학과
Dim SubjectCountOut		'학과전형수
Dim SubjectTempOut		'이전학과
Dim Division1Out		'전형
Dim Quorum2Out			'가용
Dim RegistRecordOut		'등록완료인원 (등록완료 - 환불신청자)  : 환불신청하였으나 환불완료하지 않은 인원도 적용
Dim BankRecordOut		'납부인원 (납부인원 - 등록완료자) : 납부하였으나 등록완료 처리 하지 않은 인원
Dim RefundRecordOut		'환불신청인원 : 환불신청 하였으나 환불완료처리 하지 않은 인원
Dim ShortfallOut		'미달인원
Dim AchievementRateOut	'달성률
Dim ResourceOut			'자원(정원내만 METIS의 자원을 보여준다. 정원외는 데이터 없음)

'현황 총계 변수(정원내)
Dim SumQuorum2				:	SumQuorum2 = 0				'가용 합
Dim SumRegistRecord			:	SumRegistRecord = 0			'등록완료인원 합 (등록완료 - 환불신청자)  : 환불신청하였으나 환불완료하지 않은 인원도 적용
Dim SumBankRecord			:	SumBankRecord = 0			'납부인원 합 (납부인원 - 등록완료자) : 납부하였으나 등록완료 처리 하지 않은 인원
Dim SumRefundRecord			:	SumRefundRecord = 0			'환불신청인원 합 : 환불신청 하였으나 환불완료처리 하지 않은 인원
Dim SumShortfall			:	SumShortfall = 0			'미달인원 
Dim AvgAchievementRate		:	AvgAchievementRate = 0		'달성률 평균
Dim SumResource				:	SumResource = 0				'자원 합(정원내만 METIS의 자원을 보여준다. 정원외는 데이터 없음)

'현황 총계 변수(정원외)
Dim SumQuorum2Out			:	SumQuorum2Out = 0			'가용 합
Dim SumRegistRecordOut		:	SumRegistRecordOut = 0		'등록완료인원 합 (등록완료 - 환불신청자)  : 환불신청하였으나 환불완료하지 않은 인원도 적용
Dim SumBankRecordOut		:	SumBankRecordOut = 0		'납부인원 합 (납부인원 - 등록완료자) : 납부하였으나 등록완료 처리 하지 않은 인원
Dim SumRefundRecordOut		:	SumRefundRecordOut = 0		'환불신청인원 합 : 환불신청 하였으나 환불완료처리 하지 않은 인원
Dim SumShortfallOut			:	SumShortfallOut = 0			'미달인원 합
Dim AvgAchievementRateOut	:	AvgAchievementRateOut = 0	'달성률 평균
Dim SumResourceOut			:	SumResourceOut = 0			'자원 합(정원내만 METIS의 자원을 보여준다. 정원외는 데이터 없음)

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB

SQL = ""
SQL = SQL & vbCrLf & " exec XTEB_OSC_Test.dbo.UP_RealtimeStatusReportDetails "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)

SQL = ""
SQL = SQL & vbCrLf & " exec XTEB_OSC_Test.dbo.UP_RealtimeStatusReportOthersDetails "

'objDB.blnDebug = true
arrParams = objDB.fnGetArray
AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams)

Set objDB	= Nothing
%>

<div class="row">
	<!-- 실시간 사정현황(정원내) -->
	<div class="col-lg-6">
		<div class="ibox float-e-margins">

			<div class="pad_t10"></div>
			
			<div class="ibox-title">
				<h3>2020 수시 실시간 사정현황(정원내)</h3>				
				<div style="float:right;">
				</div>
			</div>			

			<div class="ibox-content">
				<div class="pad_5" style="height:700px;">
					<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0">
						<thead>			                
							<tr>
								<th data-hide="phone">년도</th>    
								<th>학과</th>
								<th>전형</th>
								<th>정원</th>  
								<th data-hide="phone">등록</th>   
								<th>납부</th>                           
								<th>환불신청</th>                            
								<th>미달</th>                            
								<th>달성률</th>                            
								<th data-hide="phone,tablet">자원</th>							
							</tr>
						</thead>
						<tbody>
						<%
						If isArray(AryHash) Then
							For i = 0 to ubound(AryHash,1)
								Subject = AryHash(i).Item("Subject")								
								Division1 = AryHash(i).Item("Division1") 
								Quorum2 = AryHash(i).Item("Quorum2")
								RegistRecord = AryHash(i).Item("RegistRecord")
								BankRecord = AryHash(i).Item("BankRecord")
								RefundRecord = AryHash(i).Item("RefundRecord")
								Shortfall = AryHash(i).Item("Shortfall")
								Resource = AryHash(i).Item("Resource")
								SubjectCount = AryHash(i).Item("SubjectCount")

								
								'달성률은 소수점 2자리 까지
								AchievementRate = AryHash(i).Item("AchievementRate")
								AchievementRate = FormatNumber(AchievementRate,2)
								
								'달성률의 따라 텍스트 색깔 구분
								If AchievementRate  = 100 Then
									TextColor = "style='color: Black;'"
								ElseIf AchievementRate > 100 Then
									TextColor = "style='color: Red;'"
								ElseIf AchievementRate >= 95 Then
									TextColor = "style='color: Blue;'"
								ElseIf AchievementRate < 95 And AchievementRate > 0  Then
									TextColor = "style='color: Green;'"
								Else
									TextColor = "style='color: Black;'"
								End if		
						%>

							<tr <%=TextColor%> class="viewDetail" IDX="<%= AryHash(i).Item("IDX") %>">
								<%'새로운 학과일 때만 td 생성
								If Subject <> SubjectTemp Then%>
									<td style="vertical-align:middle;" rowspan=<%=SubjectCount%>><%= AryHash(i).Item("Myear") %></td>								
									<td style="vertical-align:middle;" rowspan=<%=SubjectCount%>><%= Subject %></td>
								<%End If%>
								<td><%= Division1 %></td>
								<td><%= Quorum2 %></td>
								<td><%= RegistRecord %></td>									
								<td><%= BankRecord %></td>
								<td><%= RefundRecord %></td>
								<td><%= Shortfall %></td>
								<td><%= AchievementRate %></td>
								<td><%= Resource %></td>
							</tr>
						<%
								'새로운 학과가 되면 비교학과에 새로운 학과 넣어주기
								If Subject <> SubjectTemp Then
									SubjectTemp = Subject
								End If

								'총합구하기
								SumQuorum2 = SumQuorum2 + Quorum2
								SumRegistRecord = SumRegistRecord + RegistRecord
								SumBankRecord = SumBankRecord + BankRecord
								SumRefundRecord = SumRefundRecord + RefundRecord
								SumShortfall = SumShortfall + Shortfall
								SumResource = SumResource + Resource
							Next
							If SumQuorum2 <> 0 then	'// 0으로 나누면 오버플로 오류 발생됨
								AvgAchievementRate =  (SumRegistRecord + SumBankRecord) / SumQuorum2 * 100
								AvgAchievementRate = FormatNumber(AvgAchievementRate,2)
							End If
						Else
						%>
							<tr>
								<td colspan="10" style="height:50px; vertical-align: middle;">환불서버, 충원서버에 모두 등록되어 있어야 정상 집계됩니다.</td>
							</tr>
						<%
							end if
						%>
							<tr class="viewDetail">
								<td colspan="3">총 계</td>
								<td><%= SumQuorum2 %></td>
								<td><%= SumRegistRecord %></td>									
								<td><%= SumBankRecord %></td>
								<td><%= SumRefundRecord %></td>
								<td><%= SumShortfall %></td>
								<td><%= AvgAchievementRate %></td>
								<td><%= SumResource %></td>
							</tr>
						</tbody>
					</table>
				</div>		
			</div>			

			<div class="pad_t10"></div>

		</div>
	</div>
	<!-- 실시간 사정현황(정원내) 끝 -->

	<!-- 실시간 사정현황(정원외) -->
	<div class="col-lg-6">
		<div class="ibox float-e-margins">

			<div class="pad_t10"></div>
			
			<div class="ibox-title">
				<h3>2020 수시 실시간 사정현황(정원외)</h3>				
				<div style="float:right;">
				</div>
			</div>			

			<div class="ibox-content">
				<div class="pad_5" style="height:700px;">
					<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" >
						<thead>			                
							<tr>
								<th data-hide="phone">년도</th>    
								<th>학과</th>
								<th>전형</th>
								<th>정원</th>  
								<th data-hide="phone">등록</th>   
								<th>납부</th>                           
								<th>환불신청</th>                            
								<th>미달</th>                            
								<th>달성률</th>    
							</tr>
						</thead>
						<tbody>
						<%
						If isArray(AryHash2) Then
							For i = 0 to ubound(AryHash2,1)
								SubjectOut = AryHash2(i).Item("Subject")								
								Division1Out = AryHash2(i).Item("Division1") 
								Quorum2Out = AryHash2(i).Item("Quorum2")
								RegistRecordOut = AryHash2(i).Item("RegistRecord")
								BankRecordOut = AryHash2(i).Item("BankRecord")
								RefundRecordOut = AryHash2(i).Item("RefundRecord")
								ShortfallOut = AryHash2(i).Item("Shortfall")
								SubjectCountOut = AryHash2(i).Item("SubjectCount")

								'달성률은 소수점 2자리 까지
								AchievementRateOut = AryHash2(i).Item("AchievementRate")
								AchievementRateOut = FormatNumber(AchievementRateOut,2)
								
								'달성률의 따라 텍스트 색깔 구분
								If AchievementRateOut  = 100 Then
									TextColor = "style='color: Black;'"
								ElseIf AchievementRateOut > 100 Then
									TextColor = "style='color: Red;'"
								ElseIf AchievementRateOut >= 95 Then
									TextColor = "style='color: Blue;'"
								ElseIf AchievementRateOut < 95 And AchievementRateOut > 0  Then
									TextColor = "style='color: Green;'"
								Else
									TextColor = "style='color: Black;'"
								End if		
						%>

							<tr <%=TextColor%> class="viewDetail" IDX="<%= AryHash2(i).Item("IDX") %>">
								<%'새로운 학과일 때만 td 생성
								If SubjectOut <> SubjectTempOut Then%>
									<td style="vertical-align:middle;" rowspan=<%=SubjectCountOut%>><%= AryHash2(i).Item("Myear") %></td>								
									<td style="vertical-align:middle;" rowspan=<%=SubjectCountOut%>><%= SubjectOut %></td>
								<%End If%>
								<td><%= Division1Out %></td>
								<td>?</td>
								<td><%= RegistRecordOut %></td>									
								<td><%= BankRecordOut %></td>
								<td><%= RefundRecordOut %></td>
								<td>?</td>
								<td>?</td>
							</tr>
						<%
								'새로운 학과가 되면 비교학과에 새로운 학과 넣우주기
								If SubjectOut <> SubjectTempOut Then
									SubjectTempOut = SubjectOut
								End If

								'총합구하기
								SumQuorum2Out = SumQuorum2Out + Quorum2Out
								SumRegistRecordOut = SumRegistRecordOut + RegistRecordOut
								SumBankRecordOut = SumBankRecordOut + BankRecordOut
								SumRefundRecordOut = SumRefundRecordOut + RefundRecordOut
								SumShortfallOut = SumShortfallOut + ShortfallOut
							Next
							If SumQuorum2Out <> 0 then	'// 0으로 나누면 오버플로 오류 발생됨
								AvgAchievementRateOut =  (SumRegistRecordOut + SumBankRecordOut) / SumQuorum2Out * 100
								AvgAchievementRateOut = FormatNumber(AvgAchievementRateOut,2)
							End If
						Else
						%>
							<tr>
								<td colspan="9" style="height:50px; vertical-align: middle;">환불서버, 충원서버에 모두 등록되어 있어야 정상 집계됩니다.</td>
							</tr>
						<%
							end if
						%>
							<tr class="viewDetail">
								<td colspan="3">총 계</td>
								<td>?</td>
								<td><%= SumRegistRecordOut %></td>									
								<td><%= SumBankRecordOut %></td>
								<td><%= SumRefundRecordOut %></td>
								<td>?</td>
								<td>?</td>
							</tr>
						</tbody>
					</table>
				</div>			
			</div>
			<!-- 테이블 -->

			<div class="pad_t10"></div>

		</div>
	</div>
	<!-- 실시간 사정현황(정원외) 끝 -->
</div>

<!-- #InClude Virtual = "/Common/Bottom.asp" -->