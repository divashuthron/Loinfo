<%@  codepage="65001" language="VBScript" %>
<%
Dim TopMenuSeq : TopMenuSeq = 100
Dim LeftMenuCode : LeftMenuCode = ""
Dim LeftMenuName : LeftMenuName = "Home / 계산 테스트"
Dim LeftMenuNameDetail : LeftMenuNameDetail = "계산 테스트"
%>
<!-- #include virtual="/Common/Header.asp" -->
<%
Dim objDB, SQL, arrParams, aryList, AryHash, strWhere, AryHash2, arrParams2

Dim StrURL			: StrURL = "/.asp"

'DBopen
Set objDB = New clsDBHelper
objDB.strConnectionString = strDBConnString
objDB.sbConnectDB



%>

<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			<div class="pad_t10"></div>

			<div class="ibox-title">				

				<%
				Dim Semester1, Semester2

				SQL = ""
				SQL = SQL & vbCrLf & " Select StudentNumber, Semester  "
				SQL = SQL & vbCrLf & " from ApplicationTable "
				SQL = SQL & vbCrLf & " WHERE 1 = 1  "
				SQL = SQL & vbCrLf & " And StudentNumber = '2120083' "

				'objDB.blnDebug = TRUE
				arrParams = objDB.fnGetArray
				AryHash = objDB.fnExecSQLGetHashMap(SQL, arrParams)













				'================ 1. 반영학기 체크 =========================

				Select case AryHash(0).Item("Semester")
					Case "1" 
						Semester1 = "1"
						Semester2 = "1"
					Case "2" 
						Semester1 = "1"
						Semester2 = "2"
					Case "3" 
						Semester1 = "2"
						Semester2 = "1"
					Case "4" 
						Semester1 = "2"
						Semester2 = "2"
					Case "5" 
						Semester1 = "3"
						Semester2 = "1"
				End Select


				'=============== 2. 반영학기 이용하여 생기부/교과성적에서 해당학기의 성적 뽑기 ===============
						
						SQL = ""
						SQL = SQL & vbCrLf & " Select CORS_NAME, ADPT_AVRG, RANK_GRAD, CMPT_UNIT, STDD_DEVI, ORGL_SCOR  "
						SQL = SQL & vbCrLf & " from IPSI213  "
						SQL = SQL & vbCrLf & " WHERE 1 = 1  "
						SQL = SQL & vbCrLf & " And Exam_Numb =   " & AryHash(0).Item("StudentNumber") 
						SQL = SQL & vbCrLf & " And STDT_YEAR =   " & Semester1
						SQL = SQL & vbCrLf & " And SCHL_TERM =   " & Semester2

						'Call objDB.sbSetArray("@Exam_Numb", adVarchar, adParamInput, 4, AryHash(0).Item("StudentNumber"))
						'Call objDB.sbSetArray("@STDT_YEAR", adVarchar, adParamInput, 4, Semester1)
						'Call objDB.sbSetArray("@SCHL_TERM", adVarchar, adParamInput, 4, Semester2)

						'objDB.blnDebug = TRUE
						arrParams2 = objDB.fnGetArray
						AryHash2 = objDB.fnExecSQLGetHashMap(SQL, arrParams2)

										
						Dim ScoreDim, i, totScore, Score, count, RankingGrad, totRankingGrad, Complete, totComplete, ChoiceRankingGrad, AugScore
						Dim Deviation, ChangeRankingGrad, totComplete2
						count = 1
						For i = 0 To ubound(AryHash2,1)
							If Not(isnull(AryHash2(i).Item("ORGL_SCOR"))) Then 
								Score = Cint(AryHash2(i).Item("ADPT_AVRG"))
								AugScore = Cint(AryHash2(i).Item("ORGL_SCOR"))
								RankingGrad = Cint(AryHash2(i).Item("RANK_GRAD"))
								Complete = Cint(AryHash2(i).Item("CMPT_UNIT"))
								Deviation = CInt(AryHash2(i).Item("STDD_DEVI"))
								'Response.Write(AryHash2(i).Item("CORS_NAME") & " : " & Score) 
								'Response.Write("<br>") 								
								'Response.Write("석차등급 : " & RankingGrad) 
								'Response.Write("<br>") 
								'Response.Write("이수학점 : " & Complete) 
								'Response.Write("<br>") 
								'Response.Write("원점수 : " & Score) 
								'Response.Write("<br>") 
								'Response.Write("평균점수 : " & AugScore) 
								'Response.Write("<br>") 
								'Response.Write("표준편차 : " & Deviation) 
								'Response.Write("<br>") 
								'Response.Write("<br>") 
								totScore =  totScore + Score								
								totRankingGrad = totRankingGrad + RankingGrad
								If isnull(Complete) Then
									Complete = 5
								End If
								totComplete = totComplete + Complete
								If isnull(Complete) Then
									Complete = 1
								End If
								totComplete2 = totComplete2 + Complete
								ChangeRankingGrad = ChangeRankingGrad + (RankingGrad * Complete)
								ChoiceRankingGrad = ChoiceRankingGrad + (RankingGrad * Complete)	
								Deviation = (Score - AugScore) / Deviation
								Response.write("교과별 환산등급 정규분포 " & FormatNumber(Deviation,5))
								Response.Write("<br>") 



								count = count + 1
							End If
						Next 
						
						ScoreDim = "Dim AugMyScore : AugMyScore = " & totScore & "/" & count
						execute(ScoreDim) 
						ScoreDim = "Dim AugRankingGrad : AugRankingGrad = " & totRankingGrad & "/" & count
						execute(ScoreDim)
						ScoreDim = "Dim AugComplete : AugComplete = " & totComplete2 & "/" & count
						execute(ScoreDim) 						
						Response.Write("<br>") 
						Response.Write("합계점수 : " & totScore) 
						Response.Write("<br>") 
						Response.Write("평균점수 : " & AugMyScore) 
						Response.Write("<br>") 
						Response.Write("총석차등급 : " & totRankingGrad) 
						Response.Write("<br>") 
						Response.Write("평균석차등급 : " & FormatNumber(AugRankingGrad,6)) 
						Response.Write("<br>") 						
						Response.Write("총이수학점 : " & totComplete2) 
						Response.Write("<br>") 
						Response.Write("평균이수학점 : " & AugComplete) 
						Response.Write("<br>") 
						Response.Write("반영학기 : " & Semester1 & "-" & Semester2 & "학기") 
						Response.Write("<br>") 

						Response.Write("<br>") 
						Response.Write("<br>") 
						Response.Write("과목환산 : " & ChangeRankingGrad) 
						Response.Write("<br>") 

						ScoreDim = "Dim AugChoiceRankingGrad : AugChoiceRankingGrad = " & ChoiceRankingGrad & " / " & totComplete
						execute(ScoreDim) 
						Response.Write("<br>") 
						Response.Write("<br>") 
						Response.Write("선택한학기 등급*이수단위 합 : " & ChoiceRankingGrad) 
						Response.Write("<br>") 
						AugChoiceRankingGrad = FormatNumber(AugChoiceRankingGrad,6)
						Response.Write("선택한학기 평균등급(6자리까지) : " & AugChoiceRankingGrad) 
						Response.Write("<br>") 

						AugChoiceRankingGrad = FormatNumber(CDbl(AugChoiceRankingGrad),5)
						Response.Write("선택한학기 평균등급(5자리로 반올림) : " & AugChoiceRankingGrad) 
						Response.Write("<br>") 

						Response.Write("<br>") 
						ScoreDim = "Dim Point : Point = 312 + ((9 - " & AugChoiceRankingGrad & ") * 11)"
						execute(ScoreDim) 
						Response.Write("<br>")
						Response.Write("1. 2008년 2월 졸업자 ~ 2020년 2월 졸업예정자")
						Response.Write("<br>")
						Response.Write("수시환산점수 : " & Point) 

						ScoreDim = "Dim Point : Point = 236 + ((9 - " & AugChoiceRankingGrad & ") * 11)"
						execute(ScoreDim) 
						Response.Write("<br>")
						Response.Write("정시환산점수 : " & Point) 

						Response.Write("<br>")
						Response.Write("<br>")
						Response.Write("2. 1998년 2월 졸업자 ~ 2007년 2월 졸업자")
						Response.Write("<br>")


						
						








				Set objDB = Nothing
				%>
	
			</div>
			<!-- 공식입력란 끝-->	
			<!-- 테이블 -->
			<div class="pad_t10"></div>
		</div>		
	</div>
</div>
<!-- #InClude Virtual = "/Common/Bottom.asp" -->