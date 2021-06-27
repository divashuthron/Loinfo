SELECT     
	------------------------------------------------------------------------------------
	-- 모든 기준은 등록 완료 시점으로 설정
	-- 수납 (XTEB)	: 납부현황 -> 등록완료처리
	-- 환불 (XTEB)	: 재무승인 -> 환불완효처리
	-- 충원 (METIS)	: 결과관리 -> 데이터내보내기 에서 XTEB 가상계좌가동관리 -> 데이터 가져오기
	------------------------------------------------------------------------------------
	------------------------------------------------------------------------------------
	-- 여주대 요청 필드 ----------------------------------------------------------------
	------------------------------------------------------------------------------------
	A.Myear AS YR     
	, '10' AS TERM_CD
	, LEFT(A.SUBJECTCODE,2) AS OPEN_CD     
    , SUBSTRING(A.SubjectCode,3,2) AS DEPT_CD
    , SUBSTRING(A.SUBJECTCODE,5,2) AS SELEC_CD
    , SUBSTRING(A.SUBJECTCODE,7,2) AS LECT_CD
	, SUBSTRING(A.SUBJECTCODE,9,2) AS  CORS_CD_PRAC
	, A.StudentNumber AS SEAT_NUM, A.StudentName AS NM
	, A.Degree1 AS PASS_DIFF									-- 합격차수
	, A.Citizen1 AS REG_NUM1										-- 주민번호 앞6자리
	-- PASS_DT 합격일자 ---------------------------------------------------------------
	-- 최초합격자 : 데이터 가져온 일자 - StudentTable 참조
	-- 추가합격자 : 충원작업을 통해 예비자에서 추가합격자가 된 일자 - RegistRecord 참조
	, CASE
		WHEN A.Result1 = '6' THEN								-- 합격여부 6 : 합격자, 추가합격자
			CASE
				-- 차수가 0이면 최초 합격자
				WHEN A.Degree1 = '0' THEN CONVERT(VARCHAR(10), A.InsertTime, 120)
				-- 차수가 0이 아니면 추가합격자 (충원 시 등록예정 입력 시간)
				ELSE CONVERT(VARCHAR(10), METIS_C.InsertTime, 120)
			END			
		ELSE NULL														-- 예비자는 합격일자 없음
	END AS PASS_DT
	------------------------------------------------------------------------------------
	-- PASS_CD 합격구분 ---------------------------------------------------------------
	-- 최초합격자 : 10
	-- 충원대상자 : 11 (추가합격자의 경우 무조건 11 넣어주고, 포기하겠다 하면 PASS_DTL_CD 11)
	-- 이 외 : 20
	, CASE
		WHEN A.Result1 = '6' THEN								-- 합격여부 6 : 합격자, 추가합격자
			CASE
				WHEN A.Degree1 = '0' THEN '10'					-- 차수가 0이면 최초 합격자
				ELSE '11'												-- 차수가 0이 아니면 추가합격자 (추가 합격 대상으로 등록예정인 학생)
			END
		WHEN METIS_C.RESULT = '3' THEN '11'				-- 충원대상자 였으나 충원작업(METIS)에서 "포기"를 선택한 학생 
																			-- (상담원은 등록예정,포기,미결정,미연결만 선택할 수 있으므로 포기한 학생을 고정해도 상관 없을듯 함)
		WHEN METIS_C.RESULT = '6' THEN '11'				-- 충원대상자이고 "등록예정" 이지만 아직 XTEB에서 데이터 가져오기가(추가합격자) 되지 않은 학생
																			-- XTEB에는 포기값을 보여주지 않기 때문에 수정할 곳이 많음
		ELSE '20'														-- 합격여부 00 : 예비자(불합격)
	END AS PASS_CD
	------------------------------------------------------------------------------------
	-- PASS_DTL_CD 포기구분상세 ------------------------------------------------------
	-- 11 : 포기 : 충원대상자이나 충원작업에서 포기를 선택한 학생
	-- 12 : 환불 : 환불 신청 후 입학, 회계 승인 거쳐 환불이 완료된 학생
	-- 13 : 미등록 : 합격&추가합격자 중 마감 후 미등록한 학생
	, CASE
		WHEN A.Result1 = '6' THEN								-- 환불자 및 미등록자는 합격자에 한해 있음
			CASE
				WHEN C.RESULT = '7' THEN '13'					-- 미등록 처리 되었을 경우
				WHEN C.RESULT = '10' THEN '12'					-- 환불 처리 되었을 경우
				--WHEN E.IDX IS NOT NULL THEN '12'			-- 환불 데이터 있을 경우 (C.RESULT = '10' 으로 하면 환불완료 전까지 환불자로 안뜨기 때문에 신청한 내역으로 처리)
				ELSE NULL												-- 환불자도 미등록자도 아니면 공백
			END
		WHEN METIS_C.RESULT = '3' THEN '11'				-- 충원대상자 였으나 충원작업(METIS)에서 포기를 선택한 학생 
		ELSE NULL														-- 위 경우 아니면 공백
	END AS PASS_DTL_CD
	------------------------------------------------------------------------------------
	-- REG_CD 등록구분 ----------------------------------------------------------------
	-- 10 : 등록(정시 등록완료)
	-- 20 : 미등록
	-- 21 : 등록_예치금(수시만사용)
	-- 30 : 환불
	-- 34 : 환불_예치금
	-- 최초 NULL
	, CASE
		WHEN C.RESULT = '2' THEN								-- 등록
			CASE
				WHEN C.RF1 = 0 THEN '21'							-- 예치금 (입학금 = 0)
				ELSE '10'												-- 본등록금
			END
		WHEN C.RESULT = '10' THEN								-- 환불
			CASE
				WHEN C.RF1 = 0 THEN '34'							-- 예치금 환불 (입학금 = 0)
				ELSE '30'												-- 본등록금 환불
			END
		WHEN C.RESULT = '7' THEN NULL						-- 미등록 (RESULT = '7')
		ELSE NULL														-- 미등록 (RESULT IS NULL)
	END AS REG_CD
	------------------------------------------------------------------------------------
	-- BANK_LOG 등록금 수납구분-------------------------------------------------------
	, CASE
		WHEN D.IDX IS NOT NULL THEN 'Y'
		ELSE 'N'
	END BANK_LOG
	------------------------------------------------------------------------------------
	-- RANKING 석차 -------------------------------------------------------------------
	, ISNULL(A.ETC2, 0) AS RANKING
	------------------------------------------------------------------------------------
	
	------------------------------------------------------------------------------------
	-- 내부적으로  사용할 필드 ---------------------------------------------------------
	------------------------------------------------------------------------------------
	-- StudentTable
	, A.Result1 AS StudentResult, A.Degree1 AS StudentDegree
	, '1' AS SCHL_YR													-- 학년
	, A.VAccountNumber AS IMA_ACCOUNT_NUM			-- 가상계좌
	, A.RF1 AS ADMS_AMT											-- 입학금
--	, CASE
--		WHEN A.RF1 = 0 THEN A.RF9								-- 입학금이 0원이면 예치금 상황
--		ELSE A.RF2														-- 입학금이 0이 아니면 본등록금 상황
--	 END AS TUITION_FEE											-- 예치금 & 수업료
	, A.RF2 AS TUITION_FEE										-- 수업료 (예치금 항목 (KEEP_FEE) 만들어서 따로 뺌)
	, 0 AS REDU_ENT													-- 감면_입학금
	, 0 AS REDU_AMT													-- 감면_수업료
	, 0 AS SCHOLSHIP_ADMS_AMT								-- 장학_입학금
	, 0 AS SCHOLSHIP_TUITION_FEE								-- 장학_수업료
--	, A.RF8 AS SCHOLSHIP_TUITION_FEE						-- 장학_수업료
	, A.RF9 AS KEEP_FEE											-- 예치금
	, A.RF6 AS KEEP_REG											-- 기납부_예치금
	, A.RF4 AS RF4													-- 학생회비 (T01)
	, A.RF7 AS RF7													-- 신문방송비 (T02)
	, A.InsertTime AS CONF_DT									-- 생성일자
	, '10' AS REG_BANK_CD											-- 납부은행
	
	-- RegistRecord
	, C.RESULT AS XtebResult
	, C.RF1 AS RegidtADMS_AMT									-- 입학금
	, C.RF2 AS RegidtTUITION_FEE								-- 수업료 (예치금 항목 (KEEP_FEE) 만들어서 따로 뺌)
	, 0 AS RegidtREDU_ENT											-- 감면_입학금
	, 0 AS RegidtREDU_AMT										-- 감면_수업료
	, 0 AS RegidtSCHOLSHIP_ADMS_AMT						-- 장학_입학금
	, 0 AS RegidtSCHOLSHIP_TUITION_FEE					-- 장학_수업료
--	, C.RF8 AS RegidtSCHOLSHIP_TUITION_FEE				-- 장학_수업료
	, C.RF9 AS RegistKEEP_FEE									-- 예치금
	, C.RF6 AS RegistKEEP_REG									-- 기납부_예치금
	, C.RF4 AS RegistRF4											-- 학생회비 (T01)
	, C.RF7 AS RegistRF7											-- 신문방송비 (T02)
	, CASE
		WHEN C.RESULT = '2' THEN C.InsertTime				-- 등록일때만 데이터 나오도록
		ELSE NULL														-- 미등록은 NULL
	END AS REG_DT													-- 등록일자
	--	, C.InsertTime AS REG_DT									-- 등록일자

	-- METIS.RegistRecord
	, METIS_C.RESULT AS MetisResult
	
	-- RefundRecord
	--, E.InsertTime AS REFUN_DT								-- 환불일자
	, CASE
		WHEN C.RESULT = '10' THEN C.InsertTime			-- 환불일때만 데이터 나오도록
		ELSE NULL														-- 환불 아닐때는  NULL
	END AS REFUN_DT													-- 등록일자
	, E.ApproveEnter, E.ApproveBursary, E.ApproveFinance
	------------------------------------------------------------------------------------
FROM StudentTable AS A 
	INNER JOIN	SubjectTable AS B ON A.SubjectCode = B.SubjectCode
	-- XTEB RegistRecord Join (등록금 수납 완료 내역)
	LEFT OUTER JOIN (
		SELECT 
			CA.IDX, CA.StudentNumber, ISNULL(CA.RESULT, '0') AS RESULT
			, CA.RF1, CA.RF2, CA.RF3, CA.RF4, CA.RF5, CA.RF6, CA.RF7, CA.RF8, CA.RF9, CA.RF10, CA.RF11
			, CA.InsertTime
		FROM RegistRecord AS CA
		INNER JOIN (
			SELECT StudentNumber, MAX(IDX) AS MaxIDX , COUNT(*) AS CallCount , MAX(SaveFIle) AS MaxSaveFIle
			FROM RegistRecord
			GROUP BY StudentNumber
		) AS CB ON CA.StudentNumber = CB.StudentNumber AND CA.IDX = CB.MaxIDX
	) AS C  ON A.StudentNumber = C.StudentNumber
	-- METIS RegistRecord Join (충원 작업 내역)
	LEFT OUTER JOIN (
		SELECT 
			M_CA.IDX, M_CA.StudentNumber, ISNULL(M_CA.RESULT, '0') AS RESULT
			, M_CA.InsertTime
		FROM METIS.METIS.dbo.RegistRecord AS M_CA
		INNER JOIN (
			SELECT StudentNumber, MAX(IDX) AS MaxIDX , COUNT(*) AS CallCount , MAX(SaveFIle) AS MaxSaveFIle
			FROM METIS.METIS.dbo.RegistRecord
			GROUP BY StudentNumber
		) AS M_CB ON M_CA.StudentNumber = M_CB.StudentNumber AND M_CA.IDX = M_CB.MaxIDX
	) AS METIS_C  ON A.StudentNumber = METIS_C.StudentNumber
	-- XTEB BankRecord Join
	LEFT OUTER JOIN BankRecord AS D ON A.StudentNumber = D.StudentNumber
	-- XTEB RefundRecord Join
	LEFT OUTER JOIN RefundRecord AS E ON A.StudentNumber = E.StudentNumber
--	LEFT OUTER JOIN METIS.METIS.dbo.StudentTable AS F ON A.StudentNumber = F.StudentNumber
--------------------------------------------------------------

