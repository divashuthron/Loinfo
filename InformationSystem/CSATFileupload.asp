<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
	dim dxUp
	set dxUp =  Server.CreateObject("DEXT.FileUpload")
	dxUp.DefaultPath = Server.MapPath ("/upload/Excel/") & "\"

	dim filename1
	filename1=trim(dxUp("callbackfile"))
	'파일저장
	dxUp("callbackfile").Save, False

	'response.write dxUp.LastSavedFileName'중복검사후 새이름으로 저장된 파일이름 

	server.scripttimeout = 400
	Dim SavedFileName, FileExtention
	SavedFileName = dxUp.LastSavedFileName
	FileExtention = Split(SavedFileName,".")

	select case LCase(FileExtention(UBound(FileExtention,1)))
		Case "xls", "xlsx" : 'Load Excel
			LoadXls()
	End Select

	Function LoadXls()
		'On Error Resume Next
		Dim Dbcon
		Set Dbcon = createobject("ADODB.connection")
		Dbcon.ConnectionTimeout = 200
		Dbcon.CommandTimeout = 600
		server.scripttimeout = 600

		Dim DBConnectionString
		DBConnectionString = "Provider=SQLOLEDB;Data Source=SQLMISS;Initial Catalog=InformationSystem;user ID=InterViewMng;password=east12!@;"

		Dbcon.open DBConnectionString
		'Dbcon.BeginTrans
		dim oCon, path
			
		path = server.MapPath("/upload/Excel/") & "\" + SavedFileName

		Set oCon = Createobject("ADODB.connection")
		oCon.open "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties='Excel 12.0;HDR=YES';Data Source=" + path + ""
		dim oCmd, oRs

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.CursorLocation = 3 
		oRs.CursorType = 3 
		oRs.LockType = 3 

		set oCmd = Server.CreateObject("ADODB.Command")
		oCmd.ActiveConnection = Dbcon
		oCmd.CommandType = 1

		'저장한 파일의 sheet1 읽기
		oRs.Open "select * from [sheet1$]", oCon

		Response.ContentType = "text/xml"
		'response.write "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "euc-kr" & Chr(34) & "?>" & vbCrLf
		
		'입력
		Dim INPT_USID    			: INPT_USID = SessionUserID
		Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
											
		If oRs.EOF=False Then

			Dim totalCount
			totalCount = oRs.RecordCount

			dim sql, i
			i = 0

			'임시테이블 생성
			sql = "IF OBJECT_ID('tempdb..##CSATTable') IS NOT NULL drop table ##CSATTable CREATE TABLE [##CSATTable]( [IDX] [int] IDENTITY(1,1) NOT NULL,		[SCHL_YEAR] [varchar](4) NOT NULL,	[COLL_FLAG] [varchar](6) NOT NULL,	[EXAM_NUMB] [varchar](10) NOT NULL,	[LGFD_EXFG] [int] NULL,	[LGFD_SDSC] [int] NULL,	[LGFD_CENT] [int] NULL,	[LGFD_GRAD] [int] NULL,	[MTFD_EXFG] [int] NULL,	[MTFD_EXTP] [int] NULL,	[MTFD_SDSC] [int] NULL,	[MTFD_CENT] [int] NULL,	[MTFD_GRAD] [int] NULL,	[FLFD_EXFG] [int] NULL,	[FLFD_SDSC] [int] NULL,	[FLFD_CENT] [int] NULL,	[FLFD_GRAD] [int] NULL,	[RSFD_EXFG] [int] NULL,	[RSFD_FLAG] [int] NULL,	[RSFD_CCCT] [int] NULL,	[RSFD_SBJ1] [varchar](6) NULL,	[RSFD_SCR1] [int] NULL,	[RSFD_CNT1] [int] NULL,	[RSFD_GRD1] [int] NULL,	[RSFD_SBJ2] [varchar](6) NULL,	[RSFD_SCR2] [int] NULL,	[RSFD_CNT2] [int] NULL,	[RSFD_GRD2] [int] NULL,	[RSFD_SBJ3] [int] NULL,	[RSFD_SCR3] [int] NULL,	[RSFD_CNT3] [int] NULL,	[RSFD_GRD3] [int] NULL,	[RSFD_SBJ4] [int] NULL,	[RSFD_SCR4] [int] NULL,	[RSFD_CNT4] [int] NULL,	[RSFD_GRD4] [int] NULL,	[SCFL_EXFG] [int] NULL,	[SCFL_SBJT] [int] NULL,	[SCFL_SDSC] [int] NULL,	[SCFL_CENT] [int] NULL,	[SCFL_GRAD] [int] NULL,	[REMK_TEXT] [varchar](4000) NULL,	[INPT_USID] [varchar](20) NULL,	[INPT_DATE] [datetime] NULL,	[INPT_ADDR] [varchar](20) NULL, [InsertTime] [datetime] NOT NULL CONSTRAINT [DF_SubjectTableTemp_InsertTime]  DEFAULT (getdate())) " & vbCrLf
			'oCmd.CommandText = sql
			'oCmd.Execute()
			dbcon.Execute sql

			'임시테이블에 업로드한 파일 isnert
			sql = "INSERT INTO [##CSATTable](SCHL_YEAR,COLL_FLAG,EXAM_NUMB,LGFD_EXFG,LGFD_SDSC,LGFD_CENT,LGFD_GRAD,MTFD_EXFG,MTFD_EXTP,MTFD_SDSC,MTFD_CENT,MTFD_GRAD,FLFD_EXFG,FLFD_SDSC,FLFD_CENT,FLFD_GRAD,RSFD_EXFG,RSFD_FLAG,RSFD_CCCT,RSFD_SBJ1,RSFD_SCR1,RSFD_CNT1,RSFD_GRD1,RSFD_SBJ2,RSFD_SCR2,RSFD_CNT2,RSFD_GRD2,RSFD_SBJ3,RSFD_SCR3,RSFD_CNT3,RSFD_GRD3,RSFD_SBJ4,RSFD_SCR4,RSFD_CNT4,RSFD_GRD4,SCFL_EXFG,SCFL_SBJT,SCFL_SDSC,SCFL_CENT,SCFL_GRAD,REMK_TEXT,INPT_USID,INPT_DATE,INPT_ADDR,InsertTime) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, getdate(), ?, getdate())"

		'	oCmd.CommandText = "TheStoredProc"
		'	oCmd.CommandType = 4
		'	oCmd.Parameters.Append 'something'
		'	oCmd.Parameters.Append 'something else'
		'	oCmd.ActiveConnection = ConnOb

			oCmd.CommandText = sql

			oCmd.Parameters.Append oCmd.CreateParameter("SCHL_YEAR", adVarChar, adParamInput, 4 )		'년도
			oCmd.Parameters.Append oCmd.CreateParameter("COLL_FLAG", adVarChar, adParamInput, 6 )		'모집구분
			oCmd.Parameters.Append oCmd.CreateParameter("EXAM_NUMB", adVarChar, adParamInput, 10 )		'수험번호

			oCmd.Parameters.Append oCmd.CreateParameter("LGFD_EXFG", adVarChar, adParamInput, 20 )		'언어영역응시구분
			oCmd.Parameters.Append oCmd.CreateParameter("LGFD_SDSC", adVarChar, adParamInput, 20 )		'언어영역표준점수
			oCmd.Parameters.Append oCmd.CreateParameter("LGFD_CENT", adVarChar, adParamInput, 20 )		'언어영역백분위
			oCmd.Parameters.Append oCmd.CreateParameter("LGFD_GRAD", adVarChar, adParamInput, 20 )		'언어영역등급

			oCmd.Parameters.Append oCmd.CreateParameter("MTFD_EXFG", adVarChar, adParamInput, 20 )		'수리영역응시구분
			oCmd.Parameters.Append oCmd.CreateParameter("MTFD_EXTP", adVarChar, adParamInput, 20 )		'수리영역응시유형
			oCmd.Parameters.Append oCmd.CreateParameter("MTFD_SDSC", adVarChar, adParamInput, 20 )		'수리영역표준점수
			oCmd.Parameters.Append oCmd.CreateParameter("MTFD_CENT", adVarChar, adParamInput, 20 )		'수리영역백분위
			oCmd.Parameters.Append oCmd.CreateParameter("MTFD_GRAD", adVarChar, adParamInput, 20 )		'수리영역등급

			oCmd.Parameters.Append oCmd.CreateParameter("FLFD_EXFG", adVarChar, adParamInput, 20 )		'외국어영역응시구분
			oCmd.Parameters.Append oCmd.CreateParameter("FLFD_SDSC", adVarChar, adParamInput, 20 )		'외국어영역표준점수
			oCmd.Parameters.Append oCmd.CreateParameter("FLFD_CENT", adVarChar, adParamInput, 20 )		'외국어영역백분위
			oCmd.Parameters.Append oCmd.CreateParameter("FLFD_GRAD", adVarChar, adParamInput, 20 )		'외국어영역등급

			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_EXFG", adVarChar, adParamInput, 20 )		'탐구영역응시구분
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_FLAG", adVarChar, adParamInput, 20 )		'탐구영역구분
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_CCCT", adVarChar, adParamInput, 20 )		'탐구영역선택과목수
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SBJ1", adVarChar, adParamInput, 6 )		'탐구영역과목1
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SCR1", adVarChar, adParamInput, 20 )		'탐구영역표준점수1
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_CNT1", adVarChar, adParamInput, 20 )		'탐구영역백분위1
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_GRD1", adVarChar, adParamInput, 20 )		'탐구영역등급1
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SBJ2", adVarChar, adParamInput, 6 )		'탐구영역과목2
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SCR2", adVarChar, adParamInput, 20 )		'탐구영역표준점수2
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_CNT2", adVarChar, adParamInput, 20 )		'탐구영역백분위2
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_GRD2", adVarChar, adParamInput, 20 )		'탐구영역등급2
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SBJ3", adVarChar, adParamInput, 20 )		'탐구영역과목3
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SCR3", adVarChar, adParamInput, 20 )		'탐구영역표준점수3
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_CNT3", adVarChar, adParamInput, 20 )		'탐구영역백분위3
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_GRD3", adVarChar, adParamInput, 20 )		'탐구영역등급3
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SBJ4", adVarChar, adParamInput, 20 )		'탐구영역과목4
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_SCR4", adVarChar, adParamInput, 20 )		'탐구영역표준점수4
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_CNT4", adVarChar, adParamInput, 20 )		'탐구영역백분위4
			oCmd.Parameters.Append oCmd.CreateParameter("RSFD_GRD4", adVarChar, adParamInput, 20 )		'탐구영역등급4

			oCmd.Parameters.Append oCmd.CreateParameter("SCFL_EXFG", adVarChar, adParamInput, 20 )		'제2외국어영역응시구분
			oCmd.Parameters.Append oCmd.CreateParameter("SCFL_SBJT", adVarChar, adParamInput, 20 )		'제2외국어영역과목
			oCmd.Parameters.Append oCmd.CreateParameter("SCFL_SDSC", adVarChar, adParamInput, 20 )		'제2외국어표준점수
			oCmd.Parameters.Append oCmd.CreateParameter("SCFL_CENT", adVarChar, adParamInput, 20 )		'제2외국어백분위
			oCmd.Parameters.Append oCmd.CreateParameter("SCFL_GRAD", adVarChar, adParamInput, 20 )		'제2외국어등급
			oCmd.Parameters.Append oCmd.CreateParameter("REMK_TEXT", adVarChar, adParamInput, 4000 )	'비고

			oCmd.Parameters.Append oCmd.CreateParameter("INPT_USID", adVarChar, adParamInput, 20 )		
			oCmd.Parameters.Append oCmd.CreateParameter("INPT_ADDR", adVarChar, adParamInput, 20 )		

			do until oRs.eof
				i = i + 1
				'첫 주석 패스
				If oRS(0) <> "년도" Then
						
					OCmd.Parameters("SCHL_YEAR") = oRS(0)
					OCmd.Parameters("COLL_FLAG") = oRS(1)
					OCmd.Parameters("EXAM_NUMB") = oRS(2)

					OCmd.Parameters("LGFD_EXFG") = oRS(3)
					OCmd.Parameters("LGFD_SDSC") = oRS(4)
					OCmd.Parameters("LGFD_CENT") = oRS(5)
					OCmd.Parameters("LGFD_GRAD") = oRS(6)

					OCmd.Parameters("MTFD_EXFG") = oRS(7)
					OCmd.Parameters("MTFD_EXTP") = oRS(8)
					OCmd.Parameters("MTFD_SDSC") = oRS(9)
					OCmd.Parameters("MTFD_CENT") = oRS(10)
					OCmd.Parameters("MTFD_GRAD") = oRS(11)

					OCmd.Parameters("FLFD_EXFG") = oRS(12)
					OCmd.Parameters("FLFD_SDSC") = oRS(13)
					OCmd.Parameters("FLFD_CENT") = oRS(14)
					OCmd.Parameters("FLFD_GRAD") = oRS(15)

					OCmd.Parameters("RSFD_EXFG") = oRS(16)
					OCmd.Parameters("RSFD_FLAG") = oRS(17)
					OCmd.Parameters("RSFD_CCCT") = oRS(18)
					OCmd.Parameters("RSFD_SBJ1") = oRS(19)
					OCmd.Parameters("RSFD_SCR1") = oRS(20)
					OCmd.Parameters("RSFD_CNT1") = oRS(21)
					OCmd.Parameters("RSFD_GRD1") = oRS(22)
					OCmd.Parameters("RSFD_SBJ2") = oRS(23)
					OCmd.Parameters("RSFD_SCR2") = oRS(24)
					OCmd.Parameters("RSFD_CNT2") = oRS(25)
					OCmd.Parameters("RSFD_GRD2") = oRS(26)
					OCmd.Parameters("RSFD_SBJ3") = oRS(27)
					OCmd.Parameters("RSFD_SCR3") = oRS(28)
					OCmd.Parameters("RSFD_CNT3") = oRS(29)
					OCmd.Parameters("RSFD_GRD3") = oRS(30)
					OCmd.Parameters("RSFD_SBJ4") = oRS(31)
					OCmd.Parameters("RSFD_SCR4") = oRS(32)
					OCmd.Parameters("RSFD_CNT4") = oRS(33)
					OCmd.Parameters("RSFD_GRD4") = oRS(34)

					OCmd.Parameters("SCFL_EXFG") = oRS(35)
					OCmd.Parameters("SCFL_SBJT") = oRS(36)
					OCmd.Parameters("SCFL_SDSC") = oRS(37)
					OCmd.Parameters("SCFL_CENT") = oRS(38)
					OCmd.Parameters("SCFL_GRAD") = oRS(39)
					OCmd.Parameters("REMK_TEXT") = oRS(40)

					oCmd.Parameters("INPT_USID") = INPT_USID
					oCmd.Parameters("INPT_ADDR") = INPT_ADDR

					'print_sql(oCmd)
					'response.end

					if Err.Description = "" Then
					oCmd.Execute()
					End If

					'테이블 리턴
					if Err.Description = "" Then
						response.write "<tr class='viewDetail'>"
						response.write "<td>" & oRS(0)  & "</td>"
						response.write "<td>" & oRS(1)  & "</td>"
						response.write "<td>" & oRS(2)  & "</td>"
						response.write "<td>" & oRS(3)  & "</td>"
						response.write "<td>" & oRS(4)  & "</td>"
						response.write "<td>" & oRS(5)  & "</td>"
						response.write "<td>" & oRS(6)  & "</td>"
						response.write "<td>" & oRS(7)  & "</td>"
						response.write "<td>" & oRS(8)  & "</td>"
						response.write "<td>" & oRS(9)  & "</td>"
						response.write "<td>" & oRS(10)  & "</td>"
						response.write "<td>" & oRS(11)  & "</td>"
						response.write "<td>" & oRS(12)  & "</td>"
						response.write "<td>" & oRS(13)  & "</td>"
						response.write "<td>" & oRS(14)  & "</td>"
						response.write "<td>" & oRS(15)  & "</td>"
						response.write "<td>" & oRS(16)  & "</td>"
						response.write "<td>" & oRS(17)  & "</td>"
						response.write "<td>" & oRS(18)  & "</td>"
						response.write "<td>" & oRS(19)  & "</td>"
						response.write "<td>" & oRS(20)  & "</td>"
						response.write "<td>" & oRS(21)  & "</td>"
						response.write "<td>" & oRS(22)  & "</td>"
						response.write "<td>" & oRS(23)  & "</td>"
						response.write "<td>" & oRS(24)  & "</td>"
						response.write "<td>" & oRS(25)  & "</td>"
						response.write "<td>" & oRS(26)  & "</td>"
						response.write "<td>" & oRS(27)  & "</td>"
						response.write "<td>" & oRS(28)  & "</td>"
						response.write "<td>" & oRS(29)  & "</td>"
						response.write "<td>" & oRS(30)  & "</td>"
						response.write "<td>" & oRS(31)  & "</td>"
						response.write "<td>" & oRS(32)  & "</td>"
						response.write "<td>" & oRS(33)  & "</td>"
						response.write "<td>" & oRS(34)  & "</td>"
						response.write "<td>" & oRS(35)  & "</td>"
						response.write "<td>" & oRS(36)  & "</td>"
						response.write "<td>" & oRS(37)  & "</td>"
						response.write "<td>" & oRS(38)  & "</td>"
						response.write "<td>" & oRS(39)  & "</td>"
						response.write "<td>" & oRS(40)  & "</td>"
						response.write "</tr>"
					End If
				End If			

				If i Mod 1000 = 999 Then Response.Flush

				oRs.movenext
			Loop

			'건수 리턴 
			'<count>으로 구분
			response.write "<count>"& i		
			
		End If
			
		set oCmd = Nothing
			
		oRs.close
		set oRs = nothing

		oCon.close
		set oCon = nothing

		'Dbcon.CommitTrans
		Dbcon.close
		set Dbcon = nothing

				
		'Upload된 파일 삭제
		'dim dFile

		'set dFile = Server.CreateObject("Scripting.FileSystemObject")
		'dFile.DeleteFile(path)
		'set dFile = nothing
	End Function

	'text 파일 사용 할 시(사용 안 함)
	Function LoadTxt()

	End Function
%>