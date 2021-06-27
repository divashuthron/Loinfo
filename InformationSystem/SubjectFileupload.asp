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

		'등록금
		Dim RF1, RF2, RF3, RF4, RF5, RF6, RF7, RF8, RF9, RF10, RF11  			
		
		'입력
		Dim INPT_USID    			: INPT_USID = SessionUserID
		Dim INPT_ADDR    			: INPT_ADDR = ASP_USER_IP
											
		If oRs.EOF=False Then

			Dim totalCount
			totalCount = oRs.RecordCount

			dim sql, i
			i = 0

			'임시테이블 생성
			sql = "IF OBJECT_ID('tempdb..##SubjectTableTemp') IS NOT NULL drop table ##SubjectTableTemp CREATE TABLE [##SubjectTableTemp]( [IDX] [int] IDENTITY(1,1) NOT NULL,	[MYear] [varchar](4) NOT NULL,	[SubjectCode] [varchar](20) NOT NULL,	[Division0] [varchar](20) NOT NULL,	[Subject] [varchar](50) NOT NULL,	[Division1] [varchar](50) NULL,	[Division2] [varchar](20) NULL,	[Division3] [varchar](20) NULL,	[Quorum] [smallint] NULL,	[QuorumFix] [smallint] NULL,	[RF1] [int] NULL,	[RF2] [int] NULL,	[RF3] [int] NULL,	[RF4] [int] NULL,	[RF5] [int] NULL,	[RF6] [int] NULL,	[RF7] [int] NULL,	[RF8] [int] NULL,	[RF9] [int] NULL,	[RF10] [int] NULL,	[RF11] [int] NULL, [INPT_USID] [varchar](20) NULL,	[INPT_DATE] [datetime] NULL,	[INPT_ADDR] [varchar](20) NULL,	[InsertTime] [datetime] NOT NULL CONSTRAINT [DF_SubjectTableTemp_InsertTime]  DEFAULT (getdate())) " & vbCrLf
			'oCmd.CommandText = sql
			'oCmd.Execute()
			dbcon.Execute sql

			'임시테이블에 업로드한 파일 isnert
			sql = "INSERT INTO [##SubjectTableTemp](MYear,SubjectCode,Division0,Subject,Division1,Division2,Division3,Quorum,QuorumFix,RF1,RF2,RF3,RF4,RF5,RF6,RF7,RF8,RF9,RF10,RF11,INPT_USID,INPT_DATE,INPT_ADDR,InsertTime) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, getdate(), ?, getdate())"

		'	oCmd.CommandText = "TheStoredProc"
		'	oCmd.CommandType = 4
		'	oCmd.Parameters.Append 'something'
		'	oCmd.Parameters.Append 'something else'
		'	oCmd.ActiveConnection = ConnOb

			oCmd.CommandText = sql

			oCmd.Parameters.Append oCmd.CreateParameter("MYear", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("SubjectCode", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division0", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 50 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division1", adVarChar, adParamInput, 50 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division2", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division3", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Quorum", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("QuorumFix", adInteger, adParamInput, 0 )

			oCmd.Parameters.Append oCmd.CreateParameter("RF1", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF2", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF3", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF4", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF5", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF6", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF7", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF8", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF9", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF10", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RF11", adInteger, adParamInput, 0 )

			oCmd.Parameters.Append oCmd.CreateParameter("INPT_USID", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("INPT_ADDR", adVarChar, adParamInput, 20 )

			do until oRs.eof
				RF1 = oRS(8)
				RF2 = oRS(9) 
				RF4 = oRS(10)
				RF6 = 0
				RF7 = oRS(11)
				RF8 = oRS(12)
				RF9 = oRS(13)
				RF3 = RF1 + RF2 
				RF10 = RF4 + RF7
				RF5 = RF3 + RF9 - RF8 - RF6
				RF11 = RF5 + RF6 + RF10

				i = i + 1
				'첫 주석 패스
				If oRS(0) <> "년도" Then
						
					OCmd.Parameters("MYear") = oRS(0)
					oCmd.Parameters("SubjectCode") = oRS(2) & oRS(1) & oRS(3)
					oCmd.Parameters("Division0") = oRS(1)
					oCmd.Parameters("Subject") = oRS(2)
					oCmd.Parameters("Division1") = oRS(3)
					oCmd.Parameters("Division2") = oRS(4)
					oCmd.Parameters("Division3") = oRS(5)
					OCmd.Parameters("Quorum") = oRS(6)
					oCmd.Parameters("QuorumFix") = oRS(7)

					oCmd.Parameters("RF1") = RF1
					oCmd.Parameters("RF2") = RF2
					oCmd.Parameters("RF3") = RF3
					oCmd.Parameters("RF4") = RF4
					oCmd.Parameters("RF5") = RF5
					oCmd.Parameters("RF6") = RF6
					oCmd.Parameters("RF7") = RF7
					oCmd.Parameters("RF8") = RF8
					oCmd.Parameters("RF9") = RF9
					oCmd.Parameters("RF10") = RF10
					oCmd.Parameters("RF11") = RF11

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
						response.write "<td>" & RF1  & "</td>"
						response.write "<td>" & RF2  & "</td>"
						response.write "<td>" & RF4  & "</td>"
						response.write "<td>" & RF7  & "</td>"
						response.write "<td>" & RF8  & "</td>"	
						response.write "<td>" & RF9  & "</td>"	
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