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

		Dim Time
											
		If oRs.EOF=False Then

			Dim totalCount
			totalCount = oRs.RecordCount

			dim sql, i
			i = 0

			'임시테이블 생성
			sql = "IF OBJECT_ID('tempdb..##ApplicationTable') IS NOT NULL drop table ##ApplicationTable CREATE TABLE [##ApplicationTable]( [IDX] [int] IDENTITY(1,1) NOT NULL,	[Myear] [varchar](4) NOT NULL,	[Division0] [varchar](60) NOT NULL,	[Subject] [varchar](50) NOT NULL,	[Division1] [varchar](60) NULL,	[Division2] [varchar](60) NULL,	[Division3] [varchar](60) NULL,	[Division4] [varchar](60) NULL,	[StudentNumber] [varchar](10) NOT NULL,	[StudentNameKor] [varchar](30) NOT NULL,	[StudentNameUsa] [varchar](30) NULL,	[StudentNameChi] [varchar](30) NULL,	[Citizen1] [char](6) NULL,	[Citizen2] [char](7) NULL,	[Sex] [smallint] NULL,	[HighGraduationYear] [varchar](4) NULL,	[HighCode] [varchar](10) NULL,	[HighSubject] [varchar](30) NULL,	[HighGraduationDivision] [smallint] NULL,	[Qualification] [varchar](4) NULL,	[QualificationYear] [varchar](4) NULL,	[QualificationAreaCode] [varchar](30) NULL,	[Semester] [varchar](20) NULL,	[UniversityName] [varchar](60) NULL,	[AugScore] [varchar](10) NULL,	[PerfectScore] [varchar](10) NULL,	[Credit] [smallint] NULL,	[HighDivision] [varchar](60) NULL,	[RefundDivision] [smallint] NULL,	[RefundAccountHolder] [varchar](50) NULL,	[RefundBankCode] [varchar](50) NULL,	[RefundAccount] [varchar](50) NULL,	[Tel1] [varchar](20) NULL,	[Tel2] [varchar](20) NULL,	[Tel3] [varchar](20) NULL,	[Email] [varchar](60) NULL,	[ZipCode] [varchar](6) NULL,	[Address1] [varchar](100) NULL,	[Address2] [varchar](100) NULL,	[StudentNameAgreement] [varchar](30) NULL,	[StudentRecordAgreement] [varchar](1) NULL,	[QualificationAgreement] [varchar](1) NULL,	[CSATAgreement] [varchar](1) NULL, [PersonalCollectionAgreement] [varchar](1) NULL, [UniquelyAgreement] [varchar](1) NULL, [PersonalTrustAgreement] [varchar](1) NULL, [PersonalofferAgreement] [varchar](1) NULL,	[StudentAgreement] [varchar](1) NULL,	[ReceiptDate] [datetime] NULL,	[ReceiptTime] [datetime] NULL,	[INPT_USID] [varchar](20) NULL,	[INPT_DATE] [datetime] NULL,	[INPT_ADDR] [varchar](20) NULL, [InsertTime] [datetime] NOT NULL CONSTRAINT [DF_SubjectTableTemp_InsertTime]  DEFAULT (getdate())) " & vbCrLf
			'oCmd.CommandText = sql
			'oCmd.Execute()
			dbcon.Execute sql

			'임시테이블에 업로드한 파일 isnert
			sql = "INSERT INTO [##ApplicationTable](Myear,Division0,Subject,Division1,Division2,Division3,Division4,StudentNumber,StudentNameKor,StudentNameUsa,StudentNameChi,Citizen1,Citizen2,Sex,HighGraduationYear,HighCode,HighSubject,HighGraduationDivision,Qualification,QualificationYear,QualificationAreaCode,Semester,UniversityName,AugScore,PerfectScore,Credit,HighDivision,RefundDivision,RefundAccountHolder,RefundBankCode,RefundAccount,Tel1,Tel2,Tel3,Email,ZipCode,Address1,Address2,StudentNameAgreement,StudentRecordAgreement,QualificationAgreement,CSATAgreement,PersonalCollectionAgreement,UniquelyAgreement,PersonalTrustAgreement,PersonalofferAgreement,StudentAgreement,ReceiptDate,ReceiptTime,INPT_USID,INPT_DATE,INPT_ADDR,InsertTime) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, getdate(), ?, getdate())"

		'	oCmd.CommandText = "TheStoredProc"
		'	oCmd.CommandType = 4
		'	oCmd.Parameters.Append 'something'
		'	oCmd.Parameters.Append 'something else'
		'	oCmd.ActiveConnection = ConnOb

			oCmd.CommandText = sql

			oCmd.Parameters.Append oCmd.CreateParameter("MYear", adVarChar, adParamInput, 4 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division0", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 50 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division1", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division2", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division3", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("Division4", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("StudentNumber", adVarChar, adParamInput, 10 )
			oCmd.Parameters.Append oCmd.CreateParameter("StudentNameKor", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("StudentNameUsa", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("StudentNameChi", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("Citizen1", adVarChar, adParamInput, 6 )
			oCmd.Parameters.Append oCmd.CreateParameter("Citizen2", adVarChar, adParamInput, 7 )
			oCmd.Parameters.Append oCmd.CreateParameter("Sex", adInteger, adParamInput, 0 )									  

			oCmd.Parameters.Append oCmd.CreateParameter("HighGraduationYear", adVarChar, adParamInput, 4 )
			oCmd.Parameters.Append oCmd.CreateParameter("HighCode", adVarChar, adParamInput, 10 )
			oCmd.Parameters.Append oCmd.CreateParameter("HighSubject", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("HighGraduationDivision", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("Qualification", adVarChar, adParamInput, 4 )
			oCmd.Parameters.Append oCmd.CreateParameter("QualificationYear", adVarChar, adParamInput, 4 )
			oCmd.Parameters.Append oCmd.CreateParameter("QualificationAreaCode", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("Semester", adVarChar, adParamInput, 20 )

			oCmd.Parameters.Append oCmd.CreateParameter("UniversityName", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("AugScore", adVarChar, adParamInput, 10 )
			oCmd.Parameters.Append oCmd.CreateParameter("PerfectScore", adVarChar, adParamInput, 10 )
			oCmd.Parameters.Append oCmd.CreateParameter("Credit", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("HighDivision", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("RefundDivision", adInteger, adParamInput, 0 )
			oCmd.Parameters.Append oCmd.CreateParameter("RefundAccountHolder", adVarChar, adParamInput, 50 )
			oCmd.Parameters.Append oCmd.CreateParameter("RefundBankCode", adVarChar, adParamInput, 50 )
			oCmd.Parameters.Append oCmd.CreateParameter("RefundAccount", adVarChar, adParamInput, 50 )

			oCmd.Parameters.Append oCmd.CreateParameter("Tel1", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Tel2", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Tel3", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("Email", adVarChar, adParamInput, 60 )
			oCmd.Parameters.Append oCmd.CreateParameter("ZipCode", adVarChar, adParamInput, 6 )
			oCmd.Parameters.Append oCmd.CreateParameter("Address1", adVarChar, adParamInput, 100 )
			oCmd.Parameters.Append oCmd.CreateParameter("Address2", adVarChar, adParamInput, 100 )

			oCmd.Parameters.Append oCmd.CreateParameter("StudentNameAgreement", adVarChar, adParamInput, 30 )
			oCmd.Parameters.Append oCmd.CreateParameter("StudentRecordAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("QualificationAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("CSATAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("PersonalCollectionAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("UniquelyAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("PersonalTrustAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("PersonalofferAgreement", adVarChar, adParamInput, 1 )

			oCmd.Parameters.Append oCmd.CreateParameter("StudentAgreement", adVarChar, adParamInput, 1 )
			oCmd.Parameters.Append oCmd.CreateParameter("ReceiptDate", adVarChar, adParamInput, 255 )
			oCmd.Parameters.Append oCmd.CreateParameter("ReceiptTime", adVarChar, adParamInput, 255 )

			oCmd.Parameters.Append oCmd.CreateParameter("INPT_USID", adVarChar, adParamInput, 20 )
			oCmd.Parameters.Append oCmd.CreateParameter("INPT_ADDR", adVarChar, adParamInput, 20 )

			do until oRs.eof
				If Not(isnull(oRS(48))) Then
					Time	= FormatDateTime(oRS(48),4)
				End If

				i = i + 1
				'첫 주석 패스
				If oRS(0) <> "년도" Then
						
					OCmd.Parameters("MYear") = oRS(0)
					OCmd.Parameters("Division0") = oRS(1)
					OCmd.Parameters("Subject") = oRS(2)
					OCmd.Parameters("Division1") = oRS(3)
					OCmd.Parameters("Division2") = oRS(4)
					OCmd.Parameters("Division3") = oRS(5)
					OCmd.Parameters("Division4") = oRS(6)
					OCmd.Parameters("StudentNumber") = oRS(7)
					OCmd.Parameters("StudentNameKor") = oRS(8)
					OCmd.Parameters("StudentNameUsa") = oRS(9)
					OCmd.Parameters("StudentNameChi") = oRS(10)
					OCmd.Parameters("Citizen1") = oRS(11)
					OCmd.Parameters("Citizen2") = oRS(12)
					OCmd.Parameters("Sex") = oRS(13)

					OCmd.Parameters("HighGraduationYear") = oRS(14)
					OCmd.Parameters("HighCode") = oRS(15)
					OCmd.Parameters("HighSubject") = oRS(16)
					OCmd.Parameters("HighGraduationDivision") = oRS(17)
					OCmd.Parameters("Qualification") = oRS(18)
					OCmd.Parameters("QualificationYear") = oRS(19)
					OCmd.Parameters("QualificationAreaCode") = oRS(20)
					OCmd.Parameters("Semester") = oRS(21)

					OCmd.Parameters("UniversityName") = oRS(22)
					OCmd.Parameters("AugScore") = oRS(23)
					OCmd.Parameters("PerfectScore") = oRS(24)
					OCmd.Parameters("Credit") = oRS(25)
					OCmd.Parameters("HighDivision") = oRS(26)
					OCmd.Parameters("RefundDivision") = oRS(27)
					OCmd.Parameters("RefundAccountHolder") = oRS(28)
					OCmd.Parameters("RefundBankCode") = oRS(29)
					OCmd.Parameters("RefundAccount") = oRS(30)

					OCmd.Parameters("Tel1") = oRS(31)
					OCmd.Parameters("Tel2") = oRS(32)
					OCmd.Parameters("Tel3") = oRS(33)
					OCmd.Parameters("Email") = oRS(34)
					OCmd.Parameters("ZipCode") = oRS(35)
					OCmd.Parameters("Address1") = oRS(36)
					OCmd.Parameters("Address2") = oRS(37)

					OCmd.Parameters("StudentNameAgreement") = oRS(38)
					OCmd.Parameters("StudentRecordAgreement") = oRS(39)
					OCmd.Parameters("QualificationAgreement") = oRS(40)
					OCmd.Parameters("CSATAgreement") = oRS(41)
					OCmd.Parameters("PersonalCollectionAgreement") = oRS(42)
					OCmd.Parameters("UniquelyAgreement") = oRS(43)
					OCmd.Parameters("PersonalTrustAgreement") = oRS(44)
					OCmd.Parameters("PersonalofferAgreement") = oRS(45)

					OCmd.Parameters("StudentAgreement") = oRS(46)
					OCmd.Parameters("ReceiptDate") = oRS(47)
					OCmd.Parameters("ReceiptTime") = Time

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
						response.write "<td>" & oRS(41)  & "</td>"
						response.write "<td>" & oRS(42)  & "</td>"
						response.write "<td>" & oRS(43)  & "</td>"
						response.write "<td>" & oRS(44)  & "</td>"
						response.write "<td>" & oRS(45)  & "</td>"
						response.write "<td>" & oRS(46)  & "</td>"
						response.write "<td>" & oRS(47)  & "</td>"
						response.write "<td>" & Time  & "</td>"
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