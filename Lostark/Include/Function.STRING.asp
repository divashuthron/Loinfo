<%
'=========== 페이지 설정 Proc ========================================
	Sub sbNoCacheProc()
		Response.Expires = -1
		Response.Buffer = true
		Response.AddHeader "pragma", "no-cache"
		Response.CacheControl = "no-cache"
		Response.Charset = "UTF-8"
		Session.CodePage = 65001
		Server.ScriptTimeOut = 1440
		Session.Timeout = 60
	End Sub

	Sub sbNoCacheProc2()
		Response.Expires = 0
		Response.Buffer = True
		Response.ExpiresAbsolute = now() - 1
		Response.AddHeader "pragma", "no-cache"
		Response.AddHeader "cache-control","private"
		Response.CacheControl = "no-cache"
		Response.Charset = "UTF-8"
		Session.CodePage = 65001
		Server.ScriptTimeOut = 1440
		Session.Timeout = 60
	End Sub

'=========== Response.Write Proc ======================================
	Sub W(strTemp)
		Response.Write strTemp & vbcrlf
	End Sub

'=========== Response.end Proc =======================================
	Sub EE()
		Response.end
	End Sub

'=========== Response.end Proc =======================================
	Sub RW(strTemp)
		If IsE(strTemp) Then strTemp = Empty End if
		Response.Write "<font color=red>[값] </font>"& Replace(strTemp, Chr(13), "<br/>") &" <font color=red>[타입] </font> : "& TypeName(strTemp) &"<br>"
	End Sub

'=========== Response.end Proc =======================================
	Sub RWXML(Tag, strTemp)
		If IsE(strTemp) Then strTemp = Empty End if
		Response.Write "<"& Tag &"><![CDATA["& strTemp &")]]></"& Tag &">"& vbcrlf
	End Sub

'=========== Response.end Proc =======================================
	Sub RWE(strTemp)
		Call RW(strTemp)
		Response.End()
	End Sub

'=========== 1차원 배열을 출력 =======================================
	Sub Array_RW1(byref Object)
		If Not IsE(Object) Then
			Dim i
			Response.Write "<table border=1>"
			If UBound(Object) >= 0 Then
				Response.Write "<tr valign=top>"
				For i=0 To UBound(Object)
					Response.Write "<td><font color=red>"& i &"</font> 열<br><br>"& Object(i) &"</td>"
				Next
				Response.Write "</tr>"
			End if
			Response.Write "</table>"
		Else
			Response.Write " Empty "
		End if
	End Sub

'=========== 2차원 배열을 출력 =======================================
	Sub Array_RW2(byref Object)
		If Not IsE(Object) Then
			Dim i, j
			Response.Write "<table border=1>"
			If UBound(Object,2) >= 0 Then
				For i=0 To UBound(Object, 2)
					Response.Write "<tr valign=top>"
					For j=0 To UBound(Object)
						Response.Write "<td><font color=red>"& i &"</font> 행<font color=red>"& j &"</font> 열</font><br><br>"& Object(j,i) &"</td>"
					Next
					Response.Write "</tr>"
				Next
			End if
			Response.Write "</table>"
		Else
			Response.Write " Empty "
		End if
	End Sub


'=========== Request.Form 값 반환 Proc =================================
	Function fnRF(sElement)
		fnRF = Trim(Request.Form(sElement))
	End Function

'=========== Request.QueryString 값 반환 Proc ============================
	Function fnRQ(sElement)
		fnRQ = Trim(Request.QueryString(sElement))
	End Function

'=========== Request 값 반환 Proc =====================================
	Function fnR(sElement, default)
	    If (IsE(Request(sElement))) Then
		    fnR = default
	    Else
		    fnR = Replace(Request(sElement), "'", "''")
	    End If
	End Function

'=========== Request 값 확인 Proc =====================================
	Sub fnGetValue(Separator)
		Dim sElement
		
		Response.Write "<b>Server Variable</b><br>"
		For Each sElement in Request.ServerVariables
			Response.Write sElement & " = " & Request.ServerVariables(sElement) & Separator
		Next
		
		Response.Write "<b>GET 데이터</b><br>"
		For Each sElement in Request.QueryString
			Response.Write sElement & " = " & Request.QueryString(sElement) & Separator
		Next
		
		Response.Write "<b>POST 데이터</b><br>"
		For Each sElement in Request.Form
			Response.Write sElement & " = " & Request.Form(sElement) & Separator
		Next
		
		Response.End
	End Sub

'=========== 빈값 표현 Proc ==========================================
	Function IsE(Value)
		If isArray(Value) AND Not isEmpty(Value) Then	'배열이면서 빈값이 아니면
			IsE = False
			Exit function
		Else
			IsE = True
		End if
	
		If isEmpty(Value) OR isNull(Value) OR Value = "" Then
			IsE = True
		Else
			IsE = False
		End if
	End Function

	Function IsEnV(v, default)
		If IsE(v) Then
			IsEnV = default
		Else
			IsEnV = V
		End If
	End Function

'=========== 문자 변환 ==============================================
	Function R(getStr, oldStr, newStr)
		If (IsE(getStr)) Then
			R = ""
		Else
			R = replace(getStr, oldStr, newStr)
		End If
	End Function

	'대소문자 구분 구분없이 변환
	Function UL_ReplaceText(ByVal allText, ByVal findText, ByVal replaceText)
		Dim regObj
		Set regObj = New RegExp
		regObj.Pattern = findText		'패턴 설정 
		regObj.IgnoreCase = True		'대소문자 구분 여부 
		regObj.Global = True				'전체 문서에서 검색
		UL_ReplaceText = regObj.Replace(allText, replaceText)
	End Function 
	
'=========== HTML 태그를 Text 형태로 변환하는 Proc =======================
	Function fnGetHtmlReplaceTextProc(strTemp)
		strTemp = R(strTemp, "&", "&amp;")
		strTemp = R(strTemp, "<", "&lt;")
		strTemp = R(strTemp, ">", "&gt;")
		fnGetHtmlReplaceTextProc = strTemp
	End Function

'=========== Text 형태의 데이터를 HTML 형태로 변환하는 Proc ================
	Function fnGetTextReplaceHtmlProc(strTemp)
		strTemp = R(strTemp, "&amp;", "&")
		strTemp = R(strTemp, "&lt;", "<")
		strTemp = R(strTemp, "&gt;", ">")
		fnGetTextReplaceHtmlProc = strTemp
	End Function

'=========== Textarea 또는 Chr(13)으로 들어오는 값 <br> 태그로 Change Proc ====
	Function fnGetChrReplaceBrProc(strTemp)
		If strTemp = "" Or IsNull(strTemp) Then
			strTemp = ""
		Else
			strTemp = R(strTemp, Chr(13) & Chr(10), "<br>")
		End If
		fnGetChrReplaceBrProc = strTemp
	End Function

'=========== Script Value 처리 Proc ====================================
	Function fnGetScriptValueProc(sValue)
		If sValue <> "" Then
			sValue = R(sValue, "“", "\“")
			sValue = R(sValue, "”", "\”")
			sValue = R(sValue, """", "&quot;")
			sValue = R(sValue, "<br>"& vbcrlf, "\n")
			sValue = R(sValue, vbcrlf, "\n")
			sValue = R(sValue, "'", "\'")
		End If
		fnGetScriptValueProc = sValue
	End Function

'=========== JSON Value 처리 Proc =====================================
	Function fnGetJsonValueProc(sValue)
		If sValue <> "" Then
			sValue = R(sValue, """", "\""")
			sValue = R(sValue, "\", "\\")
			sValue = R(sValue, chr(13) & chr(10), "\n")
		End If
		fnGetJsonValueProc = sValue
	End Function

'=========== 모든 테그 삭제 ==========================================
	Function removeHTML(strHTML)
		Dim rex
		Set rex				= New Regexp
		rex.IgnoreCase		= True
		rex.Global			= True
		rex.Pattern			= "<[^<|>]*>"
		removeHTML			= rex.Replace(strHTML, "")
		Set rex 				= Nothing
	End Function

'=========== 선택 테그 삭제 ==========================================
	Function StripTags(htmlDoc, removeTag)
		Dim rex
		Set rex				= new Regexp
		rex.Pattern			= "<"& removeTag &">"
		rex.IgnoreCase		= True
		rex.Global			= True
		htmlDoc				= rex.Replace(htmlDoc,"")
		rex.Pattern			= "</"& removeTag &">"
		StripTags				= rex.Replace(htmlDoc,"")
		Set rex				= Nothing
	End Function

'=========== IMG 테그 경로 추출 =======================================
	Function StripTagsIMG(intPatrn, htmlDoc)
		Dim ObjRegExp, Matches, Match, RetStr, strPatrn , intNUM
		
		Select case intPatrn
			Case 1 : strPatrn = "<img [^<>]*>"			'이미지 태그만 추출 패턴
			Case 2 : strPatrn = "[^=']*\.(gif|jpg|bmp)"		'이미지 경로 전체 추출 패턴
			Case 3 : strPatrn = "[^='/]*\.(gif|jpg|bmp)"	'이미지 파일명만 추출 패턴
		End Select
		
		Set ObjRegExp = New RegExp
			ObjRegExp.Pattern		= strPatrn	' 정규 표현식 패턴
			ObjRegExp.Global		= True		' 문자열 전체를 검색함
			ObjRegExp.IgnoreCase	= True		' 대.소문자 구분 안함
			Set Matches			= ObjRegExp.Execute(htmlDoc)
		
			RetStr	= ""
			intNUM	= 0
			For Each Match in Matches 'Matches 컬렉션을 반복
				if intNUM = 0 then
					RetStr	= Match
					intNUM	= intNUM + 1
				else
					RetStr	= RetStr & "|//|" & Match
				end if
			Next
			
			if intPatrn = 2 then
				RetStr = R(RetStr, """", "")
			end if
			
			StripTagsIMG = RetStr
		Set ObjRegExp = Nothing
	End Function

'=========== 문자열 비교후 리턴값 돌려줌 ===============================
	Public Function CV(strV1, strV2, strReturn)
		If IsE(strV1) Then strV1 = "" End If
		If IsE(strV2) Then strV2 = "" End If
		
		If strV1 = strV2 Then CV = strReturn Else CV = strV1 
	End Function

'=========== 날짜 변환 ==============================================
	Function D(getDate, getType)
		select case getType
			case 1
				'0000-00-00 00:00 
				getDate	= year(getDate) & "-" & right("0" & month(getDate),2) & "-" & right("0" & day(getDate),2) & "  " & right("0" & hour(getDate),2) & ":" & right("0" & minute(getDate),2)
			case 2
				'00/00 00:00
				getDate	= right("0" & month(getDate),2) &"/"& right("0" & day(getDate),2) &" "& right("0" & hour(getDate),2) & ":" & right("0" & minute(getDate),2)
			case 3
				'0000-00-00
				getDate	= year(getDate) & "-" & right("0" & month(getDate),2) & "-" & right("0" & day(getDate),2)
			case 4
				'00000/00/00 00:00 
				getDate	= year(getDate) & "/" & right("0" & month(getDate),2) & "/" & right("0" & day(getDate),2) & "  " & right("0" & hour(getDate),2) & ":" & right("0" & minute(getDate),2)
			case 5
				'0000.00.00
				getDate	= year(getDate) & "." & right("0" & month(getDate),2) & "." & right("0" & day(getDate),2)
			case 6
				'0000년 0월 0일
				getDate	= year(getDate) & "년 " & month(getDate) & "월 " & day(getDate) & "일"
			case 7
				'0000년 0월 0일 00시 00분
				getDate	= year(getDate) & "년 " & month(getDate) & "월 " & day(getDate) & "일" & "  " & hour(getDate) & "시 " & minute(getDate) &"분"
			case 8
				'0000년 0월
				getDate	= year(getDate) & "년 " & month(getDate) & "월"
			case else
				'0000-00-00 00:00 
				getDate	= year(getDate) & "-" & right("0" & month(getDate),2) & "-" & right("0" & day(getDate),2) & "  " & right("0" & hour(getDate),2) & ":" & right("0" & minute(getDate),2)
		End Select
		D = getDate
	End Function

'=========== 날짜 변환 ==============================================
	Function D2(getDate)
		If IsE(getDate) Then
			D2 = ""
		else
			D2 = mid(getDate, 1, 4) &"-"& mid(getDate, 5, 2) &"-"& mid(getDate, 7, 2)
		End if
		'D2 = getDate
	End Function

'=========== 10 이하의 숫자를 2자리로 치환하는 Proc =======================
	Function fnGetCipherProc(Number)
		Dim strTemp : strTemp = Number
		If IsNumeric(strTemp) Then
			If strTemp = 0 Then
				strTemp = "00"
			ElseIf Len(strTemp) = 1 Then
				strTemp = "0" & strTemp
			End if
		End If
		fnGetCipherProc = strTemp
	End Function

'=========== 몫, 나머지 구하기 Proc ====================================
	Function fnShareReminderProc(Number,DevideNum)
		Dim Share
		Dim Reminder
		Share    = Fix(Number / DevideNum)
		Reminder = Number Mod DevideNum
		fnShareReminderProc = Share & "," & Reminder
	End Function

'=========== 숫자인지 체크 ====================================
Function CheckNumber(in_Value)
	Dim i
	Dim iValueLen : iValueLen = Len(in_Value)
	Dim chValueChr

	' 문자열이 모두 숫자인지 검사한다
	i = 1
	While(i <= iValueLen)
		chValueChr = Mid(in_Value, i, 1)
		If CStr(chValueChr) < "0" Or CStr(chValueChr) > "9" Then
			' 문자열에 숫자가 아닌 다른 문자가 들어있을 경우
			CheckNumber = False
			Exit Function
		End If

		i = i + 1
	Wend

	CheckNumber = True
End Function

'=========== True/False = 1/0으로 변환 Proc =======================
	Function fnConvertTrueFalse(strTemp)
		If strTemp = "True" Then
			fnConvertTrueFalse = "1"
		Else
			fnConvertTrueFalse = "0"
		End If
	End Function

'=========== SQL Injection String 문자 변환 Function ========================
	Function fnGetStringCheckProc(strTemp)
		strTemp = Trim(strTemp)
		strTemp = R(strTemp , "'" , "''")
		strTemp = R(strTemp , "--" , "")
		strTemp = R(strTemp , "sysobjects" , "")
		strTemp = R(strTemp , "syscolumns" , "")
		strTemp = R(strTemp , "sysindexes" , "")
		strTemp = R(strTemp , "sysusers" , "")
		strTemp = R(strTemp , "systypes" , "")
		strTemp = R(strTemp , "sysreferences" , "")
		strTemp = R(strTemp , "sysprotects" , "")
		strTemp = R(strTemp , "sysproperties" , "")
		strTemp = R(strTemp , "syspermissions" , "")
		strTemp = R(strTemp , "sysmembers" , "")
		strTemp = R(strTemp , "sysindexkeys" , "")
		strTemp = R(strTemp , "sysfulltextnotify" , "")
		strTemp = R(strTemp , "sysfulltextcatalogs" , "")
		strTemp = R(strTemp , "sysforeignkeys" , "")
		strTemp = R(strTemp , "sysfiles1" , "")
		strTemp = R(strTemp , "sysfiles" , "")
		strTemp = R(strTemp , "sysfilegroups" , "")
		strTemp = R(strTemp , "sysdepends" , "")
		strTemp = R(strTemp , "syscomments" , "")
		strTemp = R(strTemp , "<script" , "")
		strTemp = R(strTemp , "</script" , "")
		fnGetStringCheckProc = strTemp
	End Function


'=========== 랜덤 추출 (숫자+문자) ====================================
	Function RandomInput(strRandomText, intRandomLen)
		Dim strRandom : strRandom = strRandomText
		Dim intRandomLength : intRandomLength = len(strRandom)
		Dim strRandomResult, intRandom
		
		Randomize
		For intRandom = 1 to intRandomLen
			strRandomResult = strRandomResult & Mid(strRandom, Int((intRandomLength - 1 + 1) * Rnd + 1),1)
		Next
		
		RandomInput = strRandomResult
	End Function

'=========== 랜덤 추출 (숫자+문자) ====================================
	Function Random(intRandomLen)
		Dim strRandom : strRandom = "abcdefghijklmnopqrstuvwxyz0123456789"
		Dim strRandomResult, intRandom	
		
		Randomize
		For intRandom = 1 to intRandomLen
			strRandomResult = strRandomResult & Mid(strRandom, Int((36 - 1 + 1) * Rnd + 1),1)
		Next
		
		Random = strRandomResult
	End Function

'=========== 랜덤 추출 (숫자) =========================================
	Function RandomNumber(intRandomLen)
		Dim strRandom : strRandom = "0123456789"
		Dim strRandomResult, intRandom	
		
		Randomize
		For intRandom = 1 to intRandomLen
			strRandomResult = strRandomResult & Mid(strRandom, Int((10 - 1 + 1) * Rnd + 1),1)
		Next
		
		RandomNumber = strRandomResult
	End Function

'=========== 난수발생 ==============================================
	Function MakeRandomize(startnum, endnum)
		Randomize    '난수 발생기를 초기화합니다.
		MakeRandomize = Int((endnum * Rnd) + startnum)    ' startnum 에서 endnum 까지 무작위 값을 발생합니다.
	End Function

'=========== 소수점 체크 ==============================================
	Function SosuCheck(Num, SosuLen)
		Dim aryNum : aryNum = Split(Num, ".")
		
		If Ubound(aryNum) = 0 then
			SosuCheck = Num
		Else
			If Not(IsE(Num)) then
				SosuCheck = FormatNumber(Num, SosuLen)
			Else
				SosuCheck = Num
			End If
		End if
	End Function

'=========== 문자열 URLEncoding ======================================
	Function fnGetURLEncode(url)
		fnGetURLEncode = Server.URLEncode(url)
	End Function

'=========== 문자열 URLEncoding ======================================
    Function fnURLEncode(URLStr)

	    Dim sURL                '** 입력받은 URL 문자열
	    Dim sBuffer             '** Encoding 중의 URL 을 담을 Buffer 문자열
	    Dim sTemp               '** 임시 문자열
	    Dim cChar               '** URL 문자열 중의 현재 Index 의 문자

	    Dim Index

	    On Error Resume Next

	    Err.Clear
        sURL = Trim(URLStr)     '** URL 문자열을 얻는다.
        sBuffer = ""            '** 임시 Buffer 용 문자열 변수 초기화.


        '******************************************************
        '* URL Encoding 작업
        '******************************************************

        For Index = 1 To Len(sURL)
            '** 현재 Index 의 문자를 얻는다.
            cChar = Mid(sURL, Index, 1)

            If cChar = "0" Or _
               (cChar >= "1" And cChar <= "9") Or _
               (cChar >= "a" And cChar <= "z") Or _
               (cChar >= "A" And cChar <= "Z") Or _
               cChar = "-" Or _
               cChar = "_" Or _
               cChar = "." Or _
               cChar = "*" Then

                '** URL 에 허용되는 문자들 :: Buffer 문자열에 추가한다.
                sBuffer = sBuffer & cChar

            ElseIf cChar = " " Then
                '** 공백 문자 :: + 로 대체하여 Buffer 문자열에 추가한다.
                sBuffer = sBuffer & "+"
            Else
                '** URL 에 허용되지 않는 문자들 :: % 로 Encoding 해서 Buffer 문자열에 추가
                sTemp = CStr(Hex(Asc(cChar)))

                If Len(sTemp) = 4 Then
                    sBuffer = sBuffer & "%" & Left(sTemp, 2) & "%" & Mid(sTemp, 3, 2)
                ElseIf Len(sTemp) = 2 Then
                    sBuffer = sBuffer & "%" & sTemp
                End If
            End If
        Next


        '** Error 처리
        If Err.Number > 0 Then
            fnURLEncode = ""
            Exit Function
        End If

        '** 결과를 리턴한다.
        fnURLEncode = sBuffer
        Exit Function

    End Function

'=========== 문자열 URLDecoding ======================================
    Function fnURLDecode(URLStr)

        Dim sURL                '** 입력받은 URL 문자열
        Dim sBuffer             '** Decoding 중의 URL 을 담을 Buffer 문자열
        Dim cChar               '** URL 문자열 중의 현재 Index 의 문자

        Dim Index
        Dim s,bUnicode

	    On Error Resume Next

        Err.Clear
        sURL = Trim(URLStr)     '** URL 문자열을 얻는다.
        sBuffer = ""            '** 임시 Buffer 용 문자열 변수 초기화.

        '******************************************************
        '* URL Decoding 작업
        '******************************************************
        '한글이 입력되는 경우가 2가지 있음
        '	"%C0%DA" ->일반 
        '	"%uC911" ->유니코드 chr 대신 chrW 사용

        Index = 1
	
	    '?,%가 없다면 검색할 필요없음
	    if instr(1,pURL,"?",1) > 0 OR instr(1,pURL,"%",1) > 0 then
	        Do While Index <= Len(sURL)
	            cChar = Mid(sURL, Index, 1)
	            If cChar = "+" Then        
	                '** '+' 문자 :: ' ' 로 대체하여 Buffer 문자열에 추가한다.
	                sBuffer = sBuffer & " "
	                Index = Index + 1            
	            ElseIf cChar = "%" Then        
	                '** '%' 문자 :: Decoding 하여 Buffer 문자열에 추가한다.
	            
	                '유니코드인지 판단한 후에 일반문자인지 한글인지 구분함
	                bUnicode = false
	            
	                s = Mid(sURL, Index + 1, 5)
	                bUnicode = boolUnicode(s)
	            
	                if bUnicode = true then
	            	    cChar = Mid(sURL, Index + 2, 2)
	                else
	            	    cChar = Mid(sURL, Index + 1, 2)
	                end if
	
	                If CInt("&H" & cChar) < &H80 Then
	                    '** 일반 ASCII 문자
	                    sBuffer = sBuffer & Chr(CInt("&H" & cChar))
	                    Index = Index + 3
	                Else
	                    '** 2 Byte 한글 문자
	                    cChar = Replace(Mid(sURL, Index + 1, 5), "%", "")
	                
	                    '유니코드인경우 맨앞의 u자를 제거해야 함
	                    if bUnicode = true then		'C0DA
	                	    cChar = mid(cChar,2,4)
	                	    sBuffer = sBuffer & ChrW(CInt("&H" & cChar))
	                    else						'uC911
	                	    sBuffer = sBuffer & Chr(CInt("&H" & cChar))
	                    end if
	                    Index = Index + 6
	                End If
	            Else
	                '** 그 외의 일반 문자들 :: Buffer 문자열에 추가한다.
	                sBuffer = sBuffer & cChar
	                Index = Index + 1
	            End If
	        Loop
	    else
		    sBuffer = ""
	    end if
	
        '** Error 처리
        If Err.Number > 0 Then
            fnURLDecode = ""
            Exit Function
        End If

        '** 결과를 리턴한다.
        fnURLDecode = sBuffer
        Exit Function

    End Function

'=========== 한글까지 되는 URLDecode ==================================
	Function URLDecode(str)
		Dim strSource, strTemp, strResult, strchr
		Dim lngPos, AddNum, IFKor
	
		strSource = R(str, "+", " ")
	
		For lngPos = 1 To Len(strSource)
			AddNum = 2
			strTemp = Mid(strSource, lngPos, 1)
	
			If strTemp = "%" Then
				If lngPos + AddNum < Len(strSource) + 1 Then
					strchr = CInt("&H" & Mid(strSource, lngPos + 1, AddNum))
					If strchr > 130 Then
						AddNum = 5
						IFKor = Mid(strSource, lngPos + 1, AddNum)
						IFKor = R(IFKor, "%", "")
						strchr = CInt("&H" & IFKor )
					End If
					strResult = strResult & Chr(strchr)
					lngPos = lngPos + AddNum
				End If
			Else
				strResult = strResult & strTemp
			End If
		Next
	
		URLDecode = strResult
	End Function

'=========== 아스키 -> UTF8 =========================================
'---------------------------------------------------------------------
'	URLEncodeUTF8 (아스키 -> UTF8)
'	Devpia.com 고일호(n4kjy)님 - 2002-07-24
'---------------------------------------------------------------------
	Public Function URLEncodeUTF8(byVal szSource)
	
		Dim szChar, WideChar, nLength, i, result
		nLength = Len(szSource)
	
		szSource = Replace(szSource," ","+")
	
		For i = 1 To nLength
			szChar = Mid(szSource, i, 1)
	
			If Asc(szChar) < 0 Then             
				WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))
	
				If (WideChar And &HFF80) = 0 Then
					result = result & "%" & Hex(WideChar)
				ElseIf (WideChar And &HF000) = 0 Then
					result = result & _
					"%" & Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
					"%" & Hex(WideChar And &H3F Or &H80)
				Else
					result = result & _
					"%" & Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
					"%" & Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
					"%" & Hex(WideChar And &H3F Or &H80)
				End If
			Else
				result = result + szChar
			End If
		Next
		URLEncodeUTF8 = result
	End Function

'=========== UTF8 --> 아스키 =========================================
'---------------------------------------------------------------------
'	URLDecodeUTF8 (UTF8 --> 아스키 )
'	mongmong - 2003. 10 (URLEncodeUTF8 참조)
'---------------------------------------------------------------------
	Public Function URLDecodeUTF8(byVal pURL)
		Dim i, s1, s2, s3, u1, u2, result
		pURL = Replace(pURL,"+"," ")
		pURL = Replace(pURL,"%20"," ")
		pURL = Replace(pURL,"%2C",",")
		
		'?,%가 없다면 검색할 필요없음
		if instr(1,pURL,"?",1) > 0 OR instr(1,pURL,"%",1) > 0 then
				For i = 1 to Len(pURL)
					if Mid(pURL, i, 1) = "%" then
						s1 = CLng("&H" & Mid(pURL, i + 1, 2))
			
						'2바이트일 경우
						if ((s1 AND &HC0) = &HC0) AND ((s1 AND &HE0) <> &HE0) then
							s2 = CLng("&H" & Mid(pURL, i + 4, 2))
			
							u1 = (s1 AND &H1C) / &H04
							u2 = ((s1 AND &H03) * &H04 + ((s2 AND &H30) / &H10)) * &H10
							u2 = u2 + (s2 AND &H0F)
							result = result & ChrW((u1 * &H100) + u2)
							i = i + 5
			
						'3바이트일 경우
						elseif (s1 AND &HE0 = &HE0) then
							s2 = CLng("&H" & Mid(pURL, i + 4, 2))
							s3 = CLng("&H" & Mid(pURL, i + 7, 2))
			
							u1 = ((s1 AND &H0F) * &H10)
							u1 = u1 + ((s2 AND &H3C) / &H04)
							u2 = ((s2 AND &H03) * &H04 +  (s3 AND &H30) / &H10) * &H10
							u2 = u2 + (s3 AND &H0F)
							result = result & ChrW((u1 * &H100) + u2)
							i = i + 8
						end if
					else
						result = result & Mid(pURL, i, 1)
					end if
				Next
		else
			result = pURL
		end if
		URLDecodeUTF8 = result
	End Function

'=========== Base64  ===============================================
' Functions to provide encoding/decoding of strings with Base64.
' 
' Encoding: myEncodedString = base64_encode( inputString )
' Decoding: myDecodedString = base64_decode( encodedInputString )
'
' Programmed by Markus Hartsmar for ShameDesigns in 2002. 
' Email me at: mark@shamedesigns.com
' Visit our website at: http://www.shamedesigns.com/

Public Function base64_encode( byVal strIn )
	Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
	
	For n = 1 To Len( strIn ) Step 3
		c1 = Asc( Mid( strIn, n, 1 ) )
		c2 = Asc( Mid( strIn, n + 1, 1 ) + Chr(0) )
		c3 = Asc( Mid( strIn, n + 2, 1 ) + Chr(0) )
		w1 = Int( c1 / 4 ) : w2 = ( c1 And 3 ) * 16 + Int( c2 / 16 )
		If Len( strIn ) >= n + 1 Then 
			w3 = ( c2 And 15 ) * 4 + Int( c3 / 64 ) 
		Else 
			w3 = -1
		End If
		If Len( strIn ) >= n + 2 Then 
			w4 = c3 And 63 
		Else 
			w4 = -1
		End If
		strOut = strOut + mimeencode( w1 ) + mimeencode( w2 ) + _
				  mimeencode( w3 ) + mimeencode( w4 )
	Next
	base64_encode = strOut
End Function

Private Function mimeencode( byVal intIn )
	Dim Base64Chars : Base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	
	If intIn >= 0 Then 
		mimeencode = Mid( Base64Chars, intIn + 1, 1 ) 
	Else 
		mimeencode = ""
	End If
End Function	


' Function to decode string from Base64
Public Function base64_decode( byVal strIn )
	Dim w1, w2, w3, w4, n, strOut
	For n = 1 To Len( strIn ) Step 4
		w1 = mimedecode( Mid( strIn, n, 1 ) )
		w2 = mimedecode( Mid( strIn, n + 1, 1 ) )
		w3 = mimedecode( Mid( strIn, n + 2, 1 ) )
		w4 = mimedecode( Mid( strIn, n + 3, 1 ) )
		If w2 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w1 * 4 + Int( w2 / 16 ) ) And 255 ) )
		If w3 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w2 * 16 + Int( w3 / 4 ) ) And 255 ) )
		If w4 >= 0 Then _
			strOut = strOut + _
				Chr( ( ( w3 * 64 + w4 ) And 255 ) )
	Next
	base64_decode = strOut
End Function

Private Function mimedecode( byVal strIn )
	Dim Base64Chars : Base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
	
	If Len( strIn ) = 0 Then 
		mimedecode = -1 : Exit Function
	Else
		mimedecode = InStr( Base64Chars, strIn ) - 1
	End If
End Function

'=========== SELECTED 리턴 ==============================================
Function setSelected(source, compare)
	If IsNull(source) Then
		source = ""
	End If

	If IsNull(compare) Then
		compare = ""
	End If

	If CStr(source) = CStr(compare) Then
		setSelected = "SELECTED"
	else
		setSelected = ""
	End If
End Function

'=========== CHECKED 리턴 ==============================================
    Function setChecked(source, compare)
        If source = compare Then
            setChecked = "checked"
        else
			setChecked = ""
		End If
    End Function

'=========== 변수 타입 추출 ==============================================
Function GetDataTypeEnum(TypeEnum)
	'// adUserDefined 사용자 정의

	Select Case TypeEnum
		Case 0: GetDataTypeEnum = "adEmpty"
		Case 2: GetDataTypeEnum = "adSmallInt"
		Case 3: GetDataTypeEnum = "adInteger"
		Case 4: GetDataTypeEnum = "adSingle"
		Case 5: GetDataTypeEnum = "adDouble"
		Case 6: GetDataTypeEnum = "adCurrency"
		Case 7: GetDataTypeEnum = "adDate"
		Case 8: GetDataTypeEnum = "adBSTR"
		Case 9: GetDataTypeEnum = "adIDispatch"
		Case 10: GetDataTypeEnum = "adError"
		Case 11: GetDataTypeEnum = "adBoolean"
		Case 12: GetDataTypeEnum = "adVariant"
		Case 13: GetDataTypeEnum = "adIUnknown"
		Case 14: GetDataTypeEnum = "adDecimal"
		Case 16: GetDataTypeEnum = "adTinyInt"
		Case 17: GetDataTypeEnum = "adUnsignedTinyInt"
		Case 18: GetDataTypeEnum = "adUnsignedSmallInt"
		Case 19: GetDataTypeEnum = "adUnsignedInt"
		Case 20: GetDataTypeEnum = "adBigInt"
		Case 21: GetDataTypeEnum = "adUnsignedBigInt"

		Case 64: GetDataTypeEnum = "adFileTime"
		Case 72: GetDataTypeEnum = "adGUID"

		Case 128: GetDataTypeEnum = "adBinary"
		Case 129: GetDataTypeEnum = "adChar"
		Case 130: GetDataTypeEnum = "adWChar"
		Case 131: GetDataTypeEnum = "adNumeric"
		Case 132: GetDataTypeEnum = "adUserDefined"
		Case 133: GetDataTypeEnum = "adDBDate"
		Case 134: GetDataTypeEnum = "adDBTime"
		Case 135: GetDataTypeEnum = "adDBTimeStamp"
		Case 136: GetDataTypeEnum = "adChapter"
		Case 138: GetDataTypeEnum = "adPropVariant"
		Case 139: GetDataTypeEnum = "adVarNumeric"

		Case 200: GetDataTypeEnum = "adVarChar"
		Case 201: GetDataTypeEnum = "adLongVarChar"
		Case 202: GetDataTypeEnum = "adVarWChar"
		Case 203: GetDataTypeEnum = "adLongVarWChar"
		Case 204: GetDataTypeEnum = "adVarBinary"
		Case 205: GetDataTypeEnum = "adLongVarBinary"

		Case 8192 : GetDataTypeEnum = "adArray"

		Case Else: GetDataTypeEnum = "<b>변수 타입을 알수 없습니다. 프로시저를 확인해 주세요</b>"
	End Select
End function
	
'=========== 자바 스크립스트 제어 =====================================
Sub locationReplace(Url)
	Call closeDb()
	Response.Write "<script language='javascript'>location.replace('"& Url &"');</script>"
	Response.End()
End sub

Sub alert(Msg)
	Response.Write "<script language='javascript'>alert('"& Msg &"');</script>"
End sub

Sub alertB(Msg)
	Response.Write "<script language='javascript'>alert('"& Msg &"');history.back();</script>"
	Response.End
End sub

Sub alertGO(Msg, Url)
	If Msg <> "" Then
		Response.Write "<script language='javascript'>alert('"& Msg &"');location.replace('"& Url &"');</script>"
	Else
		Response.Write "<script language='javascript'>location.replace('"& Url &"');</script>"
	End if
	Response.End
End Sub

Sub alertC(Msg)
	Response.Write "<script language='javascript'>alert('"& Msg &"');self.close();</script>"
End sub

Sub alertRe(Msg)
	Response.Write "<script language='javascript'>alert('"& Msg &"');document.reload();</script>"
End sub



'////////////////////////////////////////////////////////////////////
'//////// 명령어 삽입(Command Injection) 가능성
'////////////////////////////////////////////////////////////////////
function ReplaceStr(rcontent)
	If IsNull(rcontent) Then
		rcontent = ""
	End if
	
	rcontent = replace(rcontent, "'", "''")
	rcontent = replace(rcontent, """", """""")
	rcontent = replace(rcontent, "\", "\\")
	rcontent = replace(rcontent, "#", "\#")
	rcontent = replace(rcontent, "--", "\--")
	'    rcontent = replace(rcontent, ";", "\;")
	rcontent = replace(rcontent, "%", "%%")
	ReplaceStr = rcontent
end Function

function ReplaceStr_de(rcontent)
	If IsNull(rcontent) Then
		rcontent = ""
	End if
	
	rcontent = replace(rcontent, "''", "'")
	rcontent = replace(rcontent, """""", """")
	rcontent = replace(rcontent, "\\", "\")
	rcontent = replace(rcontent, "\#", "#")
	rcontent = replace(rcontent, "\--", "--")
	'    rcontent = replace(rcontent, ";", "\;")
	rcontent = replace(rcontent, "%%", "%")
	ReplaceStr_de = rcontent
end function

function XSS(get_String)
	get_String = REPLACE( get_String, "&", "&amp;" )
	get_String = REPLACE( get_String, "<xmp", "<x-xmp")
	get_String = REPLACE( get_String, "javascript", "x-javascript")
	get_String = REPLACE( get_String, "script", "x-script")
	get_String = REPLACE( get_String, "iframe", "x-iframe")
	get_String = REPLACE( get_String, "document", "x-document")
	get_String = REPLACE( get_String, "vbscript", "x-vbscript")
	get_String = REPLACE( get_String, "applet", "x-applet")
	get_String = REPLACE( get_String, "embed", "x-embed")
	get_String = REPLACE( get_String, "object", "x-object")
	get_String = REPLACE( get_String, "frame", "x-frame")
	get_String = REPLACE( get_String, "frameset", "x-frameset")
	get_String = REPLACE( get_String, "layer", "x-layer")
	get_String = REPLACE( get_String, "bgsound", "x-bgsound")
	get_String = REPLACE( get_String, "alert", "x-alert")
	get_String = REPLACE( get_String, "onblur", "x-onblur")
	get_String = REPLACE( get_String, "onchange", "x-onchange")
	get_String = REPLACE( get_String, "onclick", "x-onclick")
	get_String = REPLACE( get_String, "ondblclick", "x-ondblclick")
	get_String = REPLACE( get_String, "onerror", "x-onerror")
	get_String = REPLACE( get_String, "onfocus", "x-onfocus")
	get_String = REPLACE( get_String, "onload", "x-onload")
	get_String = REPLACE( get_String, "onmouse", "x-onmouse")
	get_String = REPLACE( get_String, "onscroll", "x-onscroll")
	get_String = REPLACE( get_String, "onsubmit", "x-onsubmit")
	get_String = REPLACE( get_String, "onunload", "x-onunload")
	
	'   if not get_HTML then
	get_String = REPLACE( get_String, "<", "&lt;" )
	get_String = REPLACE( get_String, ">", "&gt;" )
	'   end if
	XSS = get_String
end Function

function HTML_XSS(get_String)
	get_String = REPLACE( get_String, "&", "&amp;" )
	get_String = REPLACE( get_String, "<xmp", "<x-xmp")
	get_String = REPLACE( get_String, "javascript", "x-javascript")
	get_String = REPLACE( get_String, "script", "x-script")
	get_String = REPLACE( get_String, "iframe", "x-iframe")
	get_String = REPLACE( get_String, "document", "x-document")
	get_String = REPLACE( get_String, "vbscript", "x-vbscript")
	get_String = REPLACE( get_String, "applet", "x-applet")
	get_String = REPLACE( get_String, "embed", "x-embed")
	get_String = REPLACE( get_String, "object", "x-object")
	get_String = REPLACE( get_String, "frame", "x-frame")
	get_String = REPLACE( get_String, "frameset", "x-frameset")
	get_String = REPLACE( get_String, "layer", "x-layer")
	get_String = REPLACE( get_String, "bgsound", "x-bgsound")
	get_String = REPLACE( get_String, "alert", "x-alert")
	get_String = REPLACE( get_String, "onblur", "x-onblur")
	get_String = REPLACE( get_String, "onchange", "x-onchange")
	get_String = REPLACE( get_String, "onclick", "x-onclick")
	get_String = REPLACE( get_String, "ondblclick", "x-ondblclick")
	get_String = REPLACE( get_String, "onerror", "x-onerror")
	get_String = REPLACE( get_String, "onfocus", "x-onfocus")
	get_String = REPLACE( get_String, "onload", "x-onload")
	get_String = REPLACE( get_String, "onmouse", "x-onmouse")
	get_String = REPLACE( get_String, "onscroll", "x-onscroll")
	get_String = REPLACE( get_String, "onsubmit", "x-onsubmit")
	get_String = REPLACE( get_String, "onunload", "x-onunload")
	HTML_XSS = get_String
end Function

function fnReplaceInjection(get_String)
	get_String = REPLACE(get_String, "'", "''")
	get_String = REPLACE(get_String, """", """""")
	get_String = REPLACE(get_String, "\", "\\")
	get_String = REPLACE(get_String, "#", "\#")
	get_String = REPLACE(get_String, "--", "\--")
	get_String = REPLACE(get_String, "%", "%%")
	get_String = REPLACE( get_String, "&", "&amp;" )
	get_String = REPLACE( get_String, "<xmp", "<x-xmp")
	get_String = REPLACE( get_String, "javascript", "x-javascript")
	get_String = REPLACE( get_String, "script", "x-script")
	get_String = REPLACE( get_String, "iframe", "x-iframe")
	get_String = REPLACE( get_String, "document", "x-document")
	get_String = REPLACE( get_String, "vbscript", "x-vbscript")
	get_String = REPLACE( get_String, "applet", "x-applet")
	get_String = REPLACE( get_String, "embed", "x-embed")
	get_String = REPLACE( get_String, "object", "x-object")
	get_String = REPLACE( get_String, "frame", "x-frame")
	get_String = REPLACE( get_String, "frameset", "x-frameset")
	get_String = REPLACE( get_String, "layer", "x-layer")
	get_String = REPLACE( get_String, "bgsound", "x-bgsound")
	get_String = REPLACE( get_String, "alert", "x-alert")
	get_String = REPLACE( get_String, "onblur", "x-onblur")
	get_String = REPLACE( get_String, "onchange", "x-onchange")
	get_String = REPLACE( get_String, "onclick", "x-onclick")
	get_String = REPLACE( get_String, "ondblclick", "x-ondblclick")
	get_String = REPLACE( get_String, "onerror", "x-onerror")
	get_String = REPLACE( get_String, "onfocus", "x-onfocus")
	get_String = REPLACE( get_String, "onload", "x-onload")
	get_String = REPLACE( get_String, "onmouse", "x-onmouse")
	get_String = REPLACE( get_String, "onscroll", "x-onscroll")
	get_String = REPLACE( get_String, "onsubmit", "x-onsubmit")
	get_String = REPLACE( get_String, "onunload", "x-onunload")
	fnReplaceInjection = get_String
end function

Function StringToSQL(get_String)
	Dim text : text = get_String
	text = replace(text, "'", "`")
	text = replace(text, """", "``")
	
	text = replace(text, "'", "''")
	text = replace(text, "&", "&amp;")
	text = replace(text, "<", "&lt;")
	text = replace(text, ">", "&gt;")
	text = replace(text, """", "&quot;")
	text = replace(text, "?","&#63;")
	
	StringToSQL = text
End Function

Function StringToHTML(get_String)
	Dim text : text = get_String
	  If Not(IsE(text)) Then		  
		  text = Replace(text, "''", "'")
		  text = replace(text, "&amp;", "&")
		  text = replace(text, "&lt;", "<")
		  text = replace(text, "&gt;", ">")
		  text = replace(text, "&quot;", """")
		  text = replace(text, "&#63;","?")
	  Else
		text = get_String
	  End If
	
	StringToHTML = text
End Function


Function Char2SQL(ByVal Str, ByVal xType)
	Str = Trim(Str)
	
	IF Str = "" Then
		IF Lcase(xType) = "string" Then
			Str = "Null"
		Else
			Str = "0"
		End IF
	Else
		Str = Replace(Str, "'", "''")

'		IF Lcase(xType) = "string" Then
'			Str = "'"& Str & "'"
'		End IF
	End IF	
	
	Char2SQL = Str
End Function



Function Str2Num(ByVal v)
	If IsE(v) Then v = 0 End If
	Str2Num = FormatNumber(v, 0)
End Function

'배열로 넘어온 숫자를 합산한다.
Function SumNum(ByVal arrValue)
	Dim i, resultCnt
	If IsArray(arrValue) Then
		For i=0 To UBound(arrValue)
			If IsE(arrValue(i)) Then arrValue(i) = 0 End If
			If Not IsNumeric(arrValue(i)) Then arrValue(i) = 0 End If
			resultCnt = resultCnt + arrValue(i)
		Next
		SumNum = resultCnt
	Else
		SumNum = 0
	End If	
End Function
'////////////////////////////////////////////////////////////////////
%>