<%
'기본 세팅 설정

'//////////////////////////////////////////////////////////////////////////////////
'============= DB Connection Basic Info ======================
Dim strDBServerIP
Dim strDBName
Dim strDBUserID
Dim strDBPassword
Dim strDBConnString
Dim strDBConnString2

'============= DB Connection ===============================
strDBServerIP = "localhost"
strDBName = "LoInfo"
strDBUserID = "Master"
strDBPassword	= "Harang508!"

strDBConnString = "Provider=SQLOLEDB;Data Source=" & strDBServerIP & ";Initial Catalog=" & strDBName & ";user ID=" & strDBUserID & ";password=" & strDBPassword & ";"
'=======================================================
%>