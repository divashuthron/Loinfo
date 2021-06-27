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
'strDBServerIP = "192.168.1.163"
'strDBServerIP = "210.102.151.105"
'//레알
'strDBServerIP = "210.102.151.72,4508"
'//테스트
strDBServerIP = "SQLMISS"
strDBName = "InformationSystem"
strDBUserID = "xteb"
strDBPassword	= "east12!@"

strDBConnString = "Provider=SQLOLEDB;Data Source=" & strDBServerIP & ";Initial Catalog=" & strDBName & ";user ID=" & strDBUserID & ";password=" & strDBPassword & ";"
'=======================================================
%>