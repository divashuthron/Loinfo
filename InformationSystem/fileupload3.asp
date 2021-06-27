<!-- #include virtual = "/Common/Include/CodePage65001.asp" -->
<%
	dim dxUp
	set dxUp =  Server.CreateObject("DEXT.FileUpload")
	dxUp.DefaultPath = Server.MapPath ("/upload/Editer/") & "\"
	dxUp.CodePage = 65001

	dim filename1
	filename1=trim(dxUp("callbackfile").FileName)
	dxUp("callbackfile").Save, False
	'dxUp.SaveAs(dxUp.DefaultPath & "" & filename1 , False)

	'response.write "{""file"":"""& dxUp.LastSavedFileName &"""}"	'중복검사후 새이름으로 저장된 파일이름 

	response.write "/upload/Editer/"& dxUp.LastSavedFileName
%>