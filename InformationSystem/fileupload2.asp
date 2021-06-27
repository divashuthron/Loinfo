<!-- #include virtual = "/Common/Include/CodePage65001.asp" -->
<%
	Function GetUniqueName(byRef strFileName, DirectoryPath)

		Dim strName, strExt
		strName = Mid(strFileName, 1, InstrRev(strFileName, ".") - 1) ' 확장자를 제외한 파일명을 얻는다.
		strExt = Mid(strFileName, InstrRev(strFileName, ".") + 1) '확장자를 얻는다

		Dim fso
		Set fso = Server.CreateObject("Scripting.FileSystemObject")

		Dim bExist : bExist = True 
		'우선 같은이름의 파일이 존재한다고 가정
		Dim strFileWholePath : strFileWholePath = DirectoryPath & "\" & strName & "." & strExt 
		'저장할 파일의 완전한 이름(완전한 물리적인 경로) 구성
		Dim countFileName : countFileName = 0 
		'파일이 존재할 경우, 이름 뒤에 붙일 숫자를 세팅함. 
		Do While bExist ' 우선 있다고 생각함.
			If (fso.FileExists(strFileWholePath)) Then ' 같은 이름의 파일이 있을 때
				countFileName = countFileName + 1 '파일명에 숫자를 붙인 새로운 파일 이름 생성
				strFileName = strName & "(" & countFileName & ")." & strExt
				strFileWholePath = DirectoryPath & "\" & strFileName
			Else
				bExist = False
			End If
		Loop
		'GetUniqueName = strFileWholePath
		GetUniqueName = strFileName
	End Function

	dim dxUp
	set dxUp =  Server.CreateObject("DEXT.FileUpload")
	dxUp.DefaultPath = Server.MapPath ("/upload/Files/") & "\"

	dim filename1
	'filename1=GetUniqueName(trim(dxUp("FileUpload")),dxUp.DefaultPath)
	filename1=GetUniqueName(trim(dxUp("FileUpload")),dxUp.DefaultPath)
	dxUp("FileUpload").Save, False

	Dim ReturnID
	ReturnID=trim(dxUp("ReturnID"))

'response.write "filename1: " & filename1 & "<br/>"
'response.write "ReturnID: " & ReturnID & "<br/>"
%>
<script type="text/javascript" src="/js/jquery-3.1.1.min.js"></script>
<script language="javascript" type="text/javascript"> 
var fileNm = "<%=filename1%>"
 
if (fileNm != "") {
    var ext = fileNm.slice(fileNm.lastIndexOf(".") + 1).toLowerCase();
	
	jQuery(document).ready(function() {
		parent.$("#FilesName").append("<div class='filenameClass' ><span><%=filename1%><input type='hidden' name='filename' value='<%=filename1%>' /></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='#' class='deleteBtn'><span><i class='fa fa-times-circle'></i></span></a></div>");
	});
}
</script>