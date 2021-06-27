//--------------------------------------------------------------------------------------------------
//용    도 : 파일 업로드 function
//파라미터 :
//리    턴 : 없음
//--------------------------------------------------------------------------------------------------
var PrdExcelImport = {
		fnPrdidFileUpload : function() {
		var imgPath = $("uploadMultiFile").value;
		var src = FileUtil.getFileExtension(imgPath);
		if((src.toLowerCase() == "xls")){
			var objForm = document.getElementById("frmExecute");
			document.frmExecute.target = "iframeFileUpload";
			document.frmExecute.action="/prd/popup/prdRegExcelImportPop.gs";
			document.frmExecute.submit();
		} else {
			alert("엑셀파일만 업로드가 가능합니다.");
			return;
		}
	}
};

var FileUtil = {
	/**
	* @param : 파일경로
	* @usage : FileUtil.getFileSize("path")
	* @description : 파일의 크기를 넘겨준다.
	*/
	getFileSize : function(path){
		var img = new Element('IMG');
		img.src = path;
		return img.fileSize;
	},
	/**
	* @param : 파일경로
	* @usage : FileUtil.getFileExtension("path")
	* @description : 파일의 확장자를 넘겨준다.
	*/
	getFileExtension : function(filePath){
			var extension = "";
			var lastIndex = -1;
			lastIndex = filePath.lastIndexOf('.');

		if(lastIndex != -1) {
			extension = filePath.substring( lastIndex+1, filePath.length );
		} else {
			extension = "";
		}
		return extension;
	}
};

