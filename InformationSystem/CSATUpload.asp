<!-- #InClude Virtual = "/Common/PopUp_Header.asp" -->
	<script type="text/javascript">
		//로딩 기본값
		var Uploading = false
		//파일 확인
		$(document).on("click", "#FileUpload", function(){
			if (Uploading)
			{
				alert("확인중입니다. 잠시 기다리세요.");
				return;
			}
			if(!$("#callbackfile").val()){
				alert("파일을 선택하세요.");
				$("#callbackfile").focus();
				return;
			}
			//엑셀파일만 받기
			var imgPath = $("#callbackfile").val();
			var src = FileUtil.getFileExtension(imgPath);
			if((src.toLowerCase() != "xls" && src.toLowerCase() != "xlsx")){
				alert("엑셀파일만 업로드가 가능합니다.");
				return;
			}
			var datasStr = ""
			var $objForm = $("#frmFile");
			$($objForm).ajaxForm({
				url: "/CSATFileupload.asp"
				, type: "post"
				, dataType: "text"
				, success: function (datas, state) {
					//파일내용 출력
					$(".viewTable").empty().html(datas);
					//수능성적 건수 출력
					datasStr = datas.split("<count>");
					$(".CountTable").html("<h5>목록 - 전체 " + datasStr[1] + "건</h5>");
					//입학원서 건수 넣어주기(저장 시 활용)
					$("#CSATCount").val(datasStr[1]);
					//로딩이미지 제거
					document.getElementById("Prog").style.display = "none";
					Uploading = false;	
					//수능성적 저장 버튼 출력(헷갈릴 수 있으므로, 파일 확인 시 출력)
					$("#CSATSaveBtn").css("display","block");								
				}
				, error: function (reason, e) {
					alert('파일확인에 실패했습니다. : ' + e);
				}
			});
			$objForm.submit();
			document.getElementById("Prog").style.display = "block";
			Uploading = true
		});		
		//확인 파일 저장
		$(document).on("click", "#CSATSaveBtn", function(){
			if (Uploading)
			{
				alert("등록중입니다. 잠시 기다리세요.");
				return;
			}
			var $objForm = $("#frmFile");
			$objForm.attr("enctype","");
			$($objForm).ajaxForm({
				url: "/Process/CSATfileuploadSave.asp"
				, type: "post"
				, dataType: "text"
				, success: function (datas, state) {
					alert('파일등록을 완료했습니다.');
					//로딩이미지 제거
					document.getElementById("Prog").style.display = "none";
					Uploading = false;	
					//페이지 리로드	
					opener.window.location.reload();
					//팝업창 제거
					close();
				}
				, error: function (reason, e) {
					alert('파일등록에 실패했습니다. 필수입력항목 기재 여부와 중복된 입학원서가 없는지 확인하세요. : ' + e);
				}
			});
			$objForm.submit();
			document.getElementById("Prog").style.display = "block";
			Uploading = true
		});
	</script>
</head>
<body style="overflow:hidden;">
<FORM ENCTYPE="multipart/form-data" ID="frmFile" METHOD="post" NAME="frmFile">
	<input type="hidden" id="CSATCount" name="CSATCount" value="">
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 파일확인 -->
			<div class="ibox-title">
				<h2>수능성적 엑셀 업로드</h2>
			</div>

			<div class="ibox-content">
				<div>
					<div class="row show-grid" >
						<div class="col-md-5 col-xs-1 ">
	                        <input type="file" name="callbackfile" id="callbackfile" style="width: 440px; margin: 0px; padding: 0px; padding: 0px; border: 1px solid; height: 25px;"/>
						</div>
						<div class="col-md-1 col-xs-1">
							<span class="btnBasic btnTypeSave" id="FileUpload">파일확인</span>
						</div>
					</div>
				</div>
			</div>
			<!-- 파일확인 끝-->

			<!-- 뷰어 -->
			<div class="ibox-title">
				<div class="CountTable"></div>
				<div style="float:right;">
					<span style="display:none;" id="CSATSaveBtn" class="btnBasic btnTypeComplete">입학원서 등록</span>
				</div>
			</div>				
			<div style="position: absolute; z-index:100; display: none; border:0px solid black; top: 240px; left: 0px; width: 99%; text-align: center;" id="Prog" name="Prog" >
              <img src="/Img/AjaxLoding.gif" width="32" height="32" border="0" alt="">
            </div>

			<div class="ibox-content"> 
				<div class="pad_5" style="overflow:auto; height:387px;">
					<form id="ListForm" method="post">
						<table id="dt_basic_none" class="table table-striped table-bordered table-hover margin_0" width="100%">
							<thead>			                
								<tr>
									<th data-hide="phone">년도</th>           
									<th>모집구분</th>                         
									<th>수험번호</th>    
									
									<th>언어영역응시구분</th>                            
									<th>언어영역표준점수</th>                            
									<th>언어영역백분위</th>                            
									<th>언어영역등급</th>

									<th>수리영역응시구분</th>
									<th>수리영역응시유형</th>         
									<th>수리영역표준점수</th>
									<th>수리영역백분위</th>    
									<th>수리영역등급</th>   
									
									<th>외국어영역응시구분</th>    
									<th>외국어영역표준점수</th>  									
									<th>외국어영역백분위</th>                            
									<th>외국어영역등급</th>  
									
									<th>탐구영역응시구분</th>
									<th>탐구영역구분</th>
									<th>탐구영역선택과목수</th>         
									<th>탐구영역과목1</th>
									<th>탐구영역표준점수1</th>    
									<th>탐구영역백분위1</th>   									
									<th>탐구영역등급1</th>    
									<th>탐구영역과목2</th>                            
									<th>탐구영역표준점수2</th>                            
									<th>탐구영역백분위2</th>
									<th>탐구영역등급2</th>
									<th>탐구영역과목3</th>         
									<th>탐구영역표준점수3</th>
									<th>탐구영역백분위3</th>    
									<th>탐구영역등급3</th>  									
									<th>탐구영역과목4</th>                            
									<th>탐구영역표준점수4</th>                            
									<th>탐구영역백분위4</th>
									<th>탐구영역등급4</th>

									<th>제2외국어영역응시구분</th>         
									<th>제2외국어영역과목</th>
									<th>제2외국어표준점수</th>  
									<th>제2외국어백분위</th>                            
									<th>제2외국어등급</th>                            
									<th>비고</th>
								</tr>
							</thead>
							<tbody class='viewTable'>
								<tr>
									<td colspan="13" style="height:50px; vertical-align: middle;">파일을 업로드하세요. 파일 크기에 따라 몇 분정도 걸릴 수 있습니다.</td>
									<td colspan="28" style="height:50px; vertical-align: middle;"></td>
								</tr>								
							</tbody>
						</table>
					</form>

				</div>				
			</div>
			<!-- 뷰어 끝 -->

			<div class="pad_t10"></div>

		</div>		
	</div>
</div>
</FORM>
<!-- #InClude Virtual = "/Common/PopUp_Bottom.asp" -->