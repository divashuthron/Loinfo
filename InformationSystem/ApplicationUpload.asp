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
			var datasStr = "";
			var $objForm = $("#frmFile");
			$($objForm).ajaxForm({
				url: "/ApplicationFileupload.asp"
				, type: "post"
				, dataType: "text"
				, success: function (datas, state) {
					//파일내용 출력
					$(".viewTable").empty().html(datas);
					//입학원서 건수 출력
					datasStr = datas.split("<count>");
					$(".CountTable").html("<h5>목록 - 전체 " + datasStr[1] + "건</h5>");
					//입학원서 건수 넣어주기(저장 시 활용)
					$("#ApplicationCount").val(datasStr[1]);
					//로딩이미지 제거
					document.getElementById("Prog").style.display = "none";
					Uploading = false;	
					//입학원서 저장 버튼 출력(헷갈릴 수 있으므로, 파일 확인 시 출력)
					$("#ApplicationSaveBtn").css("display","block");								
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
		$(document).on("click", "#ApplicationSaveBtn", function(){
			if (Uploading)
			{
				alert("등록중입니다. 잠시 기다리세요.");
				return;
			}
			var $objForm = $("#frmFile");
			$objForm.attr("enctype","");
			$($objForm).ajaxForm({
				url: "/Process/ApplicationfileuploadSave.asp"
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
	<input type="hidden" id="ApplicationCount" name="ApplicationCount" value="">
<div class="row">
	<div class="col-lg-12">
		<div class="ibox float-e-margins">
			
			<!-- 파일확인 -->
			<div class="ibox-title">
				<h2>입학원서 엑셀 업로드</h2>
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
					<span style="display:none;" id="ApplicationSaveBtn" class="btnBasic btnTypeComplete">입학원서 등록</span>
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
									<th>모집시기</th>                         
									<th>지망학과</th>                           
									<th>지망전형1</th>                            
									<th>지망구분2</th>                            
									<th>지망구분3</th>                            
									<th>지망구분4</th>
									<th>수험번호</th>
									<th>성명(한글)</th>         
									<th>성명(영문)</th>
									<th>성명(한문)</th>    
									<th>주민번호앞</th>    
									<th>주민번호뒤</th>    
									<th>성별</th>  
									
									<th>고교졸업년도</th>                            
									<th>고교코드</th>                            
									<th>고교학과</th>
									<th>고교졸업여부</th>
									<th>검정고시여부</th>         
									<th>검정고시졸업년도</th>
									<th>검정고시합격지구</th>    
									<th>생기부반영학기</th>   
									
									<th>출신대학코드</th>    
									<th>대학평균점수</th>                            
									<th>대학만점점수</th>                            
									<th>이수학점</th>
									<th>고교(과정)구분</th>
									<th>환불방법</th>         
									<th>환불예금주</th>
									<th>환불은행</th>    
									<th>환불계좌</th>  
									
									<th>자택전화</th>                            
									<th>지원자전화</th>                            
									<th>보호자전화</th>
									<th>이메일</th>
									<th>우편번호</th>         
									<th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;기본주소
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
									<th>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;상세주소
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>  
									<th>지원자확인성명</th>                            
									<th>생기부동의</th>                            
									<th>검정고시동의</th>
									<th>수능동의</th>
									<th>수험생확인동의</th>   

									<th>개인정보수집동의</th>                            
									<th>고유식별수집동의</th>
									<th>개인정보위탁동의</th>
									<th>개인정보3자제공동의</th>   									

									<th>원서접수일자</th>
									<th>원서접수시간</th>  
								</tr>
							</thead>
							<tbody class='viewTable'>
								<tr>
									<td colspan="18" style="height:50px; vertical-align: middle;">파일을 업로드하세요. 파일 크기에 따라 몇 분정도 걸릴 수 있습니다.</td>
									<td colspan="27" style="height:50px; vertical-align: middle;"></td>
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