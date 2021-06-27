$(function () {
	// Data Table 설정
	//pageSetUp();

	/* // DOM Position key index //

	l - Length changing (dropdown)
	f - Filtering input (search)
	t - The Table! (datatable)
	i - Information (records)
	p - Pagination (paging)
	r - pRocessing 
	< and > - div elements
	<"#id" and > - div with an id
	<"class" and > - div with a class
	<"#id.class" and > - div with an id and class

	Also see: http://legacy.datatables.net/usage/features
	*/	

	/* dataTable BASIC ;*/
	var responsiveHelper_dt_basic = undefined;
	var responsiveHelper_dt_basic_Search = undefined;
	var responsiveHelper_dt_basic_sub = undefined;
	var responsiveHelper_dt_basic_sub2 = undefined;
	var responsiveHelper_dt_basi_popup = undefined;
	var responsiveHelper_dt_basic_horizontal = undefined;

	var breakpointDefinition = {
		tablet : 1024,
		phone : 480
	};

	$("#dt_basic").dataTable({
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'i><'col-sm-6 col-xs-12 hidden-xs'f>r>" +
			"t"+
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basic) {
				responsiveHelper_dt_basic = new ResponsiveDatatablesHelper($('#dt_basic'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basic.createExpandIcon(nRow);
		}
		, "drawCallback" : function(oSettings) {
			responsiveHelper_dt_basic.respond();
		}
		, "pageLength": 15
		, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		//, "bSortClasses": false		// 정렬
		// ,"sScrollY": "200px"			// 스크롤
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		//,"paging": false				// 페이징
		//, "info": false				// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	});

	var dt_basic_Search = $("#dt_basic_Search").dataTable({
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'i><'col-sm-6 col-xs-12 hidden-xs'f>r>" +
			"t"+
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basic_Search) {
				responsiveHelper_dt_basic_Search = new ResponsiveDatatablesHelper($('#dt_basic_Search'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basic_Search.createExpandIcon(nRow);
		}
		, "drawCallback" : function(oSettings) {
			responsiveHelper_dt_basic_Search.respond();
		}
		, "pageLength": 15
		, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		//, "bSortClasses": false		// 정렬
		// ,"sScrollY": "200px"			// 스크롤
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		//,"paging": false				// 페이징
		//, "info": false				// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	});

	$(".dt_basic_sub").dataTable({
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-5'i><'col-sm-7 col-xs-12 hidden-xs'>r>" +
			"t"
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			+"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basic_sub) {
				responsiveHelper_dt_basic_sub = new ResponsiveDatatablesHelper($('.dt_basic_sub'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basic_sub.createExpandIcon(nRow);
		}
		, "drawCallback" : function(oSettings) {
			responsiveHelper_dt_basic_sub.respond();
		}
		//, "pageLength": 15
		, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		//, "bSortClasses": false		// 정렬
		, "scrollY": "285px"			// 스크롤
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		,"paging": false				// 페이징
		, "info": false					// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	});

	$(".dt_basic_sub2").dataTable({
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-5'i><'col-sm-7 col-xs-12 hidden-xs'>r>" +
			"t"
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			+"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basic_sub2) {
				responsiveHelper_dt_basic_sub2 = new ResponsiveDatatablesHelper($('.dt_basic_sub2'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basic_sub2.createExpandIcon(nRow);
		}
		, "drawCallback" : function(oSettings) {
			responsiveHelper_dt_basic_sub2.respond();
		}
		, "pageLength": 15
		, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		//, "bSortClasses": false		// 정렬
		//, "scrollY": "60px"				// 스크롤
		, "bScrollCollapse": true
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		,"paging": false				// 페이징
		, "bPaginate": false
		, "info": false					// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	});

	// POPUP 창 Data Table 설정
	$("#dt_basic_popup").dataTable({
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'i><'col-sm-6 col-xs-12'f>r>" +
			"t"+
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basi_popup) {
				responsiveHelper_dt_basi_popup = new ResponsiveDatatablesHelper($('#dt_basic_popup'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basi_popup.createExpandIcon(nRow);
		}
		//, "pageLength": 15
		//, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		//, "bSortClasses": false		// 정렬
		// ,"sScrollY": "200px"			// 스크롤
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		//,"paging": false				// 페이징
		//, "info": false				// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	});

	$("#dt_basic_horizontal").DataTable( {
		// "sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'f><'col-sm-6 col-xs-12 hidden-xs'l>r>" +
		//"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6 hidden-xs'f><'col-sm-6 col-xs-12 hidden-xs'<'toolbar'>>r>" +
		"sDom": "<'dt-toolbar'<'col-xs-12 col-sm-6'i><'col-sm-6 col-xs-12 hidden-xs'f>r>" +
			"t"+
			//"<'dt-toolbar-footer'<'col-sm-6 col-xs-12 hidden-xs'i><'col-xs-12 col-sm-6'p>>",
			"<'dt-toolbar-footer'<'col-xs-12 col-sm-12'p>>"
		, "oLanguage": {
			"sSearch": '<span class="input-group-addon"><i class="glyphicon glyphicon-search"></i></span>'
		}
		, "preDrawCallback" : function() {
			// Initialize the responsive datatables helper once.
			if (!responsiveHelper_dt_basic_horizontal) {
				responsiveHelper_dt_basic_horizontal = new ResponsiveDatatablesHelper($('#dt_basic_horizontal'), breakpointDefinition);
			}
		}
		, "rowCallback" : function(nRow) {
			responsiveHelper_dt_basic_horizontal.createExpandIcon(nRow);
		}
		, "drawCallback" : function(oSettings) {
			responsiveHelper_dt_basic_horizontal.respond();
		}
		, "pageLength": 10
		, "autoWidth" : true
		, "bSort": false				// 정렬
		, "ordering": false				// 정렬
		, "scrollX": true				// 가로 스크롤
		//, "bSortClasses": false		// 정렬
		// ,"sScrollY": "200px"			// 스크롤
		//, "scrollCollapse": true		// 스크롤뷰 자동
		//, "bPaginate": false			// 페이징
		//,"paging": false				// 페이징
		//, "info": false				// 상단 인포
		, "stateSave": true				// 현재 상태 유지
	} );

	/* dataTable BASIC END */

	// Form Submit
	$(".btnSubmit").click(function() {
		var blnChkResult = true;
		var $ParentForm = $(this).parents("form");
		var $objInput = $ParentForm.find(".form-control");

		$objInput.each(function()  {
			var alertMSG = $(this).attr("alert");
			
			if (alertMSG != undefined && alertMSG)  {
				//alert(alertMSG);
				blnChkResult = blnChkResult && $.chkInputValue($(this), alertMSG)
				return blnChkResult;
			}
		});

		if (blnChkResult) {
			$ParentForm.attr("method", "get");
			$ParentForm.attr("target", "_self");
			$ParentForm.attr("action", $.url().attr('file'));
			$ParentForm.find("[name='Page']").val("1");
			$ParentForm.submit();
		}
	});

	// 엑셀 저장
	$(".btnExcel").click(function() {
		var blnChkResult = true;
		var $ParentForm = $(this).parents("form");
		var $objInput = $ParentForm.find(".form-control");

		$objInput.each(function()  {
			var alertMSG = $(this).attr("alert");
			
			if (alertMSG != undefined && alertMSG)  {
				//alert(alertMSG);
				blnChkResult = blnChkResult && $.chkInputValue($(this), alertMSG)
				return blnChkResult;
			}
		});

		if (blnChkResult) {
			$ParentForm.attr("method", "post");
			$ParentForm.attr("target", "ExcelFrame");
			$ParentForm.attr("action", "/System/ExcelDownload.asp?process="+ $(this).attr("id"));
			$ParentForm.submit();
		}
	});

	// 내용 상세보기 URL 이동 & 색상변경
	//$(document).delegate("#dt_basic tbody tr", "click", function() {
	//$(document).delegate("tr.viewDetail", "click", function() {
	$(document).on("click", "tr.viewDetail", function(){
		// 색상 입력
		$(this).siblings().css({ backgroundColor:""});	// 형제 노드 초기화
		$(this).css({ backgroundColor:"#F5F5DC"});		// 선택 노트 색생 변경
	});

	// 내용 상세보기 값 입력 2 (DataField li 사용)
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
		$.setDataField_2($(this), "InputForm");
	});


/*
	// 내용 상세보기 값 입력 1 (DataField input 사용)
	$(document).on("click", "tr.viewDetail_SetDate_1", function(){
	//$(document).delegate("tr.viewDetail_SetDate_1", "click", function() {
		$.setDataField_1($(this), "InputForm");
	});

	// 내용 상세보기 값 입력 2 (DataField li 사용)
	$(document).on("click", "tr.viewDetail_SetDate_2", function(){
	//$(document).delegate("tr.viewDetail_SetDate_2", "click", function() {
		$.setDataField_2($(this), "InputForm");
	});
*/
	// 내용 상세보기 값 입력 (DataField input 사용)
	$.setDataField_1 = function ($Obj, FormID) {
		var $DataField = $Obj.find("div.DataField>input[type='hidden']");
		var ColumnName, ColumnType, ColumnValue, TagName, TagType;

		// 내용 입력
		$DataField.each(function()  {
			ColumnName = $(this).attr("ColumnName");
			ColumnType = $(this).attr("ColumnType");
			ColumnValue = $(this).val();
			$objColumn = $("form[id='"+ FormID +"'] [name='"+ ColumnName +"']");
			$objColumnID = $("form[id='"+ FormID +"'] [id='"+ ColumnName +"']");
			TagName = String($objColumn.prop("tagName")).toLowerCase();
			TagType = String($objColumn.attr("type")).toLowerCase();
			
			if (TagName != "undefined") {
				if (TagName == "input") {
					if (TagType == "text" || TagType == "password" || TagType == "hidden") {
						// 인풋 박스
						if (ColumnType == "Date") {
							if (ColumnValue != "") {
								$objColumnID.text($.left(ColumnValue, 4) +"-"+ $.right($.left(ColumnValue, 6), 2) +"-"+ $.right(ColumnValue, 2));
							} else {
								$objColumnID.text("");
							}
							$objColumn.val(ColumnValue);
						} else if (ColumnType == "Money") {
							$objColumn.val($.commaSplit(ColumnValue));
						} else {
							$objColumn.val(ColumnValue);
						}
					} else if (TagType == "checkbox" || TagType == "radio") {
						// 체크 박스
						$objColumn.each(function() {
							$(this).prop("checked", ($(this).val() == ColumnValue));
						});
					}
				} else if (TagName == "select") {
					// 셀렉트 박스
					//$objColumn.val(ColumnValue);
					//$objColumn.val(ColumnValue).trigger("change");
					$objColumn.val(ColumnValue).trigger("chosen:updated");
				} else if (TagName == "textarea") {
					// 텍스트 에어리어
					$objColumn.text(ColumnValue);
				}
			}
		});
		
		// 색상 입력
		$Obj.siblings().css({ backgroundColor:""});	// 형제 노드 초기화
		$Obj.css({ backgroundColor:"#F5F5DC"});	// 선택 노트 색생 변경
	}

	// 내용 상세보기 값 입력 (DataField li 사용)
	$.setDataField_2 = function ($Obj, FormID) {
		var $DataField = $Obj.find("div.DataField>li");
		var ColumnName, ColumnType, ColumnValue, TagName, TagType;

		// 내용 입력
		$DataField.each(function()  {
			ColumnName = $(this).attr("ColumnName");
			ColumnType = $(this).attr("ColumnType");
			ColumnValue = $(this).attr("Columnvalue");

			$objColumn = $("form[id='"+ FormID +"'] [name='"+ ColumnName +"']");
			$objColumnID = $("form[id='"+ FormID +"'] [id='"+ ColumnName +"']");
			
			TagName = String($objColumn.prop("tagName")).toLowerCase();
			TagType = String($objColumn.attr("type")).toLowerCase();
			
			if (TagName != "undefined") {
				if (TagName == "input") {
					if (TagType == "text" || TagType == "password" || TagType == "hidden") {
						// 인풋 박스
						if (ColumnType == "Date") {
							if (ColumnValue != "") {
								$objColumnID.text($.left(ColumnValue, 4) +"-"+ $.right($.left(ColumnValue, 6), 2) +"-"+ $.right(ColumnValue, 2));
							} else {
								$objColumnID.text("");
							}
							$objColumn.val(ColumnValue);
						} else if (ColumnType == "Money") {
							$objColumn.val($.commaSplit(ColumnValue));
						} else {
							$objColumn.val(ColumnValue);
						}
					} else if (TagType == "checkbox" || TagType == "radio") {
						// 체크 박스
						$objColumn.each(function() {
							$(this).prop("checked", ($(this).val() == ColumnValue));
							/*
							if ($(this).val() == ColumnValue) {
								$(this).prop("checked", true);
							} else {
								$(this).prop("checked", false);
							}
							*/
						});
					}
				} else if (TagName == "select") {
					// 셀렉트 박스
					//$objColumn.val(ColumnValue);
					//$objColumn.val(ColumnValue).trigger("change");
					$objColumn.val(ColumnValue).trigger("chosen:updated");
				} else if (TagName == "textarea") {
					// 텍스트 에어리어
					$objColumn.text(ColumnValue);
				}
			}
		});

		// 색상 입력
		$Obj.siblings().css({ backgroundColor:""});	// 형제 노드 초기화
		$Obj.css({ backgroundColor:"#F5F5DC"});	// 선택 노트 색생 변경
	}

	// 체크박스 전체 선택/ 해제
	$(".checkedAll").click(function() {
		var SubCheckboxName = $(this).attr("SubCheckboxName");
		if ($(this).prop("checked")) {
			$("input[name='"+ SubCheckboxName +"']:checkbox").each(function() {
				$(this).prop("checked", true);
			});
		} else {
			$("input[name='"+ SubCheckboxName +"']:checkbox").each(function() {
				$(this).prop("checked", false);
			});
		}
	});
	
	// 폼 입력값 중 필수값 체크
	$.setValidation = function ($ParentForm) {
		var blnChkResult = true;
		var $objInput = $ParentForm.find(".form-control");

		$objInput.each(function()  {
			var alertMSG = $(this).attr("alert");
			if (alertMSG != undefined && alertMSG) {
				//alert(alertMSG);
				blnChkResult = blnChkResult && $.chkInputValue($(this), alertMSG)
				return blnChkResult;
			}
		});

		return blnChkResult;
	}

	// Form Reset
	// 달력 초기화 시 기본값 없으면 오늘날짜로 처리
	$.FormReset = function ($objForm) {
		// Form 리셋
		$objForm[0].reset();

		// SelectBox 리셋
		//$objForm.find("select").each(function() { $(this).trigger("change"); });
		$objForm.find("select").each(function() { $(this).trigger("chosen:updated"); });

		// 달력 리셋
		$objForm.find(".viewCalendarBtn").each(function() {
			var CalendarName = $(this).attr("Obj");
			var CalendarValue = $("[name='"+ CalendarName +"']").val();
			
			if (CalendarValue != "") {
				$("#"+ CalendarName).text($.left(CalendarValue, 4) +"-"+ $.right($.left(CalendarValue, 6), 2) +"-"+ $.right(CalendarValue, 2));
			} else {
				var d = new Date();
				$("#"+ CalendarName).text(d.getFullYear() +"-"+ $.right("000"+ (d.getMonth() + 1), 2) +"-"+ $.right("000"+ d.getDate(), 2));
				$("[name='"+ CalendarName +"']").val(d.getFullYear() +""+ $.right("000"+ (d.getMonth() + 1), 2) +""+ $.right("000"+ d.getDate(), 2));
			}
		});
	}

	// Form Reset2
	// 달력 초기화 시 기본값 없으면 공백으로 처리
	$.FormReset2 = function ($objForm) {
		// Form 리셋
		$objForm[0].reset();

		// SelectBox 리셋
		//$objForm.find("select").each(function() { $(this).trigger("change"); });
		$objForm.find("select").each(function() { $(this).trigger("chosen:updated"); });

		// 달력 리셋
		$objForm.find(".viewCalendarBtn").each(function() {
			var CalendarName = $(this).attr("Obj");
			var CalendarValue = $("[name='"+ CalendarName +"']").val();
			
			if (CalendarValue != "") {
				$("#"+ CalendarName).text($.left(CalendarValue, 4) +"-"+ $.right($.left(CalendarValue, 6), 2) +"-"+ $.right(CalendarValue, 2));
			} else {
				var d = new Date();
				$("#"+ CalendarName).text("");
			}
		});
	}

	// Ajax4Form 전송
	//$.Ajax4FormSubmit($("#ListForm"), "처리 되었습니다.");
	$.Ajax4FormSubmit = function ($objForm, alertMSG, goURL) {
		var objOpt = {"url":"","param":"","dataType":"xml","before":"","success":"$.Ajax4FormResult(datas, '"+ alertMSG +"', '"+ goURL +"')","complete":"","clear":"","reset":""};
		objOpt["url"] = $objForm.attr("action");
		$.Ajax4Form($objForm, objOpt);
		$objForm.submit();
	}

	// Ajax4Form 결과 처리
	$.Ajax4FormResult = function (datas, alertMSG, goURL) {
		var $objList = $(datas).find("List");
		var strMSG;
		
		if ($objList.find("Result").text() == "Complete") {
			if (alertMSG == "undefined" || alertMSG == "") {
				//alert("알림창 없음");
			} else {
				alert(alertMSG);
			}
			
			//if (goURL == undefined || !goURL) { //goURL이 객체가 아닌 텍스트 형식으로 넘어옴
			if (goURL == undefined || !goURL || goURL == "undefined" || goURL == "") {
				document.location.reload();
			} else {
				document.location.href = goURL;
				
			}

			//setTimeout(function() { document.location.reload(); }, 500);
		} else {
			alert($objList.find("ReturnMSG").text());
			return;
		}
	}

	// Ajax4Form 전송 - 처리 후 실행 Function 지정
	// $.Ajax4FormSubmitSetFunction($("#InputForm_SUB_5"), "기타소득 삭제가 완료되었습니다.", "SubDetailProcComplete");
	$.Ajax4FormSubmitSetFunction = function ($objForm, alertMSG, RunFunction, Loding) {
		var objOpt = {"url":"","param":"","dataType":"xml","before":"","success":"$."+ RunFunction +"(datas, '"+ alertMSG +"')","complete":"","clear":"","reset":"","Loding":Loding};
		objOpt["url"] = $objForm.attr("action");
		$.Ajax4Form($objForm, objOpt);
		$objForm.submit();
	}

	// 찾기 버튼 엔터 적용
	$("div.input-group input").keyup(function (event) {
		// 엔터키일 때 실행
		if (event.keyCode == 13) {
			$(this).parent().next().children("button").click();
		}
	});
});