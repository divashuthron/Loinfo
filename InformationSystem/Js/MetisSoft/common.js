//============================= jquery 바인딩 =====================================
/*
함수 참조

- 부모 요소 선택
	parent() 직속부모 단일 요소만 선택
	parents([selector]) 모든 부모요소 선택
	closest(selector) selector에 해당하는 가장 가까운 요소 하나 선택

- 자식 요소 선택
	children() 직계자식 요소만 선택
	find() 직계자식을 시작으로 모든 자손 선택

- 기타 요소 선택
	siblings() 형제 요소 모두 선택
	prev() 이전의 요소 선택
	next() 다음 요소 선택
*/
$(function () {
    //=== 초기 세팅 ============================
    $(window).load(function () {
		/*
		// DOM의 standard 이벤트
		// html의 로딩이 끝난 후에 시작
		// 화면에 필요한 모든 요소(css, js, image, iframe etc..)들이 웹 브라우저 메모리에 모두 올려진 다음에 실행됨
		// 화면이 모두 그려진 다음의 메세징이나 이미지가 관련 요소가 모두 올려진 다음의 애니메이션에 적합함
		// 실행 순서를 보면 document.ready() > window.load() > body onload 이벤트 순서대로 실행됨
		*/
		//var $URL = $.url();
    });

	//== SelectBox 플러그인 적용 ===================
	$(document).ready(function() {
		/*
		// 외부 리소스. 이미지와는 상관 없이 브라우저가 DOM (document object model) 트리를 생성한 직후 실행
		// window.load() 보다 더 빠르게 실행되고 중복 사용하여 실행해도 선언한 순서대로 
		*/
		////$.fn.select2.defaults.set("theme", "classic");
		////$.fn.select2.defaults.set("dropdownAutoWidth", true);
		//$.fn.select2.defaults.set("width", "100%");
		//$(".select2").select2();
		//$("select").select2();
		$(".select2").chosen({
			no_results_text:"자료가 없습니다."
			, width:"100%"
			//, disable_search_threshold: 10
		});

		$(".i-checks").iCheck({
			checkboxClass: "icheckbox_square-green",
			radioClass: "iradio_square-green",
		});
	});

    //=== 손모양 세팅 ===========================
    $("img[class]").css("cursor", "pointer");
    $("img[id]").css("cursor", "pointer");
    $("td[id]").css("cursor", "pointer");
    $("i[class^='fa fa-']").css("cursor", "pointer");

    //=== 페이지 이동 ===========================
    // 뒤로가기
    $("#GoBack").click(function () {
        var strURL = $(this).attr("url");

        if (strURL == undefined || !strURL) {
            history.back();
        } else {
            document.location.href = strURL;
        }
    });
    // 위로가기
    $("#GoTop").click(function () {
        $("body").animate(
			{ scrollTop: "0px" }, "fast"
		);
    });

    // bootstrap 버튼
	// 버튼 생성 ex) <span class="btnBasic btnTypeSearch">검색</span>
	$(".btnBasic").each(function () {
        var strHtml = "";
        var strTitle = $(this).html();
		var strURL = $(this).attr("url");
		var ButtonColor = "btn-info";
		var ButtonIcon = "glyphicon-search";

		/*
		btnTypeSearch			// 조회
		btnTypePrint			// 인쇄
		btnTypeAccept			// 전자결재
		btnTypeNew				// 신규
		btnTypeSave				// 저장
		btnTypeDelete			// 삭제
		btnTypeAdd				// 추가
		btnTypeEdit				// 수정
		btnTypeComplete			// 최종완료
		btnTypeClose			// 닫기
		btnTypeExcel			// 엑셀
		btnTypeCancel			// 취소
		btnTypePeport			// 리포트보기
		*/

		if ($(this).hasClass("btnTypeSearch") == true) {
			// 조회
			ButtonColor = "btn-info";
			ButtonIcon = "glyphicon-search";
		} else if ($(this).hasClass("btnTypePrint") == true) {
			// 인쇄
			//ButtonColor = "btn-primary";
			ButtonColor = "bg-color-orange txt-color-white";
			ButtonIcon = "glyphicon-print";
		} else if ($(this).hasClass("btnTypeAccept") == true) {
			// 전자결재
			ButtonColor = "btn-success";
			ButtonIcon = "glyphicon-share";
		} else if ($(this).hasClass("btnTypeNew") == true) {
			// 신규
			ButtonColor = "bg-color-purple txt-color-white";
			ButtonIcon = "glyphicon-repeat";
		} else if ($(this).hasClass("btnTypeSave") == true) {
			// 저장
			ButtonColor = "btn-primary";
			ButtonIcon = "glyphicon-check";
		} else if ($(this).hasClass("btnTypeDelete") == true) {
			// 삭제
			ButtonColor = "btn-danger";
			ButtonIcon = "glyphicon-trash";
		} else if ($(this).hasClass("btnTypeAdd") == true) {
			// 추가
			ButtonColor = "bg-color-purple txt-color-white";
			ButtonIcon = "glyphicon-plus-sign";
		} else if ($(this).hasClass("btnTypeEdit") == true) {
			// 수정
			ButtonColor = "btn-warning";
			ButtonIcon = "glyphicon-edit";
		} else if ($(this).hasClass("btnTypeComplete") == true) {
			// 최종완료
			ButtonColor = "btn-success";
			ButtonIcon = "glyphicon-floppy-disk";
		} else if ($(this).hasClass("btnTypeClose") == true) {
			// 닫기
			ButtonColor = "bg-color-yellow txt-color-white";
			ButtonIcon = "glyphicon-home";
		} else if ($(this).hasClass("btnTypeExcel") == true) {
			// 엑셀
			ButtonColor = "bg-color-yellow txt-color-white";
			ButtonIcon = "glyphicon-list-alt";
		} else if ($(this).hasClass("btnTypeCancel") == true) {
			// 취소
			ButtonColor = "bg-color-redLight txt-color-white";
			ButtonIcon = "glyphicon-random";
		} else if ($(this).hasClass("btnTypePeport") == true) {
			// 리포트보기
			ButtonColor = "btn-info";
			ButtonIcon = "glyphicon-print";
		} else if ($(this).hasClass("btnTypeConfirm") == true) {
			// 확정
			ButtonColor = "btn-danger";
			ButtonIcon = "glyphicon-print";
		}
		
		if ($(this).hasClass("NoneIcon") == true) {
			strHtml = "<a class=\"btn btn-labeled "+ ButtonColor +"\" style='padding-top:3px;'> " + strTitle + "</a>";
			//strHtml = "<a class=\"btn btn-labeled "+ ButtonColor +"\" style='padding-top:4px;'> <i class=\"glyphicon "+ ButtonIcon +"\"></i></a>";
		} else {
			strHtml = "<a class=\"btn btn-labeled "+ ButtonColor +"\"> <span class=\"btn-label\"><i class=\"glyphicon "+ ButtonIcon +"\"></i></span>" + strTitle + "</a>";
		}
		
        $(this).html(strHtml);
		if (strURL != undefined && strURL) { $(this).click(function() { $.goURL(strURL); }); }
    }).css("cursor", "pointer");

    // 이미지 버튼
	// 버튼 생성 ex) <span class="btnBasic1">검색</span>
	$(".btnBasic1").each(function () {
        var strHtml = "";
        var strTitle = $(this).html();
		var strURL = $(this).attr("url");

		strHtml += "<span><img src=\"./img/btn_basic01.gif\"></span>";
        strHtml += "<span style=\"color:#666666;line-height:25px; background:url('./img/btn_basic03.gif') repeat-x;\">" + strTitle + "</span>";
        strHtml += "<span><img src=\"./img/btn_basic02.gif\"></span>";

        $(this).html(strHtml);
        $(this).children("span").css("display", "inline-block").css("vertical-align", "middle").height(25);
		//$(this).css("display", "inline-block").css("position", "relative");
		if (strURL != undefined && strURL) { $(this).click(function() { $.goURL(strURL); }); }
    }).css("cursor", "pointer");
	
	// 아이콘 없는 이미지 버튼
	$(".btnBasic2").each(function () {
        var strHtml = "";
        var strTitle = $(this).html();
		var strURL = $(this).attr("url");

        strHtml += "<span style=\"width:7;\"><img src=\"./img/btn_basic04.gif\" style=\"display:inline;\" align=\"absmiddle\"></span>";
        strHtml += "<span style=\"color:#666666;line-height:25px; background:url('./img/btn_basic03.gif') repeat-x;\">" + strTitle + "</span>";
        strHtml += "<span style=\"width:7;\"><img src=\"./img/btn_basic05.gif\" style=\"display:inline;\" align=\"absmiddle\"></span>";

        $(this).html(strHtml);
        $(this).children("span").css("display", "inline-block").css("vertical-align", "middle").height(25);
		$(this).css("display", "inline").css("position", "relative");
		if (strURL != undefined && strURL) { $(this).click(function() { $.goURL(strURL); }); }
    }).css("cursor", "pointer");
	
    //=== 링크 이동  ============================
    $(".ALink").click(function () {
        var strURL = $(this).attr("url");
        if (strURL == undefined || !strURL) {
            //alert("준비중입니다.");
			alert(decodeURI("%EC%A4%80%EB%B9%84%EC%A4%91%EC%9E%85%EB%8B%88%EB%8B%A4."));
            return;
        }
        if (strURL == undefined || !strURL) { strURL = "/"; }
        $.goURL(strURL);
    })
	.css("cursor", "pointer");

    //=== 메인 아이콘 로그인 이동 =================
    $(".ALinkLogin").click(function () {
        if (!confirm("본 서비스는 로그인을 해야 이용하실 수 있습니다.\n로그인 메뉴로 이동하시겠습니까?")) { return; }

        var strURL = $(this).attr("url");
        if (strURL == undefined || !strURL) {
            $.goURL("/Login/login.asp");
        } else {
            $.goURL("/Login/login.asp?prevURL=" + strURL);
        }
    })
	.css("cursor", "pointer");
	
	//=== 창 닫기 버튼===========================
	$(".SelfClose").click(function() {
		top.window.close();
	})
	.css("cursor", "pointer");	
	
	$(".SelfCloseDIV").click(function() {
		$.closeMadal("2");
	})
	.css("cursor", "pointer");

	$(".ExcelDown").click(function() {
		var ActionType = $(this).attr("ActionType");
		var Query = $(this).attr("Query");
		
        if (Query == undefined || !Query) {
            Query = "";
        } else {
			Query = "&" + Query;
		}
		
		if (ActionType == "CounselList") {
			alert("문의 목록은 엑셀 다운로드 기능을 제공하지 않습니다.");
			return;
		} else {
			$.goURL("/ExcelDown.asp?ActionType="+ ActionType + Query);
		}
	});
	
    //=== 숫자만 입력  ==========================
    $(document).on("keyup", ".KeyTypeNUM", function(event){
	//$(".KeyTypeNUM").keyup(function (event) {
		switch (event.keyCode) {
            // 컨트롤 영역키                                                                                  
            case 8: break; case 9: break; case 13: break; case 16: break; case 17: break;
            case 18: break; case 20: break; case 21: break; case 25: break;
            case 27: break; case 33: break; case 34: break; case 35: break;
            case 36: break; case 37: break; case 38: break; case 39: break;
            case 40: break; case 45: break; case 46: break; case 144: break;
            case 229: break;

            // 상단 숫자키                                                                                  
            case 48: break; case 49: break; case 50: break; case 51: break; case 52: break;
            case 53: break; case 54: break; case 55: break; case 56: break; case 57: break;

            // 키패드 숫자키                                                                                  
            case 96: break; case 97: break; case 98: break; case 99: break;
            case 100: break; case 101: break; case 102: break; case 103: break;
            case 104: break; case 105: break;

			// -  and . : 입력 (-값, 소수점 받기 위해 혀용)
			case 188: break; case 189: break; case 190: break; case 186: break;

            default:
                //alert("숫자만 입력 가능 합니다.");
				//alert(decodeURI("%EC%88%AB%EC%9E%90%EB%A7%8C %EC%9E%85%EB%A0%A5 %EA%B0%80%EB%8A%A5 %ED%95%A9%EB%8B%88%EB%8B%A4."));
                
				$(this).val($.getNumberOnly($(this).val()));
				//$(this).val("");
                break;
        }

		if($(this).val().indexOf(".") == 0) {
			$(this).val("");
		}
    });
	
	//=== 다음 입력창으로 이동 =====================
	$(".KeyNextTab").keyup(function () {
		var intLen = $(this).attr("maxlength");
		if ($(this).val().length == parseInt(intLen)) {
			$(this).next().focus();
		}
	});
	
	//=== inputBox keyup 콤마 처리 ====================
	//$('input.input-money').on('change', function(e) {
	$("input.input-money").keyup(function () {
		$(this).val($.commaSplit($(this).val()));
	});

	//=== inputBox 자동 입력 =========================
	$(':input[title]').each(function () {
		var $this = $(this);
		
		if($this.val() === '') {
			$this.val($this.attr('title'));
		}
		
		$this.focus(function() {
			if($this.val() === $this.attr('title')) {
				$this.val('');
			}
		});
		
		$this.blur(function() {
			if($this.val() === '') {
				$this.val($this.attr('title'));
			}
		});
	});
	

	//=== inputBox keyup =========================
	//$('input[type="text"]').on('change',function(e){
	//	alert(this.value);
	//});
	$(document).on('keyup', 'input[type="text"]', function(e){
		var $this = $(this);

		if($this.val() !== $this.data('previousValue')){
			$this.trigger('change'); 
		}
		$this.data('previousValue', $this.val());
	});


	//=== textarea Maxlength 제한하기 ======================
	$(document).on('keyup change', 'textarea[maxlength]', function(){
	//$(document).delegate("textarea[maxlength]", "keyup change", function() {
	//$('textarea[maxlength]').live('keyup change', function() {
		var $Textarea = $(this);
		var TextareaVal = $Textarea.val();
		var Maxlength = ($Textarea.attr('maxlength') == undefined) ? 2000 : parseInt($Textarea.attr('maxlength'));
		
		if (TextareaVal.length > Maxlength) {
			alert("최대 "+ Maxlength +"자 까지 입력 가능 합니다.\n입력하신 글자수는 "+ TextareaVal.length +"자 입니다.");
			$Textarea.focus();
			$Textarea.val(TextareaVal.substr(0, Maxlength));
			return false;
		}
	});
	
	//=== 탭 메뉴 설정 =========================
	$(".ChangeTab").click(function() {
		//
		$(".ChangeTab").each(function() {
			$(this).attr("src", function() { return this.src.replace("_on.gif", ".gif"); });
			$(this).attr("src2", "");
		});
		$(this).attr("src", function() { return this.src.replace(".gif", "_on.gif"); });
		$(this).attr("src2", $(this).attr("src"));
	})
	.mouseover(function() {
		var reg	= /_on/gi;
		var src	= $(this).attr("src");
		
		if(!reg.test(src)) {
			$(this).attr("src", function() { return this.src.replace(".gif", "_on.gif"); });
		} else {
			$(this).attr("src2", $(this).attr("src"));
		}
	})
	.mouseout(function() {
		var reg	= /_on/gi;
		var src2	= $(this).attr("src2");
		
		if(!reg.test(src2)) {
			$(this).attr("src", function() { return this.src.replace("_on.gif", ".gif"); });
		}
	});
	
	//=== 테이블 배경색 변경 =========================
	$(".BoardListTR").hover(
		function () {
			$(this).css({ backgroundColor:"#FAFAFA"});
		},
		function () {
			$(this).css({ backgroundColor:"#FFFFFF"});
		}
	);
	
    //=== 함수 시작 [ 사용법 : $.goURL("URL"); ] ==================
    $.extend({
        //=== 스크롤 TOP ==================================
        setTopScroll: function () {
            setTimeout(scrollTo, 0, 0, 2);
            $("#GoTop").click();
        },

        //=== window open popup ============================
        windowOpen: function (strURL, intWidth, intHeight, strScroll, strWinID) {
            var strTempWin = "";
            if (strWinID == undefined || !strWinID) {
                strTempWin = "Window_" + intWidth + "x" + intHeight;
            } else {
                strTempWin = strWinID;
            }

            //$.windowOpen("/sample3/list.asp", "500", "700", "yes");
            var win = window.open(strURL, strTempWin, "width=" + intWidth + ", height=" + intHeight + ", scrollbars=" + strScroll);
            win.focus();
        },

        //=== window popup ReSize ============================
        windowReSize: function (w, h, scroll, center) {
            if (scroll) {
                w = w + 17;
            }

            if (center) {
                var winl = (screen.width / 2) - (w / 2);
                var wint = (screen.height / 2) - (h / 2);
                winl = winl - 10;
                wint = (isopera) ? wint - 130 : wint - 30;
                window.moveTo(winl, wint);
            }

            if (!scroll) {
                document.documentElement.style.overflow = 'hidden';
            } else if (!isie) {
                document.documentElement.style.overflow = 'auto';
                document.documentElement.style.overflowX = 'hidden';
            }
            var nw = (!scroll) ? document.documentElement.clientWidth : document.documentElement.clientWidth + 17;
            var nh = document.documentElement.clientHeight;
            if ((nw != w && (nw - 1) != w && (nw + 1) != w) || nh != h) window.resizeBy(w - nw, h - nh);
        },

        //=== 숫자 콤마 찍기 ==================================
        commaSplit: function (intValue) {
            var rxSplit = new RegExp('([0-9])([0-9][0-9][0-9][,.])');
            var arrNumber = intValue.replace(/\,/g,'').split('.');

            arrNumber[0] += '.';
            do {
                arrNumber[0] = arrNumber[0].replace(rxSplit, '$1,$2');
            } while (rxSplit.test(arrNumber[0]));

            if (arrNumber.length > 1) {
                return arrNumber.join('');
            }
            else {
                return arrNumber[0].split('.')[0];
            }
        },

        //=== 숫자 콤마 찍기 2 =================================
        addComma: function (str) {
            var reg = /(^[+-]?\d+)(\d{3})/;
            str += '';

            while (reg.test(str)) {
                str = str.replace(reg, '$1,$2');
            }
            return str;
        },

        //=== 숫자만 반환 =================================
        getNumberOnly: function (str) {
			var val = str;
			val = new String(val);
			var regex = /[^0-9-.]/g;	// 숫자 - . 허용
			//var regex = /[^0-9.]/g;	// 숫자 . 허용
			val = val.replace(regex, '');

            return val;
        },

        //=== 테이블 rowspan ==================================
        tableRowspan: function (objTable) {
            $('tr:eq(0) > td', objTable).each(function (colIdx) {
				$.eachRowspan(objTable, colIdx);
            });
			
        },

        //=== 테이블 rowspan (row 선택) ==================================
        eachRowspan: function (objTable, colIdx) {
            var that;
            $('tr', objTable).each(function (row) {
                $('td:eq(' + colIdx + ')', $(this)).each(function (col) {
                    if ($(this).html() == $(that).html()) {
                        rowspan = $(that).attr("rowSpan");
                        if (rowspan == undefined) {
                            $(that).attr("rowSpan", 1);
                            rowspan = $(that).attr("rowSpan");
                        }
                        rowspan = Number(rowspan) + 1;
                        $(that).attr("rowSpan", rowspan);
                        $(this).hide();
                    } else {
                        that = $(this);
                    }
                    that = (that == null) ? $(this) : that;
                });
            });
        },
		
		//=== 10 이하의 숫자를 2자리로 변경 ======================
		getCipher: function(num) {
			return parseInt(num) > 9 ? num : "0" + num
		},

        //=== 입력된 값이 정해진 문자열로만 이루어졌는지 확인 (확인할 문자열, 제한할 문자열들)=============
        containsCharsOnly: function (str, chars) {
            for (var i = 0; i < str.length; i++) {
                if (chars.indexOf(str.charAt(i)) == -1)
                    return false;
            }
            return true;
        },
		
		//=== Left 함수 ====================================
		left: function (str, n) {
			if (n <= 0) {
				return "";
			} else if (n > String(str).length) {
				return str;
			} else {
				return String(str).substring(0,n);
			}
		},
		
		//=== Right 함수 ====================================
		right: function (str, n) {
			if (n <= 0) {
				return "";
			} else if (n > String(str).length) {
				return str;
			} else {
				var iLen = String(str).length;
				return String(str).substring(iLen, iLen - n);
			}
		},
		
		//=== Random 함수 ====================================
		randomRange: function(Min, Max) {
			return Math.floor((Math.random() * (Max - Min + 1)) + Min);
		},

        //=== 접속 URL(도메인) 체크 ==========================
        chkDNS: function () {
            var strDns = location.href
            strDns = strDns.split("//");
            strDns = strDns[1].substr(0, strDns[1].indexOf("/"));

            return strDns;
        },
        
		//=== Input 유효성 체크 =============================
		chkInputValue: function (obj, returnMSG) {
			var TagName = String(obj.prop("tagName")).toLowerCase();
			var chkCount = 0;
			
			if (TagName == "input") {
				if (obj.attr("type") == "text") {
					// 텍스트 박스
					if (!obj.val()) {
						alert(returnMSG);
						obj.focus();
						return false;
					}
				} else if (obj.attr("type") == "number") {
					// 숫자 박스
					if (!obj.val()) {
						alert(returnMSG);
						obj.focus();
						return false;
					}
				} else if (obj.attr("type") == "password") {
					// 패스워드 박스
					if (!obj.val()) {
						alert(returnMSG);
						obj.focus();
						return false;
					}
				} else if (obj.attr("type") == "hidden") {
					// 히든 박스
					if (!obj.val()) {
						alert(returnMSG);
						obj.focus();
						return false;
					}
				} else if (obj.attr("type") == "checkbox") {
					// 체크 박스
					if (!obj.is(":checked")) {
						alert(returnMSG);
						obj.eq(0).focus();
						return false;
					}
				} else if (obj.attr("type") == "radio") {
					// 라디오 박스
					obj.each(function() { if ($(this).is(":checked")) { chkCount++; } });
					
					if (chkCount == 0) {
						alert(returnMSG);
						obj.eq(0).focus();
						return false;
					}
				} else if (obj.attr("type") == "file") {
					// 파일 박스
					if (!obj.val()) {
						alert(returnMSG);
						obj.focus();
						return false;
					}
				}
			} else if (TagName == "select") {
				// 셀렉트 박스
				if (!obj.val()) {
					alert(returnMSG);
					obj.focus();
					/*
					var relayEvents = [
						'open', 'opening',
						'close', 'closing',
						'select', 'selecting',
						'unselect', 'unselecting',
						'focus'
					];
					*/
					//obj.select2('focus');
					//obj.select2("open");
					obj.trigger("chosen:activate");
					obj.trigger("chosen:open");
					return false;
				}
			} else if (TagName == "textarea") {
				// 텍스트 에어리어
				//if (!obj.text()) {
				if (!obj.val()) {
					alert(returnMSG);
					obj.focus();
					return false;
				}
			}
			
			return true;
		},
		
		//=== URL 이동 ====================================
        goURL: function (strURL) {
            document.location.href = strURL;
        },

        //=== 경고창 후 이동 ===============================
        alertGO: function (strMSG, strURL) {
            alert(strMSG);
            $.goURL(strURL);
        },

        //=== 모달창 열기 ===============================
        openMadal: function ($objModal, modalType, cursorType) {
            if (modalType == undefined) { modalType = "1"; }
            if (cursorType == undefined) { cursorType = ""; }
            /*
			if (modalType == "1") {
                //jquery.simplemodal-1.4.2.js
                $objModal.modal({
                    minWidth: $objModal.width(),
                    minHeight: $objModal.height(),
                    onOpen: function (dialog) {
                        $("#basic-modal-content").css("display", "none");
                        $("#simplemodal-overlay").css("background-color", "#000");
                        $("#simplemodal-container").css("background-color", "#FFF").css("border", "2px solid #444");
                        $("#simplemodal-container").css("padding", "10px");
                        dialog.overlay.fadeIn('fast', function () {
                            dialog.container.fadeIn('fast', function () {
                                dialog.data.show();
                                //dialog.data.fadeIn('fast');
                            });
                        });
                    },
                    onClose: function (dialog) {
                        //dialog.data.fadeOut('fast', function () {
                        dialog.container.fadeOut('fast', function () {
                            dialog.overlay.fadeOut('fast', function () {
                                $.modal.close();
                            });
                        });
                        //});
                    }
                });
            } else {
			*/
                //jquery.blockUI.js
                $.blockUI({
                    message: $objModal,
                    css: {
                        top: ($(window).height() - ($objModal.height() + 100)) / 2 + 'px',
                        left: ($(window).width() - $objModal.width()) / 2 + 'px',
                        width: $objModal.width() + 'px',
                        border: '2px solid #444444',
                        padding: '15px',
						cursor: cursorType
                    },
					baseZ: 2000
                });
            /*
			}
			*/
        },

        //=== 모달창 닫기 ===============================
        closeMadal: function (modalType) {
            if (modalType == undefined) { modalType = "1"; }
            /*
			if (modalType == "1") {
                $.modal.close();
            } else {
			*/
                $.unblockUI();
			/*
            }
			*/
        },

        showLoding: function () {
            $.findLoding();
            $.openMadal($('div.LodingPanel'));
        },

        hideLoding: function () {
            $.closeMadal();
        },

        showLoding2: function (cursorType) {
            $.findLoding();
            $.openMadal($('div.LodingPanel'), "2", cursorType);
        },

        hideLoding2: function () {
			$.closeMadal("2");
        },

        //=== 잠시만 기다려 주세요 창 세팅 ===================
        findLoding: function () {
            var $html = $("body");
            var strHtml = "";

            if ($html.find("div.LodingPanel").html() == null) {
                strHtml += "<div class='LodingPanel' style='display:none; width:300px;'>";
                strHtml += "<div style='text-align:center; width:300px; height:40px; padding:20px 0 0 0; opacity:.9;'><img src='./img/ajax-loader-1.gif'></div>";
                strHtml += "<div style='font-size:15px; text-align:center; width:300px; padding:20px 0 10px 0; color:#333333; font-weight:bold;'>잠시만 기다려 주세요.</div>";
                strHtml += "</div>";
                $html.append(strHtml);
            }
        },

        //=== 페이지 영역 설정 ==============================
		// $.makePage(페이지번호, 페이지그룹 보여지는 갯수, 총 페이지 갯수, Html 입력 엘리먼트);
		// $.makePage(11, 10, 61, ".paging");
		//========================================================
        makePage: function (intPAGE, intBLOCKPAGE, intTOTALPAGE, pagingDIV) {
            if (pagingDIV == undefined) { pagingDIV = ".paging"; }
			
			// 페이지 Html 생성
			$.makePageBlock(intPAGE, intBLOCKPAGE, intTOTALPAGE, pagingDIV);
			// 생성된 Html에 Click Event 생성
            $.setPageButton();
        },

        //== 페이지 블럭 설정 (화면이동) =======================
        makePageBlock: function (intPAGE, intBLOCKPAGE, intTOTALPAGE, pagingDIV) {
            var strHTML = "";
            var varLineCOLOR = "#30C70A";
            var intTEMP;
            var intLOOP;
			
            intPAGE = parseInt(intPAGE);
            intBLOCKPAGE = parseInt(intBLOCKPAGE);
            intTOTALPAGE = parseInt(intTOTALPAGE);
            intTEMP = parseInt((intPAGE - 1) / intBLOCKPAGE) * intBLOCKPAGE + 1;
            intLOOP = 1;

            // 이전
			if (intPAGE > intBLOCKPAGE) {
                strHTML += "<span class=\"pageNUM\" value=\"" + parseInt(intTEMP - 1) + "\">"+ decodeURI("%EC%9D%B4%EC%A0%84") +"</span>";
            }

            while (intLOOP <= intBLOCKPAGE && intTEMP <= intTOTALPAGE) {
                if (intTEMP == intPAGE) {
                    strHTML += "<span class=\"pageNUM mouseOVER\" value=\"" + intTEMP + "\">" + intTEMP + "</span>";
                } else {
                    strHTML += "<span class=\"pageNUM\" value=\"" + intTEMP + "\">" + intTEMP + "</span>"
                }
                /*
                if (intLOOP != intBLOCKPAGE && intTEMP != intTOTALPAGE) {
                strHTML += "<span> &nbsp; | &nbsp; </span>";
                }
                */
                intTEMP++;
                intLOOP++;
            }

            // 다음
			if (intTOTALPAGE < intTEMP) {

            } else {
                strHTML += "<span class=\"pageNUM\" value=\"" + intTEMP + "\">"+ decodeURI("%EB%8B%A4%EC%9D%8C") +"</span>";
            }
			
            $(pagingDIV).html(strHTML);
        },

        //=== 페이지 버튼 이벤트 설정 =========================
        setPageButton: function () {
            $(".pageNUM").click(function () {
                if ($(this).attr("value") != undefined) {
                    $("#SearchForm input[name='Page']").val($(this).attr("value"));
					$("#SearchForm").attr("method", "get");
					$("#SearchForm").attr("target", "_self");
					$("#SearchForm").attr("action", $.url().attr('file'));
					$("#SearchForm").submit();
                }
            }).css("cursor", "pointer");
        },

        //=== ajaxForm 공통모튤 ======================================
        //[url, param, dataType, before, success, complete, clear, reset]=>순서상관없음
        //========================================================
        // $.Ajax4Form("#form", obj);
        // $("#form").submit();
		//========================================================
        Ajax4Form: function (form, obj, debuge) {
            debuge = debuge || false;
            var clearForm = obj["clear"] || false;
            var resetForm = obj["reset"] || false;
			var viewLoding = obj["Loding"] || true;
			
            //-------------- alert ---------------
            if (debuge) {
                alert(" url: " + obj["url"]
				+ "\r\n param :" + obj["param"]
				+ "\r\n dataType :" + obj["dataType"]
				+ "\r\n before :" + obj["before"]
				+ "\r\n success :" + obj["success"]
				+ "\r\n complete :" + obj["complete"]
				+ "\r\n clear :" + obj["clear"]
				+ "\r\n reset :" + obj["reset"]);
            }
            //-------------- alert ---------------
            $(form).ajaxForm({
                url: obj["url"] + obj["param"], // override for form's 'action' attribute
                type: "post", 					// 'get' or 'post', override for form's 'method' attribute
                dataType: obj["dataType"], 		// 'xml', 'script', or 'json' (expected server response type)
                clearForm: clearForm, 			// clear all form fields after successful submit
                resetForm: resetForm, 			// reset the form after successful submit
                beforeSubmit: function () {
					if (viewLoding == true) { $.showLoding2(); }
                    eval(obj["before"]);
                },
                success: function (datas, state) {
                    eval(obj["success"]);
                },
                error: function (reason, e) {
					alert('서버연결에 실패했습니다. : ' + e);
                },
                complete: function () {
                    if (viewLoding == true) { $.hideLoding2(); }
                    eval(obj["complete"]);
                }
            });
        },

        //=== ajax 공통 모듈 =========================================
        //	var aryData = {"type":"setChangeDate", "dateSetType":"Stop"}
        //	$.Ajax4Get("/url.asp", aryData, "$.PostponeDateSet(datas, 'Stop')", "xml", "","", true);
        //========================================================
        Ajax4Get: function (strURL, aryData, objSuccess, strDataType, objBefore, objComplete, blnLoding, blnAsync) {
			if (aryData == undefined || !aryData) { aryData = null; }
            if (objSuccess == undefined || !objSuccess) { objSuccess = null; }
            if (strDataType == undefined || !strDataType) { strDataType = "xml"; }
            if (objBefore == undefined || !objBefore) { objBefore = null; }
            if (objComplete == undefined || !objComplete) { objComplete = null; }
            if (blnLoding == undefined) { blnLoding = true; }
			if (blnAsync == undefined) { blnAsync = true; }
			
			$.ajax({
                url: strURL,
                data: aryData,
                type: "post",
                dataType: strDataType,
                async: blnAsync,
                cache: false,
                beforeSend: function () {
					if (blnLoding == true) {
                        $.showLoding2();
                    }
                    eval(objBefore);
                },
                success: function (datas, state) {
                    eval(objSuccess);
                },
                error: function (reason, e) {
					alert('서버연결에 실패했습니다. : ' + e);
                },
                complete: function () {
                    if (blnLoding == true) {
                        $.hideLoding2();
                    }
                    eval(objComplete);
                }
            });
        },
		
		//===  HEX 인코딩 ============================
		stringToHex: function (tmp) {
			var str = "";
			var i = 0;
			var tmp_len = tmp.length;
			var c;
			for (; i < tmp_len; i += 1) {
				c = tmp.charCodeAt(i);
				//str += '\u' + c.toString(16);
				str += "\\u" + c.toString(16);
			}
			return str;
		},
		
		//===  HEX 디코딩 ============================
		hexToString: function (tmp) {
			var arr = tmp.split("\\u");
			var str = "";
			var i = 0;
			var arr_len = arr.length;
			var c;
			for (; i < arr_len; i += 1) {
				c = String.fromCharCode( parseInt(arr[i], 16) );
				str += c;
			}
			return str;
		}
    });
    //=== 함수 끝 =========================================
});

//===========================================================================
// left, right 함수
function left(str, n){
	if (n <= 0) {
		return "";
	} else if (n > String(str).length) {
		return str;
	} else {
		return String(str).substring(0,n);
	}
}
function right(str, n){
	if (n <= 0) {
		return "";
	} else if (n > String(str).length) {
		return str;
	} else {
		var iLen = String(str).length;
		return String(str).substring(iLen, iLen - n);
	}
}

//===========================================================================
// 입력된 String의 공백문자 체크
// Input 
//		- obj : 검사할 Object
// Output
//		- return true : 공백이 있을 때
//		- return false : 공백이 없을때
// 사용방법 : ch_blank(this)
function ck_blank(obj) {
	var temp;
	str = obj.val();
	len = str.length;
	
	for(k=0;k<len;k++)
	{
		temp = str.charAt(k);
		if(temp == ' ')
		{
			obj.val("");
			obj.focus();
			return true;
		}
	}
	return false;
}

//===========================================================================
// Email 형식 체크 함수
// Input 
//		- obj : 텍스트 객체
// Output
//		- return true  : 올바른 이메일 형식
//		- return false : 잘못된 이메일 형식
// 사용방법 : <input type="text" ... onblur="javascript:ck_email(this);">
function ck_email(obj)
{
	validatenum(obj);
	if(obj.val() != "")
	{ 
		// 메일 형식 체크
		if(obj.val().indexOf('@') == -1 || obj.val().indexOf('.')==-1) {
			//alert("메일주소가 정확하지 않습니다.");
			alert(decodeURI("%EB%A9%94%EC%9D%BC%EC%A3%BC%EC%86%8C%EA%B0%80 %EC%A0%95%ED%99%95%ED%95%98%EC%A7%80 %EC%95%8A%EC%8A%B5%EB%8B%88%EB%8B%A4."));
			
			obj.val("");
			//obj.focus();
			return false;
		}

		// daum 메일 체크
		checkpoint = obj.val().lastIndexOf('@');
		atpoint    = obj.val().substring(checkpoint+1, obj.val().length);
		mailtype   = atpoint.toLowerCase();

		// 공백문자 체크
		if(ck_blank(obj)) {
			//alert("이메일은 공백을 포함할 수 없습니다.");
			alert(decodeURI("%EC%9D%B4%EB%A9%94%EC%9D%BC%EC%9D%80 %EA%B3%B5%EB%B0%B1%EC%9D%84 %ED%8F%AC%ED%95%A8%ED%95%A0 %EC%88%98 %EC%97%86%EC%8A%B5%EB%8B%88%EB%8B%A4."));
			obj.val("");
			//obj.focus();
			return false;
		}

		// 한글체크
		if(not_ck_kor(obj)) {
			//alert("이메일을 한글을 포함할 수 없습니다.");
			alert(decodeURI("%EC%9D%B4%EB%A9%94%EC%9D%BC%EC%9D%84 %ED%95%9C%EA%B8%80%EC%9D%84 %ED%8F%AC%ED%95%A8%ED%95%A0 %EC%88%98 %EC%97%86%EC%8A%B5%EB%8B%88%EB%8B%A4"));
			obj.val("");
			//obj.focus();
			return false;
		}
	}

	return true;
}

// 입력된 String이 한글인지 체크
// Input 
//		- str : 검사할 스트링
// Output
//		- return true : 한글일때
//		- return false : 한글이 아닐때
// 사용방법 : ck_kor_beta("가나다라마바사")
//
// 2005-04-26 추가

function ck_kor_beta(obj) {
	var temp;
	str = obj.val();
	len = str.length;

	for(k=0;k<len;k++)
	{
		temp = str.charAt(k);

		if (escape(temp).length > 4) {
			continue;
		} else {
			//alert("한글만 입력하셔야 합니다.");
			alert(decodeURI("%ED%95%9C%EA%B8%80%EB%A7%8C %EC%9E%85%EB%A0%A5%ED%95%98%EC%85%94%EC%95%BC %ED%95%A9%EB%8B%88%EB%8B%A4."));
			obj.val("");
			obj.focus();
			return false;
		}
	}

	return true;
}

//===========================================================================
// 입력된 String이 한글이면 안되는 체크
// Input 
//		- str : 검사할 스트링
// Output
//		- return true : 한글일때
//		- return false : 한글이 아닐때
// 사용방법 : not_ck_kor("가나다라마바사")
function not_ck_kor(obj) {
	var temp;
	str = obj.val();
	len = str.length;

	for(k=0;k<len;k++)
	{
		temp = str.charAt(k);

		if (escape(temp).length > 4) {
//			alert("한글은 입력하실 수 없습니다.")
			obj.val("");
			obj.focus();
			return true;
		} else {
			continue;
		}
	}

	return false;
}

function validatenum(obj) {
	var num="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@._-";
	obj.returnValue = true;

	for (var i=0;i<obj.val().length;i++)
		if (-1 == num.indexOf(obj.val().charAt(i)))
		   obj.returnValue = false;
}

function validatenum2(obj) {
	var num="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-";

	for (var i=0;i<obj.val().length;i++) {
		if (-1 == num.indexOf(obj.val().charAt(i))) {
			obj.val("");
			obj.focus();
			return true;
		} else {
			continue;
		}
	}
	
	return false;
}

// 브라우저 체크
//================================================================
var isie = (navigator.userAgent.toLowerCase().indexOf('msie') != -1) ? true : false;
var isie6 = (navigator.userAgent.toLowerCase().indexOf('msie 6') != -1) ? true : false;
//var isie7=(navigator.userAgent.toLowerCase().indexOf('msie 7')!=-1)? true : false;
if (navigator.userAgent.toLowerCase().indexOf('msie 7') != -1) {
	isie6 = false;
	var isie7 = true;
}
if (navigator.userAgent.toLowerCase().indexOf('msie 8') != -1) {
	isie6 = false;
	var isie8 = true;
}
var isfirefox = (navigator.userAgent.toLowerCase().indexOf('firefox') != -1) ? true : false;
var isopera = (navigator.userAgent.toLowerCase().indexOf('opera') != -1) ? true : false;
//================================================================