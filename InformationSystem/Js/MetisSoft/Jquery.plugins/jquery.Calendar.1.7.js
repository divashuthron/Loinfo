/*
//== 사용법 ========================================
<input type="text" name="인풋name" id="인풋ID"/>
<span class="viewCalendarBtn btnBasic" Obj="인풋ID" CalendarDIV="레이어ID" left="left위치" top="top위치">달력</span>
// ================================================
*/
$(function() {
	//$(".viewCalendarBtn").click(function() {
	$(document).on("click", ".viewCalendarBtn", function(e) {
		var Obj = $(this).attr("Obj");
		var CalendarDIV = ($(this).attr("CalendarDIV") == undefined || !$(this).attr("CalendarDIV")) ? Obj + "CalendarDIV" : $(this).attr("CalendarDIV");
		var left = ($(this).attr("left") == undefined || !$(this).attr("left")) ? ($(this).offset().left - 60) : $(this).attr("left");
		var top = ($(this).attr("top") == undefined || !$(this).attr("top")) ? ($(this).offset().top + $(this).height() + 3) : $(this).attr("top");
		//var top = ($(this).attr("top") == undefined || !$(this).attr("top")) ? ($(this).offset().top - 300) : $(this).attr("top");
		
		// 위치 조정
		//if (parseInt(left) + 300 >= parseInt($(window).width())) { left = $(window).width() - 300; }
		if (parseInt(top) + 200 > parseInt($(window).height())) { top = top - 305; }

		$.findCalendarDIV(CalendarDIV, left, top);
		$.viewCalendar(CalendarDIV, Obj, "CalendarOpen");
	});
	
	$(document).on("click", ".viewCalendarBtn2", function(e) {
		var CalendarDIV = $(this).attr("CalendarDIV");
		var Obj = $(this).attr("Obj");
		var Date = $(this).attr("Date").split("-");
		var Year = Date[0];
		var Momth = Date[1];
		
		$.viewCalendar(CalendarDIV, Obj, "CalendarMonth", Year, Momth);
	});
	
	$(document).on("change", ".changeCalendar", function(e) {
		var CalendarDIV = $(this).attr("CalendarDIV");
		var Obj = $(this).attr("Obj");
		var Year  = $("#yearSelect_"+ CalendarDIV).val();
		var Momth  = $("#monthSelect_"+ CalendarDIV).val();
		
		$.viewCalendar(CalendarDIV, Obj, "CalendarMonth", Year, Momth);
	});
	
	$(document).on("click", ".selectCalender", function(e) {
		$("#"+ $(this).attr("Obj")).text($(this).attr("Date"));
		$("#"+ $(this).attr("Obj")).val($(this).attr("Date"));
		//$("#"+ $(this).attr("Obj") +"_REAL").val($(this).attr("Date").replace(/-/g, ""));
		$("#"+ $(this).attr("Obj") +"_REAL").val($(this).attr("Date").replace(/-/g, ""));
		$("#"+ $(this).attr("CalendarDIV")).hide();
	});
	/*
	.on("mouseenter", ".selectCalender",
		function () {
			$(this).addClass("CalenderMouseOver");
		}
	)
	.on("mouseleave", ".selectCalender",
		function () {
			$(this).removeClass("CalenderMouseOver");
		}
	);
	*/
	
	$(document).on("click", ".CloseCalendarDIV", function(e){
		$("#"+ $(this).attr("CalendarDIV")).hide();
	});

	/*
	$(document).on("focusin", ".viewCalendarBtn input[type='text']", function(e){
		$(this).closest(".viewCalendarBtn").click();
	});
	*/
	
	$.extend({
		findCalendarDIV : function (CalendarDIV, left, top) {
            var $html = $("body");
            var strHtml = "";
			
			if ($html.find("#"+ CalendarDIV).html() == null) {
				strHtml += "<div id=\""+ CalendarDIV +"\" class=\"CalendarDIV\" style=\"display:none;left:"+ left +"px;top:"+ top +"px;z-index:2000;\"></div>";
                $html.append(strHtml);
            }
        },
		// 달력 보기
		viewCalendar : function(CalendarDIV, Obj, CalendarType, Year, Momth) {
			if (CalendarType == undefined || !CalendarType) { CalendarType = "CalendarOpen"; }
			//var ObjValue = $("#"+ Obj).val();
			var ObjValue = $("#"+ Obj).text();
			var divide = "-"
			var className, styleName, currentDates, CalenderDate;
			var tempdate = new Date();
			var NowDate = tempdate.getFullYear() + divide + $.getCipher(tempdate.getMonth()+1) + divide + $.getCipher(tempdate.getDate());
			var yearCount	= 1;
			var monthCount = 1;
			var WeekCount = 0;
			var j = 0;
			
			if (CalendarType == "CalendarOpen") {
				if (!ObjValue) {
					Year = tempdate.getFullYear();
					Momth = tempdate.getMonth() + 1;
				} else {
					currentDates = ObjValue.split(divide);
					Year = parseInt(currentDates[0],10);
					Momth = parseInt(currentDates[1],10);
				}
			} else {
				Year = parseInt(Year);
				Momth = parseInt(Momth);
			}
			
			var d1 = (Year+(Year-Year%4)/4-(Year-Year%100)/100+(Year-Year%400)/400 +Momth*2+(Momth*5-Momth*5%9)/9-(Momth<3?Year%4||Year%100==0&&Year%400?2:3:4))%7;
			var arryMonth = new Array(0,31,28,31,30,31,30,31,31,30,31,30,31);
			
			if ((Year % 4 == 0) && (Year % 100 != 0) || (Year % 400 == 4)) {
				arryMonth[1]=29;
			}
			
			if (CalendarType == "CalendarMonth" || $("#"+ CalendarDIV).css("display") == "none") {
				var text_calendar = "";
				
				text_calendar += "<table width=\"280\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\">";
				text_calendar += "	<tr>";
				text_calendar += "		<td height=\"43\" align=\"center\">";
				
				text_calendar += "			<table width=\"280\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">";
				text_calendar += "				<tr>";
				text_calendar += "					<td width=\"10\" height=\"38\"></td>";
				text_calendar += "					<td width=\"30\" align=\"right\">";
				text_calendar += "						<img src=\"./img/carlendarPrev.gif\" class=\"viewCalendarBtn2\" CalendarDIV=\""+ CalendarDIV +"\" Obj=\""+ Obj +"\" Date=\""+ (Momth==1?(Year-1)+"-"+12:Year+"-"+(parseInt(Momth)-1)) +"\" alt=\"이전\" style=\"cursor:pointer;\"/>";
				text_calendar += "					</td>";
				text_calendar += "					<td width=\"160\" align=\"center\">"
				text_calendar += "						<select id=\"yearSelect_"+ CalendarDIV +"\" class=\"changeCalendar\" CalendarDIV=\""+ CalendarDIV +"\" Obj=\""+ Obj +"\" style=\"width:55px;\">";

				for (yearCount = tempdate.getFullYear() - 10; yearCount<=tempdate.getFullYear() + 10 ; yearCount++) {
					if (yearCount == Year){
						text_calendar +="					<option selected=\"selected\" value=\""+yearCount+"\">"+yearCount+"</option>\n";
					} else {
						text_calendar +="					<option value=\""+yearCount+"\">"+yearCount+"</option>\n";
					}
				}
				text_calendar += "						</select> 년\n";
				text_calendar += "						<select id=\"monthSelect_"+ CalendarDIV +"\" class=\"changeCalendar\" CalendarDIV=\""+ CalendarDIV +"\" Obj=\""+ Obj +"\" style=\"width:45px;\">";

				for (monthCount = 1; monthCount <= 12; monthCount++) {
					if (monthCount == Momth) {
						text_calendar += "					<option selected=\"selected\" value=\""+monthCount+"\">"+monthCount+"</option>";
					} else {
						text_calendar += "					<option value=\""+monthCount+"\">"+monthCount+"</option>";
					}
				}

				text_calendar += "						</select> 월\n";
				text_calendar += "					</td>";
				text_calendar += "					<td width=\"30\" align=\"left\">";
				text_calendar += "						<img src=\"./img/carlendarNext.gif\" class=\"viewCalendarBtn2\" CalendarDIV=\""+ CalendarDIV +"\" Obj=\""+ Obj +"\" Date=\""+ (Momth==12?(Year+1)+"-"+1:Year+"-"+(parseInt(Momth)+1)) +"\" alt=\"다음\" style=\"cursor:pointer;\"/>";
				text_calendar += "					</td>";
				text_calendar += "					<td width=\"50\" align=\"center\"><img src=\"./img/btn_close3.gif\" class=\"CloseCalendarDIV\" CalendarDIV=\""+ CalendarDIV +"\" style=\"cursor:pointer;\"></td>";
				text_calendar += "				</tr>";
				text_calendar += "			</table>";
				
				text_calendar += "		</td>";
				text_calendar += "	</tr>";
				text_calendar += "	<tr>";
				text_calendar += "		<td align=\"center\">";
				
				text_calendar += "			<table width=\"100%\" border=\"1\" cellpadding=\"5\" cellspacing=\"0\" class=\"TableBorder_Calendar\">";
				text_calendar += "				<tr bgcolor=\"f3f3f3\">";
				text_calendar += "					<td width=\"14%\" height=\"30\" align=\"center\"><font color=\"#d41200\"><strong>Sun</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#676767\"><strong>Mon</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#676767\"><strong>Tue</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#676767\"><strong>Wed</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#676767\"><strong>Thu</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#676767\"><strong>Fri</strong></font></td>";
				text_calendar += "					<td width=\"14%\" align=\"center\"><font color=\"#009bd0\"><strong>Sat</strong></font></td>";
				text_calendar += "				</tr>";
				text_calendar += "				<tr>";

				for (i = 1; i <= d1; i++) {
					text_calendar += "				<td height=\"30\" align=\"center\"></td>\n";
					WeekCount++
				}
				
				for (i = 1; i <= arryMonth[Momth]; i++) {
					CalenderDate = Year + divide + $.getCipher(Momth) +divide+ $.getCipher(i);
					
					if (WeekCount % 7 == 0) {
						styleName ="color:#d41200;";
					 } else if(WeekCount % 7 == 6) {
						styleName ="color:#009bd0;";
					} else {
						styleName = "";
					}
					//오늘 날짜 스타일 변경
					styleName += (CalenderDate == NowDate) ? "background-color:#fff4d5;font-weight:bold;" : "";
					//선택 날짜 스타일 변경
					styleName += (CalenderDate == ObjValue) ? "background-color:#F5B554;font-weight:bold;" : "";
					
					text_calendar += "			<td height=\"30\" align=\"center\" style=\"cursor:pointer;"+ styleName +"\" class=\"selectCalender\" CalendarDIV=\""+ CalendarDIV +"\" Obj=\""+ Obj +"\" Date=\""+ CalenderDate +"\">";
					text_calendar += "			"+ (i);
					text_calendar += "			</td>\n";
					
					WeekCount++
					
					if (WeekCount % 7 == 0 && arryMonth[Momth] != i){
						text_calendar +="		</tr>\n";
						text_calendar +="		<tr>\n";
						WeekCount = 0;
					}
				}
				
				if (WeekCount != 0) {
					for (i = WeekCount; i < 7; i++) {
						text_calendar += "				<td height=\"30\" align=\"center\"></td>\n";
					}
				}
				
				text_calendar +="				</tr>\n";
				text_calendar +="			</table>\n";
				
				text_calendar += "		</td>";
				text_calendar += "	</tr>";
				text_calendar += "	<tr>";
				text_calendar += "		<td align=\"center\" height=\"30\" style=\"padding:10px 0 7px 0;\">";
				text_calendar += "			<table width=\"100%\" border=\"1\" cellpadding=\"5\" cellspacing=\"0\" class=\"TableBorder_Calendar\">";
				text_calendar += "				<tr bgcolor=\"f3f3f3\">";
				text_calendar += "					<td height=\"25\" align=\"center\" style=\"font-size:12px;color:#676767;font-weight:bold;\">\n";
				text_calendar += "					오늘 : "+tempdate.getFullYear()+ divide +$.getCipher(tempdate.getMonth()+1)+ divide +$.getCipher(tempdate.getDate())+"\n";
				text_calendar += "					</td>\n";
				text_calendar += "				</tr>";
				text_calendar +="			</table>\n";
				text_calendar +="		</td>\n";
				text_calendar +="	</tr>\n";
				text_calendar +="</tabel>\n";
				
				$(".CalendarDIV").hide();
				$("#"+ CalendarDIV).html(text_calendar).show();
				
				$(".selectCalender").hover(
					function () { $(this).addClass("CalenderMouseOver"); },
					function () { $(this).removeClass("CalenderMouseOver");}
				);
			}else{
				$("#"+ CalendarDIV).hide();
			}
		}
	})
});