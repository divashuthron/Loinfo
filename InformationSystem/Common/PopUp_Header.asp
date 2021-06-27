<%@  codepage="65001" language="VBScript" %>
<!--#InClude Virtual = "/Common/Include/Function.asp" -->
<%
If IsE(SessionUserID) Then
	Response.Write "<script language='javascript'>"
	Response.Write "location.href='/Login.asp';"
	Response.Write "</script>"
	Response.End
End If
%>

<!DOCTYPE html>
<html>

<head>

    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">

	<title>대학종합정보시스템 - (주)메티소프트</title>

    <!-- Default Style -->
	<link href="/css/bootstrap.min.css" rel="stylesheet">
    <link href="/font-awesome/css/font-awesome.css" rel="stylesheet">

    <!-- Plugins Style-->
	<!--<link href="/css/plugins/iCheck/custom.css" rel="stylesheet">
    <link href="/css/plugins/select2/select2.min.css" rel="stylesheet">
	<link href="/css/plugins/chosen/bootstrap-chosen.css" rel="stylesheet">
	<link href="/css/plugins/dataTables/datatables.min.css" rel="stylesheet">
	<link href="/css/plugins/awesome-bootstrap-checkbox/awesome-bootstrap-checkbox.css" rel="stylesheet"> 
	<link href="/css/plugins/clockpicker/clockpicker.css" rel="stylesheet">
	
	<link href="/css/plugins/chosen/bootstrap-chosen.css" rel="stylesheet">
    <link href="/css/plugins/bootstrap-tagsinput/bootstrap-tagsinput.css" rel="stylesheet">
    <link href="/css/plugins/colorpicker/bootstrap-colorpicker.min.css" rel="stylesheet">
    <link href="/css/plugins/cropper/cropper.min.css" rel="stylesheet">
    <link href="/css/plugins/switchery/switchery.css" rel="stylesheet">
    <link href="/css/plugins/jasny/jasny-bootstrap.min.css" rel="stylesheet">
    <link href="/css/plugins/nouslider/jquery.nouislider.css" rel="stylesheet">
    <link href="/css/plugins/datapicker/datepicker3.css" rel="stylesheet">
    <link href="/css/plugins/ionRangeSlider/ion.rangeSlider.css" rel="stylesheet">
    <link href="/css/plugins/ionRangeSlider/ion.rangeSlider.skinFlat.css" rel="stylesheet">
    <link href="/css/plugins/daterangepicker/daterangepicker-bs3.css" rel="stylesheet">
	<link href="/css/plugins/touchspin/jquery.bootstrap-touchspin.min.css" rel="stylesheet">
    <link href="/css/plugins/dualListbox/bootstrap-duallistbox.min.css" rel="stylesheet">
	-->

	<!-- Basic Style -->
	<link href="/css/animate.css" rel="stylesheet">
    <link href="/css/style.css" rel="stylesheet">

	<!-- Your Style -->
	<link href="/css/your_style.css" rel="stylesheet">

	<!-- FAVICONS -->
	<link rel="shortcut icon" href="/img/favicon/favicon.ico" type="image/x-icon">
	<link rel="icon" href="/img/favicon/favicon.ico" type="image/x-icon">

	<!-- GOOGLE FONT -->
	<link rel="stylesheet" href="http://fonts.googleapis.com/css?family=Open+Sans:400italic,700italic,300,400,700">

	<!-- Link to Google CDN's jQuery + jQueryUI; fall back to local -->
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>

	<!-- daum 주소찾기 -->
	<script src="http://dmaps.daum.net/map_js_init/postcode.v2.js"></script>

	<script>
		if (!window.jQuery) {
			document.write('<script src="js/libs/jquery-2.1.1.min.js"><\/script>');
		}
	</script>

	<script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
	<script>
		if (!window.jQuery.ui) {
			document.write('<script src="js/libs/jquery-ui-1.10.3.min.js"><\/script>');
		}
	</script>