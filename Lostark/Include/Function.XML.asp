<%
'== 사용법 ===============================================
'dim oDOM, url
'  
' XML 데이터 주소
'url = "http://m.cinetong.com/movie/movieinfo.asp?cineNo=24943"
'
'Set oDOM = new XMLDOMClass
'with oDOM
'	if .LoadHTTP(url) Then
'		Response.Write "cineNo:" & .TagText("cineNo", 0) & "<br>"
'		Response.Write "etitle:" & .TagText("etitle", 0) & "<br>"
'		Response.Write "Actor:" & .TagText("Actor", 0) & "<br>"
'		Response.Write "audiences:" & .TagText("audiences", 0) & "<br>"
'		
'		For each tempNodes in .Nodes("img")
'			Response.Write "img : " & tempNodes.Text & "<br>"
'		Next
'	else
'		Response.Write "XML을 읽어오는데 실패하였습니다."
'	End if
'end with
'Set oDOM = Nothing

'== MXL 소스 ==
'<?xml version="1.0" encoding="utf-8"?>
'<cinetong>
'<movie>
'<cineNo>24943</cineNo>
'<pimg>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/20110128101110569.jpg</pimg>
'<mtitle><![CDATA[킹스 스피치]]></mtitle>
'<etitle><![CDATA[The King's Speech]]></etitle>
'<info>드라마 | 영국,오스트레일리아,미국 | 118분</info>
'<openDate>2011.03.17</openDate>
'<Score>8.00</Score>
'<Actor><![CDATA[톰 후퍼]]></Actor>
'<Person><![CDATA[헬레나 본햄 카터,가이 피어스,콜린 퍼스,마이클 갬본,제프리 러쉬,데릭 자코비,티모시 스폴,안소니 앤드류즈]]></Person>
'<Grade><![CDATA[12]]></Grade>
'<Supply><![CDATA[㈜ 화앤담이엔티]]></Supply>
'<Site><![CDATA[http://www.kingsspeech.com]]></Site>
'<audiences>808,347 명</audiences>
'<Produce><![CDATA[]]></Produce>
'<TODAY>TODAY : 2011.05.17</TODAY>
'<Synopsis><![CDATA[연합군의 비밀무기는 말더듬이 영국왕?! 
'세상을 감동시킨 국왕의 컴플렉스 도전이 시작된다! 
'때는 1939년, 세게의 스캔들을 일으키며 왕위를 포기한 형 때문에 본의 아니게 왕위에 오른 버티. 권력과 명예, 모든 것을 다 가진 그에게도 두려운 것이 있었으니 바로 마.이.크! 그는 사람들 앞에 서면 “더더더...” 말을 더듬는 컴플렉스를 가졌던 것! 국왕의 자리가 버겁기만 한 터비와 그를 지켜보는 아내 엘리자베스 왕비, 그리고 국민들도 애가 타기는 마찬가지... 게다가 지금 세계는 2차 세계대전중! 불안한 정세 속 새로운 지도자를 간절히 원하는 국민들을 위해 버티는 아내의 소개로 괴짜 언어 치료사 라이오넬 로그를 만나게 되고, 삐걱거리는 첫 만남 이후 둘을 기상천외한 치료법을 통해 말더듬증 극복에 도전하게 되는데... 세기의 선동가 히틀러와 맞선 말더듬이 영국왕... 과연 그는 국민의 마음을 감동시킬 연설에 성공할 수 있을까?
']]></Synopsis>
'</movie>
'<Stillcut>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/01.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/02.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/03.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/04.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/05.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/06.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/07.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/08.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/09.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/09.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/14.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/16.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/18.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/21.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/22.jpg</img>
'<img>http://img.cinetong.com/cineInfo/cine/Y2010/1109/20101109152944411/stillCut/6.jpg</img>
'</Stillcut>
'</cinetong>
'=========================================================

'== XML 데이터 읽기 ==========================================
Class XMLDOMClass
	Private m_DOM ' XMLDOM 객체

	' ---------------------- 생성자 -----------------------
	Private Sub Class_Initialize()
		Set m_DOM = Server.CreateObject("Microsoft.XMLDOM")
	End Sub
   
	' ---------------------- 소멸자 -----------------------
	Private Sub Class_Terminate()
		Set m_DOM = Nothing
	End Sub
   
	' ------------------- Property Get --------------------
	Public Property Get TagText(tagName, index)
		Dim Nodes
		
		Set Nodes	= m_DOM.getElementsByTagName(tagName)
		TagText		= Nodes(index).Text
		Set Nodes	= Nothing
	End Property
	
	Public Property Get SubTagText(Nodes, index)
		SubTagText	= Nodes.Childnodes(index).Text
	End Property
   
	Public Property Get Nodes(tagName)
		Set Nodes = m_DOM.getElementsByTagName(tagName)
	End Property

	' ------------------- 원격 XML 읽기 --------------------
	Public Function LoadHTTP(url)
		with m_DOM
			.async = False ' 동기식 호출
			.setProperty "ServerHTTPRequest", True ' HTTP로 XML 데이터 가져옴
			LoadHTTP = .Load(url)
		end with  
	End Function

	' ------------------- XML 파일 읽기 --------------------
	Public Function Load(file)
		with m_DOM
			.async	= False ' 동기식 호출
			Load		= .Load( Server.MapPath(file) )
		end with
	End Function
End Class
'===========================================================
%>