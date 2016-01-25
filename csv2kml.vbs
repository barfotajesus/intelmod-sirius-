Dim	strLine
Dim	arrFields
Dim	strMessage
Dim	strHeader
Dim	strFuter
Dim strTab3
strTab3 =  vbTab & vbTab & vbTab
Dim tt(9)
tt(0)=0 '名前
tt(1)=1 'MOD AXA/AXA
tt(2)=2 'Lv
tt(3)=3 'Tm
tt(4)=5 '＠
tt(5)=6 'Hack 4[300s]
tt(6)=8 'Lat 37
tt(7)=9 'Lng 140
tt(8)=10 'URL

Set Stream = CreateObject( "ADODB.Stream" )
Set Stream2 = CreateObject( "ADODB.Stream" )

Stream.Open
Stream.Type = 2	' テキスト
Stream.Charset = "UTF-8"
Stream.LoadFromFile "C:\Users\Koyama\Desktop\test.csv"
Stream.LineSeparator = 10	' LF
Stream.Position = 0

Stream2.Open
Stream2.Type = 2	' テキスト
'Stream2.Charset = "shift_jis"
Stream2.Charset = "UTF-8"
Stream2.LineSeparator = -1	' CRLF

strHeader = _
	 "<?xml version='1.0' encoding='UTF-8'?>" & vbCrLf _
	& "<kml xmlns='http://www.opengis.net/kml/2.2'>" & vbCrLf _
	& vbTab &"<Document>" & vbCrLf _
	& vbTab & vbTab &"<name>" + Cstr(Now()) + "</name>" ' & vbCrLf 最後は要らないぽい
'	& vbTab & vbTab &"<name>" + Cstr(Date()) + "</name>" ' & vbCrLf 最後は要らないぽい
strFuter = _
	  vbTab & vbTab & "<Style id='icon-P8E'>" & vbCrLf _
	& vbTab & vbTab & "<IconStyle>" & vbCrLf _
	& vbTab & vbTab & "<color>ff579D00</color>" & vbCrLf _
	& vbTab & vbTab & "<scale>1.1</scale>" & vbCrLf _
	& vbTab & vbTab & "<Icon>" & vbCrLf _
	& vbTab & vbTab & "<href>http://www.gstatic.com/mapspro/images/stock/960-wht-star-blank.png</href>" & vbCrLf _
	& vbTab & vbTab & "</Icon>" & vbCrLf _
	& vbTab & vbTab & "</IconStyle>" & vbCrLf _
	& vbTab & vbTab & "</Style>" & vbCrLf _
	& vbTab & vbTab & "<Style id='icon-P7E'>" & vbCrLf _
	& vbTab & vbTab & "<IconStyle>" & vbCrLf _
	& vbTab & vbTab & "<color>ff579D00</color>" & vbCrLf _
	& vbTab & vbTab & "<scale>1.1</scale>" & vbCrLf _
	& vbTab & vbTab & "<Icon>" & vbCrLf _
	& vbTab & vbTab & "<href>http://www.gstatic.com/mapspro/images/stock/962-wht-diamond-blank.png</href>" & vbCrLf _
	& vbTab & vbTab & "</Icon>" & vbCrLf _
	& vbTab & vbTab & "</IconStyle>" & vbCrLf _
	& vbTab & vbTab & "</Style>" & vbCrLf _
	& vbTab & vbTab & "<Style id='icon-P8R'>" & vbCrLf _
	& vbTab & vbTab & "<IconStyle>" & vbCrLf _
	& vbTab & vbTab & "<color>ffF08641</color>" & vbCrLf _
	& vbTab & vbTab & "<scale>1.1</scale>" & vbCrLf _
	& vbTab & vbTab & "<Icon>" & vbCrLf _
	& vbTab & vbTab & "<href>http://www.gstatic.com/mapspro/images/stock/960-wht-star-blank.png</href>" & vbCrLf _
	& vbTab & vbTab & "</Icon>" & vbCrLf _
	& vbTab & vbTab & "</IconStyle>" & vbCrLf _
	& vbTab & vbTab & "</Style>" & vbCrLf _
	& vbTab & vbTab & "<Style id='icon-P7R'>" & vbCrLf _
	& vbTab & vbTab & "<IconStyle>" & vbCrLf _
	& vbTab & vbTab & "<color>ffF08641</color>" & vbCrLf _
	& vbTab & vbTab & "<scale>1.1</scale>" & vbCrLf _
	& vbTab & vbTab & "<Icon>" & vbCrLf _
	& vbTab & vbTab & "<href>http://www.gstatic.com/mapspro/images/stock/962-wht-diamond-blank.png</href>" & vbCrLf _
	& vbTab & vbTab & "</Icon>" & vbCrLf _
	& vbTab & vbTab & "</IconStyle>" & vbCrLf _
	& vbTab & vbTab & "</Style>" & vbCrLf _
	& vbTab & "</Document>" & vbCrLf _
	& "</kml>" & vbCrLf _

Stream2.WriteText strHeader,1
Do While not Stream.EOS
    strMessage = strMessage & vbTab & vbTab &"<Placemark>" & vbCrLf

	' -2 は、ストリームから次の行を読み取ります
	strLine = Stream.ReadText( -2 )
'WScript.Echo strLine
	arrFields = Split(strLine,vbTab)
	if(arrFields(tt(3))="1")then
	  if(arrFields(tt(2))="8")then
	    strMessage = strMessage & strTab3 &"<styleUrl>#icon-P8R</styleUrl>" & vbCrLf
	  else
	    strMessage = strMessage & strTab3 &"<styleUrl>#icon-P7R</styleUrl>" & vbCrLf
	  end if
	else if(arrFields(tt(2))="8")then
			strMessage = strMessage & strTab3 &"<styleUrl>#icon-P8E</styleUrl>" & vbCrLf
		else
			strMessage = strMessage & strTab3 &"<styleUrl>#icon-P7E</styleUrl>" & vbCrLf
		end if
	end if
    strMessage = strMessage & strTab3 &"<name>P" & arrFields(tt(2)) & " "& arrFields(tt(0)) & ",@" & arrFields(tt(4)) & "," & arrFields(tt(5)) & "</name>" & vbCrLf
    strMessage = strMessage & strTab3 &"<description><![CDATA[" & arrFields(tt(1)) & "<br>URL> " & arrFields(tt(8)) & ",]]></description>" & vbCrLf
    strMessage = strMessage & strTab3 &"<ExtendedData></ExtendedData>" & vbCrLf
    strMessage = strMessage & strTab3 &"<Point>" & vbCrLf
    strMessage = strMessage & strTab3 & vbTab &"<coordinates>" & Left(arrFields(tt(7)),3) & "." & mid(arrFields(tt(7)),4) & "," _
		& Left(arrFields(tt(6)),2) & "." & mid(arrFields(tt(6)),3) & ",0.0</coordinates>" & vbCrLf
    strMessage = strMessage & strTab3 &"</Point>" & vbCrLf
    strMessage = strMessage & vbTab & vbTab &"</Placemark>" ' & vbCrLf 最後は要らないぽい
	
	Stream2.WriteText strMessage,1
	strMessage = ""
'	Stream2.WriteText Stream.ReadText( -2 ), 1
Loop
Stream2.WriteText strFuter,1

Stream2.SaveToFile "C:\Users\Koyama\Desktop\test.kml", 2

Stream.Close
Stream2.Close
