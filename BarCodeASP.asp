<HTML>
<BODY>
<TITLE> Testing Delphi ASP </TITLE>
<CENTER>
<H3> You should see the results of your Delphi Active Server method below </H3>
</CENTER>
<HR>
<% 
SysSN="9527"
www="2006/9/28"
qaa=162
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
   'DelphiASPObj.GenBillPrintBarCode 64,"9999","","U9-3471","073828","220073","100","950903","44","台北市交通事件裁決所",1,300,0,True,False,"2006/9/12 10:50:20"
    'DelphiASPObj.GenBillPrintBarCode 988,"RA3786441","6020303","RU-7199","163","115001","206","951223","22","台北市交通事件裁決所",0,900,0,True,False,"2006/9/12 10:50:20"
    DelphiASPObj.GenBillPrintBarCode 1481,"WW0000461","4310102","K7F-880","288627","220073","224","951230","42","台北市交通事件裁決所",0,12000,0,True,False,"2006/11/30"
%>
<HR>
</BODY>
</HTML>