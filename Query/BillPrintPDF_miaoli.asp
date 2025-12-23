<%

Set Pdf = Server.CreateObject("Persits.Pdf")   
Set Doc = Pdf.CreateDocument

Doc.ImportFromUrl "http://192.168.1.210/traffic/Query/BillBaseFastPaper_miaoli.asp?PBillSN="&trim(Request("PBillSN")), "scale=0.6; hyperlinks=true; drawbackground=true"
Filename = Doc.Save( Server.MapPath(".\img\"&Session("USER_ID")&".pdf"), true)
response.Redirect ".\img\"&Session("USER_ID")&".pdf?nowtime="&now
%>