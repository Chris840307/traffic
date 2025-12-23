  <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/OldBillDataDB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單列印-Legal Size</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family:"標楷體"; font-size: 10px;}
.style2 {font-family:"標楷體"; font-size: 12px;}
.style3 {font-family:"標楷體"; font-size: 16px;}
.style4 {font-family:"標楷體"; font-size: 22px;}
.style5 {font-family:"標楷體"; font-size: 12px; color:#ff0000; }
.style6 {font-family:"標楷體"; font-size: 14px; color:#ff0000; }
.style7 {font-family:"標楷體"; font-size: 20px;}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsxie8.cab#Version=6,5,439,50">
</object>
<%

Function GetData(Fields,table,coniditon,code)
	tmp=""
	tmpsql="Select "&Fields&" from "&table&" where "&coniditon&"='"&code&"'"
	Set rstmp=conn.execute(tmpsql)
	If Not rstmp.eof Then tmp=rstmp(0)
	Set rstmp=Nothing
	GetData=tmp
End Function

'on Error Resume Next
Function GetCDates(tdate)
tmp=""
  If Trim(tdate)<>"" Then 
  tmp=Year(tdate)-1911&"/"&Right("0"&Month(tdate),2)&"/"&Right("0"&day(tdate),2)
  Else
  tmp="//"
  End If 

  GetCDates=tmp
End Function

Function GetCDates2(tdate)
tmp=""
  If Trim(tdate)<>"" Then 
  tmp=mid(tdate,1,2)&"/"&mid(tdate,3,2)&"/"&mid(tdate,5,2)
  Else
  tmp="//"
  End If 

  GetCDates2=tmp
End Function

           function QuotedStr(Str)
                QuotedStr="'"+Str+"'"
            end function
           
			'判斷檔案是否存在
			function HaveFile(FileName)
				dim fs
				set fs=Server.CreateObject("Scripting.FileSystemObject")
				if fs.FileExists(FileName)=true then
					HaveFile = "1"
				else
					HaveFile = "0"
				end if
				set fs=nothing
			end function
            '判斷如果是0的話回傳&nbsp;
            function ReplaceSpace(value)
                if trim(value)="" then
                    ReplaceSpace =  ""
                else
                    ReplaceSpace = value 
                end if 
            end function
             
           '判斷法條第八碼，將其組合起來
           function composeLaw(value1,value2)
                if trim(value2) <> "" then
                    composeLaw = value1 & value2
                else
                    composeLaw = value1
                end if  
           end function
             
           '組小數點位數
            function composeDot(value1,value2)  
                if (trim(value1) <> "") and (trim(value2) <> "") then
                    composeDot = value1 & "." & value2
                end if 
                if (trim(value1) <> "") and (trim(value2) = "") then
                    composeDot = value1
                end if
            end  function
            
            '查詢單位名稱
            function QueryUnitName(value)
                UnitSql="Select Acc_NM from accnew where ACC_No=" & QuotedStr(trim(value))
                set UnitRs=conn.execute(UnitSql)
                if  not UnitRs.Eof then
                    QueryUnitName = UnitRs("Acc_NM")
                end if      
                UnitRs.close
            end function    
            
            '查詢操作人員名稱
            function QueryOperat(value)  
                OperatSql="Select OPName from operat where Operat=" & QuotedStr(trim(value))
                set OperatRs=conn.execute(OperatSql)
                if  not OperatRs.Eof then
                    QueryOperat = OperatRs("OPName")
                end if 
                OperatRs.close
            end  function   
            
            '查詢某一個欄位
            function SelectFld(TableName,Fld,Cond)  
                QuerySql="Select " & Fld & " from " & TableName & " where " & Cond
                set QueryRS=conn.execute(QuerySql)

                if  not QueryRS.Eof then
                    SelectFld = QueryRS(Fld)
                end if 
                QueryRS.close
            end  function   
            
            '選擇勾選條件
            function chkPS(value)  
                if trim(value)="0" then
                    chkPS="郵繳"
                else
                    chkPS="到案" 
                end if 
            end  Function
            
						'查詢DCI狀態
            function QryDCIState(value)  
                if trim(value)="00" then
                  QryDCIState="未寫入資料庫"
                ElseIf trim(value)="Y" then
									QryDCIState="寫入資料庫"
								ElseIf trim(value)="N" then
									QryDCIState="未寫入資料庫"
								ElseIf trim(value)="S" then
									QryDCIState="違規人已先繳結案"
								ElseIf trim(value)="L" then
									QryDCIState="已入案過"
								ElseIf trim(value)="n" then
									QryDCIState="不可寫入,監理單位已入案"
                end if 
            end  Function

            '查詢法條內容
            function QueryLawContent(value,RealSpeed,LimitSpeed) 
                LawSql="Select * from traffic3.rule_n where Rule_C=" & QuotedStr(trim(value))
                set LawRs=conn.execute(LawSql)
                if not LawRs.Eof then
                    if (mid(LawRs("Rule_c"),1,3)="293") and (LawRs("A_DESC")="1") then
                        QueryLawContent = replace(LawRs("Rule_D"),"重量 噸","重量 " & LimitSpeed & " 噸")
                        QueryLawContent = replace(QueryLawContent,"過磅 噸","過磅 " & RealSpeed & " 噸")
                        QueryLawContent = replace(QueryLawContent,"超載 噸","超載 " & RealSpeed-LimitSpeed & " 噸")
                    elseif (mid(LawRs("Rule_c"),1,4)="4010") and (LawRs("A_DESC")="1") then
                        QueryLawContent = replace(LawRs("Rule_D"),"限速 公里","限速 " & LimitSpeed & " 公里")
                        QueryLawContent = replace(QueryLawContent,"時速 公里","時速 " & RealSpeed & " 公里")
                        QueryLawContent = replace(QueryLawContent,"超速 公里","超速 " & RealSpeed-LimitSpeed & " 公里")
                    else
                         QueryLawContent = LawRs("Rule_D")
                    end if 
                end if 
            end  function
             
            '檢查保險證
            Function chkissure(value)
                 if trim(value)="0" then
                    chkissure="正常"
                 elseif  trim(value)="1" then
                    chkissure="未帶"
                 elseif  trim(value)="2" then
                    chkissure="肇事且未帶"
                 elseif  trim(value)="3" then
                    chkissure="逾期且未保"
                 elseif  trim(value)="4" then
                    chkissure="肇事且逾期或未帶"
                 end if     
            end function
           
            '檢查簽收情形
            Function chksigner(value)
                 if trim(value)="0" then
                    chksigner="正常"
                 elseif  trim(value)="1" then
                    chksigner="拒簽"
                 elseif  trim(value)="2" then
                    chksigner="拒收"
                 elseif  trim(value)="3" then
                    chksigner="拒簽拒收"
                 end if     
            end function   
            
            function chkBillType(BillTypeID)
                if trim(BillTypeID) <> "" then
                    Select Case  trim(BillTypeID)
                        Case "1" chkBillType="攔停"
                        Case "2" chkBillType="逕舉"
                        Case "3" chkBillType="逕舉手開單" 
                        Case "4" chkBillType="拖吊" 
                        Case "5" chkBillType="慢車行人"   
                        Case "6" chkBillType="肇事"   
                        Case "7" chkBillType="掌-攔停"   
                        Case "8" chkBillType="掌-行人"   
                        Case "9" chkBillType="掌電拖吊"   
                        Case "H" chkBillType="人工移送"   
                        Case "M" chkBillType="郵寄處理"   
                        Case "N" chkBillType="攔停逕行(未開單)"   
                        Case "D" chkBillType="註銷"   
                        Case "R" chkBillType="單退"   
                        Case "V" chkBillType="掌電拖吊(補開單)"   
                    end select       
                end if 
            end function
            
            '選取單退原因
            Function GetReturnCode(Code)
                 if trim(Code)="1" then
                    GetReturnCode="遷移不明"
                 elseif  trim(Code)="2" then
                    GetReturnCode="查無此人"
                 elseif  trim(Code)="3" then
                    GetReturnCode="地址欠詳"
                 elseif  trim(Code)="4" then
                    GetReturnCode="查無地址"
                 elseif  trim(Code)="5" then
                    GetReturnCode="招領逾期"
                elseif  trim(Code)="6" then
                    GetReturnCode="拒收"
                elseif  trim(Code)="7" then
                    GetReturnCode="投箱待領逾期"
                elseif  trim(Code)="8" then
                    GetReturnCode="其他"
                 end if     
            end function     
            
            '選取單退結果
            Function GetReturnResult(Code)
                 if trim(Code)="S" then
                    GetReturnResult="成功"
                 elseif  trim(Code)="N" then
                    GetReturnResult="找不到資料"
                 elseif  trim(Code)="n" then
                    GetReturnResult="己結案"
                 elseif  trim(Code)="k" then
                    GetReturnResult="已送達不可做未達註記"
                 elseif  trim(Code)="Y" then
                    GetReturnResult="撤銷送達"
                elseif  trim(Code)="h" then
                    GetReturnResult="已開裁決書"
                elseif  trim(Code)="B" then
                    GetReturnResult="無此車號/無此證號"
                elseif  trim(Code)="E" then
                    GetReturnResult="日期錯誤"
                 end if     
            end function       
            
            '查詢車種 
            function GetCarType(CarTypeID)
                if trim(CarTypeID) <> "" then
                    Select Case  trim(CarTypeID)
                        Case "1" GetCarType="自大客車"
                        Case "2" GetCarType="自大貨車"
                        Case "3" GetCarType="自小客(貨)" 
                        Case "4" GetCarType="營大客車" 
                        Case "5" GetCarType="營大貨車"   
                        Case "6" GetCarType="營小貨車"   
                        Case "7" GetCarType="營小客車"   
                        Case "8" GetCarType="租賃小客"   
                        Case "9" GetCarType="遊覽客車"   
                        Case "A" GetCarType="營交通車"   
                        Case "B" GetCarType="貨櫃曳引"   
                        Case "C" GetCarType="自用拖車"   
                        Case "D" GetCarType="營業拖車"   
                        Case "E" GetCarType="外賓小客"   
                        Case "F" GetCarType="外賓大客"   
                        Case "H" GetCarType="普通重機"    
                        Case "L" GetCarType="輕機"   
                        Case "p" GetCarType="併裝車"   
                        Case "x" GetCarType="動力機械"      
                        Case "Y" GetCarType="租賃小貨車"
                        Case "W" GetCarType="自小客"
                        Case "V" GetCarType="自小貨"
                        Case "G" GetCarType="大型重機250CC"
                        Case "Q" GetCarType="大型重機550CC"    
                    end select       
                end if 
            end function
            
            '組地址字串
            function composeAddress(Address,lane,alley,No,Dash)
                composeAddress=""
                composeAddress=Address
                if  trim(lane) <> "" then
                    composeAddress=composeAddress & lane & "巷"
                end if
                if  trim(alley) <> "" then
                    composeAddress=composeAddress & alley & "弄"
                end if
                if  trim(No) <> "" then
                    composeAddress=composeAddress & No & "號"
                end if
                if  trim(Dash) <> "" then
                    composeAddress=composeAddress & "之" & Dash
                end if
            end function
            
            function SetEngDate(tDate)
	            if len(tDate)=7 then
		            SetEngDate=left(tDate,3)&"年"& mid(tDate,4,2)&"月"& Right(tDate,2)&"日"
	            else
		            SetEngDate=""
	            end if
            end function

            function SetchinaDate2(tDate)
	            if len(trim(tDate))=6 then
		            SetchinaDate2=mid(tDate,1,2) &"/"& mid(tDate,3,2)&"/"& mid(tDate,5,2)
	            else
		            SetchinaDate2="0/0/0"
	            end if
            end function  

            function SetchinaDate(tDate)
	            if len(trim(tDate))=6 then
		            SetchinaDate=mid(tDate,1,2) &"年"& mid(tDate,3,2)&"月"& mid(tDate,5,2)&"日"
	            else
		            SetchinaDate=""
	            end if
            end function  

            function GetTime(ttime)
		if trim(ttime) <> "" and not isnull(ttime)  then
                W=""
                H=""
                H=left(ttime,2)
                N=right(ttime,2)
                if cdbl(H)=12 then
                  W="中午"
                elseif cdbl(H)<6  then
                  W="凌晨"
                elseif cdbl(H)>5 and cdbl(H)<12 then
                  W="早上"
                elseif cdbl(H)>12 and cdbl(H)<18 then
                  W="下午"
                elseif cdbl(H)>17 then
                  W="晚上"
		end if

		SH=0

		if H>12 then
			SH=cdbl(H)-12
		else
			SH=H		
		end if
		if len(ttime)=4 then
			GetTime=W&" "&right("00"&SH,2)&"點"&N&"分"
		else
			GetTime=""
		end if
		end if
	end function  
'高市----------------------------------------------------------------------------------------------------------------------------------------------------------------
sql="select * from traffic3.Vil_Rec where tkt_no='"&request("BillNo")&"'"
Set rs=conn.execute(sql)
If Not rs.eof Then 

Sys_BillTypeID         = ReplaceSpace(chkBillType(trim(rs("acc_tp"))))
Sys_BillNo             = ReplaceSpace(rs("tkt_no"))
Sys_CarNo              = ReplaceSpace(rs("plt_no"))


Sys_DriverID           = ReplaceSpace(trim(rs("id_num")))

If Sys_BillTypeID="1" Then 
	Sys_DriverHomeZip      = rs("DRVZIP")
	Sys_Owner      = rs("DRIVER")
	Sys_OwnerZip      = rs("DRVZIP")
	Sys_OwnerAddress       = rs("iAddr")
else
	Sys_OwnerZip           = rs("DRVZIP")
	Sys_Owner              = rs("OWNERX")
	Sys_OwnerAddress       = rs("DRVZIP")
End if

Sys_DealLineDate       = split(GetCDates(rs("DealLineDate")),"/")
Sys_StationID          = trim(rs("ArvADD"))

sql="select ARV_NM,SUPECO from traffic3.arvadd where arvadd='"&Sys_StationID&"'"
Set rstmp=conn.execute(sql)
If Not rstmp.eof Then 
	Sys_STATIONNAME        = rstmp("ARV_NM")
Sys_StationID          = replace(rstmp("SUPECO"),"D","")
End if


If rs("DealLineDate")<>"" Then 
	Sys_MailDate           = ginitdt(rs("DealLineDate"))
End if

BillPageUnit           = "高市警"
'Sys_A_Name             = rs("Brand")
'Sys_CarColor           = rs("ColorC")

sql="select ColorD from traffic3.ColorC where ColorC='"&rs("ColorC")&"'"
Set rstmp=conn.execute(sql)
If Not rstmp.eof Then 
	Sys_CarColor        = rstmp("ColorD")
'	Sys_StationTel         = rstmp("StationTel")  '監理站電話
End if

If Mid(rs("id_num"),1,2)="1" Then 
	Sys_Sex="男"
ElseIf Mid(rs("id_num"),1,2)="2" Then 
	Sys_Sex="女"
End If
'if trim(rs("birthd"))<>"" then 
Sys_DriverBirth        = Split(ReplaceSpace(SetchinaDate2(trim(rs("birthd")))),"/")
'end if

fastring               = ReplaceSpace(mid(rs("hold_c"),1,1) & "  " & SelectFld("traffic3.hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs("hold_c"),1,1))))&ReplaceSpace(mid(rs("hold_c"),2,1) & "  " & SelectFld("traffic3.hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs("hold_c"),2,1))))&ReplaceSpace(mid(rs("hold_c"),3,1) & "  " & SelectFld("traffic3.hold_c","Hold_D","Hold_C=" & QuotedStr(mid(rs("hold_c"),3,1))))
fastring=replace(fastring,"0","")
fastring=replace(fastring,"無","")

Sys_DCIRETURNCARTYPE   = GetData("CAR_NM","traffic3.CAR_TP","CAR_TP",rs("CAR_TP"))

 Address1=""
                    Address2="" 
				    if trim(rs("vilad1")) <> "" then
				        Address1 = SelectFld("traffic3.addr_c","Addr_D"," addr_c= " & QuotedStr(trim(rs("vilad1")))) 
                    end if                     
                   	if trim(rs("vilad2")) <> "" then
                        Address2 = SelectFld("traffic3.addr_c","Addr_D"," addr_c= " & QuotedStr(trim(rs("vilad2")))) 
                    end if  

Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))

Sys_IllegalSpeed       = composeDot(trim(rs("R1_SB2")),Mid(trim(rs("rece01")),3,2))
Sys_RuleSpeed          = composeDot(trim(rs("R1_SB1")),Mid(trim(rs("rece01")),1,2))
Sys_ILLEGALADDRESS     = ReplaceSpace(trim(rs("vilad1")) &"  " & composeAddress(Address1,trim(rs("vil_a1")),trim(rs("vil_b1")),trim(rs("vil_c1")),trim(rs("vil_d1"))))&" "&ReplaceSpace(trim(rs("vilad2")))


sys_Date               = split(GetCDates(rs("BillFillDate")),"/")
Sys_BillFillerMemberID = rs("Police")  '背章號碼

Sum_Level=0
Sys_Level1=""
Sys_Level2=""
Sys_Level3=""
Sys_Level4=""

Sys_Rule1               = ReplaceSpace(composeLaw(trim(rs("Rule_1")),Mid(rs("rece03"),3,1)))


Sys_IllegalRule1  = ReplaceSpace(QueryLawContent(composeLaw(trim(rs("Rule_1")),Mid(rs("rece03"),3,1)),RealSpeed1,LimitSpeed1))

Sys_Level1              = ReplaceSpace(trim(rs("money1")) & " " & "元")

If trim(rs("Rule_2"))<>"" Then 
Sys_Rule2               = ReplaceSpace(composeLaw(trim(rs("Rule_2")),Mid(rs("rece03"),4,1)))
Sys_IllegalRule2        = ReplaceSpace(QueryLawContent(composeLaw(trim(rs("Rule_2")),Mid(rs("rece03"),4,1)),RealSpeed2,LimitSpeed2))
If trim(rs("money2"))<>"0" Then 
Sys_Level2              = ReplaceSpace(trim(rs("money2")) & " " & "元")
end if
End If



End if
'-高市-----------------------------------------------------------------------------------------------


'-高縣-----------------------------------------------------------------------------------------------
sql="select * from traffic4.FMaster where FSEQ='"&request("BillNo")&"'"
Set rs=conn.execute(sql)
If Not rs.eof Then 


Sys_BillTypeID         = rs("AccUSeCode")
Sys_BillNo             = rs("FSEQ")
Sys_CarNo              = rs("CarNo")


Sys_DriverID           = rs("IIDno")

If Sys_BillTypeID="1" Then 
Sys_DriverHomeZip      = rs("iZip")
Sys_Owner      = rs("iName")
Sys_OwnerZip      = rs("iZip")
Sys_OwnerAddress       = rs("iAddr")
else
Sys_OwnerZip           = rs("OwZip")
Sys_Owner              = rs("Owname")
Sys_OwnerAddress       = rs("OwAddr")
End if

Sys_DealLineDate       = split(GetCDates(rs("DealLineDate")),"/")
Sys_StationID          = rs("SPRVSNNo")

sql="select STATIONNAME,StationTel from traffic1.Station where StationID='"&Sys_StationID&"'"
Set rstmp=conn.execute(sql)
If Not rstmp.eof Then 
Sys_STATIONNAME        = rstmp("STATIONNAME")
Sys_StationTel         = rstmp("StationTel")  '監理站電話
End if


If rs("DealLineDate")<>"" Then 
	Sys_MailDate           = ginitdt(rs("DealLineDate"))
End if

BillPageUnit           = "高縣警"
Sys_A_Name             = rs("Brand")
Sys_CarColor           = rs("color")

If Mid(rs("IIDno"),1,2)="1" Then 
Sys_Sex="男"
ElseIf Mid(rs("IIDno"),1,2)="2" Then 
Sys_Sex="女"
End If

'Sys_DriverBirth        = Split(GetCDates(rs("IBirth")),"/")

Sys_DriverBirth        = Split(GetCDates2(rs("IBirth")),"/")


fastring               = GetData("HoldName","traffic4.Hold","HoldCode",rs("HoldCode1"))

Sys_DCIRETURNCARTYPE   = GetData("CDName","traffic4.CarKind","CDType",rs("cdtype"))

Sys_IllegalDate        = split(GetCDates(rs("IllegalDate")),"/")
Sys_IllegalDate_h      = Right("0"&Hour(rs("IllegalDate")),2)
Sys_IllegalDate_m      = Right("0"&minute(rs("IllegalDate")),2)
Sys_IllegalSpeed       = request("IllegalSpeed")
Sys_RuleSpeed          = request("RuleSpeed")
Sys_ILLEGALADDRESS     = rs("irname")


sys_Date               = split(GetCDates(rs("BillFillDate")),"/")
Sys_BillFillerMemberID = rs("PCode1")  '背章號碼

Sum_Level=0
Sys_Level1=0
Sys_Level2=0
Sys_Level3=0
Sys_Level4=0

Sys_Rule1               = rs("RuleF1")
sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs("RULEF1")&"'"
set rsPCode=conn.execute(sql)
if not rspcode.eof then 
Sys_IllegalRule1  = rsPCode("RULENAME")&"&nbsp;"&trim(rs("FACTG1"))
end if
rsPcode.close
set rsPcode=nothing
Sys_Level1              = GetData("RuleAmt","traffic4.RuleF","RuleCode",rs("RuleF1"))

If trim(rs("RuleF2"))<>"" Then 
Sys_Rule2               = rs("RuleF2")
sql="select RULENAME from traffic4.RULEF where RULECODE ='"&rs("RULEF2")&"'"
set rsPCode=conn.execute(sql)
if not rspcode.eof then 
Sys_IllegalRule2  = rsPCode("RULENAME")&"&nbsp;"&trim(rs("FACTG2"))
end if
rsPcode.close
set rsPcode=nothing
Sys_Level2              = GetData("RuleAmt","traffic4.RuleF","RuleCode",rs("RuleF2"))
End If
'response.write Sys_IllegalRule2

If trim(rs("RuleF3"))<>"" Then 
Sys_Rule3               = rs("RuleF3")
Sys_IllegalRule3        = rs("FactG3")
Sys_Level3              = GetData("RuleAmt","traffic4.RuleF","RuleCode",rs("RuleF3"))
End If

If trim(rs("RuleF4"))<>"" Then 
Sys_Rule4               = rs("RuleF4")
Sys_IllegalRule4        = rs("FactG4")
Sys_Level4              = GetData("RuleAmt","traffic4.RuleF","RuleCode",rs("RuleF4"))
End if

End if

Sum_Level=Cint("0"&Replace(Sys_Level1,"元",""))+Cint("0"&Replace(Sys_Level2,"元",""))+Cint("0"&Replace(Sys_Level3,"元",""))+Cint("0"&Replace(Sys_Level4,"元",""))
Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP") 
BillSN=0
Sys_MailNumber=0


if trim(Sys_BillTypeID)="1" then
	DelphiASPObj.GenBillPrintBarCode
	PBillSN(i),Sys_BillNo,Sys_Rule1,Sys_CarNo,Sys_MailNumber,"220073",Sys_DriverHomeZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),Sys_StationID,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,800,263,36
	'response.write "DelphiASPObj.GenBillPrintBarCode"& PBillSN(i)&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,"&Sys_DriverHomeZip&","&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate
	'response.end
else
	DelphiASPObj.GenBillPrintBarCode 20,Sys_BillNo,Sys_Rule1,trim(Sys_CarNo),Sys_MailNumber,"220073","001",right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),20,"台北市交通事件裁決所",0,Sum_Level,0,True,False,Sys_MailDate,800,263,36
'DelphiASPObj.GenBillPrintBarCode 20,Sys_BillNo,Sys_Rule1,Sys_CarNo,"11111","220073",Sys_OwnerZip,right(Sys_DealLineDate(0),2)&Sys_DealLineDate(1)&Sys_DealLineDate(2),"20","台北市交通事件裁決所",0,Sum_Level,0,True,False,"1010101",800,263,36

	DelphiASPObj.GenSendStoreBillno Sys_BillNo,0,28,160,0

'	response.write "DelphiASPObj.GenBillPrintBarCode"& 1&","&Sys_BillNo&","&Sys_Rule1&","&Sys_CarNo&","&Sys_MailNumber&",220073,001,"&Sys_DealLineDate(0)&Sys_DealLineDate(1)&Sys_DealLineDate(2)&","&Sys_StationID&",台北市交通事件裁決所,0,"&Sum_Level&",0,True,False,"&Sys_MailDate&",1"
'	response.end
end if
%>
<!--#include virtual="traffic/Common/checkLaw.asp"-->
<div id="L78" style="position:relative;">

<div id="Layer50" class="style3" style="position:absolute; left:10px; top:<%=0%>px; width:400px; height:10px; z-index:5"><%=Sys_BatchNumber&"　"&SysUnit&"　"%></div>

<%if showBarCode then%>
<div id="Layer1" style="position:absolute; left:10px; top:<%=pagesum+20%>px; width:10px; height:20px; z-index:5">v</div>
<%else%>
<div id="Layer2" style="position:absolute; left:10px; top:<%=pagesum+45%>px; width:10px; height:20px; z-index:5">v</div>
<%end if%>
<%if trim(Sys_BillTypeID)="1" then%>
	<div id="Layer3" style="position:absolute; left:130px; top:<%=pagesum+35%>px; width:202px; height:36px; z-index:5">v</div>
<%else%>
	<div id="Layer4" style="position:absolute; left:130px; top:<%=pagesum+40%>px; width:202px; height:36px; z-index:5">v</div>
<%end if%>

<!--<div id="Layer5" style="position:absolute; left:185px; top:<%=pagesum+45%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%if trim(Sys_INSURANCE)="0" then%>
	<div id="Layer6" style="position:absolute; left:625px; top:<%=pagesum+25%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%elseif trim(Sys_INSURANCE)="1" then%>
	<div id="Layer7" style="position:absolute; left:625px; top:<%=pagesum+35%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%else%>
	<div id="Layer8" style="position:absolute; left:625px; top:<%=pagesum+45%>px; width:202px; height:36px; z-index:5">Ｖ</div>
<%end if%>-->
<div id="Layer9" style="position:absolute; left:10px; top:<%=pagesum+70%>px; width:202px; height:36px; z-index:5"><%
	if showBarCode then
		response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_3.jpg"">"
	else
		response.write SysUnit
	end if
%></div>
<div id="Layer9" class="style7" style="position:absolute; left:250px; top:<%=pagesum+85%>px; width:202px; height:36px; z-index:5"><%=DriverStatus%></div>
<!--<div id="Layer42" style="position:absolute; left:210px; top:<%=pagesum+70%>px; width:202px; height:36px; z-index:5"><%="<font size=1>"&SysUnit&"<br>("&SysUnitTel&")</font>"%></div>-->
<div id="Layer10" style="position:absolute; left:460px; top:<%=pagesum+70%>px; width:233px; height:32px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>
<!--<div id="Layer11" style="position:absolute; left:485px; top:<%=(i*1550+110)%>px; width:230px; height:12px; z-index:7"><font size=1>　<%=BillPageUnit%>交字第<%=Sys_BillNo%>號</font></div>-->
<div id="Layer12" class="style3" style="position:absolute; left:60px; top:<%=pagesum+130%>px; width:150px; height:11px; z-index:3"><%
	response.write "逕行舉發　"&funcCheckFont(Sys_A_Name,16,1)&"<br>"
	'response.write "逕行舉發　<br>"
	if left(trim(Sys_Rule1),2)<>"562" then response.write "<span class=""style2"">依據採證照片</span>"
	response.write "　"&Sys_CarColor
%></div>
<div id="Layer13" class="style3" style="position:absolute; left:220px; top:<%=pagesum+130%>px; width:28px; height:11px; z-index:3"><%=Sys_Sex%></div>
<div id="Layer14" class="style2" style="position:absolute; left:330px; top:<%=pagesum+130%>px; width:500px; height:10px; z-index:4"><%if showBarCode then response.write "<font size=2>*本單可至郵局或期限內至統一、全家、ok、萊爾富等超商繳納</font>"%></div>
<div id="Layer15" class="style3" style="position:absolute; left:210px; top:<%=pagesum+160%>px; width:100px; height:10px; z-index:8"><font size=2><%if trim(Sys_DriverBirth(0))<>"" then response.write Sys_DriverBirth(0)&"年"&right("0"&Sys_DriverBirth(1),2)&"月"&right("0"&Sys_DriverBirth(2),2)&"日"%></font></div>
<div id="Layer16" class="style3" style="position:absolute; left:365px; top:<%=pagesum+160%>px; width:106px; height:13px; z-index:9"><%=Sys_DriverID%></div>
<div id="Layer17" class="style3" style="position:absolute; left:560px; top:<%=pagesum+160%>px; width:99px; height:12px; z-index:10"><%=fastring%></div>
<div id="Layer18" class="style3" style="position:absolute; left:60px; top:<%=pagesum+175%>px; width:100px; height:14px; z-index:11"><B><%=Sys_CarNo%></B></div>
<div id="Layer19" class="style3" style="position:absolute; left:225px; top:<%=pagesum+175%>px; width:117px; height:20px; z-index:12"><%=Sys_DCIRETURNCARTYPE%></div>
<div id="Layer20" class="style3" style="position:absolute; left:395px; top:<%=pagesum+175%>px; width:251px; height:17px; z-index:13"><%=funcCheckFont(Sys_Owner,16,1)%></div>
<div id="Layer21" class="style3" style="position:absolute; left:60px; top:<%=pagesum+200%>px; width:800px; height:13px; z-index:14"><%
	Response.Write Sys_OwnerZip&" "&funcCheckFont(Sys_OwnerZipName&Sys_OwnerAddress,14,1)&chkaddress
	If chkIllegalDate Then Response.Write "　(車主自取)"
	
%></div>

<div id="Layer22" class="style3" style="position:absolute; left:70px; top:<%=pagesum+225%>px; width:40px; height:13px; z-index:15"><%=Sys_IllegalDate(0)%></div>
<div id="Layer23" class="style3" style="position:absolute; left:120px; top:<%=pagesum+225%>px; width:40px; height:17px; z-index:16"><%=Sys_IllegalDate(1)%></div>
<div id="Layer24" class="style3" style="position:absolute; left:170px; top:<%=pagesum+225%>px; width:40px; height:16px; z-index:17"><%=Sys_IllegalDate(2)%></div>
<div id="Layer25" class="style3" style="position:absolute; left:220px; top:<%=pagesum+225%>px; width:40px; height:16px; z-index:18"><%=right("00"&Sys_IllegalDate_h,2)%></div>
<div id="Layer26" class="style3" style="position:absolute; left:270px; top:<%=pagesum+225%>px; width:40px; height:13px; z-index:19"><%=right("00"&Sys_IllegalDate_m,2)%></div>
<div id="Layer27" class="style3" style="position:absolute; left:355px; top:<%=pagesum+220%>px; width:250px; height:31px; z-index:20"><%
	response.write "<font size=3>"
	if left(trim(Sys_Rule1),2)="40" or (int(Sys_Rule1)>4310200 and int(Sys_Rule1)<4310209) or (int(Sys_Rule1)>3310101 and int(Sys_Rule1)<3310111) then
		if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then
			response.write "限速"&Sys_RuleSpeed&"公里、經測速時速"&Sys_IllegalSpeed&"公里、<b>超速"&Sys_IllegalSpeed-Sys_RuleSpeed&"公里</b>"
'			if Sys_IllegalSpeed-Sys_RuleSpeed>=100 then
'				response.write "<br>100以上"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=80 then
'				response.write "<br>80以上未滿100"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=60 then
'				response.write "<br>60以上未滿80"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=40 then
'				response.write "<br>40以上未滿60"
'			elseif Sys_IllegalSpeed-Sys_RuleSpeed>=20 then
'				response.write "<br>20以上未滿40"
'			else
'				response.write "<br>未滿20公里"
'			end if
		end if
	else
		if int(Sys_Rule1)=4340003 then Sys_IllegalRule1=Sys_IllegalRule1&"(吊扣牌照三個月)"
		response.write Sys_IllegalRule1

		'if trim(Sys_IllegalSpeed)<>"" and trim(Sys_RuleSpeed)<>"" then response.write "(限制"&Sys_RuleSpeed&",實際"&Sys_IllegalSpeed&")"
		
	end if
	if trim(Sys_Rule4)<>"" then response.write "("&Sys_Rule4&")"
	response.write "</font>"
	if trim(Sys_Rule2)<>"" then
		'smith edit for print two law 20070621
		if (Sys_Rule2)="4340003" then Sys_IllegalRule2=Sys_IllegalRule2&"(吊扣牌照三個月)"
		response.write "<br>"&Sys_IllegalRule2
	end if
%></div>
<div id="Layer28" class="style3" style="position:absolute; left:60px; top:<%=pagesum+250%>px; width:280px; height:15px; z-index:21"><%=Sys_ILLEGALADDRESS%></div>
<div id="Layer29" class="style3" style="position:absolute; left:100px; top:<%=pagesum+290%>px; width:34px; height:11px; z-index:22"><%=Sys_DealLineDate(0)%></div>
<div id="Layer30" class="style3" style="position:absolute; left:180px; top:<%=pagesum+290%>px; width:35px; height:13px; z-index:23"><%=Sys_DealLineDate(1)%></div>
<div id="Layer31" class="style3" style="position:absolute; left:260px; top:<%=pagesum+290%>px; width:32px; height:15px; z-index:24"><%=Sys_DealLineDate(2)%></div>
<div id="Layer32" class="style3" style="position:absolute; left:370px; top:<%=pagesum+295%>px; width:400px; height:49px; z-index:29"><%
	response.write "<font size='2'>"&left(trim(Sys_Rule1),2)&"　　"
	if len(trim(Sys_Rule1))>7 then response.write "　"&right(trim(Sys_Rule1),1)
	response.write Mid(trim(Sys_Rule1),3,1)&"　　"&Mid(trim(Sys_Rule1),4,2)
	response.write "　　　　　　　　　　　　　　"&Sys_Level1
	if trim(Sys_Rule2)<>"0" then
		response.write "<br>"&left(trim(Sys_Rule2),2)&"　　"
		if len(trim(Sys_Rule2))>7 then response.write "　"&right(trim(Sys_Rule2),1)
		response.write Mid(trim(Sys_Rule2),3,1)&"　　"&Mid(trim(Sys_Rule2),4,2)
		response.write "　　　　　　　　　　　　　　"&Sys_Level2
	end if
	response.write "</font>"
%></div>

<div id="Layer34" class="style3" style="position:absolute; left:355px; top:<%=pagesum+335%>px; width:95px; height:30px; z-index:28"><font size=2><%=Sys_STATIONNAME&"<br>"&Sys_StationTel%></font></div>
<div id="Layer33" style="position:absolute; left:445px; top:<%=pagesum+330%>px; width:400px; height:30px; z-index:28"><%if showBarCode then response.write "<img src=""../BarCodeImage/"&Sys_BillNo&"_5.jpg"">"%></div>

<div id="Layer35" class="style3" style="position:absolute; left:300px; top:<%=pagesum+445%>px; width:200px; height:49px; z-index:29"><%
	If not ifnull(Request("Sys_BillPrintUnitTel")) Then
		Response.Write "<br>TEL："&Request("Sys_BillPrintUnitTel")
	end if
%></div>
<!--<div id="Layer36" style="position:absolute; left:580px; top:<%=pagesum+420%>px; width:100px; height:43px; z-index:30">主管</div>-->
<div id="Layer37" class="style3" style="position:absolute; left:625px; top:<%=pagesum+450%>px; width:200px; height:46px; z-index:31"><%=Sys_ChName%></div>
<div id="Layer38" class="style3" style="position:absolute; left:230px; top:<%=pagesum+480%>px; width:60px; height:10px; z-index:32"><%=sys_Date(0)%></div>
<div id="Layer39" class="style3" style="position:absolute; left:390px; top:<%=pagesum+480%>px; width:60px; height:13px; z-index:33"><%=sys_Date(1)%></div>
<div id="Layer40" class="style3" style="position:absolute; left:540px; top:<%=pagesum+480%>px; width:60px; height:11px; z-index:34"><%=sys_Date(2)%></div>

<div style="position:absolute; left:20px; top:<%=pagesum+600%>px;">
<table width="645" border="0">
  <tr>
	<th align="left">&nbsp;</th>
    <th align="left" valign="top"><span class="style4"><%=SysUnit&replace(SysUnitLevel3,SysUnit,"")%><br>　<%=SysAddressLevel3%></span></th>
    <td align="left" height="130" valign="top"></td>
  </tr>
  <tr>
	<td align="left" valign="top"></td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th align="left" colspan="2"><span class="style4"><%=chstr(Sys_Driver)%>　台啟</span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left"><span class="style4"><%=funcCheckFont(Sys_Owner,20,1)%>　台啟</span></th>
	<%end if%>
	<td align="left" valign="top"></td>
  </tr>
  <tr>
	<td>&nbsp;</td>
	<%if trim(Sys_BillTypeID)="1" then%>
    <th colspan="2" align="left" valign="top" nowrap><span class="style4"><%=Sys_DriverHomeZip%><br>
    <%=Sys_DriverZipName&Sys_DriverHomeAddress%></span></th>
	<%elseif trim(Sys_BillTypeID)="2" then%>
	<th align="left" valign="top" width="400"><span class="style4"><%=Sys_OwnerZip%><br>
    <%=Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,25,1)&chkaddress%></span></th>
	<%end if%>
    <td>&nbsp;</td>
  </tr>
</table>
</div>

<%If instr("BB,BC,BD",left(Sys_BillNo,2))>0 and Sys_MailNumber>0 Then%>

<div id="Layer61" class="style3" style="position:absolute; left:500px; top:<%=pagesum+850%>px; width:200px; height:32px; z-index:6">
大宗郵資已付掛號函件
</div>

<div id="Layer62" class="style3" style="position:absolute; left:540px; top:<%=pagesum+870%>px; width:200px; height:32px; z-index:6">
第<%=Sys_MailNumber%>號
</div>

<div id="Layer63" style="position:absolute; left:500px; top:<%=pagesum+890%>px; width:200px; height:32px; z-index:6">
<img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%>>
</div>

<div id="Layer64" class="style3" style="position:absolute; left:500px; top:<%=pagesum+940%>px; width:200px; height:32px; z-index:6">
<%=Sys_MAILCHKNUMBER%>
</div>
<%end if%>


<div id="Layer49" style="position:absolute; left:500px; top:<%=pagesum+660%>px; width:200px; height:32px; z-index:6"><%If chkIllegalDate Then Response.Write "<br>　　　　(車主自取)"%><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_4.jpg"""%>></div>

<div id="Layer49" class="style3" style="position:absolute; left:10px; top:<%=pagesum+1020%>px; height:32px; z-index:6"><%
If not ifnull(Request("Sys_BillPrintUnitTel")) Then
	Response.Write "申訴服務電話："&Request("Sys_BillPrintUnitTel")
	Response.Write "　上班受理時間：週一至週五&nbsp;上午8:00~12:00&nbsp;下午13:30~17:30"
end if
%>
</div>

<div id="Layer43" class="style7" style="position:absolute; left:290px; top:<%=pagesum+1080%>px; width:300px; height:12px; z-index:36"><%=Sys_RedUnitName%></div>
<div id="Layer44" class="style3" style="position:absolute; left:250px; top:<%=pagesum+1110%>px; width:800px; height:12px; z-index:36"><%
	if Sys_BillTypeID="1" then
		response.write "<font size=2>"&chstr(Sys_Driver)&"</font>"
		response.write "<font size=2>　　"&Sys_DriverHomeZip&"&nbsp;&nbsp;"&Sys_DriverZipName&Sys_DriverHomeAddress&"</font>"
	else
		response.write "<font size=2>"&funcCheckFont(Sys_Owner,10,1)&"</font>"
		response.write "<font size=2>　　"&Sys_OwnerZipDeliver&"&nbsp;&nbsp;"&Sys_OwnerZipNameDeliver&funcCheckFont(Sys_OwnerAddressDeliver,10,1)&chkaddress&"</font>"
	end if%></div>
<div id="Layer45" class="style3" style="position:absolute; left:370px; top:<%=pagesum+1130%>px; width:280px; height:12px; z-index:36"><%=Sys_BillNo%></div>

<div id="Layer46" style="position:absolute; left:450px; top:<%=pagesum+1130%>px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&".jpg"""%>></div>

<%If instr("BB,BC,BD",left(Sys_BillNo,2))>0 and Sys_MailNumber>0 Then%>

<div id="Layer46" style="position:absolute; left:440px; top:<%=pagesum+1550%>px; z-index:6"><img src=<%="""../BarCodeImage/"&Sys_BillNo&"_2.jpg"""%> width="140" height="30"></div>

<div id="Layer47" class="style3" style="position:absolute; left:580px; top:<%=pagesum+1555%>px; width:500px; z-index:36">第<%=Sys_MailNumber%>號</div> 
<%end if%>

<div id="Layer47" class="style3" style="position:absolute; left:70px; top:<%=pagesum+1555%>px; width:500px; z-index:36"><%=SysAddress%></div> 
</div>
<%
If not ifnull(errBillNo) Then errBillNo="下列車主姓名不足三個字\n"&errBillNo%>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();<%
	If Not ifnull(errBillNo) Then%>
		alert("<%=errBillNo%>");<%
	end if%>
	printWindow(true,0,5.08,0,5.08);
</script>