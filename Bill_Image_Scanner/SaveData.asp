<!--#include virtual="traffic/Common/DB.ini"-->
<%
                              	'SN§ì³Ì¤j­È
                              	sSQL = "select BillAttatchImage_seq.nextval as SN from Dual"
                              	set oRST = Conn.execute(sSQL)
                              	
                              	if not oRST.EOF then
                        	       	sMaxSN = oRST("SN")
                              	end if
                              	oRST.close
                              	
                                   '/img/scan/memberid/xxxx.jpg
                                FileDirAndName="/img/scan/" & request("theRecordMemberID") & "/YHandle/" & Request("ImageNum") & ".JPG"
                                
                                if request("RuleDateS")="" and request("RuleDateE")="" and request("RuleDateE")="" and request("IlgalAddr")="" and request("IlgalAddrID")="" then 
                                  strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID)" & _
                    				" values("&sMaxSN&",'"&FileDirAndName&"','','"& request("TypeID") &"','"& request("theRecordMemberID") &"',SYSDATE,0)"
                       			else
                                	strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID,RuleDateS,RuleDateE,IllegalAddress,IllegalAddressID)" & _
                    				" values("&sMaxSN&",'"&FileDirAndName&"','','"& request("TypeID") &"','"& request("theRecordMemberID") &"',SYSDATE,0,to_date('"& request("RuleDateS") &" 00:00:00','yyyy/MM/dd HH24:MI:SS') ,to_date('"& request("RuleDateE") &" 23:59:59','yyyy/MM/dd HH24:MI:SS'),'"& request("IlgalAddr") &"','"& request("IlgalAddrID") &"')"

                       			end if	
                            	conn.execute strInsert


%>