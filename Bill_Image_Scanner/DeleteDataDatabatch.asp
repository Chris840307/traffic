<!--#include virtual="traffic/Common/DB.ini"-->
<%
                                
                              	strDelete="Update BillAttatchImage set RecordStateID=1 where SN in (select a.SN from BillAttatchImage a,MemberData b,UnitInfo c where a.RecordMemberID=b.MemberID and b.unitid=c.unitid and a.RecordStateID=0 "&request("strwhere") &")"
								response.write strdelete
                            	conn.execute strDelete


%>
