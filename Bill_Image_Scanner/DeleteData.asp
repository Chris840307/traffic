<!--#include virtual="traffic/Common/DB.ini"-->
<%
                                
                              	strDelete="Update BillAttatchImage set RecordStateID=1 where SN=" & request("SN") 
                            	conn.execute strDelete


%>
