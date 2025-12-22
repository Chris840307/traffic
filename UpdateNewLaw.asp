<html>
<!--#include virtual="traffic/Common/DB.ini"-->

<body>
<%
Server.ScriptTimeout = 65000

'1120828=====================================================================
	strChkL2="select * from Law where itemid ='530000' and recordstateid=-1 and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="update law set Recordstateid=-1 where itemid in ('43000','530000')"
		conn.execute strInsL2
		
		strInsL2="update law set illegalrule='初次領有駕駛執照未滿一年之駕駛人，依本條例第44條第1項第2款規定應記違規點數者，加記違規點數1點' where itemid ='44000' and version='2'"
		conn.execute strInsL2

		strInsL2="update law set illegalrule='初次領有駕駛執照未滿一年之駕駛人，依本條例第53條第2項規定應記違規點數者，加記違規點數1點' where itemid ='53000' and version='2'"
		conn.execute strInsL2

		strInsL2="update law set illegalrule='初次領有駕駛執照未滿一年之駕駛人，依本條例第92條第7項規定應記違規點數者，加記違規點數1點' where itemid ='92000' and version='2'"
		conn.execute strInsL2

		strInsL2="update law set Specpunish=6 where itemid ='3210002' and version='2'"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing


'1121025=====================================================================
	strChkL2="select * from Law where itemid ='4800101' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="insert into law values('4800101','0','轉彎未注意來往行人',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800102','0','一般道路變換車道未注意來往行人',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800103','0','轉彎前未減速慢行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800104','0','載運危險物品車輛轉彎未注意來往行人',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800105','0','載運危險物品車輛一般道路變換車道未注意來往行人',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800106','0','載運危險物品車輛轉彎前未減速慢行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800201','0','轉彎或變換車道不依標誌、標線、號誌指示',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800301','0','行經交岔路口未達中心處，佔用來車道搶先左轉',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800401','0','在多車道右轉彎，不先駛入外側車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800402','0','在多車道左轉彎，不先駛入內側車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800501','0','道路設有劃分島，劃分快慢車道，在慢車道上左轉彎',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800502','0','道路設有劃分島，劃分快慢車道，在快車道上右轉彎',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4800601','3','轉彎車不讓直行車先行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800601','5','轉彎車不讓直行車先行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800601','6','轉彎車不讓直行車先行',1400,1500,1600,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800602','0','載運危險物品車輛轉彎車不讓直行車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800701','0','直行車佔用最內側轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800702','0','直行車佔用最外側轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4800703','0','直行車佔用轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="update law set illegalrule='行經設有險坡標誌之路段超車' where itemid ='4710102' and version='2'"
		conn.execute strInsL2
		
		strInsL2="update law set illegalrule='載運危險物品車輛行經設有險坡標誌之路段超車' where itemid ='4710108' and version='2'"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
%>
跑完了

insert into law values('4800101','0','轉彎未注意來往行人',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800102','0','一般道路變換車道未注意來往行人',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800103','0','轉彎前未減速慢行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800104','0','載運危險物品車輛轉彎未注意來往行人',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800105','0','載運危險物品車輛一般道路變換車道未注意來往行人',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800106','0','載運危險物品車輛轉彎前未減速慢行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800201','0','轉彎或變換車道不依標誌、標線、號誌指示',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800301','0','行經交岔路口未達中心處，佔用來車道搶先左轉',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800401','0','在多車道右轉彎，不先駛入外側車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800501','0','道路設有劃分島，劃分快慢車道，在慢車道上左轉彎',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800502','0','道路設有劃分島，劃分快慢車道，在快車道上右轉彎',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800601','3','轉彎車不讓直行車先行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800601','5','轉彎車不讓直行車先行',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800601','6','轉彎車不讓直行車先行',1400,1500,1600,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800602','0','載運危險物品車輛轉彎車不讓直行車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800701','0','直行車佔用最內側轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800702','0','直行車佔用最外側轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
insert into law values('4800703','0','直行車佔用轉彎專用車道',600,700,800,900,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)
go
update law set illegalrule='行經設有險坡標誌之路段超車' where itemid ='4710102' and version='2'
go
update law set illegalrule='載運危險物品車輛行經設有險坡標誌之路段超車' where itemid ='4710108' and version='2'

</form>
</body>
<%
conn.close
Set conn=nothing
%>
</html>