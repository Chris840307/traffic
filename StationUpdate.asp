<!--#include virtual="traffic/Common/DB.ini"-->
<%
sql="update station set stationname ='臺北市區監理所' where dcistationid='20'"
conn.execute(sql)

sql="update station set stationname ='士林監理站' where dcistationid='21'"
conn.execute(sql)

sql="update station set stationname ='連江監理站' where dcistationid='28'"
conn.execute(sql)

sql="update station set dcistationname ='連江監理站' where dcistationid='37'"
conn.execute(sql)

sql="update station set stationname ='高雄市區監理所' where dcistationid='30'"
conn.execute(sql)

sql="update station set stationname ='苓雅監理站' where dcistationid='31'"
conn.execute(sql)

sql="update station set dcistationname ='金門監理站',stationname ='金門監理站' where dcistationid='36'"
conn.execute(sql)
response.write "OK"
conn.close
set conn=nothing
%>