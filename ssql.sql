1=select sysdate as result from dual
1=select createdate as result from dcicreatefilelog where filename like '%.big' and createdate between add_months(sysdate,-1) and sysdate order by createdate desc
1=select unitname as result from unitinfo where DCIUnitID is null or DCIUnitID=''
6=select count(billno) as result from billbase where recordstateid <> -1
6=select count(billno) as result from billbase where (RecordDate >to_date('2009/01/01', 'yyyy/mm/dd') and RecordDate< to_date('2009/12/31', 'yyyy/mm/dd')) and recordstateid <> -1
6=select count(distinct(actionchname)) as result from Log where typeid=350 and (actiondate between add_months(sysdate,-1) and sysdate ) order by actionchname
6=select ceil(PERCENT_SPACE_USED) as result from v$flash_recovery_area_usage where lower(File_Type)='archivelog'