
alter database default tablespace USERS;
alter database default temporary tablespace Temp;

drop tablespace Takecar INCLUDING CONTENTS;
drop tablespace TakecarLog INCLUDING CONTENTS;

create tablespace Takecar datafile 'D:\oracleData\TakecarDB.DBF' size 1000M REUSE AUTOEXTEND ON NEXT 100M LOGGING EXTENT MANAGEMENT LOCAL SEGMENT SPACE MANAGEMENT AUTO ;
ALTER DATABASE DEFAULT TABLESPACE Takecar;
                                                                                                
CREATE TEMPORARY TABLESPACE TakecarLOG
TEMPFILE 'D:\oracleData\TakecarLog.DBF' SIZE 200M REUSE
AUTOEXTEND ON NEXT 50M MAXSIZE unlimited
EXTENT MANAGEMENT LOCAL;

ALTER DATABASE DEFAULT TEMPORARY TABLESPACE TakecarLOG;
