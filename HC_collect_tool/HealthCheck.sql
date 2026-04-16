set echo off linesize 5000 verify off feedback off trimspool on arraysize 100 null '' wrap on timing off
set serveroutput on size 300000 format trunc pagesize 50000 embedded off recsep off tab off
CLEAR BREAKS COMPUTE

ALTER SESSION SET NLS_DATE_FORMAT = 'DD.MM.YYYY';
DEF long_date_format = 'DD.MM.YYYY-HH24:MI:SS'
DEF filter_user1="'APEX_030200','APEX_PUBLIC_USER','APPQOSSYS','CTXSYS','DBSNMP','DIP','EXFSYS','FLOWS_FILES','MDDATA','MDSYS','MGMT_VIEW','OLAPSYS','ORACLE_OCM','ORDDATA','ORDPLUGINS'"
DEF filter_user2="'ORDSYS','OUTLN','OWBSYS','OWBSYS_AUDIT','PERFSTAT','RMAN','SCOTT','SI_INFORMTN_SCHEMA','SPATIAL_CSW_ADMIN_USR','SPATIAL_WFS_ADMIN_USR','SYS','SYSMAN','SYSTEM','WMSYS','XDB'"

set flush off term off

COL dbname NEW_VALUE V_dbname
COL day    NEW_VALUE V_day
SELECT LOWER(name) dbname,TO_CHAR(sysdate,'YYYY_MM_DD') day FROM v$database;
def output_file="check_oracle_&V_day._&V_dbname..log"
COL dbname CLEAR
COL day    CLEAR
undef V_dbname V_day

set term on
PROMPT
PROMPT Generating output file: &output_file....
spool &output_file
set term off

EXEC DBMS_OUTPUT.put_line('Generated at: ' || TO_CHAR(sysdate,'&long_date_format'));

--==========================================================================================================
prompt
prompt + Database/Host info
prompt

declare
    procedure p(P_name varchar2,P_val varchar2) is
    begin
        if P_val is not null then
            DBMS_OUTPUT.put_line(rpad(P_name||' ',25,'.')||' '||P_val);
        end if;
    end p;
begin
    for r in (
        select d.name,i.instance_name,i.version,d.created,d.log_mode,i.startup_time,
               substr(numtodsinterval(sysdate-i.startup_time,'day'),9,11) uptime,
               i.instance_role,i.parallel,i.instance_number,i.host_name,d.platform_name
        from   gv$database d,gv$instance i
		where  i.instance_number=1 and d.inst_id=1
    ) loop
        p('DB Name',r.name);
        p('Version',r.version);
        p('Parallel (RAC)',r.parallel);
        p('Created',r.created);
        p('Startup Time',to_char(r.startup_time,'&&long_date_format')||', Uptime '||
          case when extract(day from numtodsinterval(sysdate-r.startup_time,'DAY'))>0 then
            extract(day from numtodsinterval(sysdate-r.startup_time,'DAY'))||'-'
          end||substr(numtodsinterval(sysdate-r.startup_time,'DAY'),12,5)
        );
        p('Log Mode',r.log_mode);
		DBMS_OUTPUT.put_line(' ');
        p('Instance Name',r.instance_name);
        p('Instance Number',r.instance_number);
        p('Instance Role',r.instance_role);
		p('Host Name',r.host_name);
        p('Platform',r.platform_name);
    end loop;
	for r in (
        select
         (select value from v$osstat where stat_name='PHYSICAL_MEMORY_BYTES') phy_mem,
         (select value from v$osstat where stat_name='NUM_CPUS')              num_cpus,
         (select value from v$osstat where stat_name='NUM_CPU_SOCKETS')       num_sockets,
         (select value from v$osstat where stat_name='NUM_CPU_CORES')         num_cores
        from dual
    ) loop
        p('Physical Memory(Mb)',to_char(r.phy_mem/1048576,'fm999,999,999'));
        p('Number of CPUs',r.num_cpus);
        p('Number of CPU Sockets',r.num_sockets);
        p('Number of CPU Cores',r.num_cores);
    end loop;
	for r in (
        select d.name,i.instance_name,i.version,d.created,d.log_mode,i.startup_time,
               substr(numtodsinterval(sysdate-i.startup_time,'day'),9,11) uptime,
               i.instance_role,i.parallel,i.instance_number,i.host_name,d.platform_name
        from   gv$database d,gv$instance i
		where  i.instance_number=2 and d.inst_id=2
    ) loop
		DBMS_OUTPUT.put_line(' ');
        p('Instance Name',r.instance_name);
        p('Instance Number',r.instance_number);
        p('Instance Role',r.instance_role);
        p('Host Name',r.host_name);
        p('Platform',r.platform_name);
    end loop;
	for r in (
        select
         (select value from v$osstat where stat_name='PHYSICAL_MEMORY_BYTES') phy_mem,
         (select value from v$osstat where stat_name='NUM_CPUS')              num_cpus,
         (select value from v$osstat where stat_name='NUM_CPU_SOCKETS')       num_sockets,
         (select value from v$osstat where stat_name='NUM_CPU_CORES')         num_cores
        from dual
    ) loop
        p('Physical Memory(Mb)',to_char(r.phy_mem/1048576,'fm999,999,999'));
        p('Number of CPUs',r.num_cpus);
        p('Number of CPU Sockets',r.num_sockets);
        p('Number of CPU Cores',r.num_cores);
    end loop;
end;
/

--=================iiiiiiii=========================================================================================
PROMPT
PROMPT
PROMPT DATABASE STATUS
PROMPT =================

--select INSTANCE_NAME,STATUS,DATABASE_STATUS,ACTIVE_STATE,STARTUP_TIME from v$instance;
select INSTANCE_NAME,ARCHIVER,THREAD#,STATUS,DATABASE_STATUS,ACTIVE_STATE,STARTUP_TIME,VERSION,LOG_SWITCH_WAIT from gv$instance;

PROMPT
PROMPT DATABASE NAME AND MODE
PROMPT ========================

select name, database_role, open_mode, log_mode,flashback_on flashback, protection_mode,protection_level from v$database;

--select name,value from gv$parameter where name in ('db_name','db_unique_name','fal_server','fal_client','remote_login_passwordfile','standby_file_management');
PROMPT
PROMPT ASM INFOR
PROMPT ===============================================================
COLUMN group_name             FORMAT a20           HEAD 'Disk Group|Name'
COLUMN sector_size            FORMAT 99,999        HEAD 'Sector|Size'
COLUMN block_size             FORMAT 99,999        HEAD 'Block|Size'
COLUMN allocation_unit_size   FORMAT 999,999,999   HEAD 'Allocation|Unit Size'
COLUMN state                  FORMAT a11           HEAD 'State'
COLUMN type                   FORMAT a6            HEAD 'Type'
COLUMN total_mb               FORMAT 999,999,999   HEAD 'Total Size (MB)'
COLUMN used_mb                FORMAT 999,999,999   HEAD 'Used Size (MB)'
COLUMN pct_used               FORMAT 999.99        HEAD 'Pct. Used'



break on report on disk_group_name skip 1
compute sum label "Total: " of total_mb used_mb on report


SELECT
    name                                     group_name
  , sector_size                              sector_size
  , block_size                               block_size
  , allocation_unit_size                     allocation_unit_size
  , state                                    state
  , type                                     type
  , total_mb                                 total_mb
  , (total_mb - free_mb)                     used_mb
  , ROUND((1- (free_mb / total_mb))*100, 2)  pct_used
FROM
    v$asm_diskgroup
ORDER BY
    name;
PROMPT
PROMPT ASM Other
PROMPT ================================================================
col bytes format 999,999,999,999
col space format 999,999,999,999
col gn format 999
col name format a20
col au format 99999999
col state format a12
col type format a12
col total_mb format 999,999,999
col free_mb format 999,999,999
col od format 999
col compatibility format a12
col dn format 999
col mount_status format a12
col header_status format a12
col mode_status format a12
col mode format a12
col failgroup format a20
col label format a12
col path format a45
select group_number gn,disk_number dn, mount_status, header_status,mode_status,state, total_mb, free_mb,name, failgroup, label, path,create_date, mount_date
from v$asm_disk order by group_number, disk_number;

PROMPT
PROMPT
PROMPT DR DATABASE STATUS
PROMPT ==================
select to_char(sysdate,'mm-dd-yyyy hh24:mi:ss') "Current Time" from dual;
SELECT DB_NAME,  APPLIED_TIME, LOG_ARCHIVED-LOG_APPLIED LOG_GAP ,
(case when ((APPLIED_TIME is not null and (LOG_ARCHIVED-LOG_APPLIED) is null) or
            (APPLIED_TIME is null and (LOG_ARCHIVED-LOG_APPLIED) is not null) or
            ((LOG_ARCHIVED-LOG_APPLIED) > 5))
      then 'Error! Log Gap is '
      else 'OK!'
 end) Status
FROM
(
SELECT INSTANCE_NAME DB_NAME
FROM GV$INSTANCE
where INST_ID = 1
),
(
SELECT MAX(SEQUENCE#) LOG_ARCHIVED
FROM V$ARCHIVED_LOG WHERE DEST_ID=1 AND ARCHIVED='YES' and THREAD#=1
),
(
SELECT MAX(SEQUENCE#) LOG_APPLIED
FROM V$ARCHIVED_LOG WHERE DEST_ID=2 AND APPLIED='YES' and THREAD#=1
),
(
SELECT TO_CHAR(MAX(COMPLETION_TIME),'DD-MON/HH24:MI') APPLIED_TIME
FROM V$ARCHIVED_LOG WHERE DEST_ID=2 AND APPLIED='YES' and THREAD#=1
)
UNION
SELECT DB_NAME,  APPLIED_TIME, LOG_ARCHIVED-LOG_APPLIED LOG_GAP,
(case when ((APPLIED_TIME is not null and (LOG_ARCHIVED-LOG_APPLIED) is null) or
            (APPLIED_TIME is null and (LOG_ARCHIVED-LOG_APPLIED) is not null) or
            ((LOG_ARCHIVED-LOG_APPLIED) > 5))
      then 'Error! Log Gap is '
      else 'OK!'
 end) Status
from (
SELECT INSTANCE_NAME DB_NAME
FROM GV$INSTANCE
where INST_ID = 2
),
(
SELECT MAX(SEQUENCE#) LOG_ARCHIVED
FROM V$ARCHIVED_LOG WHERE DEST_ID=1 AND ARCHIVED='YES' and THREAD#=2
),
(
SELECT MAX(SEQUENCE#) LOG_APPLIED
FROM V$ARCHIVED_LOG WHERE DEST_ID=2 AND APPLIED='YES' and THREAD#=2
),
(
SELECT TO_CHAR(MAX(COMPLETION_TIME),'DD-MON/HH24:MI') APPLIED_TIME
FROM V$ARCHIVED_LOG WHERE DEST_ID=2 AND APPLIED='YES' and THREAD#=2
)
;

PROMPT
PROMPT
PROMPT ARCHIVE LOG GENERATION LAST WEEK
PROMPT ===========================================================
SELECT A.*,  ROUND (A.Count# * B.AVG# / 1024 / 1024) Daily_Avg_Mb
FROM
  (SELECT TO_CHAR (First_Time,'YYYY-MM-DD') DAY,COUNT (1) Count#, MIN (RECID) Min#,MAX (RECID) Max#,MIN(sequence#),MAX(sequence#)
   FROM v$log_history  GROUP BY TO_CHAR (First_Time, 'YYYY-MM-DD')  ORDER BY 1 DESC ) A,
  (SELECT AVG (BYTES) AVG#,COUNT (1) Count#,MAX (BYTES) Max_Bytes,MIN (BYTES) Min_Bytes
   FROM v$log ) B   where rownum < 8;

PROMPT
PROMPT
PROMPT DB PHYSICAL SIZE
PROMPT =====================================
select sum(bytes/1024/1024/1024) "DB Physical Size(GB)" from dba_data_files;


PROMPT
PROMPT
PROMPT DB ACTUAL SIZE
PROMPT =====================================
select sum(bytes/1024/1024/1024) "DB Actual Size(GB)" from dba_segments;

--==========================================================================================================
PROMPT
PROMPT + CPU stat

var inst_num number
var bsnap_id number

begin
    select instance_number into :inst_num from v$instance;
    select min(snap_id) into :bsnap_id
    from   dba_hist_snapshot
    where  begin_interval_time>=sysdate-numtodsinterval(51,'hour')
    and    instance_number=:inst_num;
end;
/

col cpu_graph format a50 trunc
col inst      format 999
col load      format 9999
break on begin_ts

select instance_number inst,load,
       to_char(p_usr,'990D0') p_usr,to_char(p_sys,'990D0') p_sys,to_char(p_iow,'990D0') p_iow,
       rpad('U',p_usr/2,'U')||rpad('s',p_sys/2,'s')||rpad('w',p_iow/2,'w') cpu_graph
from
(
  select bt,instance_number,flag_restart,load,
         sys_time/nullif(tot_cpu,0)*100 p_sys,
         usr_time/nullif(tot_cpu,0)*100 p_usr,
         iow_time/nullif(tot_cpu,0)*100 p_iow
  from
  (
    select s.begin_interval_time bt,s.instance_number,sys_time,usr_time,iow_time,idl_time,
           sys_time+usr_time+iow_time+idl_time tot_cpu,load,
           case when s.startup_time between s.begin_interval_time and s.end_interval_time then '*' end flag_restart
    from
    (
      select snap_id,dbid,instance_number,
             sum(decode(stat_name,'LOAD',       round(value))) load,
             sum(decode(stat_name,'SYS_TIME',   value-l_value)) sys_time,
             sum(decode(stat_name,'USER_TIME',  value-l_value)) usr_time,
             sum(decode(stat_name,'IOWAIT_TIME',value-l_value)) iow_time,
             sum(decode(stat_name,'IDLE_TIME',  value-l_value)) idl_time
      from
      (
          select snap_id,dbid,instance_number,stat_name,value,
                 lag(value) over(partition by stat_name,instance_number order by snap_id) l_value
          from   dba_hist_osstat
          where  snap_id>=:bsnap_id
          and    stat_name in ('SYS_TIME','USER_TIME','IOWAIT_TIME','IDLE_TIME','LOAD')
          and    dbid=(select dbid from v$database)
      )
      where  l_value is not null
      group  by snap_id,dbid,instance_number
    ) os,dba_hist_snapshot s
    where os.snap_id=s.snap_id and os.dbid=s.dbid and os.instance_number=s.instance_number
  )
)
order  by trunc(bt,'mi'),instance_number;

col cpu_graph clear
col inst      clear
col load      clear
clear breaks

--==========================================================================================================
PROMPT
PROMPT + System performance

set pagesize 100

col "%CPU"         format 999
col "%DB CPU"      format 990
col "Trans/s"      format 99G990
col "%RLB"         format 990
col "IO MB/s"      format 990D0
col "RedoMB/s"     format 990D0
col "WMB/s"        format 990D0
col "RMB/s"        format 990D0
col "BlkR Lat"     format 9G990D0
col "DBWait%"      format 990
col "Act.Sess"     format 999G990
col "Logon/s"      format 990D0
col "CurLogons"    format 999G999

-- assume 8KB block for physical read/writes

select 
       avg(cpu)            "%CPU",
       avg(decode(host_cpu,0,0,cpu*db_cpu/host_cpu)) "%DB CPU",
       avg(trans)          "Trans/s",
       avg(rollb_pct)      "%RLB",
       avg(io_mb_s)        "IO MB/s",
       avg(redo)/1048576   "RedoMB/s",
       avg(pwrite)*8/1024  "WMB/s",
       avg(pread)*8/1024   "RMB/s",
       avg(block_read_lat) "BlkR Lat",
       avg(dbwait_ratio)   "DBWait%",
       avg(avg_act_sess)   "Act.Sess",
       avg(logon)          "Logon/s",
       avg(logons)         "CurLogons"
from
(
  select begin_time ts1,metric_id,value v
  from   gv$sysmetric_history
  where  intsize_csec>5000
  and    metric_id in (2057,2003,2004,2006,2016,2147,2018,2144,2103,2107,2145,2155,2075,2025)
)
pivot
(
  min(v) for metric_id in
  (
    2057 cpu,2003 trans,2004 pread,2006 pwrite,
    2016 redo,2147 avg_act_sess,2018 logon,
    2144 block_read_lat,2103 logons,2107 dbwait_ratio,2145 io_mb_s,
    2155 host_cpu,2075 db_cpu,2025 rollb_pct
  )
)
group by rollup(ts1)
order by ts1;

col "%CPU"         clear
col "%DB CPU"      clear
col "Trans/s"      clear
col "%RLB"         clear
col "IO MB/s"      clear
col "RedoMB/s"     clear
col "WMB/s"        clear
col "RMB/s"        clear
col "BlkR Lat"     clear
col "DBWait%"      clear
col "Act.Sess"     clear
col "Logon/s"      clear
col "CurLogons"    clear

--==========================================================================================================
PROMPT
PROMPT + SGA INSTANCE 1

break on pool skip 1 on sumpool
col pool    format a12
col name    format a26
col mb      format 999,999
col sumpool format 999,999

select case when pool is null then '<MISC>' else pool end pool,name,mb,sumpool
from
(
    select pool,name,round(bytes/1048576) mb, round(sum(bytes) over(partition by pool)/1048576) sumpool,
           dense_rank() over(partition by pool order by bytes desc) r
    from   gv$sgastat
	where INST_ID=1
) where r<=10
order by pool,mb desc;

clear breaks
col pool    clear
col name    clear
col mb clear
col sumpool clear

PROMPT
PROMPT + PGA (size in MB) INSTANCE 1

select name,case when unit='bytes' then round(value/1048576) else value end value
from   gv$pgastat
where  INST_ID=1 and name in ('aggregate PGA target parameter','aggregate PGA auto target','total PGA allocated',
                'total PGA inuse','total PGA used for auto workareas','maximum PGA allocated',
                'maximum PGA used for auto workareas','over allocation count','cache hit percentage')
order  by name;


PROMPT
PROMPT + SGA INSTANCE 2

break on pool skip 1 on sumpool
col pool    format a12
col name    format a26
col mb      format 999,999
col sumpool format 999,999

select case when pool is null then '<MISC>' else pool end pool,name,mb,sumpool
from
(
    select pool,name,round(bytes/1048576) mb, round(sum(bytes) over(partition by pool)/1048576) sumpool,
           dense_rank() over(partition by pool order by bytes desc) r
    from   gv$sgastat
) where  r<=10
order by pool,mb desc;

clear breaks
col pool    clear
col name    clear
col mb clear
col sumpool clear

PROMPT
PROMPT + PGA (size in MB) INSTANCE 2

select name,case when unit='bytes' then round(value/1048576) else value end value
from   gv$pgastat
where  inst_id=2 and name in ('aggregate PGA target parameter','aggregate PGA auto target','total PGA allocated',
                'total PGA inuse','total PGA used for auto workareas','maximum PGA allocated',
                'maximum PGA used for auto workareas','over allocation count','cache hit percentage')
order  by name;

--==========================================================================================================
PROMPT
PROMPT + IO stat
PROMPT

column NAME format A20
column PERC_READS format A20
column PERC_WRITES format A20
column PERC_TOTAL format A20

SELECT T.NAME, SUM(Physical_READS) Physical_READS, 
ROUND((RATIO_TO_REPORT(SUM(Physical_READS)) OVER ())*100, 2) || '%' PERC_READS, 
SUM(Physical_WRITES) Physical_WRITES, 
ROUND((RATIO_TO_REPORT(SUM(Physical_WRITES)) OVER ())*100, 2) || '%' PERC_WRITES, 
SUM(total) total, ROUND((RATIO_TO_REPORT(SUM(total)) OVER ())*100, 2) || '%' PERC_TOTAL 
FROM (SELECT ts#, NAME, phyrds Physical_READS, phywrts Physical_WRITES,
phyrds + phywrts total
FROM v$datafile df, v$filestat fs
WHERE df.FILE# = fs.FILE#
ORDER BY phyrds DESC ) A, sys.ts$ T
WHERE A.ts# = T.ts#
GROUP BY T.NAME 
ORDER BY Physical_READS DESC;

--==========================================================================================================
PROMPT
PROMPT + Redo logs

COL member          FORMAT A100
col mbytes          format 999,999

SELECT l.thread#,l.group#,round(bytes/1048576) mbytes,f.member
FROM   v$log l, v$logfile f
WHERE  l.group# = f.group#
ORDER  BY thread#,group#,member;

COL member          CLEAR

set pagesize 0
PROMPT
PROMPT  all threads
PROMPT

with sw as
(
  select to_number(first_time-lag(first_time)
           over (partition by thread# order by first_time)) elapsed,
         first_time
  from   v$loghist
  where  first_time >= sysdate-2
)
select     'from         : '||to_char(min(first_time),'&&long_date_format')
||chr(10)||'to           : '||to_char(max(first_time),'&&long_date_format')
||chr(10)||'#log switches: '||count(*)
from  sw
where elapsed is not null
union all
select 'decile'||to_char(decile,'99')||'    : '||
       case when max_el >= 1 then to_char(max_el,'fm9999999')||'-'
       end||to_char(to_date(1,'j')+max_el,'hh24:mi:ss')
from (
    select decile,max(elapsed) max_el
    from (
        select elapsed,
               ntile(10) over (order by elapsed) decile
        from   sw
        where  elapsed is not null
    )
    group by decile
    order by decile
);

set pagesize 5000

PROMPT
PROMPT Logs switches

select to_char(ts_,'&&long_date_format') ts,
       nullif(t1,0) thread1,nullif(t2,0) thread2,nullif(t3,0) thread3,nullif(t4,0) thread4
from
(
  select trunc(first_time,'hh') ts_,thread#
  from   v$loghist
  where  first_time>=sysdate-2
)
pivot
(
  count(*) for thread# in (1 t1,2 t2,3 as t3,4 as t4)
)
order by ts_;

--==========================================================================================================
PROMPT
PROMPT + Cluster statistic

SELECT * FROM v$CLUSTER_INTERCONNECTS;

--==========================================================================================================
PROMPT
PROMPT + Top SQL statements (by elapsed time)
PROMPT

var esnap_id number
var rc1 refcursor
var rc2 refcursor

declare
  V_check pls_integer;
begin
  select max(snap_id),min(snap_id) into :esnap_id,:bsnap_id
  from   dba_hist_snapshot
  where  begin_interval_time>=trunc(sysdate,'hh')-to_dsinterval('PT24H5M');

  begin
    select 1 into V_check from dba_hist_sqlstat
    where  snap_id between :bsnap_id and :esnap_id
    and    rownum=1;

    open :rc1 for q'!
with q as (
  select (select min(begin_interval_time) from dba_hist_snapshot where snap_id=:bsnap_id) bt,
         (select max(end_interval_time)   from dba_hist_snapshot where snap_id=:esnap_id) et
  from dual
)
select 'Top SQL statements'
from q!' using :bsnap_id,:esnap_id;

    open :rc2 for
q'!select 'Exec:'    ||to_char(execu  ,'FM9,999,999,999')
     ||' Pc:'     ||to_char(pcall  ,'FM9,999,999,999')
     ||' Fetch:'  ||to_char(fe     ,'FM9,999,999,999')
     ||' LIO:'    ||to_char(buffg  ,'FM9,999,999,999')
     ||' PIO:'    ||to_char(diskr  ,'FM9,999,999,999')
     ||' Row:'    ||to_char(rowp   ,'FM9,999,999,999')
     ||' Sort:'   ||to_char(sorts  ,'FM9,999,999,999')
     ||' LIO/r:'  ||case when rowp<>0 then to_char(buffg/rowp,'FM9,999,999,990.0')   else '-' end
     ||' LIO/e:'  ||case when execu<>0 then to_char(buffg/execu,'FM9,999,999,990.0') else '-' end h1,
       'El:'      ||to_char(elapsed/1000,'FM9,999,999,990')
     ||' El/e:'   ||case when execu<>0 then to_char(elapsed/execu/1000,'FM9,999,999,990') else '-' end
     ||case when elapsed<>0 and round(cpu/elapsed*100)>0 then ' %cpu:'||to_char(cpu/elapsed*100,'FM990') end
     ||case when elapsed<>0 and round(iowait/elapsed*100)>0 then ' %io:'||to_char(iowait/elapsed*100,'FM990') end
     ||case when elapsed<>0 and round(ccwait/elapsed*100)>0 then ' %cc:'||to_char(ccwait/elapsed*100,'FM990') end
     ||case when elapsed<>0 and round(apwait/elapsed*100)>0 then ' %ap:'||to_char(apwait/elapsed*100,'FM990') end
     ||case when elapsed<>0 and round(clwait/elapsed*100)>0 then ' %cl:'||to_char(clwait/elapsed*100,'FM990') end
     ||case when elapsed<>0 and round(pls/elapsed*100)>0    then ' %pl:'||to_char(pls/elapsed*100,'FM990') end
     ||'  '||schema_name||','||nvl(substr(module,1,25),'-')||','||sql_id,sql_text
from
(
    select st.*,sq.sql_text
    from
    (
        select sql_id,max(module) module,
               max(parsing_schema_name) schema_name,
               sum(executions_delta)     execu,
               sum(disk_reads_delta)     diskr,
               sum(buffer_gets_delta)    buffg,
               sum(rows_processed_delta) rowp,
               sum(elapsed_time_delta)   elapsed,
               sum(parse_calls_delta)    pcall,
               sum(fetches_delta)        fe,
               sum(sorts_delta)          sorts,
               sum(cpu_time_delta)       cpu,
               sum(iowait_delta)         iowait,
               sum(clwait_delta)         clwait,
               sum(apwait_delta)         apwait,
               sum(ccwait_delta)         ccwait,
               sum(plsexec_time_delta)   pls
        from   dba_hist_sqlstat
        where  snap_id between :bsnap_id and :esnap_id
        group  by sql_id
    ) st,dba_hist_sqltext sq
    where st.sql_id=sq.sql_id
    and   st.module<>'check_oracle_11g'
    order by st.elapsed desc
)
where rownum<=5!' using :bsnap_id,:esnap_id;

  exception
  when no_data_found then
    open :rc1 for q'!select 'From '||to_char(startup_time,'&long_date_format')||' (instance startup)' from v$instance!';
    open :rc2 for
q'!select * from
(
  select 'Exec:'    ||to_char(executions    ,'FM9,999,999,999')
       ||' Pc:'     ||to_char(parse_calls   ,'FM9,999,999,999')
       ||' Fetch:'  ||to_char(fetches       ,'FM9,999,999,999')
       ||' LIO:'    ||to_char(buffer_gets   ,'FM9,999,999,999')
       ||' PIO:'    ||to_char(disk_reads    ,'FM9,999,999,999')
       ||' Row:'    ||to_char(rows_processed,'FM9,999,999,999')
       ||' Sort:'   ||to_char(sorts         ,'FM9,999,999,999')
       ||' LIO/r:'  ||case when rows_processed<>0 then to_char(buffer_gets/rows_processed,'FM9,999,999,990.0') else '-' end
       ||' LIO/e:'  ||case when executions<>0 then to_char(buffer_gets/executions,'FM9,999,999,990.0') else '-' end h1,
         'El:'      ||to_char(elapsed_time/1000,'FM9,999,999,990')
       ||' El/e:'   ||case when executions<>0 then to_char(elapsed_time/executions/1000,'FM9,999,999,990') else '-' end
       ||case when elapsed_time<>0 and round(cpu_time/elapsed_time*100)>0 then ' %cpu:'||to_char(cpu_time/elapsed_time*100,'FM990') end
       ||case when elapsed_time<>0 and round(user_io_wait_time/elapsed_time*100)>0     then ' %io:'||to_char(user_io_wait_time/elapsed_time*100,'FM990') end
       ||case when elapsed_time<>0 and round(cluster_wait_time/elapsed_time*100)>0     then ' %cc:'||to_char(cluster_wait_time/elapsed_time*100,'FM990') end
       ||case when elapsed_time<>0 and round(application_wait_time/elapsed_time*100)>0 then ' %ap:'||to_char(application_wait_time/elapsed_time*100,'FM990') end
       ||case when elapsed_time<>0 and round(cluster_wait_time/elapsed_time*100)>0     then ' %cl:'||to_char(cluster_wait_time/elapsed_time*100,'FM990') end
       ||case when elapsed_time<>0 and round(plsql_exec_time/elapsed_time*100)>0       then ' %pl:'||to_char(plsql_exec_time/elapsed_time*100,'FM990') end
       ||'  '||parsing_schema_name||','||nvl(substr(module,1,25),'-')||','||sql_id,sql_text
  from   v$sqlarea
  where  module<>'check_oracle_10g'
  order by elapsed_time desc
)
where rownum<=5!';
  end;
end;
/

set pagesize 0
print rc1
prompt

set linesize 120 recsepchar '-' recsep each long 65536 trimspool off

col sql_text        format a120 wrapped fold_before
col h1              fold_before

print rc2

col sql_text clear
col h1       clear
set recsep off recsepchar ' ' pagesize 50000 linesize 5000 trimspool on

--==========================================================================================================
PROMPT
PROMPT + Check Object Statistic 
PROMPT
prompt - Top 10 Biggest Table

column owner format A20
column SEGMENT_NAME format A50
SELECT * FROM (
  SELECT
    OWNER, SEGMENT_NAME, BYTES/1024/1024 SIZE_MB
  FROM
    DBA_SEGMENTS
WHERE
      SEGMENT_TYPE = 'TABLE'
  ORDER BY
    BYTES/1024/1024  DESC ) WHERE ROWNUM <= 10;
	
prompt Top 10 Biggest Index 
	
column SEGMENT_NAME format A50
SELECT * FROM (
  SELECT
    OWNER, SEGMENT_NAME, BYTES/1024/1024 SIZE_MB
  FROM
    DBA_SEGMENTS
WHERE
      SEGMENT_TYPE = 'INDEX'
  ORDER BY
    BYTES/1024/1024  DESC ) WHERE ROWNUM <= 10
/

PROMPT
prompt - Top 10 Biggest Segment (for Partitioned Object)

SELECT * FROM (
  SELECT
    SEGMENT_NAME, SEGMENT_TYPE, SUM(BYTES)/1024/1024 SIZE_MB
  FROM
    DBA_SEGMENTS
  WHERE
	SEGMENT_TYPE LIKE '%PARTI%'
  GROUP BY
      SEGMENT_NAME, SEGMENT_TYPE
  ORDER BY
    SUM(BYTES)/1024/1024  DESC ) WHERE ROWNUM <= 10
/

PROMPT
prompt - Separation of Tables and Indexes into Different Tablespaces 


select * from
(select owner, tablespace_name,segment_name,segment_type from dba_segments where owner in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS')) order by owner, tablespace_name)
pivot (COUNT(segment_name) for (segment_type) in ('TABLE' as Tables,'INDEX' as Indexes))
order by 3 desc, 4 desc;


PROMPT
prompt - Users with Objects in Tablespace SYSTEM, SYSAUX 


col OWNER format a30
col segment_name format a40
col tablespace_name format a30
col bytes format a20

select owner "User", segment_name "Segment Name", segment_type "Type", tablespace_name "TABLESPACE", bytes/1024/1024 "Size (MB)"
from dba_segments
where tablespace_name in ('SYSTEM','SYSAUX') and owner in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS'))
order by owner, bytes desc;


PROMPT
prompt - Tables without Indexes 


WITH TABLE_NO_INDEX AS 
(
select  OWNER,
 TABLE_NAME
from  dba_tables
minus
select  TABLE_OWNER,
 TABLE_NAME
from  dba_indexes
)
select  T.OWNER, T.TABLE_NAME, S.PARTITION_NAME, S.BYTES/1024/1024 "Size (MB)"
FROM TABLE_NO_INDEX T LEFT JOIN DBA_SEGMENTS S
ON (T.OWNER=S.OWNER and T.TABLE_NAME=S.SEGMENT_NAME)
where T.OWNER in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS'))
order  by S.BYTES desc, T.OWNER, T.TABLE_NAME, S.PARTITION_NAME;

PROMPT
prompt - Tables with More Than Five Indexes 

select 	Owner, TABLE_NAME "Table", COUNT(*) "Number of Indexes"
from  	dba_indexes
where  	OWNER in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS'))
group  	by OWNER, TABLE_NAME
having  COUNT(*) > 5
order 	by COUNT(*) desc, OWNER, TABLE_NAME;

PROMPT
prompt - List Table not analyzed in 60 days or nerver analyzed

select owner, table_name,NVL(to_char(last_analyzed,'dd-MON-YYYY'),'Nerver Analyze') 
from dba_tables 
where (last_analyzed is null or last_analyzed < (sysdate - 60)) 
and owner in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS')) 
order by owner;


PROMPT
prompt - List Index not analyzed in 60 days or nerver analyzed

select owner, table_name, index_name, NVL(to_char(last_analyzed,'dd-MON-YYYY'),'Nerver Analyze') 
from dba_indexes
where (last_analyzed is null or last_analyzed < (sysdate - 60)) 
and owner in (select username from dba_users where account_status='OPEN' and trunc(created) > (select trunc(created) from dba_users where username='SYS')) 
order by owner;


--==========================================================================================================
PROMPT
PROMPT + TOP foreground events

begin
    select min(snap_id) into :bsnap_id
    from   dba_hist_snapshot
    where  begin_interval_time>=sysdate-numtodsinterval(10,'hour')
    and    instance_number=:inst_num;
end;
/

col event_name format a30 trunc
col wait_class format a20 trunc
col waits      format 99G999G999
col w_per_s    format 999G999
col wt_sec     format 999G999
col wt_per_s   format 9G990D0
col avg_wt_ms  format 999G999
col "%DBt"     format 990
break on begin_ts skip 1

set linesize 200 pagesize 9000

select to_char(sna.begin_interval_time,'mm.yyyy') begin_ts,
       event_name,waits,waits/nullif( (cast(sna.end_interval_time as date)-cast(sna.begin_interval_time as date))*86400,0 ) w_per_s,
       wt/1e6 wt_sec,wt/nullif(waits,0)/1e3 avg_wt_ms,
       wt/nullif( (cast(sna.end_interval_time as date)-cast(sna.begin_interval_time as date))*86400,0 )/1e6 wt_per_s,
       100*wt/nullif(dbt.db_time-dbt.l_db_time,0) "%DBt",wait_class
from
(
    select snap_id,event_name,wait_class,
           total_waits_fg-l_total_waits_fg waits,
           time_waited_micro_fg-l_time_waited_micro_fg wt,
           row_number() over(partition by snap_id order by time_waited_micro_fg-l_time_waited_micro_fg desc) rn
    from
    (
        select snap_id,event_name,wait_class,
               total_waits_fg,time_waited_micro_fg,
               lag(total_waits_fg,1)       over (partition by event_name order by snap_id) l_total_waits_fg,
               lag(time_waited_micro_fg,1) over (partition by event_name order by snap_id) l_time_waited_micro_fg
        from
        (
            select snap_id,event_name,wait_class,
                   total_waits_fg,time_waited_micro_fg
            from   dba_hist_system_event
            where  snap_id>=:bsnap_id
            and    instance_number=:inst_num
            and    wait_class<>'Idle'
            union all
            select snap_id,stat_name,to_char(null) wait_class,
                   to_number(null),value
            from   dba_hist_sys_time_model
            where  snap_id>=:bsnap_id
            and    instance_number=:inst_num
            and    stat_name='DB CPU'
       )
    )
    where l_time_waited_micro_fg is not null
) evt,
(
    select snap_id,db_time,lag(db_time,1,0) over(order by snap_id) l_db_time
    from
    (
        select snap_id,value db_time
        from   dba_hist_sys_time_model
        where  snap_id>=:bsnap_id
        and    stat_name='DB time'
        and    instance_number=:inst_num
    )
) dbt,dba_hist_snapshot sna
where rn<=5
and   evt.snap_id=dbt.snap_id
and   evt.snap_id=sna.snap_id
and   sna.instance_number=:inst_num
and   sna.startup_time not between sna.begin_interval_time and sna.end_interval_time
order  by evt.snap_id,evt.wt desc;

col event_name clear
col wait_class clear
col waits      clear
col w_per_s    clear
col wt_sec     clear
col wt_per_s   clear
col avg_wt_ms  clear
col "%DBt"     clear
clear breaks


--==========================================================================================================
PROMPT
PROMPT + Check Backup

col START_TIME format a30
col END_TIME format a30
col DOW format a10
set lines 300
set pages 1000
col cf for 9,999
col df for 9,999
col elapsed_seconds heading "ELAPSED|SECONDS"
col i0 for 9,999
col i1 for 9,999
col l for 9,999
col output_mbytes for 9,999,999 heading "OUTPUT|MBYTES"
col session_recid for 999999 heading "SESSION|RECID"
col session_stamp for 99999999999 heading "SESSION|STAMP"
col status for a10 trunc
col time_taken_display for a10 heading "TIME|TAKEN"
col output_instance for 9999 heading "OUT|INST"
select
  to_char(j.start_time, 'yyyy-mm-dd hh24:mi:ss') start_time,
  to_char(j.end_time, 'yyyy-mm-dd hh24:mi:ss') end_time,
  (j.output_bytes/1024/1024) output_mbytes, j.status, j.input_type,
  decode(to_char(j.start_time, 'd'), 1, 'Sunday', 2, 'Monday',
                                     3, 'Tuesday', 4, 'Wednesday',
                                     5, 'Thursday', 6, 'Friday',
                                     7, 'Saturday') dow,
  j.elapsed_seconds, j.time_taken_display
from V$RMAN_BACKUP_JOB_DETAILS j
  left outer join (select
                     d.session_recid, d.session_stamp,
                     sum(case when d.controlfile_included = 'YES' then d.pieces else 0 end) CF,
                     sum(case when d.controlfile_included = 'NO'
                               and d.backup_type||d.incremental_level = 'D' then d.pieces else 0 end) DF,
                     sum(case when d.backup_type||d.incremental_level = 'D0' then d.pieces else 0 end) I0,
                     sum(case when d.backup_type||d.incremental_level = 'I1' then d.pieces else 0 end) I1,
                     sum(case when d.backup_type = 'L' then d.pieces else 0 end) L
                   from
                     V$BACKUP_SET_DETAILS d
                     join V$BACKUP_SET s on s.set_stamp = d.set_stamp and s.set_count = d.set_count
                   where s.input_file_scan_only = 'NO'
                   group by d.session_recid, d.session_stamp) x
    on x.session_recid = j.session_recid and x.session_stamp = j.session_stamp
  left outer join (select o.session_recid, o.session_stamp, min(inst_id) inst_id
                   from GV$RMAN_OUTPUT o
                   group by o.session_recid, o.session_stamp)
    ro on ro.session_recid = j.session_recid and ro.session_stamp = j.session_stamp
where j.start_time > trunc(sysdate)-30
order by j.start_time;

--==========================================================================================================
PROMPT
PROMPT + Check Job Gather Statistic
SET lines 300
SELECT client_name, status FROM dba_autotask_operation;
SELECT CLIENT_NAME,
       STATUS,
       MEAN_INCOMING_TASKS_7_DAYS,
       MEAN_INCOMING_TASKS_30_DAYS
FROM   DBA_AUTOTASK_CLIENT
WHERE  CLIENT_NAME = 'auto optimizer stats collection'
/

col job_name FOR a30
SET lines 150
SELECT * FROM
     (SELECT log_date,job_name,status,actual_start_date,run_duration
     FROM DBA_SCHEDULER_JOB_RUN_DETAILS
     WHERE job_name='GATHER_STATS_JOB'
     ORDER BY log_id DESC)
     WHERE rownum<=30;

ALTER SESSION SET NLS_DATE_FORMAT='DD-MON-YYYY HH24:MI:SS';
SET LINESIZE 150
COLUMN task_name FORMAT A25
COLUMN execution_name FORMAT A20
COLUMN execution_end FORMAT A20
COLUMN execution_type FORMAT A20

SELECT task_name,
       execution_name,
       execution_start,
       execution_end,
       execution_type,
       status
FROM   dba_advisor_executions
WHERE  task_name = 'AUTO_STATS_ADVISOR_TASK'
AND    execution_end >= SYSDATE-30
ORDER BY 3;

PROMPT
PROMPT
PROMPT Check Tablespace including Autoextend
column Tablespace format a30
column "Total GB" format a15
column "Free GB" format a15
column "Pct Free" format a15

SELECT fa.tablespace_name "Tablespace",
       TO_CHAR (df.size_GB, '99990D99') "Total GB",
       TO_CHAR (fa.free_auto_GB + f.free_GB, '99990D99') "Free GB",
       TO_CHAR ( (fa.free_auto_GB + f.free_GB) / df.size_GB * 100, '990D99') "Pct Free"
FROM ( SELECT f.tablespace_name,
              SUM (DECODE (autoextensible, 'YES', (maxbytes - f.bytes) / 1024 / 1024 / 1024,0)) free_auto_GB
       FROM dba_data_files f
       GROUP BY f.tablespace_name) fa,
     ( SELECT s.tablespace_name,
              SUM (s.bytes / 1024 / 1024 / 1024) free_GB
       FROM dba_free_space s
       GROUP BY s.tablespace_name) f,
     ( SELECT df.tablespace_name,
              SUM (DECODE (autoextensible, 'YES', maxbytes, bytes)) / 1024 / 1024 / 1024 size_GB
       FROM dba_data_files df
       GROUP BY df.tablespace_name) df
WHERE fa.tablespace_name = f.tablespace_name
AND   df.tablespace_name = f.tablespace_name
ORDER BY 4;

============================
PROMPT
PROMPT
PROMPT Check Tablespace A DUONG
set linesize 800
col tablespace_name format a16
col Usage format 999,990
col sizeMB format 999,999,990
col sizeMB format 999,999,990
col usedMb format 999,999,990
col Pct_Free format 990
col Pct_used format 990
col MAX format 999,999,990
col Used_pct_of_Max format 990
SELECT a.tablespace_name,
       ROUND (100 * (a.bytes_alloc - NVL (b.bytes_free, 0)) / a.bytes_alloc,
              0)
          Usage,
       ROUND (a.bytes_alloc / 1024 / 1024) sizeMB,
       ROUND (NVL (b.bytes_free, 0) / 1024 / 1024) sizeMB,
       ROUND ( (a.bytes_alloc - NVL (b.bytes_free, 0)) / 1024 / 1024) usedMb,
       ROUND ( (NVL (b.bytes_free, 0) / a.bytes_alloc) * 100) Pct_Free,
       100 - ROUND ( (NVL (b.bytes_free, 0) / a.bytes_alloc) * 100) Pct_used,
       ROUND (maxbytes / 1048576) MAX,
       ROUND (100 * (a.bytes_alloc - NVL (b.bytes_free, 0)) / a.maxbytes, 0)
          Used_pct_of_Max
  FROM (  SELECT f.tablespace_name,
                 SUM (f.bytes) bytes_alloc,
                 SUM (
                    DECODE (f.autoextensible,
                            'YES', f.maxbytes,
                            'NO', f.bytes))
                    maxbytes
            FROM dba_data_files f
        GROUP BY tablespace_name) a,
       (  SELECT f.tablespace_name, SUM (f.bytes) bytes_free
            FROM dba_free_space f
        GROUP BY tablespace_name) b
 WHERE     a.tablespace_name = b.tablespace_name(+)
UNION ALL
  SELECT h.tablespace_name,
         ROUND (
            SUM (NVL (p.bytes_used, 0)) / SUM (h.bytes_free + h.bytes_used),
            0)
            Usage,
         ROUND (SUM (h.bytes_free + h.bytes_used) / 1048576) megs_alloc,
         ROUND (
              SUM ( (h.bytes_free + h.bytes_used) - NVL (p.bytes_used, 0))
            / 1048576)
            megs_free,
         ROUND (SUM (NVL (p.bytes_used, 0)) / 1048576) megs_used,
         ROUND (
              (  SUM ( (h.bytes_free + h.bytes_used) - NVL (p.bytes_used, 0))
               / SUM (h.bytes_used + h.bytes_free))
            * 100)
            Pct_Free,
           100
         - ROUND (
                (  SUM ( (h.bytes_free + h.bytes_used) - NVL (p.bytes_used, 0))
                 / SUM (h.bytes_used + h.bytes_free))
              * 100)
            pct_used,
         ROUND (
              SUM (DECODE (f.AUTOEXTENSIBLE, 'YES', f.maxbytes, f.bytes))
            / 1048576,
            0)
            MAX,
         ROUND (
              100
            * SUM (NVL (p.bytes_used, 0))
            / SUM (DECODE (f.AUTOEXTENSIBLE, 'YES', f.maxbytes, f.bytes)),
            0)
            Used_pct_of_Max
    FROM sys.v_$TEMP_SPACE_HEADER h,
         sys.v_$Temp_extent_pool p,
         dba_temp_files f
   WHERE     p.file_id(+) = h.file_id
         AND p.tablespace_name(+) = h.tablespace_name
         AND f.file_id = h.file_id
         AND f.tablespace_name = h.tablespace_name
GROUP BY h.tablespace_name
ORDER BY 1;

PROMPT =================================
PROMPT SHRINK RECOMMEND

set linesize 1200
col tablespace_name format a30
col segment_owner format a10
col segment_name format a30
col segment_type format a20
col partition_name format a10
col aLLOCATED_SPACE format 9999999999
col USED_SPACE  format 9999999999
col RECLAIMABLE_SPACE format 9999999999

select /*+  NO_MERGE  */ /* 2b.57 */  tablespace_name, segment_owner, segment_name, segment_type, partition_name , aLLOCATED_SPACE ,USED_SPACE ,RECLAIMABLE_SPACE
FROM TABLE(dbms_space.asa_recommendations())
Where segment_owner not in ('ANONYMOUS','APEX_030200','APEX_040000','APEX_SSO','APPQOSSYS','CTXSYS','DBSNMP','DIP','EXFSYS','FLOWS_FILES','MDSYS','OLAPSYS','ORACLE_OCM','ORDDATA','ORDPLUGINS','ORDSYS','OUTLN','OWBSYS') and
   segment_owner not in ('SI_INFORMTN_SCHEMA','SQLTXADMIN','SQLTXPLAIN','SYS','SYSMAN','SYSTEM','TRCANLZR','WMSYS','XDB','XS$NULL','PERFSTAT','STDBYPERF','MGDSYS','OJVMSYS')
order by reclaimable_space desc;


PROMPT =================================
PROMPT CHECK PATCH

COLUMN action_time FORMAT A20
COLUMN action FORMAT A10
COLUMN status FORMAT A10
COLUMN description FORMAT A40
COLUMN version FORMAT A10
COLUMN bundle_series FORMAT A10
SELECT TO_CHAR(action_time, 'DD-MON-YYYY HH24:MI:SS') AS action_time,
       action,
       namespace,
       version,
       id,
       comments
FROM   sys.registry$history
ORDER by action_time;


PROMPT
PROMPT
PROMPT NO of USER CONNECTED
PROMPT ===========================================================
select count(distinct username) "No. of users Connected" from v$session where username is not null;



PROMPT
PROMPT
PROMPT NO of SESSIONS CONNECTED
PROMPT ===========================================================
Select count(*) AS "No of Sessions connected" from v$session where username is not null;


PROMPT
PROMPT
PROMPT DISTINCT USERNAME CONNECTED
PROMPT ===========================================================
Select distinct(username) AS "USERNAME" from v$session;

PROMPT
PROMPT
PROMPT DICTIONARY HIT RATIO. THIS VALUE SHOULD BE GREATER 85%
PROMPT ==========================================================
select   (  1 - ( sum (decode (name, 'physical reads', value, 0)) / (  sum (decode (name, 'db block gets',value, 0)) + sum (decode (name, 'consistent gets', value, 0))))) * 100 "Buffer Hit Ratio" from v$sysstat;

PROMPT
PROMPT
PROMPT LIBRARY CACHE HIT RATIO. THIS VALUE SHOULD BE GREATER 90%
PROMPT ===========================================================
select (sum(pins)/(sum(pins)+sum(reloads))) * 100 "Library Cache Hit Ratio" from v$librarycache;


PROMPT
PROMPT
PROMPT BLOCKING QUERY
PROMPT ===========================================================
select s1.username || '@' || s1.machine|| ' ( SID=' || s1.sid || ' )  is blocking '|| s2.username || '@' || s2.machine || ' ( SID=' || s2.sid || ' ) ' AS blocking_status from v$lock l1, v$session s1, v$lock l2, v$session s2 where s1.sid=l1.sid and s2.sid=l2.sid and l1.BLOCK=1 and l2.request > 0 and l1.id1 = l2.id1 and l2.id2 = l2.id2;


PROMPT
PROMPT
PROMPT INVALID OBJECT COUNT LIST
PROMPT ===========================================================
select count(*),owner ,  object_type from dba_objects where status='INVALID' group by owner ,  object_type order by owner;

PROMPT
PROMPT
PROMPT PARAMETER INFO
PROMPT ==========================================================
show parameter;

spool off

set term on flush on pagesize 24
PROMPT check_oracle complete
PROMPT
PROMPT

EXEC DBMS_APPLICATION_INFO.set_module(NULL,NULL)


exit

