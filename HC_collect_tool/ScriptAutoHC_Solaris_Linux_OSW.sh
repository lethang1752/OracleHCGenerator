#!/bin/ksh

#-----Declare variable for script

diag="V\$DIAG_INFO"
database="V\$DATABASE"
parameter="V\$PARAMETER"
datafile="V\$DATAFILE_HEADER"
instance="V\$INSTANCE"
backupjob="V\$RMAN_BACKUP_JOB_DETAILS"
host=$(hostname)
os=$(uname)

#-----Start script

echo
echo "MAKE SURE SERVER HAS KSH (KORNSHELL)*"
echo "Set variable for the process..."
echo "===============>>"
if [[ "$os" == 'Linux' ]]; then
	read -p " <> ORACLE_SID($ORACLE_SID)  : " sid
	[[ $sid != "" ]] && export ORACLE_SID=$sid
	read -p " <> ORACLE_HOME($ORACLE_HOME) : " home
	[[ $home != "" ]] && export ORACLE_HOME=$home
else
	read sid?" <> ORACLE_SID($ORACLE_SID)  : "
	[[ $sid != "" ]] && export ORACLE_SID=$sid
	read home?" <> ORACLE_HOME($ORACLE_HOME) : "
	[[ $home != "" ]] && export ORACLE_HOME=$home
fi
export PATH=$ORACLE_HOME/bin:$PATH
echo "<<==============="

#-----Declare variable for OS Machine

if [[ "$os" == 'Linux' ]]; then
	grep='grep'
	awk='awk'
	java_check="version"
else
	grep='ggrep'
	awk='nawk'
	java_check="version"
fi

#-----Get Date now

time=$(date +'%d_%m_%Y')

#-----Get Grid home

grid_file=/etc/init.d/init.ohasd
if [[ -f "$grid_file" ]]; then
	grid=$($awk -F = ' $1 ~ /^ORA_CRS_HOME$/ {print $2} ' /etc/init.d/init.ohasd)
	export PATH=$PATH:$grid/bin
else
	echo "None Grid Environment."
	grid='N/A'
fi

#-----Get pwd script

SCRIPT=$(readlink -f "$0")
pwd=$(dirname "$SCRIPT")

#-----Check connection

if
	echo "exit;" | sqlplus / as sysdba 2>&1 | grep -q "Connected to:"
then
	echo Connect Database OK
else
	echo Connect Database FAIL
	exit
fi

#-----Get db_name

dbname=$(
	sqlplus -s / as sysdba <<EOF
set head off
set feed off
select name from $database;
exit
EOF
)
dbname=$(echo $dbname | tr -d '[:space:]')

#-----Get instance_name

insname=$(
	sqlplus -s / as sysdba <<EOF
set head off
set feed off
show parameter instance_name;
exit
EOF
)
insname=$(echo $insname | awk -v FS=' ' '{print $3}')

#-----Create folder

mkdir -p $pwd/${insname}
cd $pwd/${insname}

#-----Get Trace Log path

spwd=$(
	sqlplus -s / as sysdba <<EOF
set head off
set feed off
set line 150 
select value from $diag where name='Diag Trace';
exit
EOF
)

Head() {
	echo
	echo ">>--------------------------- *** ---------------------------<<"
	echo "<<========================<<  MPS  >>========================>>"
	echo ">>--------------------------- *** ---------------------------<<"
	echo
	echo " <> OS Machine :" $os
	echo " <> Date Time  :" $time
	echo " <> Script     :" $pwd
	echo " <> DB Name    :" $dbname
	echo " <> Instance   :" $ORACLE_SID
	echo " <> Grid Home  :" $grid
	echo " <> DB Home    :" $ORACLE_HOME
	echo " <> Alert Log  :" $spwd
	echo " <> OSWbb Log  :" $pwd/oswbb_log_MPS_$host
	echo
	echo "|<<=======================<<  ***  >>=======================>>|"
	echo "|                                                             |"
	echo "|                    ---------------------                    |"
	echo "|   <<===>>       << HEALTH-CHECK-DATABASE >>       <<===>>   |"
	echo "|                    ---------------------                    |"
	echo "|                                                             |"
	echo "| ==>> 1. Get Database Information:                           |"
	echo "|                                                             |"
	echo "|         # Disk, Alert_log, Rman, Opatch, Grid               |"
	echo "|         # Backup, Dbinfo, HealthCheck, Awrrpt               |"
	echo "|                                                             |"
	echo "| ==>> 2. Run OSWatcher.                                      |"
	echo "|                                                             |"
	echo "| ==>> 3. Cancel Script.                                      |"
	echo "|                                                             |"
	echo "|                                                             |"
	echo "|                                         Ver_1.4.0_VictorLe  |"
	echo "|<<========================<< *** >>========================>>|"
	echo
	if [[ "$os" == 'Linux' ]]; then
		read -p "Option: " option
	else
		read option?"Option: "
	fi
}

Body() {
	#==#==Option Failed.

if [[ -z $option || $option < 1 || $option > 3 ]]; then
	echo
	echo "=> #Error! Choose again."

#==#==Option 1.(Run collect log)

elif [ $option == 1 ]; then
	echo
	echo "Processing..."
	echo

#-----Alert_log

cp $spwd/alert_$ORACLE_SID.log .
echo
echo "************************"
echo "* Copy Alert_log done. *"
echo "************************"
echo

#-----HealthCheck

sqlplus / as sysdba <<EOF
@$pwd/HealthCheck.sql
EOF
echo
echo "*************************"
echo "* Get HealthCheck done. *"
echo "*************************"
echo

#-----database_information

sqlplus / as sysdba <<EOF
@$pwd/database_information.sql
EOF
echo
echo "**********************************"
echo "* Get Database Information done. *"
echo "**********************************"
echo

#-----OS_Command

file_name='database_information.html'
unamestr=$(uname)
if [[ "$unamestr" == 'AIX' ]]; then
	disk_command='df -g'
else
	disk_command='df -h'
fi

#-----Disk_Usage

echo "<p>+ DISK_USAGE</p>" >>$file_name
$disk_command -P | $grep -v ^none | (
read header
echo "$header"
sort -rn -k 5 ) | $awk 'BEGIN{print("<table WIDTH='90%' BORDER='1'><tr><th>'FILESYSTEM'</th><th>'SIZE'</th><th>'USED'</th><th>'AVAIL'</th><th>'USE%'</th><th>'MOUNTED_ON'</th></tr>")}
{
	if ($2!="0K" && $2!="Size") {
		print("<tr><td>",$1,"</td><td>",$2,"</td><td>",$3,"</td><td>",$4,"</td><td>",$5,"</td><td>",$6,$7,"</td></tr>")
	}
}
END{
	print("</table><p><p>")
}' >>$file_name

#-----Check_Listener

echo "<p>+ CHECK_LISTENER</p>" >>$file_name
lsnrctl stat | awk 'BEGIN{
print("<p><table WIDTH='90%' BORDER='1'><tr><th>'LISTENER_STATUS'</th></tr><tr><td>")}
{
	if ($0!=NULL) {
		print($0,"<br>")
	}
}
	END{
		print("</tr></td></table><p><p>")
	}' >>$file_name

#-----Check_Patches

echo "<p>+ CHECK_PATCHES</p>" >>$file_name
$ORACLE_HOME/OPatch/opatch lsinventory | $grep -B 2 "Patch description" | grep -v "Unique" | $awk -v hs=$host -v oraclehome=$ORACLE_HOME 'BEGIN{print("<p><table WIDTH='90%' BORDER='1'><tr><th>SERVER</th><th>ORACLE_HOME</th><th>PATCH INFORMATION</th></tr><tr><td>",hs,"</td><td>",oraclehome,"</td><td>")}
{
	if ($0!=NULL) {
		print($0,"<br>")
	}
}
END{
	print("</tr></td></table><p><p>")
}' >>$file_name

#-----Backup_Policy

echo "<p>+ BACKUP_POLICY</p>" >>$file_name
echo "<table WIDTH='90%' BORDER='1'><tr><th>RMAN_RETENTION</th></tr><tr><td>" >>$file_name
rman target / <<EOF | grep CONFIGURE >>$file_name
show retention policy;
EOF
echo "</tr></td><tr><td>NULL</td></tr></table>" >>$file_name

#-----No Grid Option

if [ "$grid" == "N/A" ]; then
	echo "<p>+ RESOURCE_CRS</p>" >>$file_name
	echo "<table WIDTH='90%' BORDER='1'>" >>$file_name
	echo "<tr><th>NAME</th><th>TARGET</th><th>STATE</th><th>TARGET_SERVER</th><th>STATE_DETAILS</th></tr>" >>$file_name
	echo "<tr><td>NULL</td><td>NULL</td><td>NULL</td><td>NULL</td><td>NULL</td></tr></table>" >>$file_name
	echo "<p>+ CHECK_CLUSTER</p>" >>$file_name
	echo "<table WIDTH='90%' BORDER='1'>" >>$file_name
	echo "<tr><th>HOST_NAME</th><th>CLUSTER_SERVICE</th></tr><tr><td>NULL</td><td>NULL</td></tr></table>" >>$file_name
else

#-----Resource_Crs

echo "<p>+ RESOURCE_CRS<p>" >>$file_name
crsctl status resource -v |
egrep -e "NAME|TARGET|STATE|LAST_SERVER|STATE_DETAILS" |
/bin/gawk 'BEGIN {FS="=";}
{
if ($1=="NAME")  resname=$2; else
if ($1=="TARGET") restrg=$2; else
if ($1=="STATE")   resst=$2; else
if ($1=="LAST_SERVER") {
resser=$2
} else
if ($1=="STATE_DETAILS") {
resdet=$2;
if(length($3)!=0) { resdet=resdet"="$3 }
idxx1=index(resst, " "); tat=substr(resst, 0, idxx1);
if (tat=="") {tat="OFFLINE"};
printf "%-35s %-20s %-25s %-20s %-10s\n", resname, restrg, tat, resser, resdet}
}' | $awk 'BEGIN{print("<table WIDTH='90%' BORDER='1'><tr><th>'NAME'</th><th>'TARGET'</th><th>'STATE'</th><th>'LAST_SERVER'</th><th>'STATE_DETAILS'</th></tr>")}
{
	if ($4!=NULL) {
		print("<tr><td>",$1,"</td><td>",$2,"</td><td>",$3,"</td><td>",$4,"</td><td>",$5,$6,$7,$8,"</td></tr>")
	}
}
END{
	print("</table>")
}' >>$file_name

#-----Check_Cluster

echo "<p>+ CHECK_CLUSTER<p>" >>$file_name
# Determine which check command to use
if crsctl check crs 2>&1 | grep -q "invalid argument"; then
    crs_cmd="crsctl check has"
else
    crs_cmd="crsctl check crs"
fi
$crs_cmd | $awk -v hs=$host 'BEGIN{print("<p><table WIDTH='90%' BORDER='1'><tr><th>HOST_NAME</th><th>CLUSTER_SERVICE</th></tr><tr><td>",hs"</td><td>")}
{
	if ($0!=NULL) {
		print($0,"<br>")
	}
}
END{
	print("</td></tr></table>")
}' >>$file_name
fi

echo
echo "************************"
echo "* Get OS_Command done. *"
echo "************************"
echo

#-----Awrrpt

end_snap_id=$(
	sqlplus -s / as sysdba <<EOF
set head off
set feed off
SELECT
    snap_id
FROM
    (
        SELECT
            snap.instance_number,
            snap.snap_id,
            delta,
            begin_interval_time,
            end_interval_time,
            startup_time
        FROM
            (
                SELECT
                    *
                FROM
                    dba_hist_snapshot
                WHERE
                    dbid = (
                        SELECT
                            dbid
                        FROM
                            $database
                    )
            ) snap,
            (
                SELECT
                    instance_number
                FROM
                    $instance
            ) inst,
            (
                SELECT
                    instance_number,
                    snap_id,
                    sum_val,
                    sum_val - LAG(sum_val)
                              OVER(
                        ORDER BY
                            instance_number, snap_id
                              ) delta
                FROM
                    (
                        SELECT
                            instance_number,
                            snap_id,
                            SUM(value) sum_val
                        FROM
                            dba_hist_service_stat
                        WHERE
                                stat_name = 'DB time'
                            AND dbid = (
                                SELECT
                                    dbid
                                FROM
                                    $database
                            )
                        GROUP BY
                            instance_number,
                            snap_id
                    )
            ) stat
        WHERE
                snap.instance_number = stat.instance_number
            AND snap.instance_number = inst.instance_number
            AND snap.snap_id = stat.snap_id
            AND begin_interval_time > sysdate - 7
            AND delta IS NOT NULL
        ORDER BY
            delta DESC
    )
WHERE
    ROWNUM <= 1;
exit
EOF
)
end_snap_id=$(echo $end_snap_id | tr -d '[:space:]')

begin_snap_id=$(
	sqlplus -s / as sysdba <<EOF
set head off
set feed off
SELECT
    snap_id - 1
FROM
    (
        SELECT
            snap.instance_number,
            snap.snap_id,
            delta,
            begin_interval_time,
            end_interval_time,
            startup_time
        FROM
            (
                SELECT
                    *
                FROM
                    dba_hist_snapshot
                WHERE
                    dbid = (
                        SELECT
                            dbid
                        FROM
                            $database
                    )
            ) snap,
            (
                SELECT
                    instance_number
                FROM
                    $instance
            ) inst,
            (
                SELECT
                    instance_number,
                    snap_id,
                    sum_val,
                    sum_val - LAG(sum_val)
                              OVER(
                        ORDER BY
                            instance_number, snap_id
                              ) delta
                FROM
                    (
                        SELECT
                            instance_number,
                            snap_id,
                            SUM(value) sum_val
                        FROM
                            dba_hist_service_stat
                        WHERE
                                stat_name = 'DB time'
                            AND dbid = (
                                SELECT
                                    dbid
                                FROM
                                    $database
                            )
                        GROUP BY
                            instance_number,
                            snap_id
                    )
            ) stat
        WHERE
                snap.instance_number = stat.instance_number
            AND snap.instance_number = inst.instance_number
            AND snap.snap_id = stat.snap_id
            AND begin_interval_time > sysdate - 7
            AND delta IS NOT NULL
        ORDER BY
            delta DESC
    )
WHERE
    ROWNUM <= 1;
exit
EOF
)
begin_snap_id=$(echo $begin_snap_id | tr -d '[:space:]')

TEMPFILE=/tmp/tmpawr.sql
echo "<<=====================================================>>"
echo "AWRRPT: @$ORACLE_HOME/rdbms/admin/awrrpt.sql"
echo "<<=====================================================>>"
echo
sqlplus -s / as sysdba <<EOF >/dev/null
set echo off
set head off
set feed off
spool ${TEMPFILE}
select 'define begin_snap = ${begin_snap_id}' from dual;
select 'define end_snap = ${end_snap_id}' from dual;
select 'define report_type = ' || '''html''' from dual;
select 'define inst_name = ' || INSTANCE_NAME from v\$instance;
select 'define db_name = ' || name from v\$database;
select 'define dbid = ' || dbid from v\$database;
select 'define inst_num = ' || INSTANCE_NUMBER from v\$instance;
select 'define num_days = 7' from dual;
select 'define report_name = awrrpt_${time}_' || INSTANCE_NAME from v\$instance;
select '@$ORACLE_HOME/rdbms/admin/awrrpt.sql' from dual;
spool off
exit
EOF

sqlplus -s / as sysdba <<EOF >/dev/null
@$TEMPFILE
exit
EOF

if [ -f "$TEMPFILE" ]; then
	rm $TEMPFILE
fi
echo
echo "********************"
echo "* Get Awrrpt done. *"
echo "********************"

#-----Get information for report file

# Name, HA/Standalone, Hardware, File system, Archiving, Flashback, Version,Patch, DB size, Backup status

# Name: dbname

# HA/Standalone

rp_ha=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
SELECT VALUE FROM $parameter WHERE NAME = 'cluster_database';
exit
EOF
)

rp_ha=$(echo $rp_ha | tr -d '[:space:]')
if [[ "$rp_ha" == 'TRUE' ]]; then
	rp_ha_last='RAC'
else
	rp_ha_last='Stand Alone'
fi

# OS

os1=$(uname -o)
os2=$(uname -m)
os_last="$os1 $os2"

#Hardware

if [[ "$os" == 'Linux' ]]; then
	cpu=$(lscpu | grep -E '^CPU\(s\):' | tr -dc '0-9')
	ram=$(cat /proc/meminfo | grep MemTotal | tr -dc '0-9')
	ram_last=$(expr "$ram" / 1048576)
else
	cpu=$(psrinfo -p)
	ram=$(prtconf | grep Mem | ggrep -E -o '[0-9]+')
	ram_last=$(expr "$ram" / 1024)
fi
hw_last="CPU: $cpu cores, RAM: $ram_last GB"

#Filesystem

dtf=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select name from $datafile where name like '+%' and rownum=1;
exit
EOF
)

if [[ $dtf == *[+]* ]]; then
	dtf_last="ASM"
else
	dtf_last="File System"
fi

#Archiving Enabled

arc=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select log_mode from $database;
exit
EOF
)

if [[ "$arc" == *"ARCHIVELOG"* ]]; then
	arc_last='Yes'
else
	arc_last='No'
fi

#Flashback Enabled

fls=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select FLASHBACK_ON from $database;
exit
EOF
)

if [[ "$fls" == *"YES"* ]]; then
	fls_last='Yes'
else
	fls_last='No'
fi

#Version
ver_last=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select version from $instance;
exit
EOF
)

#Patch

pat_last=$($ORACLE_HOME/OPatch/opatch lspatches | $grep "Database" | $awk -v FS=';' '{print $2}')

#DBsize

size=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select round(sum(bytes)/1024/1024/1024,2) from $datafile;
exit
EOF
)

size_last="$size GB"

#Backup status

bkp_last=$(
sqlplus -s / as sysdba <<EOF
set head off
set feed off
select status from $backupjob where rownum=1;
exit
EOF
)

#Print table

echo "<p>+ REPORT DETAILS</p>" >>$file_name
echo "<table WIDTH='90%' BORDER='1'>" >>$file_name
echo "<tr><th>ITEMS</th><th>INFORMATION</th></tr>" >>$file_name
echo "<tr><td>Name</td><td>$dbname</td></tr>" >>$file_name
echo "<tr><td>HA/Stand Alone</td><td>$rp_ha_last</td></tr>" >>$file_name
echo "<tr><td>OS Version</td><td>$os_last</td></tr>" >>$file_name
echo "<tr><td>Hardware (CPU,RAM)</td><td>$hw_last</td></tr>" >>$file_name
echo "<tr><td>File System/raw devices</td><td>$dtf_last</td></tr>" >>$file_name
echo "<tr><td>Archiving Enabled</td><td>$arc_last</td></tr>" >>$file_name
echo "<tr><td>Flashback Enabled</td><td>$fls_last</td></tr>" >>$file_name
echo "<tr><td>Version</td><td>$ver_last</td></tr>" >>$file_name
echo "<tr><td>Patch</td><td>$pat_last</td></tr>" >>$file_name
echo "<tr><td>DB Size</td><td>$size_last</td></tr>" >>$file_name
echo "<tr><td>Backup status</td><td>$bkp_last</td></tr></table>" >>$file_name
echo
echo "**************************"
echo "* Get Report Infor done. *"
echo "**************************"
echo
echo "End of process!"

#==#==Option 2 (OSWatcher).

elif [ $option == 2 ]; then
	echo
	echo "Setup OSWatcher..."
	echo

#-----Check Java Installation

if java -version 2>&1 >/dev/null | $grep "$java_check"; then
	echo "Java installed!"

	#-----Tar file oswbb840.tar

	tar -xf $pwd/oswbb840.tar -C $pwd/.

	#-----Create folder for oswbb_log

	mkdir -p $pwd/oswbb_log_MPS_$host

	#-----Start oswbb840

	cd $pwd/oswbb
	nohup ./startOSWbb.sh 300 120 None $pwd/oswbb_log_MPS_$host/ >nohup.out 2>&1 &
	echo
	echo

else
	echo "Java NOT installed!"
	echo "Prepare to Install Java!"
fi

	#==#==Option 3 (Exit).

	elif [ $option == 3 ]; then
		echo
		exit
	fi
}

#=====MainStream

Head
while :; do
	Body
	Head
done