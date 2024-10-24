<#
File		:	Get-SQLHealth.ps1

Parameters	:
			IN 	:	-Config
            	OUT	:	

Summary	:	Daily Health Check
#>

param([string]$config="config.xml", [string]$scriptType = "HEALTHCHECK", [string]$scriptSubtype = "FULL", [string]$scriptMessage = "No message", [int]$testFlag = $false)


$globalScriptType
$globalScriptSubType
$globalScriptMessage
## Modified 10/01/19 - GFF
## Added global variable for command timeouts.
$globalCommandTimeout
## Modified 10/01/19 - GFF
## Added global variable for script timeouts.
$globalScriptTimeout
$globalTestFlag


function writeTestMessage ($testFlag, $testMessage)
{
    if ($testFlag){
        Write-Host ""
        Write-Host $testMessage
    }
}


function SetScriptType ($scriptType) 
{ 
    $global:globalScriptType = $scriptType 
}


function SetScriptSubType ($scriptSubtype) 
{ 
    $global:globalScriptSubType = $scriptSubtype 
}


function SetScriptMessage ($scriptMessage) 
{ 
    $global:globalScriptMessage = $scriptMessage 
}


function DeleteFiles ($path, $ext, $days){
    
    #Write-Host 'DeleteFiles'

    #Write-Host 'path ' $path
    $ext = '*' + $ext
    #Write-Host 'Ext' $ext

    #----- get current date ----#
    $Now = Get-Date

    #----- define amount of days ----#
    #$Days = "7"

    #----- define folder where files are located ----#
    $SourceFolder = $path
    #Write-Host 'Folder' $SourceFolder 

    #----- define LastWriteTime parameter based on $Days ---#
    $LastWrite = $Now.AddDays(-1 * $Days)
    #Write-Host 'Date ' $LastWrite
 
    #----- get files based on lastwrite filter and specified folder ---#
    $Files = Get-Childitem $SourceFolder -Include $ext -Recurse | Where {$_.LastWriteTime -lt $LastWrite}

    foreach ($File in $Files)
        {
        if ($File -ne $NULL)
            {
                #write-host ("File to delete: " + $File)
                #----- delete the file from the archive folder ---#
                Remove-Item $File
                #write-host("")
            }
        }

}


function PrintMeOut ($Heading, $StuffToPrint)
{
    Write-Host '$Heading - $StuffToPrint'
}


function Format-DateTime ($datetime){
	if(([DBNull]::Value).Equals($datetime) -Or !($datetime)){
		""
	} else {
		Get-Date -Date $datetime -Format "dd-MMM-yy HH:mm"
	}
	
}


function Format-Boolean ($bool){
	if($bool){
		return "Yes"
	} else {
		return "No"
	}
	
}


## Updated 31/05/2023 Gordon F
## Check for Office 365 smtp server
function Send-Mail ($smtpServer,$smtpPort,$ssl,$user,$pwd,$from,$to,$subject,$body){

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;

    if ($smtpServer -eq "smtp.office365.com"){
        $EncryptedPasswordFile = "$directory\keys\EmailKey"
        $SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
        $EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $from,$SecureStringPassword
        $PSEmailServer = $smtpServer
        Send-MailMessage -UseSsl -From $from -To $to -Subject $subject -BodyAsHtml $body -Port $smtpPort -Credential $EmailCredential
    }
    else{
	    $SMTPClient = New-Object Net.Mail.SmtpClient($smtpServer,$smtpPort) 
	    $msg = New-Object Net.Mail.MailMessage
	
	    $msg.From = New-Object Net.Mail.MailAddress($from)
	    $msg.To.Add($to); 
	    $msg.IsBodyHTML = $true
	    $msg.Body = $body
	    $msg.Subject = $subject
 	
	    if($ssl){
		    $SMTPClient.EnableSsl = $true 
	    } else {
		    $SMTPClient.EnableSsl = $false 
	    }
	
	    if(![string]::IsNullOrEmpty($user)){
		    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($user,$pwd); 
	    }
		
        $SMTPClient.Send($msg)
    }
    
}


function Query-SQL ($instance,$query){
	
	$connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = "Server = $instance; Database = master; Integrated Security = True; Application Name = SQL Health Check;"
	$connection.Open()

	$command = New-Object System.Data.SqlClient.SqlCommand
	$command.CommandText = $query
	$command.Connection = $connection

    ## Modified 10/01/19 - GFF
    ## Added global variable for command timeouts.

    $command.CommandTimeout = $globalCommandTimeout

    $table = New-Object System.Data.DataTable
	    

    ## Modified 11/08/18 - GFF
    ## Capture query errors and return them to the calling function.
    try {	        
        $result = $command.ExecuteReader()
        $table.Load($result)
        $connection.Close()	
	    return $table	
    }

    catch {
        #Write-Host "Error in Query-SQL for $instance"
	    $ErrorMessage = $_.Exception.Message
        $ErrorMessage = "Error in Query-SQL for $instance`r`n" + @($ErrorMessage) + ": `r`n$query"
	    #Write-Host $ErrorMessage 
        #Write-Host $query
	    #Write-Error $ErrorMessage 
        #Write-Error $query
      
        $connection.Close()	
	    #return $table  
        throw $ErrorMessage
	
    }
    	
}


function Get-SQLLocalDateTime ($instance) {
	
	$server_date = $null
	
	Query-SQL $instance "SELECT GETDATE() AS server_date" | ForEach-Object {
    	$server_date = $_.server_date
	}
	
	return $server_date

}


function Get-SQLLocalDateMidnight ($instance) {
	
	$server_date = $null
	Query-SQL $instance "SELECT DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()), 0) AS server_date" | ForEach-Object {
    	$server_date = $_.server_date
	}
	
	return $server_date

}


function Test-SQLConnect ($instance){
	
	try {
		Query-SQL $instance "SELECT 1" | Out-Null
		return $true
	}
	
	catch {
		Write-Error -Message "User can't connect to SQL server $instance"
		return $false
	}
	
}


function Test-SQLPermissions ($instance){
	
	$(if($(Query-SQL $instance "SELECT IS_SRVROLEMEMBER ('sysadmin')")[0] -eq 1){
		return $true
	} 
	else {
		Write-Error -Message "User does not have a 'sysadmin' privelege on "+$instance
		return $false
	})
	
}


function Test-WMIConnect ($server){
	if($(Get-WmiObject -Class Win32_Service -Computer $server -Filter "Name='RemoteRegistry'" -ErrorAction SilentlyContinue)){
		return $true
	}
	else{
		Write-Error -Message "User can't connect to WMI service on $server"
		return $false
	}
}


function Get-SQLVersion ($instance){
	
	$version = Query-SQL $instance "SELECT SERVERPROPERTY('ProductVersion') AS version"
	
	$version = switch -wildcard ($version[0]) 
		{ 
			"8.*" {2000} 
			"9.*" {2005} 
			"10.0*" {2008} 
			"10.5*" {2008.5} 
			"11.*" {2012} 
			"12.*" {2014} 
			"13.*" {2016} 
			"15.*" {2017} 
			"16.*" {2019}

	}
	
	return $version

}


## Added 20/07/2018 IanH
## Used with the maintenance window setting
## Convert a day of the week to an integer
function DOW-ToInt ($day){
    
    ## If invalid day entered in the config then we want to flag this up
    $dayINT = 99 

    if ( ($day -eq 'Monday')    -Or ($day -eq 'Mon')   ) { $dayINT = 1 }
    if ( ($day -eq 'Tuesday')   -Or ($day -eq 'Tues') -Or ($day -eq 'Tue') )   { $dayINT = 2 }
    if ( ($day -eq 'Wednesday') -Or ($day -eq 'Wed')   ) { $dayINT = 3 }
    if ( ($day -eq 'Thursday')  -Or ($day -eq 'Thurs') ) { $dayINT = 4 }
    if ( ($day -eq 'Friday')    -Or ($day -eq 'Fri')   ) { $dayINT = 5 }
    if ( ($day -eq 'Saturday')  -Or ($day -eq 'Sat')   ) { $dayINT = 6 }
    if ( ($day -eq 'Sunday')    -Or ($day -eq 'Sun')   ) { $dayINT = 7 }

    return $dayINT
}


## Added 21/06/2023 GFF
## Retrieves SQL Server builds from an Azure REST API
function Get-SQLBuildsAPI ($apiBuild, $versionNo, $productVersion, $GDR, $directory){

    ## Get latest version to see if there are any updates available
        
    $apiPath = $apiBuild.Path
    $setTLS = $apiBuild.setTLS
    $setProxy = $apiBuild.setProxy

    <#
    Write-Host ("apiPath: $apiPath")
    Write-Host ("versionNo: $versionNo")
    Write-Host ("productVersion: $productVersion")
    Write-Host ("GDR: $GDR")
    Write-Host ("setTLS: $setTLS")
    Write-Host ("setProxy: $setProxy")
    Write-Host ("directory: $directory")
    #>
                
    $errorStatus = $false
    
    if ($setTLS -eq 1){
        ## Enable Tls 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        ## Enable Tls 1, 1.1 and 1.2
        #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    }
    
    $dest = $apiPath + "?Version=$versionNo"
    if ($GDR -eq 1){        
        $dest = $dest + "&GDROnly=1"
    }
    #Write-Host ("dest: $dest")
    #Write-Host ("")

    if ($setProxy -eq 1){
        $proxy = ([System.Net.WebRequest]::GetSystemWebproxy()).GetProxy($dest)
            
        try{
            Invoke-WebRequest -Uri $dest -Proxy $proxy -ProxyUseDefaultCredentials | ConvertFrom-Json | Out-File $directory\builds.json
        }

        catch [System.Net.WebException]{
            $errorStatus = $true
        }
    }
    else{
        try{

            Invoke-RestMethod -Uri $dest | Out-File $directory\builds.json
        }

        catch [System.Net.WebException]{
            $errorStatus = $true 
        }

    }

    $Error.Clear()

    if (!$errorStatus){
        $builds = Get-Content $directory\builds.json | Out-String | ConvertFrom-Json
        return $builds
    }   
}


## Modified 28/12/17 - Gordon Feeney
## Instance query returns version name as well as build
## Added replicas query to retun available AG replicas where appropriate

## Modified 31/10/18 - Gordon Feeney
## Added Edition to instance summary and tweaked version due to extraneous ).

## Modified 19/11/18 - Gordon Feeney
## Check for version in Get-SQLInstance should have been > 10.5 and not >= 10.5.	    

## Modified 19/12/18 - Gordon Feeney
## Fixed bug in Version string in Get-SQLInstance

## Modified 21/05/21 - Gordon Feeney
## Added new items to the summary and amended the overall query
                
## Instance Summary
function Get-SQLInstance ($instance, $version, $config, $flags, $production, $buildAPI, $directory){

    ## Modified 24/07/18 - IanH 
    ## Get the maintenance window settings from config.xml "maintenance" section
    $startDay = $config.startday
    $startTime = $config.starttime
    $endDay = $config.endday
    $endTime = $config.endtime

    $database_mail_error_days = 7
    $flags.flag | Where-Object {$_.name -eq "database_mail_error_days"} | ForEach {
        $database_mail_error_days = $_.value
    }

    
    ## For maintenance window - get today and yesterday as DOW to compare with values in config.xml    $todayDOW = ( (Get-Date).DayOfWeek ) 
    $yesterdayDOW = ( (Get-Date).AddDays(-1).DayOfWeek )

    ## Check if a maintenance window is defined in the config.xml
    if ( 
        ($startDay) -And ($startTime) -And ($endDay) -And ($endTime) 
       ) 
       {
            $maintWindow = 1
            $maintWindowNote = "Instance restarted during a maintenance window (" + $startDay + " " + "$startTime" + ":00" + " to " + $endDay + " " + $endTime + ":00) see config.xml"
       }
    else 
        {$maintWindow = 0} 

    
	$server_date = (Get-SQLLocalDateTime $instance)
	

    ## Modified 20/19/18 - Gordon Feeney
    #Fixed bug whereby search for SQL Agent service occassionally return false result.

    ## Modified 31/10/18 - Gordon Feeney
    ## Added Edition to instance summary and tweaked version due to extraneous ).

    ## Modified 19/11/18 - Gordon Feeney    
    ## Check for version in Get-SQLInstance should have been > 10.5 and not >= 10.5.

    ## Modified 19/12/18 - Gordon Feeney
    ## Fixed bug in Version string in Get-SQLInstance
	    
    $table = Query-SQL $instance "        
        DECLARE
	        @windows_server_version varchar(150), 
			@instance_name sysname, 
	        @instance_start datetime, 
	        @productVersion varchar(250), 
			@version varchar(250), 
			@versionNo varchar(4),
	        @gdr bit, 
	        @edition varchar(250), 
	        @is_clustered bit, 	
	        @ha_enabled bit, 
	        @sql_agent bit, 
	        @ram smallint, 
	        @max_server_memory int, 
	        @min_server_memory int, 
	        @maxdop tinyint, 
	        @ctfp smallint, 
	        @ad_hoc bit, 
	        @database_mail bit,
			@monitoring_installed datetime, 
			@CPUCount tinyint, 
			@CoreCount tinyint, 
			@CoresInUse tinyint,
			@ServerType int, 
            @EMailErrorDate datetime,
			@EMailErrors bit, 
			@EMailErrorMsg varchar(500);

		DECLARE @TSQL nvarchar(1000);
		
		-----------------------------------------------------------------------------------------------------------------------

		--Windows Server version and edition
		IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS varchar), 4) AS decimal(4, 1))) BETWEEN 9 AND 10
			--SQL Server 2005/2008
			BEGIN
				IF CHARINDEX('Windows NT', @@VERSION) > 0
					BEGIN
						SELECT @windows_server_version = 
							CASE RIGHT(SUBSTRING(@@VERSION, CHARINDEX('Windows NT', @@VERSION), 14), 3)
								WHEN '5.0' THEN 'Windows 2000'
								WHEN '5.1' THEN 'Windows XP'
								WHEN '5.2' THEN 'Windows Server 2003 or 2003 R2'
								WHEN '6.0' THEN 'Windows Server 2008'
								WHEN '6.1' THEN 'Windows Server 2008 R2'
								WHEN '6.2' THEN 'Windows Server 2012'
								WHEN '6.3' THEN 'Windows Server 2012 R2'
								WHEN '10.0' THEN 'Windows Server 2016 or Windows Server 2019'
								ELSE 'Not known'
							END 
					END
				 ELSE
					SELECT @windows_server_version = 'Not known'
				--END IF	
			END
		ELSE IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS varchar), 4) AS decimal(4, 1))) BETWEEN 10.5 AND 13
			--SQL Server 2008 R2 - SQL Server 2016
			BEGIN
				SET @TSQL = N'SELECT @Result = CASE windows_release
						WHEN ''5.0'' THEN ''Windows 2000''
						WHEN ''5.1'' THEN ''Windows XP''
						WHEN ''5.2'' THEN ''Windows Server 2003 or 2003 R2''
						WHEN ''6.0'' THEN ''Windows Server 2008 or Windows Vista''
						WHEN ''6.1'' THEN ''Windows Server 2008 R2''
						WHEN ''6.2'' THEN ''Windows Server 2012''
						WHEN ''6.3'' THEN ''Windows Server 2012 R2''
						WHEN ''10.0'' THEN ''Windows Server 2016 or Windows Server 2019''
                        ELSE ''Not known''
					END + 
					'' ('' + 
					CASE windows_sku
						WHEN ''4'' THEN ''Enterprise Edition''
						WHEN ''7'' THEN ''Standard Server Edition''
						WHEN ''8'' THEN ''Datacenter Server Edition''
						WHEN ''10'' THEN ''Enterprise Server Edition''
						WHEN ''48'' THEN ''Professional Edition''
						WHEN ''161'' THEN ''Pro for Workstations''
                        ELSE ''Not known''
					END + 
					'')''
				FROM sys.dm_os_windows_info;';

				EXEC sp_executesql 
					@query = @TSQL, 
					@params = N'@Result nvarchar(150) OUTPUT', 
					@Result = @windows_server_version OUTPUT;

			END
		ELSE IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS varchar), 4) AS decimal(4, 1))) >= 14
			--SQL Server 2017 and above
			BEGIN
				SET @TSQL = N'SELECT @Result = host_distribution FROM sys.dm_os_host_info';

				EXEC sp_executesql 
					@query = @TSQL, 
					@params = N'@Result nvarchar(150) OUTPUT', 
					@Result = @windows_server_version OUTPUT;
			END
		ELSE

			SET @windows_server_version = 'Not known'
		--END IF			

        -----------------------------------------------------------------------------------------------------------------------

		--Properties and config values
		SELECT @instance_name = @@SERVERNAME;

        IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') AS varchar), 4) AS decimal(5, 3))) > 10.5
			BEGIN
				SELECT @instance_start = CAST(last_startup_time AS datetime) FROM sys.dm_server_services WHERE servicename LIKE 'SQL Server (%';
				SELECT @sql_agent = CASE (SELECT status_desc FROM sys.dm_server_services WHERE servicename LIKE 'SQL Server Agent%') WHEN 'stopped' THEN 0 ELSE 1 END
			END
        ELSE
			BEGIN
				SELECT @instance_start = login_time FROM master.dbo.sysprocesses WHERE spid = 1;
				SELECT @sql_agent = CASE (SELECT COUNT(*) FROM master.dbo.sysprocesses WHERE program_name LIKE N'SQLAgent%') WHEN 0 THEN 0 ELSE 1 END
			END
        --END IF

		SELECT @productVersion = CAST(SERVERPROPERTY('ProductVersion') AS varchar(20));

		--SELECT @version = SUBSTRING ((SELECT @@VERSION), 1, CHARINDEX('(',(SELECT @@VERSION)) - 2) + ' (' + 
  --                       CAST(SERVERPROPERTY('ProductLevel') AS varchar) + ', build ' + 
  --                       CAST(SERVERPROPERTY('ProductVersion') AS varchar) + ')';

		SELECT @version = SUBSTRING (@@VERSION, 1, CHARINDEX('(',@@VERSION) - 2) + ' (' + 
			SUBSTRING (@@VERSION, CHARINDEX('(', @@VERSION) + 1, CHARINDEX(')', @@VERSION) - CHARINDEX('(', @@VERSION) - 1) + ', build ' + 
			CAST(SERVERPROPERTY('ProductVersion') AS varchar) + ')';

		--SELECT @versionNo = CAST(SERVERPROPERTY('ProductMajorVersion') AS varchar(4));
        IF CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS varchar), 4) AS numeric(5, 2)) < 10
	        --2005 and below
	        SELECT @versionNo = CAST(SERVERPROPERTY('ProductVersion') AS varchar(1))
        ELSE
	        IF CAST(LEFT(CAST(SERVERPROPERTY('ProductVersion') AS varchar), 4) AS numeric(5, 2)) = 10.5
		        --2008 R2
		        SELECT @versionNo = CAST(SERVERPROPERTY('ProductVersion') AS varchar(4))
	        ELSE
		        --2008 and 2012 and above
		        SELECT @versionNo = CAST(SERVERPROPERTY('ProductVersion') AS varchar(2))
	        --END IF
        --END IF

        IF PATINDEX('%GDR%', @@VERSION) > 0
			SET @gdr = 1
		ELSE
			SET @gdr = 0
		--END IF

	    SELECT @edition = CAST(SERVERPROPERTY('Edition') AS varchar);

        SELECT @is_clustered = CAST(SERVERPROPERTY('IsClustered') AS varchar);

        SELECT @ha_enabled = ISNULL(CAST(SERVERPROPERTY('IsHadrEnabled') AS bit), 0);

        SELECT @max_server_memory = CAST(value AS int) FROM sys.configurations WHERE name = 'max server memory (MB)';

        SELECT @min_server_memory = CAST(value AS int) FROM sys.configurations WHERE name = 'min server memory (MB)';

        SELECT @maxdop = CAST(value AS int) FROM sys.configurations WHERE name = 'max degree of parallelism';

        SELECT @ctfp = CAST(value AS int) FROM sys.configurations WHERE name = 'cost threshold for parallelism';

        SELECT @ad_hoc = CAST(value AS int) FROM sys.configurations WHERE name = 'optimize for ad hoc workloads';

		-----------------------------------------------------------------------------------------------------------------------

		--RAM. Throws an error in all version of SQL Server because of the difference in column names so the 
        --easiest way to deal with it is through dynamic SQL.
		
		IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') AS varchar), 4) AS decimal(5, 3))) > 10.5
			SET @TSQL = 'SELECT @Result = ROUND(physical_memory_kb / 1024.0 / 1024.0, 0) FROM sys.dm_os_sys_info'
		ELSE
			SET @TSQL = 'SELECT @Result = ROUND(physical_memory_in_bytes / 1024.0 / 1024.0 / 1024.0, 0) FROM sys.dm_os_sys_info'
		--END  IF

		EXEC sp_executesql 
			@query = @TSQL, 
			@params = N'@Result int OUTPUT', 
			@Result = @ram OUTPUT;

		-----------------------------------------------------------------------------------------------------------------------
       
		--Database Mail Profile
        DECLARE @res TABLE  
        (  
            Value VARCHAR(255)  
            , Data VARCHAR(255)  
        );  

        INSERT INTO @res  
        EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent', N'UseDatabaseMail';  
        INSERT INTO @res  
        EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', N'SOFTWARE\Microsoft\MSSQLServer\SQLServerAgent', N'DatabaseMailProfile';  

        IF (  
                SELECT COUNT(*)  
                FROM @res r  
                WHERE r.Value = 'UseDatabaseMail' AND r.Data = 1  
            ) = 1 AND   
            (  
                SELECT COUNT(*)  
                FROM @res r  
                WHERE r.Value = 'DatabaseMailProfile' AND r.Data IS NOT NULL  
            ) = 1  
	        SET @database_mail = 1 
        ELSE  
	        SET @database_mail = 0;

		-----------------------------------------------------------------------------------------------------------------------

        --Is monitoring installed?
		SET @monitoring_installed = NULL;
        
		IF EXISTS(SELECT * FROM sys.databases WHERE name = 'ProDBA')
			IF EXISTS(SELECT * FROM ProDBA.dbo.sysobjects WHERE name = 'DBFileStats')
				IF EXISTS (SELECT * FROM msdb.dbo.sysjobs WHERE name LIKE '(Pro-DBA)%')
					SET @monitoring_installed = (SELECT date_created FROM msdb.dbo.sysjobs WHERE name = '(Pro-DBA) DB Growth Monitor')
				--END IF
			--END IF
		--END IF

		-----------------------------------------------------------------------------------------------------------------------

        --Any index or statistics or index maintenance or databased integrity checks?
		DECLARE @index_maint_flag tinyint;
		DECLARE @stats_maint_flag tinyint;
		DECLARE @dbcc_flag tinyint;

        SELECT @index_maint_flag = 0, @stats_maint_flag = 0, @dbcc_flag = 0;


		SET @index_maint_flag = CASE WHEN NOT EXISTS(SELECT task_detail_id FROM msdb.dbo.sysmaintplan_logdetail WHERE line1 LIKE 'Reorganize index%' OR line1 LIKE 'Rebuild index%') THEN '1' ELSE '0' END;
		SET @stats_maint_flag = CASE WHEN NOT EXISTS(SELECT task_detail_id  FROM msdb.dbo.sysmaintplan_logdetail WHERE line1 LIKE 'Update Statistics%') THEN '1' ELSE '0' END;
		SET @dbcc_flag = CASE WHEN NOT EXISTS(SELECT task_detail_id  FROM msdb.dbo.sysmaintplan_logdetail WHERE line1 LIKE 'Check Database integrity%') THEN '1' ELSE '0' END;

		IF EXISTS(SELECT * from master.sys.tables WHERE name = 'CommandLog') AND EXISTS(SELECT * from msdb.dbo.sysjobs WHERE name LIKE 'IndexOptimize -%')
			SET @index_maint_flag = @index_maint_flag & CASE WHEN NOT EXISTS(SELECT TOP 1 CommandType FROM master.dbo.CommandLog WHERE CommandType IN ('ALTER_INDEX')) THEN '1' ELSE '0' END;
					
		IF EXISTS(SELECT * from master.sys.tables WHERE name = 'CommandLog') AND EXISTS(SELECT * from msdb.dbo.sysjobs WHERE name LIKE 'IndexOptimize -%')
			SET @stats_maint_flag = @stats_maint_flag & CASE WHEN NOT EXISTS(SELECT TOP 1 CommandType FROM master.dbo.CommandLog WHERE CommandType IN ('UPDATE_STATISTICS')) THEN '1' ELSE '0' END;
					
		IF EXISTS(SELECT * from master.sys.tables WHERE name = 'CommandLog') AND EXISTS(SELECT * from msdb.dbo.sysjobs WHERE name LIKE 'DatabaseIntegrityCheck -%')
			SET @dbcc_flag = @dbcc_flag & CASE WHEN NOT EXISTS(SELECT TOP 1 CommandType FROM master.dbo.CommandLog WHERE CommandType IN ('DBCC_CHECKDB')) THEN '1' ELSE '0' END;
		
		IF EXISTS(SELECT * from msdb.dbo.sysjobs WHERE name LIKE '(Pro-DBA) Index Defrag%' AND enabled = 1)
			SET @index_maint_flag = '0';
							
		-----------------------------------------------------------------------------------------------------------------------

		IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') AS varchar), 4) AS decimal(5, 3))) > 10
			SET @TSQL = N'SELECT @Result = virtual_machine_type FROM sys.dm_os_sys_info'
		ELSE
			SELECT @TSQL = N'SELECT @Result = -1';
		--END IF

		EXEC sp_executesql 
			@query = @TSQL, 
			@params = N'@Result int OUTPUT', 
			@Result = @ServerType OUTPUT;

		SELECT @CPUCount = CASE hyperthread_ratio WHEN 0 THEN cpu_count ELSE cpu_count/hyperthread_ratio END FROM sys.dm_os_sys_info
		SELECT @CoreCount = COUNT(*) FROM sys.dm_os_schedulers AS CoreCount WHERE scheduler_id < 255 
		SELECT @CoresInUse = COUNT(*) FROM sys.dm_os_schedulers AS CoresInUse WHERE scheduler_id < 255 AND is_online = 1
		
		-----------------------------------------------------------------------------------------------------------------------

        --Database mail errors
		SELECT TOP 1
			@EMailErrorDate = ai.sent_date,
			@EMailErrorMsg = el.[description]
		FROM
			msdb.dbo.sysmail_allitems ai 
			INNER JOIN msdb.dbo.sysmail_profile p ON p.profile_id = ai.profile_id 
			LEFT OUTER JOIN msdb.dbo.sysmail_event_log AS el ON ai.mailitem_id = el.mailitem_id 
		WHERE 1 = 1
			AND p.name =  'ProDBA'
            AND ai.sent_status = 'failed'
			AND ai.sent_date >= DATEADD(DAY, -$database_mail_error_days, GETDATE())
		ORDER BY  ai.sent_date DESC;

		 IF @EMailErrorDate IS NULL
            SELECT @EMailErrors = 0
        ELSE
            SET @EMailErrors = 1;
        --END IF

		-----------------------------------------------------------------------------------------------------------------------

        SELECT 
			@windows_server_version AS windows_server_version, 
	        @instance_name AS instance_name, 
	        @instance_start AS instance_start, 
			@productVersion AS productVersion,
	        @version AS [version], 
			@versionNo AS [versionNo], 
	        @gdr AS gdr,
	        @edition AS edition, 
	        @is_clustered As is_clustered, 
	        @ha_enabled AS ha_enabled, 
	        @sql_agent AS sql_agent, 
	        @ram AS ram, 
	        @max_server_memory AS max_server_memory, 
	        @min_server_memory AS min_server_memory, 
	        @maxdop AS [maxdop], 
	        @ctfp AS ctfp, 
	        @ad_hoc AS ad_hoc, 
	        @database_mail AS mail_profile_enabled,
			@monitoring_installed AS monitoring_installed, 
			@index_maint_flag AS index_maint_flag, 
			@stats_maint_flag AS stats_maint_flag, 
			@dbcc_flag AS dbcc_flag, 
			@ServerType AS server_type, 
			@CPUCount As cpu_count, 
			@CoreCount AS core_count, 
			@CoresInUse AS cores_in_use,
			@EMailErrorDate AS email_error_date,
			CASE WHEN LEN(@EMailErrorMsg) >= 500 THEN @EMailErrorMsg + ' .....' ELSE @EMailErrorMsg END AS email_error_message,
            @EMailErrors AS email_errors;
	"
	
	$table | ForEach-Object {
	    $windows_server_version = $_.windows_server_version
		$instance_name = $_.instance_name
		$instance_start = $_.instance_start
        $productVersion = $_.productVersion
		$version = $_.version
        $versionNo = $_.versionNo
        $GDR = $_.gdr
        $edition = $_.edition
		$is_clustered = [bool]$_.is_clustered
		$ha_enabled = [bool]$_.ha_enabled
        $sql_agent = [bool]$_.sql_agent
        $ram = $_.RAM
        $max_server_memory = $_.max_server_memory
        $min_server_memory = $_.min_server_memory
        $maxdop = $_.maxdop
        $ctfp = $_.ctfp
        $ad_hoc = $_.ad_hoc
        $mail_profile_enabled = $_.mail_profile_enabled
        $monitoring_installed = $_.monitoring_installed
        $index_maint_flag = $_.index_maint_flag
		$stats_maint_flag = $_.stats_maint_flag
		$dbcc_flag = $_.dbcc_flag
        $server_type = $_.server_type
        $cpu_count = $_.cpu_count
        $core_count = $_.core_count
        $cores_in_use = $_.cores_in_use
        $email_error_date = $_.email_error_date
        $email_error_message = $_.email_error_message
        $email_errors = $_.email_errors

        ## Get any flags
        $ignore_index_maint_alert = "0"
        $ignore_stats_maint_alert = "0"
        $ignore_dbcc_alert = "0"
        $ignore_version_check = "0"
        $ignore_database_mail_error_check = "0"
                            
        $flags.flag | ForEach {
            if ($_.name -eq "ignore_index_maint_alert"){
                $ignore_index_maint_alert = $_.Value
            }
            elseif ($_.name -eq "ignore_stats_maint_alert"){
                $ignore_stats_maint_alert = $_.Value
            }
            elseif ($_.name -eq "ignore_dbcc_alert"){
                $ignore_index_dbcc_alert = $_.Value
            }
            elseif ($_.name -eq "ignore_version_check"){
                $ignore_version_check = $_.Value
            }
            elseif ($_.name -eq "ignore_database_mail_error_check"){
                $ignore_database_mail_error_check = $_.Value
            }
            elseif ($_.name -eq "database_mail_error_days"){
                $database_mail_error_days = $_.Value
            }
            
	    }

        
        if (($index_maint_flag -eq "1") -and ($ignore_index_maint_alert -eq "1")){
            $index_maint_flag = "2"
        }

        if (($stats_maint_flag -eq "1") -and ($ignore_stats_maint_alert -eq "1")){
            $stats_maint_flag = "2"
        }
        
        if (($dbcc_flag -eq "1") -and ($ignore_index_dbcc_alert -eq "1")){
            $dbcc_flag = "2"
        }                
        
        if (($email_errors -eq "1") -and ($ignore_database_mail_error_check -eq "1")){
            $email_errors = "0"
        }
        
                
        if ($is_clustered){

			$nodes = Query-SQL $instance "
			SELECT 
			    [NodeName] AS [node]
			    ,CASE [NodeName] 
			        WHEN (SELECT SERVERPROPERTY('ComputerNamePhysicalNetBIOS')) THEN 1
			        ELSE 0
			    END AS [is_current_owner]
			FROM sys.dm_os_cluster_nodes"
			
			$nodesarray = $nodes | ForEach-Object {
			
				New-Object PSObject -Property @{
					Node = $_.node
					Owner = [bool]$_.is_current_owner
				}
				
			}
        }

        ## Added distinct to replica query to avoid duplication - v2.30 IanH

        $replica_array = @()

        if ($ha_enabled){

			$replicas = Query-SQL $instance "
			DECLARE @SQL nvarchar(max);

            IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') as varchar), 4) AS DECIMAL(5, 3))) >= 11
	            SET @SQL = '
		            SELECT 
			            distinct UPPER(ISNULL(ar.replica_server_name, '''')) AS replica, 
			            CASE 
				            WHEN ISNULL(ags.primary_replica, '''') = ar.replica_server_name THEN 1
				            ELSE 0
			            END AS is_primary_replica,
                        UPPER(ISNULL(ag.name,'''')) as agname
		            FROM 
			            sys.availability_replicas ar
				            LEFT OUTER JOIN 
			            sys.dm_hadr_availability_group_states ags ON ar.replica_server_name = ags.primary_replica AND ar.group_id = ags.group_id
                        join master.sys.availability_groups ag on ar.group_id = ag.group_id
                    order by agname , replica'
            ELSE
	            SET @SQL = '
		            SELECT NULL AS replica_server_name, NULL AS is_primary_replica, NULL AS ag_name'
            --END IF

            EXEC sp_executesql @SQL;
            "
			
			$replica_array = $replicas | ForEach-Object {
			
                New-Object PSObject -Property @{
					Replica = $_.replica
					IsPrimaryReplica = [bool]$_.is_primary_replica
                    AGName = $_.agname 
				}
				
			}
			
		}
<# REMOVE 

$AGLoopCounter = [int]0
$AGNewNameFlag = [int]1
$AGNameTemp = $null

 $AGNameList = @()
    $replica_array | ForEach {
        
        if ($AGLoopCounter -gt 0) 
        {   
            $AGNameTemp = $_.agname 
            $AGNameList | ForEach {
                if($_.agname -eq $AGNameTemp){$AGNewNameFlag = 0}
            }
        }
        if ($AGLoopCounter -eq 0 -or $AGNewNameFlag -eq 1)
        {
            $AGNameListRow = New-Object PSObject -Property @{
					agname = $_.agname
                }
            $AGNameList += $AGNameListRow
        }
        $AGLoopCounter++
        $AGNewNameFlag = 1
    }

## Global variable - i.e. used in the database section     
$AGNumber =  $AGNameList.Count 
#>
   

## Modified 20/07/18 IanH
## UptimeAlert (which flags if an instance has been restarted in the last 24 hours)
## modified to include logic to check for a maintenance window defined in the config.xml
## If it occurs during a maintenance window then the restart is classed as a headsup in
## FormatHTML-SQLInstance (UptimeAlert = 2) 

        if ($ignore_version_check -eq "1"){
            ## We can bypass the check for version updates
            $versionAvailability = -1
        }
        else{
            $builds = Get-SQLBuildsAPI $buildAPI $versionNo $productVersion $GDR $directory
            if ($builds.Count -gt 0){

                #Write-Host "Builds returned"

                $build = $builds | Select-Object -First 1
                $buildVersion = $build.Build
                $buildDesc = $build.Description
            
                <#
                0: Up-to-date
                1: Version available
                2: Build table not up-to-date
                3: Anomaly in build table
                4: Can't access build table
                #>
            
                # Replace dots in build/version numbers as SQL Server 2005 has a slightly different product version number to subsequent versions.
                $buildVersionRepl = $buildVersion.Replace('.', '') 
                $productVersionRepl = $productVersion.Replace('.', '') 

                if ([string]$buildVersionRepl -eq [string]$productVersionRepl){
                    $versionAvailability = 0
                }
                elseif([string]$buildVersionRepl -gt [string]$productVersionRepl){
                    $versionAvailability = 1
                }
                elseif([string]$buildVersionRepl -lt [string]$productVersionRepl){
                    $versionAvailability = 2
                }
                else{
                    $versionAvailability = 3
                }
            }
            else{
                $versionAvailability = 4
            }
        }
        
        New-Object PSObject -Property @{
            WindowsServerVersion = $windows_server_version
			InstanceName = $instance_name
			ServerDate = $server_date
			InstanceStart = $instance_start
			Version = $version
            VersionNo = $versionNo
            VersionAvailability = $versionAvailability
            LatestVersion = $buildVersion
            LatestVersionDesc = $buildDesc            
            Edition = $edition
			IsClustered = $is_clustered
			Nodes = $nodesarray
            IsHADREnabled = $ha_enabled
            Replicas = $replica_array			
			SQLAgent = $sql_agent
            RAM = $ram
            MaxServerMemory = $max_server_memory
            MinServerMemory = $min_server_memory
            MAXDOP = $maxdop
            CTFP = $ctfp
            AdHoc = $ad_hoc
            MailProfileEnabled = $mail_profile_enabled
            MaintWindow = $maintWindow 
            MaintWindowNote = $maintWindowNote
            MonitoringInstalled = $monitoring_installed
            IndexMaintAlert = $index_maint_flag
            StatsMaintAlert = $stats_maint_flag
            DBCCAlert = $dbcc_flag
            ServerType = $server_type
            CPUCount = $cpu_count
            CoreCount = $core_count
            CoresInUse = $cores_in_use
            EmailErrorDate = $email_error_date
            EmailErrorMsg = $email_error_message
            EmailErrors = $email_errors
		} |  Select-Object `
		WindowsServerVersion `
        ,InstanceName `
		,ServerDate `
		,InstanceStart `
		,Version `
        ,VersionNo `
        ,VersionAvailability `
        ,LatestVersion `
		,LatestVersionDesc `
		,Edition `
        ,IsClustered `
		,Nodes `
        ,IsHADREnabled `
        ,Replicas `
		,SQLAgent `
        ,RAM `
        ,MaxServerMemory `
        ,MinServerMemory `
        ,MAXDOP `
        ,CTFP `
        ,AdHoc `
        ,MailProfileEnabled `
        ,MaintWindow `
        ,MaintWindowNote `
        ,MonitoringInstalled `
		,IndexMaintAlert `
        ,StatsMaintAlert `
        ,DBCCAlert `
        ,ServerType `
        ,CPUCount `
        ,CoreCount `
        ,CoresInUse `
        ,EmailErrorDate `
        ,EmailErrorMsg `
        ,EmailErrors `
		,@{Name="UptimeAlert";Expression=
		{ 

        if($_.InstanceStart -gt ($_.ServerDate.AddDays(-1)) )
            {
                if ($maintWindow -eq 1) ## if a mainteance window defined check if restart fell within it
                {
                    $todayInt = (Get-Date).DayOfWeek.value__
                    $startDayInt = DOW-ToInt $startDay
                    $endDayInt = DOW-ToInt $endDay

                 ## Check the start and end times are valid integers
                try
                {
                    $startTimeInt = [int]$startTime
                }
                catch  ## Something wrong with the starttime in the config.xml so set to 99
                {
                    $startTimeInt = 99
                }

                try  
                {
                    $endTimeInt = [int]$endTime
                }
                catch   ## Something wrong with the endtime in the config.xml so set to 99
                {
                    $endTimeInt = 99 
                } 
    
                if ( `
                        ($startTime.Length -eq 2 ) -And ($endTime.Length -eq 2) `
                        -And ($startTimeInt -ge 0) -And ($startTimeInt -le 23)  `
                        -And ($endTimeInt -ge 0) -And ($endTimeInt -le 23) `
                        -And ($startDayInt -ne 99) -And ($endDayInt -ne 99)`
                   )
                     {
                       ## Calculate the date for the current week using today's date as your basis
                       ## Use the DOW as an integer to add or subtract the correct number of days. 
                       if ( $todayInt -ge $startDayInt) 
                       {
                           $daysDiff = $todayInt - $startDayInt
                           $daysDiff = $daysDiff - ($daysDiff * 2) 
                       }
                       if ( $todayInt -lt $startDayInt) 
                       {
                           $daysDiff = $todayInt - $startDayInt + 7
                       }
                       [datetime]$startDT = Get-Date  $((Get-Date).AddDays($daysDiff)) -Hour $startTime -Minute 0 -Second 0

                       if ( $todayInt -ge $endDayInt) 
                       {
                           $daysDiff = $todayInt - $endDayInt
                           $daysDiff = $daysDiff - ($daysDiff * 2) 
                       }
                       if ( $todayInt -lt $endDayInt) 
                       {
                           $daysDiff = $todayInt - $endDayInt + 7
                       }
                       [datetime]$endDT = Get-Date  $((Get-Date).AddDays($daysDiff)) -Hour $endTime -Minute 0 -Second 0


                       if ( ( $_.InstanceStart -gt $startDT) -And ( $_.InstanceStart -lt $endDT ) )
                       {
                            [int]2     ## Inside maintenance window
                       }
                       else
                       {
                            [int]1    ## Outside maintenance window
                       }
                    }
                    else  
                    {
                        [int]1      ## Invalid day values entered in config.xml so treat as no maintenance window
                    }

                }
                else
                {
                         [int]1    ## No maintance window defined so same as outside maintenance window
                }
            }
            else{
				    [int]0       ## No restart in the last 24 hours
            }

	    }  ## End of UptimeAlert expression
		} `
        ,@{Name="MonitoringInstalledHeadsup";Expression=
			
			{ 
			if(!([string]::IsNullOrEmpty($_.MonitoringInstalled))){
				[bool]0
			}
			
			elseif($_.Edition -like "Express Edition*" -or $production -eq 0){
				[bool]1
			}

            else{
                [bool]0
            }
			}
					
		}`
		,@{Name="AgentAlert";Expression=
			
			{ 
			
			if($_.SQLAgent -or $_.Edition -like "Express Edition*"){
				[bool]0
			}
			
			else{
				[bool]1
			}
			}
		} `
		,@{Name="AgentHeadsup";Expression=
			
			{ 
			
			if($_.SQLAgent){
				[bool]0
			}
			elseif($_.Edition -like "Express Edition*"){
                [bool]1
            }
			else{
				[bool]0
			}
			}
		} `
        ,@{Name="MonitoringInstalledAlert";Expression=
			
			{ 
			
			if(!([string]::IsNullOrEmpty($_.MonitoringInstalled)) -or $_.Edition -like "Express Edition*" -or $production -eq 0){
				[bool]0
			}
			
			else{
				[bool]1
			}
			}
					
		}`
        ,@{Name="MailProfileAlert";Expression=
			
			{ 
			
			if($_.MailProfileEnabled -or $_.Edition -like "Express Edition*" -or $production -eq 0){
				[bool]0
			}
			
			else{
				[bool]1
			}
			}
					
		}`
        ,@{Name="MailProfileHeadsup";Expression=
			
			{ 
			
			if($_.MailProfileEnabled ){
				[bool]0
			}
			
			elseif($_.Edition -like "Express Edition*" -or $production -eq 0){
				[bool]1
			}
            else{
				[bool]0
			}
			}
		 }`
        ,@{Name="ExpressEditionAlert";Expression=
			
			{ 
			
			if($_.Edition -like "Express Edition*"){
				[bool]1
			}
			
			else{
				[bool]0
			}
		}
		}`
        ,@{Name="NonProductionAlert";Expression=
			
			{ 
			
			if($production -eq 0){
				[bool]1
			}
			else{
				[bool]0
			}
		}
		}`
        ,@{Name="MaintenanceAlert";Expression=
			
			{ 			
			    if(($_.IndexMaintAlert -eq "1") -or ($_.StatsMaintAlert -eq "1") -or ($_.DBCCAlert -eq "1")){
				    #[bool]1
                    1
			    }
                if(($_.IndexMaintAlert -eq "2") -and ($_.StatsMaintAlert -eq "2") -and ($_.DBCCAlert -eq "2")){
                    2
                }
			    else{
				    #[bool]0
                    0
			    }
		    }
		}`
        ,@{Name="CoreHeadsUp";Expression=
			
			{ 
			
			if($cores_in_use -ne $core_count){
				[bool]1
			}
			else{
				[bool]0
			}
		}
		}`
	}     ## End :- $table | ForEach-Object {	    

    #Write-Host "------------------------------------------------------------------------------------------------------"
    #Write-Host ""
}   ## End Get-SQLInstance


## IanH 11/11/2020
## Turns "messy" DOW (thurs or Thurs or THURS) into nice tidy version (Thursday)
function DOWTidyUp ($day) 
{
    $tidyDOW = 'invalid' 

    if ( ($day -eq 'Monday')    -Or ($day -eq 'Mon') ) { $tidyDOW = 'Monday' }
    if ( ($day -eq 'Tuesday')   -Or ($day -eq 'Tues') -Or ($day -eq 'Tue') )   { $tidyDOW = 'Tuesday' }
    if ( ($day -eq 'Wednesday') -Or ($day -eq 'Wed')   ) { $tidyDOW= 'Wednesday' }
    if ( ($day -eq 'Thursday')  -Or ($day -eq 'Thurs') -Or ($day -eq 'Thur')) { $tidyDOW = 'Thursday' }
    if ( ($day -eq 'Friday')    -Or ($day -eq 'Fri')   ) { $tidyDOW = 'Friday' }
    if ( ($day -eq 'Saturday')  -Or ($day -eq 'Sat')   ) { $tidyDOW = 'Saturday' }
    if ( ($day -eq 'Sunday')    -Or ($day -eq 'Sun')   ) { $tidyDOW = 'Sunday' }

    return $tidyDOW
}




## IanH 11/11/2020
## Some clients only run backups at the weekend, or only on weekdays. This bit allows us to specify that we only check 
## backups on certain days, so avoid always having backups shown as failed on a Monday (or whenever). 
## Verify that the values set in config.xml for a backup check day range are valid
## If they are then see if today falls in that range 

function BackupsCheckedToday ($backupdaystart, $backupdayend) 
{

            ## by default we check backups 
            $backupCheckedTodayFlag = 1
            
            ## turn the DOW to an int
            [int]$backupdaystartInt = DOW-ToInt ($backupdaystart)
            [int]$backupdayendInt = DOW-ToInt ($backupdayend)

            ## Bug fix : DOW-ToInt sets Sunday to 7. Powershell's Get-Date sets Sunday to 0. Will use Sunday = 0 here

            if($backupdaystartInt -eq 7) { $backupdaystartInt = 0 }
            if($backupdayendInt   -eq 7) { $backupdayendInt  = 0 }

            
            ## if config.xml settings are valid, see if today is a day on which we check the backups
            if($backupdaystartInt -ne 99 -and $backupdayendInt -ne 99)
            {

                ## Get today                 
                [Int]$todayDOW = (Get-Date).DayOfWeek.value__

                ## Get range of days between start and end 
                if($backupdaystartInt -le $backupdayendInt) {[int[]]$dowArray = $backupdaystartInt .. $backupdayendInt }
                if($backupdaystartInt -gt $backupdayendInt) 
                {  
                  [int[]]$dowArray =  0 .. $backupdayendInt
                  $dowArray += $backupdaystartInt .. 6
                } 

                ## If today is not included in the range of days in which we check backups then we don't check backups
                if($dowArray.Contains($todayDOW) -eq $false ) 
                { 
                    $backupCheckedTodayFlag = 0 
                }
            }
            return $backupCheckedTodayFlag
}





## Modified 29/12/17 - Gordon Feeney - Added AG replicas section

## Modified 21/05/21 - Gordon Feeney
## Added new items to the summary
                		
function FormatHTML-SQLInstance ($obj){
	

	$html = "
	<table class='summary'>
	<tr>
		<th>Property</th> 
		<th>Value</th>
	</tr>
	"

## Modified 26/07/18 IanH
## Variables used to determine if a maintenance window, Mail Profile and Monitoring Installed notes should appear
    $MaintWindowNoteFlag = $false
    $MailProfileAlertFlag = $false 
    $MonitoringNotInstalledAlertFlag = $false 
    $ExpressEditionAlertFlag = $false 
    $NonProductionAlertFlag = $false 
 
## Modified 20/07/18 - Ian H
## Added "headsup" to instance restarts where this occurs during a defined maintenance window	
    $alert = $false
    	
    $obj | ForEach {

        ## Modified 26/07/18 IanH
        ## Check if an instance restart occurred during a maintenance window
        ## if it has we want a note below our Instance table
        if($_.UptimeAlert -eq 2) 
        { 
            $MaintWindowNoteFlag = $true
            $MaintWindowNote = $_.MaintWindowNote
        }

        if($_.UptimeAlert -eq 1) 
        { 
            $InstanceRestartAlertFlag = $true
        }

        if($_.MailProfileAlert -eq 1)
        {
            $MailProfileAlertFlag = $true 
        }

        if($_.MonitoringInstalledAlert -eq 1)
        {
            $MonitoringNotInstalledAlertFlag = $true
        }
        if($_.ExpressEditionAlert -eq 1)
        {
            $ExpressEditionAlertFlag = $true
        }
        if($_.NonProductionAlert -eq 1)
        {
            $NonProductionAlertFlag = $true
        }
 
		$IndexMaintAlert = $_.IndexMaintAlert
        $StatsMaintAlert = $_.StatsMaintAlert
        $DBCCAlert = $_.DBCCAlert
        $MaintAlert = ($IndexMaintAlert -or $StatsMaintAlert -or $DBCCAlert)


        <#
        0: Up-to-date
        1: Version available
        2: Build table not up-to-date
        3: Anomaly in build table
        4: Can't access build table
        #>
        $versionAvailability = $_.versionAvailability
        $latestVersion = $_.latestVersion
        $latestVersionDesc = $_.latestVersionDesc

        $emailErrorsAlert = if($_.EmailErrors -eq 1){$true} else {$false}
        if ($emailErrorsAlert -eq $true){
            $emailErrorDate = $_.EmailErrorDate
            $emailErrorMsg = $_.EmailErrorMsg
            $emailErrorDate = $emailErrorDate.ToString("dd/MM/yyyy") + " at " + $emailErrorDate.ToString("HH:mm:ss")
        }
        
        
        if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){

            $html += "
		    <tr><td>Windows Server</td><td>"+$_.WindowsServerVersion+"</td></tr>
		    <tr><td>SQL Server</td><td>"+$_.InstanceName+"</td></tr>
            <tr><td>Server Date</td><td>"+(Format-DateTime $_.ServerDate)+"</td></tr>
		    <tr"+$(if($_.UptimeAlert -eq 1){" class='headsup'"} elseif($_.UptimeAlert -eq 2){" class='headsup'"})+"><td>Server Start</td><td>"+(Format-DateTime $_.InstanceStart)+"</td></tr>
		    <tr"+$(if($_.versionAvailability -eq 1){" class='headsup'"} elseif($_.versionAvailability -gt 1){" class='warning'"})+"><td>Version</td><td>"+$_.Version+"</td></tr>
            <tr><td>Edition</td><td>"+$_.Edition+"</td></tr>
            
            <tr><td>Server Type</td><td>"+ $(if($_.ServerType -eq 0){'Physical'} elseif($_.ServerType -eq 1 -or $_.ServerType -eq 2){'Virtual'} else{'Not known'}) + "</td></tr>

		    <tr><td>Clustered</td><td>"+(Format-Boolean $_.IsClustered)+"</td></tr>
		    "+$(if($_.IsClustered){
		
			    $firstRow = $true
			
			    $_.Nodes | ForEach {
				    "<tr><td>"+$(if($firstRow){"Nodes"} else {""})+"</td><td>"+$(if($_.Owner){"<b>"+$_.Node+"</b>"} else {$_.Node})+"</td></tr>"
				    $firstRow = $false
			    }
			
		
		    })+"
            <tr><td>Always On</td><td>"+(Format-Boolean $_.IsHADREnabled)+"</td></tr>
            "+$(if($_.IsHADREnabled){
		
			    $firstRow = $true
			
			    $_.Replicas | ForEach {
				    "<tr><td>"+$(if($firstRow){"Replicas"} else {""})+"</td><td>"+$(if($_.IsPrimaryReplica){"<b>"+$_.Replica+ " (" + $_.AGName + ")" + "</b>"} else {$_.Replica + " (" + $_.AGName + ")" }) +  "</td></tr>"
				    $firstRow = $false
			    }					
		    })+"
		    <tr"+$(if($_.AgentAlert){" class='warning'"}elseif($_.AgentHeadsup){" class='headsup'"})+"><td>SQL Agent</td><td>"+$(if(($_.SQLAgent)){"Running"} else {"Not Running"})+"</td></tr>
            <tr><td>RAM</td><td>"+$_.RAM + "GB</td></tr>
            <tr><td>Max Server Memory</td><td>"+$_.MaxServerMemory+"MB</td></tr>
            <tr><td>Min Server Memory</td><td>"+$_.MinServerMemory+"MB</td></tr>
            <tr><td>Max Degree of Parallelism</td><td>"+$_.MAXDOP+"</td></tr>
            <tr><td>Cost Threshold for Parallelism</td><td>"+$_.CTFP+"</td></tr>
            <tr><td>CPU Count</td><td>"+$_.CPUCount+"</td></tr>
            <tr><td>Core Count</td><td>"+$_.CoreCount+"</td></tr>
            <tr"+$(if($_.CoresInUse -ne $_.CoreCount){" class='headsup'"})+"><td>Cores in Use</td><td>"+$_.CoresInUse+"</td></tr>            
            <tr><td>Optimise For Ad Hoc</td><td>"+ $(if($_.AdHoc -eq 0){'No'} else{'Yes'}) + "</td></tr>
            <tr"+$(if($_.MailProfileAlert){" class='warning'"}elseif($_.MailProfileHeadsup){" class='headsup'"})+"><td>SQL Agent Mail Profile Enabled</td><td>"+$(if(($_.MailProfileEnabled)){"Yes"} else {"No"})+"</td></tr>
            <tr"+$(if($_.MonitoringInstalledAlert -eq 1){" class='warning'"}elseif($_.MonitoringInstalledHeadsup){" class='headsup'"})+"><td>Monitoring Installed</td><td>"+$(if([string]::IsNullOrEmpty($_.MonitoringInstalled)){"No"} else {$_.MonitoringInstalled.toString("dd-MMM-yyyy")})+"</td></tr>
            <tr"+$(if($_.IndexMaintAlert -eq 1){" class='headsup'"})+"><td>Index maintenance</td><td>"+ $(if($_.IndexMaintAlert -eq 1){'No'} elseif($_.IndexMaintAlert -eq 2){'N/A'} else{'Yes'}) + "</td></tr>
            <tr"+$(if($_.StatsMaintAlert -eq 1){" class='headsup'"})+"><td>Statistics updates</td><td>"+ $(if($_.StatsMaintAlert -eq 1){'No'} elseif($_.StatsMaintAlert -eq 2){'N/A'} else{'Yes'}) + "</td></tr>
            <tr"+$(if($_.DBCCAlert -eq 1){" class='headsup'"})+"><td>Database integrity checks</td><td>"+ $(if($_.DBCCAlert -eq 1){'No'} elseif($_.DBCCAlert -eq 2){'N/A'} else{'Yes'}) + "</td></tr>
            <tr"+$(if($emailErrorsAlert){" class='warning'"})+"><td>Database Mail</td><td>"+ $(if($emailErrorsAlert){'Errors'} else{'OK'}) + "</td></tr>
		    "		
        }
        else{
            #if (($_.UptimeAlert -eq 1) -or ($_.UptimeAlert -eq 2)){
            if ($_.UptimeAlert -eq 1){
                $alert = $true
                $html += "
                <tr"+$(if($_.UptimeAlert -eq 1){" class='headsup'"} elseif($_.UptimeAlert -eq 2){" class='headsup'"})+"><td>Server Start</td><td>"+(Format-DateTime $_.InstanceStart)+"</td></tr>"
            } 
            if($_.AgentAlert){
                $alert = $true
                $html += "
                <tr"+$(if($_.AgentAlert){" class='warning'"})+"><td>SQL Agent</td><td>"+$(if(($_.SQLAgent)){"Running"} else {"Not Running"})+"</td></tr>"
            }
            
            <#
            if (!$alert ){
                $html += "
                <tr><td colspan='2'>No instance issues</td></tr>"
            } 
            #>   
        }
        
	}
    	
	$html += "</table>"

## Modified 26/07/18 IanH
## Add note if restart during a maintenance window
	<#
    if($MaintWindowNoteFlag){

    		$html += "<div class='headsupnote'><span>"
		    $html+="Notes"
            $html += "</span><table class='disabled'>"

            $html += "<tr class='headsup'>NOTE: " + $MaintWindowNote + "</div>"
	}
#>	

    

         if($MaintWindowNoteFlag -or $MailProfileAlertFlag -or $MonitoringNotInstalledAlertFlag -or $InstanceRestartAlertFlag -or $ExpressEditionAlertFlag -or $NonProductionAlertFlag -or $MaintAlert -or ($versionAvailability -gt 0) -or $emailErrorsAlert){
            $html += "<div class='headsupnote'><span>"
		    $html+="<b>Notes</b>"
            $html += "</span><table class='notenormal'>"
            if($MaintWindowNoteFlag) {
                $html += "<tr class='headsup'>"
		        $html+="<td>" + $MaintWindowNote  + "</td>"
		        $html += "</tr>"
            }
            if($InstanceRestartAlertFlag){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Instance restarted in the last 24 hours"  + "</td>"
		        $html += "</tr>"
            }
            if($versionAvailability -gt 0){
                $html += "<tr class='headsup'>"
                if ($versionAvailability -eq 1){
                    $html+="<td>" + "Version " + $latestVersion + " ($latestVersionDesc)" + " is available"  + "</td>"
                }
                elseif ($versionAvailability -eq 2){
                    $html+="<td style='color: red'>" + "The version data source is not up-to-date. Please inform the SQL Support Team."  + "</td>"
                }
		        elseif ($versionAvailability -gt 2){
                    $html+="<td style='color: red'>" + "Can't retrieve SQL Server versions data. Please inform the SQL Support Team."  + "</td>"
                }
		        $html += "</tr>"
            }
            if($MailProfileAlertFlag){
             	$html += "<tr class='headsup'>"
		        $html+="<td>" + "Mail Profile Not Enabled : ProDBA should investigate before emailing the client"  + "</td>"
		        $html += "</tr>"
            }
            if($MonitoringNotInstalledAlertFlag){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Monitoring not Installed : ProDBA should investigate before emailing the client"  + "</td>"
		        $html += "</tr>"
            }
            if($ExpressEditionAlertFlag){
                $html += "<tr class='headsup'>"
	        $html+="<td>" + "SQL Express Edition so no SQL Agent, DBMail or ProDBA Monitoring"  + "</td>"
		        $html += "</tr>"
            }
            if($NonProductionAlertFlag){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Non-Production instance, so no ProDBA Monitoring or DBMail "  + "</td>"
		        $html += "</tr>"
            }
		    if($IndexMaintAlert){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Index maintenance may not be taking place : ProDBA should investigate before recommending to the client that this be done"  + "</td>"
		        $html += "</tr>"
            }
		    if($StatsMaintAlert){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Statistics updates may not be taking place : ProDBA should investigate before recommending to the client that this be done"  + "</td>"
		        $html += "</tr>"
            }
		    if($DBCCAlert){
                $html += "<tr class='headsup'>"
		        $html+="<td>" + "Database integrity checks may not be taking place : ProDBA should investigate before recommending to the client that this be done"  + "</td>"
		        $html += "</tr>"
            }
		    if($emailErrorsAlert){
                $html += "<tr>"
		        $html+="<td style='color: red'>" + "Database Mail error: $emailErrorMsg.<br/>Last error: $emailErrorDate. Please check the Database Mail log."  + "</td>"
		        $html += "</tr>"
            }
		    $html += "</table>"
            $html += "</div>"
        }


    return $html
	
}


## Modified 17/08/2017 - Ian Harris
## Added $backupdefault and $logdefault parameters 

## Modified 28/12/17 - Gordon Feeney
## Updated query to retun AG status and health

## Modified 21/05/21 - Gordon Feeney
## Added new items to the summary
                
function Get-SQLDatabaseSummary($instance,$version,$config){

    $server_date = (Get-SQLLocalDateMidnight $instance)
  
    
    ## Default backup thresholds
    if($config.backupdefault) {
            $backupdefault = $config.backupdefault
    }
    else
    {
            $backupdefault = 0
    }


    ## IanH 11/11/2020
    ## Allow for a day range to be added to config.xml for days when we check backups
    ## Useful when a client has only weekend (or weekday) backups so backup failures are ignored on other days

    $backupsCheckedTodayFlag = 1 
	

    if($config.backupdaystart -And $config.backupdayend)
    {

           $backupdaystart = $config.backupdaystart
           $backupdayend = $config.backupdayend 

           $backupsCheckedTodayFlag = BackupsCheckedToday -backupdaystart $backupdaystart -backupdayend $backupdayend


           if($backupsCheckedTodayFlag -eq 0) 
           {
             ## Neater to always have the full day name in the message
             $backupdaystartTidy = DOWTidyUp -day $backupdaystart
             $backupdayendTidy = DOWTidyUp -day $backupdayend 
			 $backupsMessage = "<div class='disabled'>"  + "As it's " + (Get-Date).DayOfWeek + " we're not checking backups today" + "</div>" + `
                        "<div class='disabled'>" + "(Days on which backups are checked as defined in config.xml : from $backupdaystartTidy to $backupdayendTidy)" + "</div>"
           }
           else
           {
             $backupsMessage = "na" 
           }

    }


    if ($config.logdefault) 
    {
        $logdefault = $config.logdefault
    }
    else
    {
        $logdefault = 0
    }

    if($config.enabled) {
            $enableddefault = [bool]$config.enabled
    }
    else
    {
            $enableddefault = [bool]1
    }
    ## AG Backup Defaults
    ## IanH 15/05/2020
    
    if($config.AGFullDefault) {
            $AGFullDefault = $config.AGFullDefault
    }
    else
    {
        $AGFullDefault = $null 
    }
    
    if($config.AGDiffDefault) {
            $AGDiffDefault = $config.AGDiffDefault
    }
    else
    {
            $AGDiffDefault = $null 
    }

    
    if($config.AGLogDefault) {
            $AGLogDefault = $config.AGLogDefault
    }
    else
    {
            $AGLogDefault = $null 
    }

    ## Can have a backup history cutoff date for slow servers or servers with huge msdb backup history tables
    ## v2.28 IanH

    if($config.BackupHistoryCutoff) {
             $BackupHistoryCutoffStr = " and bs.backup_finish_date > dateadd(dd, -" + $config.BackupHistoryCutoff.ToString() + ", getdate())" 
             $backupHistoryCutoffNote  = $config.BackupHistoryCutoff
    }
    else
    {
            $BackupHistoryCutoffStr = " -- do nothing" 
            $backupHistoryCutoffNote  = 0
    }


    
    $config = $config.database | ForEach {

		New-Object PSObject -Property @{
			DatabaseName = $_.name
			BackupHours = $_.backup
			LogBackupMins = $_.log
			SimpleOverride = $(if(!([string]::IsNullOrEmpty($_."simple-override"))) { if([int]$_."simple-override" -eq 1) { $true } else { $false } } else { $false } )
            LogOnlyBkup = $(if(!([string]::IsNullOrEmpty($_."log-only-backup"))) { if([int]$_."log-only-backup" -eq 1) { $true } else { $false } } else { $false } )
			Enabled = [bool]([int]$_.enabled)
			Notes = $_.notes
		}
	}


## Added the backup query as a separate string variable to allow us to insert an optional "backup history cutoff" date Where clause 
## for slow running servers or large msdb backup history tables where the query would otherwise time out - v2.28 IanH

  $backupQueryString = "
        DECLARE @SQL nvarchar(max);
        
        -- Returns the following columns
        -- database_name, is_backup_instance, is_ha, recovery_model, last_backup, last_backup_type, device_type,
        -- native_backup, last_log_backup, state, mirror_role, mirror_state, AG_role, AG_state, backup_preferred_replica

        IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') as varchar), 4) AS DECIMAL(5, 3))) >= 11
	        SET @SQL = '		        
		        SELECT 
                    d.name AS database_name
			        ,CASE 
	                    WHEN mirroring_role_desc IS NOT NULL THEN
		                    CASE mirroring_role_desc 
			                    WHEN ''PRINCIPAL'' THEN 1
			                    ELSE 0
		                    END
	                    ELSE 
		                    CASE 
			                    WHEN AG_info.role_desc IS NOT NULL THEN master.sys.fn_hadr_backup_is_preferred_replica(d.name)
			                    ELSE 1
		                    END
                    END AS is_backup_instance
                    ,CASE 
						WHEN mirroring_role_desc IS NULL AND AG_info.role_desc IS NULL THEN 0
						ELSE 1
					END AS is_ha
					,d.recovery_model_desc AS recovery_model
			        ,b.backup_start_date AS last_backup
					,DATEDIFF(SECOND, b.backup_start_date, b.backup_finish_date) AS last_backup_duration
			        ,b.type AS last_backup_type
					,COALESCE(b.device_type, l.device_type) AS device_type
                    ,b.native_backup
			        ,l.backup_start_date AS last_log_backup
			        ,DATEDIFF(SECOND, l.backup_start_date, l.backup_finish_date) AS last_log_backup_duration
			        ,CASE d.is_in_standby
						WHEN 1 THEN ''STANDBY''
						WHEN 0 THEN d.state_desc 
			        END AS [state]
					,m.mirroring_role_desc AS [mirror_role]
					,m.mirroring_state_desc AS [mirror_state]
			        ,AG_info.role_desc AS [ag_role]
					,AG_info.synchronization_health_desc AS [ag_state]
					,CASE 
						WHEN AG_info.role_desc IS NOT NULL THEN master.sys.fn_hadr_backup_is_preferred_replica(d.name)
						ELSE NULL
					END AS backup_preferred_replica
                    ,AG_info.name as [ag_name]
                    ,d.compatibility_level
					,d.is_auto_shrink_on
					,fsize.file_size_mb AS data_file_size
					,lsize.file_size_mb AS log_file_size 
					,b.compression_ratio AS full_backup_compression_ratio
					,l.compression_ratio AS log_backup_compression_ratio
                    ,d.create_date as db_create_date
		        FROM sys.databases d
					LEFT OUTER JOIN 
				(
					SELECT 
						database_name
						,backup_start_date
						,backup_finish_date
						,CASE bs.type
							WHEN ''D'' THEN ''FULL''
							WHEN ''I'' THEN ''DIFF''
							WHEN ''F'' THEN ''FILE/FILEGROUP''
							WHEN ''G'' THEN ''DIFF FILE/FILEGROUP''
							WHEN ''P'' THEN ''PARTIAL''
							WHEN ''Q'' THEN ''DIFF PARTIAL''
						END AS [type]
						,CASE bmf.device_type
							WHEN 2 THEN ''Disk''
				            WHEN 3 THEN ''Diskette''
							WHEN 5 THEN ''Tape''
							WHEN 6 THEN ''Pipe''
							WHEN 7 THEN ''Virtual''
                            WHEN 9 THEN ''Azure Storage''
							ELSE ''Unknown''
						END AS device_type
						,CASE 
							WHEN bmf.device_type IN (2, 3, 5, 6, 9) THEN 1
							ELSE 0
						END AS native_backup
						,CONVERT (NUMERIC (20, 2), (CONVERT (FLOAT, bs.backup_size) / CONVERT (FLOAT, bs.compressed_backup_size))) compression_ratio
						,ROW_NUMBER() OVER (PARTITION BY database_name ORDER BY backup_start_date DESC) AS [rn]
					FROM 
						msdb.dbo.backupset bs
							INNER JOIN 
						msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
					WHERE bs.[type] NOT IN (''L'')
				) b ON (d.name COLLATE Latin1_General_CI_AS = b.database_name COLLATE Latin1_General_CI_AS AND b.rn = 1)
					LEFT OUTER JOIN 
				(
					SELECT 
						database_name
						,backup_start_date
						,backup_finish_date
						,CASE bmf.device_type
							WHEN 2 THEN ''Disk''
							WHEN 3 THEN ''Diskette''
							WHEN 5 THEN ''Tape''
							WHEN 6 THEN ''Pipe''
							WHEN 7 THEN ''Virtual''
							WHEN 9 THEN ''Azure Storage''
							ELSE ''Unknown''
						END AS device_type
						,CONVERT (NUMERIC (20, 2), (CONVERT (FLOAT, bs.backup_size) / CONVERT (FLOAT, bs.compressed_backup_size))) compression_ratio
						,ROW_NUMBER() OVER (PARTITION BY database_name ORDER BY backup_start_date DESC) AS [rn]
					FROM 
						msdb.dbo.backupset bs
							INNER JOIN 
						msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
					WHERE bs.[type] = ''L''
				) l ON (d.name COLLATE Latin1_General_CI_AS = l.database_name COLLATE Latin1_General_CI_AS AND l.rn=1)
					LEFT OUTER JOIN  
				sys.database_mirroring m ON m.database_id = d.database_id AND m.mirroring_guid IS NOT NULL
					LEFT OUTER JOIN 
				(
					SELECT drs.database_id, drs.synchronization_health_desc, ars.role_desc, ag.name
					FROM 
						sys.dm_hadr_database_replica_states drs 
							LEFT OUTER JOIN 
					    sys.dm_hadr_availability_replica_states AS ars ON drs.replica_id = ars.replica_id AND ars.is_local = 1 
                        join master.sys.availability_groups ag on ars.group_id = ag.group_id
					WHERE ars.role_desc IS NOT NULL
				) AG_info ON AG_info.database_id = d.database_id 
					INNER JOIN 
				(
					SELECT database_id, (SUM(cast(size as bigint)) * 8) / 1024 AS file_size_mb
					FROM sys.master_files
					WHERE type = 0
					GROUP BY database_id
				) AS fsize ON fsize.database_id = d.database_id
					INNER JOIN 
				(
					SELECT database_id, (SUM(cast(size as bigint)) * 8) / 1024 AS file_size_mb
					FROM sys.master_files
					WHERE type = 1
					GROUP BY database_id
				) AS lsize ON lsize.database_id = d.database_id
			WHERE d.name != ''tempdb'' 
                --AND d.name = ''RedGate'' 		        
			ORDER BY d.name
            OPTION (RECOMPILE);'
        ELSE
	        SET @SQL = '
		        SELECT 
			        d.name AS database_name
			        ,CASE 
				        WHEN mirroring_role_desc IS NOT NULL THEN
					        CASE mirroring_role_desc 
						        WHEN ''PRINCIPAL'' THEN 1
						        ELSE 0
					        END	
				        ELSE 1							
			        END AS is_backup_instance
					,CASE 
						WHEN mirroring_role_desc IS NULL THEN 0
						ELSE 1
					END AS is_ha
					,d.recovery_model_desc AS recovery_model
			        ,b.backup_start_date AS last_backup
			        ,DATEDIFF(SECOND, b.backup_start_date, b.backup_finish_date) AS last_backup_duration
			        ,b.type AS last_backup_type
					,COALESCE(b.device_type, l.device_type) AS device_type
			        ,b.native_backup
			        ,l.backup_start_date AS last_log_backup
			        ,DATEDIFF(SECOND, l.backup_start_date, l.backup_finish_date) AS last_log_backup_duration
			        ,CASE d.is_in_standby
						WHEN 1 THEN ''STANDBY''
						WHEN 0 THEN d.state_desc 
			        END AS [state]
			        ,m.mirroring_role_desc AS [mirror_role]
					,m.mirroring_state_desc AS [mirror_state]
			        ,NULL AS [AG_role], NULL [AG_role]
			        ,NULL AS [AG_role], NULL [AG_state]
					,NULL AS backup_preferred_replica
                    ,NULL AS [AG_name]
                    ,d.compatibility_level
					,d.is_auto_shrink_on
					,fsize.file_size_mb AS data_file_size
					,lsize.file_size_mb AS log_file_size 
					,b.compression_ratio AS full_backup_compression_ratio
					,l.compression_ratio AS log_backup_compression_ratio
                    ,d.create_date as db_create_date
		        FROM sys.databases d
					LEFT OUTER JOIN 
				(
					SELECT 
						database_name
						,backup_start_date
						,backup_finish_date
						,CASE bs.type
							WHEN ''D'' THEN ''FULL''
							WHEN ''I'' THEN ''DIFF''
							WHEN ''F'' THEN ''FILE/FILEGROUP''
							WHEN ''G'' THEN ''DIFF FILE/FILEGROUP''
							WHEN ''P'' THEN ''PARTIAL''
							WHEN ''Q'' THEN ''DIFF PARTIAL''
						END AS [type]
						,CASE bmf.device_type
							WHEN 2 THEN ''Disk''
							WHEN 3 THEN ''Diskette''
							WHEN 5 THEN ''Tape''
							WHEN 6 THEN ''Pipe''
							WHEN 7 THEN ''Virtual''
							WHEN 9 THEN ''Azure Storage''
							ELSE ''Unknown''
						END AS device_type
						,CASE 
							WHEN bmf.device_type IN (2, 3, 5, 6, 9) THEN 1
							ELSE 0
						END AS native_backup ' + 
						CASE 
							WHEN CAST(LEFT(CAST(SERVERPROPERTY('productversion') as varchar), 4) AS DECIMAL(5, 3)) > 10 THEN 
								',CONVERT (NUMERIC (20, 2), (CONVERT (FLOAT, bs.backup_size) / CONVERT (FLOAT, bs.compressed_backup_size))) compression_ratio'
							ELSE ',NULL AS compression_ratio'
						END + '
						,ROW_NUMBER() OVER (PARTITION BY database_name ORDER BY backup_start_date DESC) AS [rn]
					FROM 
						msdb.dbo.backupset bs
							INNER JOIN 
						msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
					WHERE bs.[type] NOT IN (''L'') 
				) b ON (d.name COLLATE Latin1_General_CI_AS = b.database_name COLLATE Latin1_General_CI_AS AND b.rn = 1)
					LEFT OUTER JOIN 
				(
					SELECT 
						database_name
						,backup_start_date
						,backup_finish_date
						,CASE bmf.device_type
							WHEN 2 THEN ''Disk''
							WHEN 3 THEN ''Diskette''
							WHEN 5 THEN ''Tape''
							WHEN 6 THEN ''Pipe''
							WHEN 7 THEN ''Virtual''
							WHEN 9 THEN ''Azure Storage''
							ELSE ''Unknown''
						END AS device_type ' + 
						CASE 
							WHEN CAST(LEFT(CAST(SERVERPROPERTY('productversion') as varchar), 4) AS DECIMAL(5, 3)) > 10 THEN 
								',CONVERT (NUMERIC (20, 2), (CONVERT (FLOAT, bs.backup_size) / CONVERT (FLOAT, bs.compressed_backup_size))) compression_ratio'
							ELSE ',NULL AS compression_ratio'
						END + '
						,ROW_NUMBER() OVER (PARTITION BY database_name ORDER BY backup_start_date DESC) AS [rn]
					FROM 
						msdb.dbo.backupset bs
							INNER JOIN 
						msdb.dbo.backupmediafamily bmf ON bs.media_set_id = bmf.media_set_id
					WHERE bs.[type] = ''L'' 
				) l ON (d.name COLLATE Latin1_General_CI_AS = l.database_name COLLATE Latin1_General_CI_AS AND l.rn=1)
					LEFT OUTER JOIN  
				sys.database_mirroring m ON m.database_id = d.database_id AND m.mirroring_guid IS NOT NULL
					INNER JOIN 
				(
					SELECT database_id, (SUM(size) * 8) / 1024 AS file_size_mb
					FROM sys.master_files
					WHERE type = 0
					GROUP BY database_id
				) AS fsize ON fsize.database_id = d.database_id
					INNER JOIN 
				(
					SELECT database_id, (SUM(size) * 8) / 1024 AS file_size_mb
					FROM sys.master_files
					WHERE type = 1
					GROUP BY database_id
				) AS lsize ON lsize.database_id = d.database_id
		        WHERE d.name != ''tempdb'' 
                    --AND d.name = ''RedGate''
                ORDER BY d.name
                OPTION (RECOMPILE);
		        '

        EXEC sp_executesql @SQL;
		--PRINT @SQL;
	"


    ## Modified 13/09/18 - GFF
    ## Encapsulated query in try-catch block, created separate object in Catch block and added QueryError value to both Try and Catch objects    

    ## Modified 07/02/19 - GFF
    ## Added OPTION (RECOMPILE) to backups query due to intermittent timeouts.  

    try{
        $table =  Query-SQL $instance $backupQueryString
        
       
     $table | ForEach-Object {        

	    $dbname = $_.database_name
        
        $recoverymodel = $_.recovery_model		
        $lastbackup = $_.last_backup
        $lastbackupduration = $_.last_backup_duration
        
    ## Modified 17/08/2017 - Ian Harris
    ## Option to specify default backup and log backup thresholds

    ## Modified 28/122017 - Gordon Feeney
    ## Added $is_pref_backup variable to help determine if backups flagged as alerts can be ignored
        
        $lastbackupthreshold = if($backupdefault -eq 0) {$null} else {$server_date.AddHours(-1*$backupdefault)}
        $lastbackuptype = $_.last_backup_type     
	    $devicetype = $_.device_type       
	    $lastlogbackup = $_.last_log_backup
        $lastlogbackupduration = $_.last_log_backup_duration
	    $lastlogbackupthreshold = if($logdefault -eq 0) {$null} else {$server_date.AddMinutes(-1*$logdefault)} 
	    $isbackupinstance = $_.is_backup_instance #Will always be 1 for non-AG and non-mirror instances
        $nativebackup = [bool]$_.native_backup
        $isha = $_.is_ha
        $state = $_.state
	    $mirrorrole = $_.mirror_role
	    $mirrorstate = $_.mirror_state
	    $agrole = $_.ag_role
	    $agstate = $_.ag_state
	    $prefreplica = if($_.backup_preferred_replica -eq 1){"Yes"}elseif($_.backup_preferred_replica -eq 0){"No"}else{""}
        $agname = $_.ag_name
        $complevel = $_.compatibility_level
        $autoshrink = $_.is_auto_shrink_on
        $enabled = $enableddefault
	    $notes = $null
        $BackupHistoryCutoffNote = $backupHistoryCutoffNote
	    $simpleOverride = $false
        $logonlyBackup = $false
        $AGFull = if(![string]::IsNullOrEmpty($AGFullDefault)) {$AGFullDefault} 
        $AGDiff = if(![string]::IsNullOrEmpty($AGDiffDefault)) {$AGDiffDefault}
        $AGLog = if(![string]::IsNullOrEmpty($AGLogDefault))  {$AGLogDefault}
        $dataFileSize = $_.data_file_size
        $logFileSize = $_.log_file_size
        $fullBackupCompRatio = $_.full_backup_compression_ratio
        $logBackupCompRatio = $_.full_backup_compression_ratio
        $dbCreateDate = $_.db_create_date
        

## IanH 11/11/2020 Added for "weekend only" or "weekday only " backups
        $backupsCheckedToday = $backupsCheckedTodayFlag
        $backupsCheckedTodayMessage = $backupsMessage

        $config |  Where-Object {$_.DatabaseName -eq $dbname} | ForEach {

            if(![string]::IsNullOrEmpty($_.BackupHours)){
        	    $lastbackupthreshold = $server_date.AddHours(-1*$_.BackupHours)
		    }
            
		    if(![string]::IsNullOrEmpty($_.LogBackupMins)){
                $lastlogbackupthreshold = $server_date.AddMinutes(-1*$_.LogBackupMins)
		    }
			
		    if(![string]::IsNullOrEmpty($_.Enabled) -And $_.Enabled){
			    $enabled = $true
		    } elseif (![string]::IsNullOrEmpty($_.Enabled) -And !($_.Enabled)){
			    $enabled = $false
		    }
			
            if($_.SimpleOverride){
			    $simpleOverride = $true
		    }

            if($_.LogOnlyBkup){
			    $logonlyBackup = $true

		    }
            if(![string]::IsNullOrEmpty($_.Notes)){
                        $notes = $_.Notes + " (" + $_.DatabaseName + ")" 
            }
            else {
                $notes = $null
            }
			
	    }
	    
        
            New-Object PSObject -Property @{
                QueryError = $false
			    DatabaseName = $dbname
			    RecoveryModel = $recoverymodel
			    LastBackup = $lastbackup
                LastBackupDuration = $lastbackupDuration
			    LastBackupType = $lastbackuptype
			    DeviceType = $devicetype
			    LastBackupThreshold = $lastbackupthreshold
			    LastLogBackup = $lastlogbackup
                LastLogBackupduration = $lastlogbackupduration
			    LastLogBackupThreshold = $lastlogbackupthreshold
			    IsBackupInstance = $isbackupinstance
                NativeBackup = $nativebackup
                IsHA = $isha
                State = $state
			    MirrorRole = $mirrorrole
			    MirrorState = $mirrorstate
			    AGRole = $agrole
			    AGState = $agstate
                AGFull = $AGFull
                AGDiff = $AGDiff
                AGLog = $AGLog
                AGName = $agname
			    PrefReplica = $prefreplica
                CompLevel = $complevel
                AutoShrink = $autoshrink
			    Enabled = $enabled 
			    Notes = $notes
                BackupHistoryCutoffNote=$BackupHistoryCutoffNote 
			    SimpleOverride = $simpleOverride
                LogOnlyBackup = $logonlyBackup
                backupsCheckedToday = $backupsCheckedToday
                backupsCheckedTodayMessage = $backupsCheckedTodayMessage
                DataFileSize = $dataFileSize
                LogfileSize= $logFileSize 
                FullBackupCompRatio = $fullBackupCompRatio
                LogBackupCompRatio = $logBackupCompRatio
                DBCreateDate = $dbCreateDate
 		    } | Select-Object `
		    QueryError `
            ,DatabaseName `
		    ,RecoveryModel `
		    ,LastBackup `
		    ,LastBackupDuration `
		    ,LastBackupType `
		    ,DeviceType `
		    ,LastBackupThreshold `
		    ,LastLogBackup `
		    ,LastLogBackupDuration `
		    ,LastLogBackupThreshold `
		    ,IsBackupInstance `
            ,NativeBackup `
            ,IsHA `
		    ,State `
		    ,MirrorRole `
		    ,MirrorState `
		    ,AGRole `
		    ,AGState `
            ,AGFull `
            ,AGDiff `
            ,AGLog `
            ,AGName `
		    ,PrefReplica `
            ,CompLevel `
            ,AutoShrink `
		    ,Enabled `
		    ,Notes `
            ,BackupHistoryCutoffNote `
            ,SimpleOverride `
            ,backupsCheckedToday `
            ,backupsCheckedTodayMessage `
            ,DataFileSize `
            ,LogFileSize `
            ,FullBackupCompRatio `
            ,LogBackupCompRatio `
            ,DBCreateDate `
		    ,@{Name="BackupAlert";Expression=
			    { 
			    if($_.State -ne "ONLINE" -Or $_.LogOnlyBackup -Or $_.backupsCheckedToday -eq 0 -Or $_.DBCreateDate -gt $_.LastBackupThreshold ){   
                    [bool]0
			    }
			    elseif( !( ([DBNull]::Value).Equals($_.LastBackup) -Or [string]::IsNullOrEmpty($_.LastBackupThreshold) ) -And ($_.LastBackup -gt $_.LastBackupThreshold) ){                    
                    [bool]0
			    }
			    else{
                    [bool]1
<#
                    ## Only alert if IsBackupInstance, but also check AGFull and AGDiff to see if Full / Diff backups run on this replica
                    ## IanH 13/05/2020
                     ## If this DB has an AG Role and at least one of AGFull or AGDiff are defined
                    if ( ![string]::IsNullOrEmpty($_.AGRole) -And ( !([string]::IsNullOrEmpty($_.AGFull)) -Or !([string]::IsNullOrEmpty($_.AGDiff)) )`
                             -And ($_.lastbackuptype -eq "DIFF" -Or $_.lastbackuptype -eq "FULL") ) 
                    {
                        if($_.AGRole -eq $_.AGFull -And $_.lastbackuptype -eq "FULL")
                        {
                            [bool]1
                        }
                        elseif($_.AGRole -eq $_.AGDiff -And $_.lastbackuptype -eq "DIFF")
                        {
                            [bool]1
                        }
                        elseif($_.AGRole -ne $_.AGFull -And $_.lastbackuptype -eq "FULL")
                        {
                            [bool]0
                        }
                        elseif($_.AGRole -ne $_.AGDiff -And $_.lastbackuptype -eq "DIFF")
                        {
                            [bool]0
                        }
                        elseif($_.PrefReplica -eq "Yes" -And $_.AGFull -eq "PREFERRED" -And $_.lastbackuptype -eq "FULL" )
                        {
                            [bool]1
                        }
                        elseif ($_.PrefReplica -eq "Yes"  -And $_.AGDiff -eq "PREFERRED" -And $_.lastbackuptype -eq "DIFF" )
                        {
                            [bool]1
                        }
                        elseif ($_.IsBackupInstance)
                        {
                            [bool]1
                        }
                        else
                        {
                            [bool]0
                        }
    		        }
                    ## This isn't an AG DB or we're using neither AGFull nor AGDiff
			        elseif ($_.IsBackupInstance){
                           [bool]1
                    }
                    else{
                           [bool]0
                    }
					#>
			    }
                }
		
		    } `
		    ,@{Name="LogBackupAlert";Expression=
			    { 
			
			    if($_.State -ne "ONLINE" -Or $_.RecoveryModel -eq "SIMPLE" -Or $_.SimpleOverride -Or  $_.backupsCheckedToday -eq 0 -Or $_.DBCreateDate -gt $_.LastBackupThreshold){   ## IanH 11/11/2020 Added backupsCheckedToday
				    [bool]0
  			    }
			    elseif( !(([DBNull]::Value).Equals($_.LastLogBackup) -Or [string]::IsNullOrEmpty($_.LastLogBackupThreshold)) -And ($_.LastLogBackup -gt $_.LastLogBackupThreshold) ){
				    [bool]0
			    }
			    else{
                    [bool]1
<#
                    ## If this DB has and AG Role and AGLog is defined
			        ## Negation exclamation mark in the wrong place when processing Log backups.
                    if (![string]::IsNullOrEmpty($_.AGRole) -And ![string]::IsNullOrEmpty($_.AGLog))
                    {
                        if($_.AGRole -eq $_.AGLog -or ($_.PrefReplica -eq "Yes" -And $_.AGLog -eq "PREFERRED" ) )
                        {
                            [bool]1
                        }
                        elseif($_.AGRole -ne $_.AGLog -and !($_.PrefReplica -eq "Yes" -And $_.AGLog -eq "PREFERRED" ) )
                        {
                            [bool]0
                        }
                        elseif($_.PrefReplica -eq "Yes" -And $_.AGLog -eq "PREFERRED" )
                        {
                            [bool]1
                        }
				        elseif ($_.IsBackupInstance){
				            [bool]1
                        }
                        else{
				            [bool]0
                        }
			        }
				    elseif ($_.IsBackupInstance){
				          [bool]1
                    }
                    else{
				          [bool]0
                    }
#>
			      }
					
			    }
		     } `
            ,@{Name="MirrorAlert";Expression=
			
			    { 			
			        if(![string]::IsNullOrEmpty($_.MirrorState) -and $_.MirrorState -ne "SYNCHRONIZED" -and $_.MirrorState -ne "SYNCHRONIZING"){
				        [bool]1
			        }
			        else{
				        [bool]0
			        }
			    }
		    } `
		    ,@{Name="AGAlert";Expression=
			
			    { 			
			        if(![string]::IsNullOrEmpty($_.AGState) -and $_.AGState -ne "HEALTHY" ){
				        [bool]1
			        }
			        else{
				        [bool]0
			        }
			    }
		    } `
          ,@{Name="DodgyDBAlert";Expression=
          {
           if(($_.State -eq "SUSPECT" -or $_.State -eq "RECOVERY_PENDING") -And (($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled)))
                    {
                        [bool]1
                    }
                    else  
                    {
				        [bool]0
			        }
            }	
          }`
            ,@{Name="AutoShrinkAlert";Expression=
			
			    { 			
			        if(![string]::IsNullOrEmpty($_.AutoShrink) -and $_.AutoShrink -eq "True" ){
				        [bool]1
			        }
			        else{
				        [bool]0
			        }
			    }
		    } `
            ,@{Name="RecentlyCreatedDBAlert";Expression=
			
			    { 			
			        if(![string]::IsNullOrEmpty($_.DBCreateDate) -and $_.DBCreateDate -gt $_.LastBackupThreshold ){
				        [bool]1
			        }
			        else{
				        [bool]0
			        }
			    }
		    } | Select-Object `
		    QueryError `
            ,DatabaseName `
		    ,RecoveryModel `
		    ,LastBackup `
		    ,LastBackupDuration `
		    ,LastBackupType `
		    ,DeviceType `
		    ,LastBackupThreshold `
		    ,LastLogBackup `
		    ,LastLogBackupDuration `
		    ,LastLogBackupThreshold `
		    ,IsBackupInstance `
            ,NativeBackup `
        	,IsHA `
		    ,State `
		    ,MirrorRole `
		    ,MirrorState `
		    ,AGRole `
		    ,AGState `
            ,AGFull `
            ,AGDiff `
            ,AGLog `
            ,AGName `
		    ,PrefReplica `
            ,CompLevel `
            ,AutoShrink `
		    ,Enabled `
		    ,Notes `
            ,BackupHistoryCutoffNote `
            ,SimpleOverride `
            ,backupsCheckedToday `
            ,backupsCheckedTodayMessage `
		    ,DataFileSize `
            ,LogFileSize `
            ,FullBackupCompRatio `
            ,LogBackupCompRatio `
            ,DBCreateDate `
            ,BackupAlert `
            ,LogBackupAlert `
            ,DodgyDBAlert `
            ,AGAlert `
            ,MirrorAlert `
		    ,@{Name="Alert";Expression=			
			    { 			
			        #if(($_.BackupAlert -Or $_.LogBackupAlert) -And (($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled)))
                    if( ($_.BackupAlert -Or $_.LogBackupAlert -Or $_.DodgyDBAlert -Or $_.AGAlert ) -And ( ($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled) ) )
                    {
                        [bool]1
			        } 
                    else  
                    {
				        [bool]0
			        }						
			    }

		    }`
		    ,@{Name="GenericBackupAlert";Expression=			
			    { 			
                    if( ($_.BackupAlert -Or $_.LogBackupAlert) -And ( ($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled) ) )
                    {
                        [bool]1
			        } 
                    else  
                    {
				        [bool]0
			        }						
			    }
           }`
          ,@{Name="AGNotFullReplica";Expression=			## If an AG and fulls not run here then 1 otherwise 0 
			    { 		
                    if([string]::IsNullOrEmpty($_.AGRole))   ## not an AG so just set to 0
                    {
                       [bool]0
                    }                       
	                elseif(($_.AGRole -eq $_.AGFull) -or ($_.PrefReplica -eq "Yes" -And $_.AGFull -eq "PREFERRED" ) `
                            -or ([string]::IsNullOrEmpty($_.AGFull) -and $_.PrefReplica -eq "Yes" ) )  ## Is an AG and full backups are run here
                    {
                        [bool]0
                    }
                    else
                    {
                        [bool]1
                    } 
                }               
		    }`
          ,@{Name="AGNotDiffReplica";Expression=			## If an AG and diffs not run here then 1 otherwise 0 
			    { 		
                    if([string]::IsNullOrEmpty($_.AGRole))   ## not an AG so just set to 0
                    {
                       [bool]0
                    }
	                elseif(($_.AGRole -eq $_.AGDiff) -or ($_.PrefReplica -eq "Yes" -And $_.AGDiff -eq "PREFERRED" ) `
                                -or ([string]::IsNullOrEmpty($_.AGDiff) -and $_.PrefReplica -eq "Yes" ) )   ## Is an AG and diff backups are run here
                    {
                        [bool]0
                    }
                    else
                    {
                        [bool]1
                    }
                 }                
		    }`
            ,@{Name="AGNotLogReplica";Expression=			## If an AG and logs not run here then 1 otherwise 0 
			    { 		
                    if([string]::IsNullOrEmpty($_.AGRole))   ## not an AG so just set to 0
                    {
                       [bool]0
                    }
	                elseif($_.AGRole -eq $_.AGLog -or ($_.PrefReplica -eq "Yes" -And $_.AGLog -eq "PREFERRED" ) `
                                   -or ([string]::IsNullOrEmpty($_.AGLog) -and $_.PrefReplica -eq "Yes" )  )   ## Is an AG and logs are backed up here
                    {
                        [bool]0
                    }
                    else
                    {
                        [bool]1
                    } 
                 }               
		    }`
            ,@{Name="HAAlert";Expression=			
			    { 			
			        if(($_.MirrorAlert -or $_.AGAlert))
                    {
				        [bool]1
			        } 
                    else 
                    {
				        [bool]0
			        }						
			    }

		    }`
            ,@{Name="AutoShrinkAlert";Expression=			
			    { 			
			        if(($_.AutoShrinkAlert))
                    {
				        [bool]1
			        } 
                    else 
                    {
				        [bool]0
			        }						
			    }

		    }`
            ,@{Name="RecentlyCreatedDBAlert";Expression=			
			    { 			
			        if(($_.RecentlyCreatedDBAlert )  )
                    {
				        [bool]1
			        } 
                    else 
                    {
				        [bool]0
			        }						
			    }

		    }`
        } | Select-Object `
		    QueryError `
            ,DatabaseName `
		    ,RecoveryModel `
		    ,LastBackup `
		    ,LastBackupDuration `
		    ,LastBackupType `
		    ,DeviceType `
		    ,LastBackupThreshold `
		    ,LastLogBackup `
		    ,LastLogBackupDuration `
		    ,LastLogBackupThreshold `
		    ,IsBackupInstance `
            ,NativeBackup `
        	,IsHA `
		    ,State `
		    ,MirrorRole `
		    ,MirrorState `
		    ,AGRole `
		    ,AGState `
            ,AGFull `
            ,AGDiff `
            ,AGLog `
            ,AGName `
		    ,PrefReplica `
            ,CompLevel `
            ,AutoShrink `
		    ,Enabled `
		    ,Notes `
            ,BackupHistoryCutoffNote `
            ,SimpleOverride `
            ,backupsCheckedToday `
            ,backupsCheckedTodayMessage `
		    ,DataFileSize `
            ,LogFileSize `
            ,FullBackupCompRatio `
            ,LogBackupCompRatio `
            ,DBCreateDate `
            ,DodgyDBAlert `
            ,AGAlert `
            ,MirrorAlert `
            ,Alert `
            ,BackupAlert `
            ,LogBackupAlert `
            ,GenericBackupAlert `
            ,AGNotFullReplica `
            ,AGNotDiffReplica `
            ,AGNotLogReplica `
            ,HAAlert `
            ,AutoShrinkAlert `
            ,RecentlyCreatedDBAlert `
            ,@{Name="RedBackupAlert";Expression=			
			    { 			
			        if( ([string]::IsNullOrEmpty($_.enabled) -or $_.enabled -eq 1) `
                        -and
                        (
                          (
                            $_.BackupAlert `
                            -and ( 
                                    ($_.LastBackupType -eq "FULL" -and $_.AGNotFullReplica -eq 0) `
                                -or ($_.LastBackupType -eq "DIFF" -and $_.AGNotDiffReplica -eq 0) `
                                -or ( ([DBNull]::Value).Equals($_.LastBackup) -and ($_.AGNotFullReplica -eq 0 -or $_.AGNotDiffReplica -eq 0))  
                                ) 
                         ) `
                     -or 
                         (
                            $_.LogBackupAlert `
                            -and $_.AGNotLogReplica -eq 0
                         ) `
                     -or $_.AGAlert `
                     -or $_.DodgyDBAlert `
                     -or $_.MirrorAlert)
                     )
                    {
				        [bool]1
			        } 
                    else 
                    {
				        [bool]0
			        }						
			    }

		    }`
            ,@{Name="AmberBackupAlert";Expression=			
			    { 			
			        if( ([string]::IsNullOrEmpty($_.enabled) -or $_.enabled -eq 1) `
                        -and
                        (          
                            (
                                $_.BackupAlert `
                                -and (
                                        ($_.LastBackupType -eq "FULL" -and $_.AGNotFullReplica -eq 1) `
                                    -or ($_.LastBackupType -eq "DIFF" -and $_.AGNotDiffReplica -eq 1) `
                                ) `
                            ) `
                        -or
                            (
                                $_.LogBackupAlert `
                                -and $_.AGNotLogReplica -eq 1
                            ) `
                        -or
                            (
                                $_.AutoShrinkAlert  `
                                -or ($_.LogFileSize -gt $_.DataFileSize -and $_.LogFileSize -gt 1000) `
                                -or $_.RecentlyCreatedDBAlert
                            )`
                        -and $_.State -ne "OFFLINE" 
                        )
                    )
                    {
				        [bool]1
			        } 
                    else 
                    {
				        [bool]0
			        }						
			    }

		    }`

    }
    

	catch{
        New-Object PSObject -Property @{
                QueryError = $true
			    DatabaseName = $_.Exception.Message			    
		    } | Select-Object `
		    QueryError `
            ,DatabaseName `
            ,@{Name="RedBackupAlert";Expression={[bool]1}}
##          ,@{Name="Alert";Expression={[bool]1}}
    }		
}


## Modified 21/05/21 - Gordon Feeney
## Added new columns to the summary
                
function FormatHTML-SQLDatabaseSummary($obj){

##$obj | Select-Object  ## Useful troubleshooting tip - shows the whole array of objects passed to this function
    

    ## If we use a backup history cutoff then we want to indicate this with a note below the table - v2.28 IanH
    $BackupHistoryCutoffNoteFlag = 0
    
    ## The backup history cutoff note is set at the instance backup level - but this foreach is the only way I could work 
    ## out how to pass it through to this function - v2.28 Ianh   
    ## (IanH 11/11/2020 - Same for the "backups checked today" flag and message)

    
    $obj |  Where-Object {($_.BackupHistoryCutoffNote -gt 0) } | ForEach {
               $BackupHistoryCutoffNoteFlag = 1 
               $BackupHistoryCutoffNoteValue = $_.BackupHistoryCutoffNote
        }
    
    $obj |  Where-Object {($_.backupsCheckedToday -eq 0) } | ForEach {
               $backupsCheckedTodayFlag = 0 
               $backupsCheckedTodayMessage = $_.backupsCheckedTodayMessage
        }



	$notesCounter = [int]0
	$notes = @()
		
	$obj | Where-Object {!($_.Notes -eq $null) } | ForEach {
		
		$notesCounter += 1
			
		$note = New-Object PSObject -Property @{
			ID = $notesCounter
			Database = $_.DatabaseName
			Note = $_.Notes
		}
		
		$notes += $note
    }

 ## Let's not show the backup error for all AG backups - too wordy
 
    $AGNotesFlag = [int]0
	$AGBackupAlertNotes = @()

	$obj | Where-Object {!($_.AGBackupAlertNotes -eq $null) -And $_.AGBackupAlertNotes -ne "NoAGBackupAlert"} | ForEach {
		
		$AGBackupAlertNotesCounter += 1
			
		$AGBackupAlertNote = New-Object PSObject -Property @{
			ID = $AGBackupAlertNotesCounter
			Database = $_.DatabaseName
			AGBackupAlertNote = $_.AGBackupAlertNotes
		}
		
		$AGBackupAlertNotes += $AGBackupAlertNote				
	}


## Commented out $logFileSizeAlert and $AutoShrinkAlert settings as this seemed to lead to all databases being flagged as headsup
## IanH 05/05/2022

    # Build a LogFileSizeAlert array
    $LogFileSizeAlertNotesCounter = [int]0
	$LogFileSizeAlertNotes = @()
    ##$logFileSizeAlert = $false

	$obj | Where-Object {($_.logFileSize -gt $_.dataFileSize -and $_.logfilesize -gt 1000 -and $_.State -ne "OFFLINE") } | ForEach {
		
		$LogFileSizeAlertNotesCounter += 1
			
		$LogFileSizeAlertNote = New-Object PSObject -Property @{
			ID = $LogFileSizeAlertNotesCounter
			Database = $_.DatabaseName
			DataFileSize = [int]($_.DataFileSize / 1024)
            LogFileSize = [int]($_.LogFileSize / 1024 )
		}
		
		$LogFileSizeAlertNotes += $LogFileSizeAlertNote
        ##$logFileSizeAlert = $true
	}

    # Build an AutoShrinkAlert array
    $AutoShrinkAlertNotesCounter = [int]0
	$AutoShrinkAlertNotes = @()
    ##$AutoShrinkAlert = $false

	$obj | Where-Object {($_.AutoShrinkAlert -eq $true) } | ForEach {
		
		$AutoShrinkAlertNotesCounter += 1
			
		$AutoShrinkAlertNote = New-Object PSObject -Property @{
			ID = $AutoShrinkAlertNotesCounter
			Database = $_.DatabaseName
		}
		
		$AutoShrinkAlertNotes += $AutoShrinkAlertNote
        ##$AutoShrinkAlert = $true
	}

    # Build a RecentlyCreatedDBAlert array
    $RecentlyCreatedDBAlertNotesCounter = [int]0
	$RecentlyCreatedAlertDBNotes = @()

    $obj | Where-Object {($_.RecentlyCreatedDBAlert -eq $true) } | ForEach {
		
		$RecentlyCreatedDBAlertNotesCounter += 1
        
        $db = $_.DatabaseName
		$RecentlyCreatedDBAlertNote = New-Object PSObject -Property @{
			ID = $RecentlyCreatedDBAlertNotesCounter
			Database = $_.DatabaseName
            CreateDate = $_.DBCreateDate 
		}
		
		$RecentlyCreatedAlertDBNotes += $RecentlyCreatedDBAlertNote
	}

    # Build a Suspect or Recovery Pending Alert array
    $DodgyDBAlertNotesCounter = [int]0
	$DodgyDBAlertNotes = @()

	$obj | Where-Object {($_.DodgyDBAlert -eq $true) } | ForEach {
		
		$DodgyDBAlertNotesCounter += 1
			
		$DodgyDBAlertNote = New-Object PSObject -Property @{
			ID = $DodgyDBAlertNotesCounter
			Database = $_.DatabaseName
            Status = $_.State 
		}
		
		$DodgyDBAlertNotes += $DodgyDBAlertNote
	}

## Convoluted way of getting all the AG names from the AGName property of each database backup row
## Used when determining whether or not to use any AG backup default settings in config.xml
## If more than one AG then we ignore any config.xml settings as it gets too convoluted. 


$AGLoopCounter = [int]0 
$AGNewNameFlag = [int]0
$AGNameTemp = $null 
$AGNameList = @()

    $obj | Where-Object {![string]::IsNullOrEmpty($_.AGName) } | ForEach {
	
        $AGNameTemp = $_.AGName
      
        if($AGLoopCounter -gt 0){
            $AGNewNameFlag = 0
    
            $AGNameList | ForEach {
                if($_.AGName -eq $AGNameTemp){}
                else { $AGNewNameFlag = 1 }
      
            }  
        } 

     	if($AGLoopCounter -eq 0 -or $AGNewNameFlag -eq 1) {
            #Write-Host "AG Name : " $AGNameTemp
            $AGNameObject = New-Object PSObject -Property @{ 
				AGName = $AGNameTemp
		    }   
            $AGNameList += $AGNameObject
        }
		$AGLoopCounter += 1

	}



##    Write-Host "AG Notes Counter : " $AGBackupAlertNotesCounter
		
	$html = "
	<h5>Backups alerts are based on a threshold (usually midnight on the previous day) so any database not backed up between that threshold and the time the Health Check is generated is flagged as a potential issue.<br/></h5>
	<table class='summary'>
	<tr>
		<th>Database</th>
		<th>Recovery Model</th>
		<th>Last Backup</th>
		<th>Last Backup Type</th>
        <th>Last Backup Duration (mins)</th>
		<th>Device Type</th>
		<th>Backup Threshold</th>
		<th>Last Log Backup</th>
		<th>Log Backup Threshold</th>
		<th>State</th>
        <th>Compatibility Level</th>
        <th>Auto Shrink</th>
		<th>Mirror Role</th>
		<th>Mirror State</th>
        <th>AG Role</th>
		<th>AG State</th>
		<th>Preferred replica</th>
	</tr>
	"
	
## Modified 29/12/17 - Gordon Feeney
## Added AGRole and AGState columns

    $IsBackupInstanceFlag = $False
    $alert = $false
    $headsup = $false
##    $row = $False
    $databaseEnabled = $true
    $nonNativeBackups = $false
    $haAlert = $false

    ## Flags used to determine if we display notes at the bottom related to disabling dbs in config.xml
    ## and show note regarding AG database backups
    $atLeastOneDBDisabled = $false
    $atLeastOneAGDatabase = $false

    $AGFullDef = $null 
    $AGDiffDef = $null 
    $AGLogDef = $null 
        
    $obj | ForEach {
    
##  If we're flagging a backup with a warning, then we want to highlight which columns 
##  have caused the alert. - IanH 05/05/2022
        $highlightbackupissue = $false 
        $highlightAGissue = $false 
        $highlightDBissue = $false ## DB state is SUSPECT or RECOVERY_PENDING
        $highlightMirrorissue = $false ## Alert if mirror state is not SYNCHRONIZED or SYNCHRONIZING

        $db = $_.DatabaseName 
        $nonNativeBackups = $nonNativeBackups -or !$_.NativeBackup
        $haAlert = $haAlert -or $_.HAAlert
        #$autoShrinkAlert = $autoShrinkAlert -or $_.AutoShrinkAlert        

        if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){$atLeastOneDBDisabled = $true }

        ## Put in check for just a single AG. If more than one AG then ignore any AG backup prefs set in config.xml   
  #### REMOVE
  
        if((($_.Enabled) -Or ([string]::IsNullOrEmpty($_.Enabled))  ) -And !([string]::IsNullOrEmpty($_.AGRole)) -And $AGNameList.Count -ge 1 ){
            $atLeastOneAGDatabase = $true 
            if(!([string]::IsNullOrEmpty($_.AGFull))){
                $AGFullDef = $_.AGFull
            }
            if(!([string]::IsNullOrEmpty($_.AGDiff))){
                $AGDiffDef = $_.AGDiff
            }
            if(!([string]::IsNullOrEmpty($_.AGLog))){
                $AGLogDef = $_.AGLog
            }
        }


                
        ## Modified 13/09/18 - GFF
        ## Test for new QueryError value created in Get-SQLDatabaseSummary.            
        if ($_.QueryError -eq $true){
            $html += "<tr class='warning'>
                <td colspan='14' class='disabled'" + ">" + $db + "</td>"
        }
        else{
            #If .....
            if ($_.NativeBackup -and $_.IsHA -and !$_.IsBackupInstance)
	        {
                #Even if only one database has native backups then flag it for a later heads-up below.
		        $IsBackupInstanceFlag = $True
	        }

            
##            $startrow = $false
            $databaseIsEnabled = $databaseIsEnabled -and $_.Enabled

            

            #Only display the row if it's: (a) a Full Health Check; or (b) a Headline Healtch Check and it's an alert                
            if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and $_.Alert)){
##                $row = $true
##                Write-Host "Alert for database : " $_.DatabaseName " GBA " $_.GenericBackupAlert " AGA " $_.AGAlert " DDA " $_.DodgyDBAlert " MA " $_.MirrorAlert " IBA " $_.IsBackupInstance

                ## We want to be able to highlight which column caused the warning so use the issue flags
                if( $_.RedBackupAlert ) `
  <#                 ##     $_.BackupAlert `
                        -and ( 
                                ($_.LastBackupType -eq "FULL" -and $_.AGNotFullReplica -eq 0) `
                            -or ($_.LastBackupType -eq "DIFF" -and $_.AGNotDiffReplica -eq 0) `
                            -or ([DBNull]::Value).Equals($_.LastBackup) 
                            ) 
                     ) `
                     -or (
                        $_.LogBackupAlert `
                        -and $_.AGNotLogReplica -eq 0
                     ) `
                     -or $_.AGAlert `
                     -or $_.DodgyDBAlert `
                     -or $_.MirrorAlert)#>{
			        $html += "<tr class='warning'>" 
                    if($_.GenericBackupAlert){$highlightbackupissue = $true}               
                    if($_.AGAlert){$highlightAGissue = $true}
                    if($_.DodgyDBAlert){$highlightDBissue = $true}
                    if($_.MirrorAlert){$highlightMirrorissue = $true}
		        } 
                ## Bugfix for $AutoShrinkAlert and $LogFileSizeAlert which were setting everything to headsup - IanH 05/05/2022
 ##               elseif(( $_.NativeBackup -and  !$_.IsBackupInstance) -or $_.AutoShrinkAlert -or ($_.LogFileSize -gt $_.DataFileSize) )  ## remove NativeBackup check
                elseif( $_.AmberBackupAlert ) 
<#                        (
                            $_.BackupAlert `
                            -and (
                                    ($_.LastBackupType -eq "FULL" -and $_.AGNotFullReplica -eq 1) `
                                -or ($_.LastBackupType -eq "DIFF" -and $_.AGNotDiffReplica -eq 1) `
                            ) `
                        ) `
                        -or
                        (
                            $_.LogBackupAlert `
                            -and $_.AGNotLogReplica -eq 1
                        ) `
                        -or
                        (
                            $_.AutoShrinkAlert  `
                            -or ($_.LogFileSize -gt $_.DataFileSize -and $_.LogFileSize -gt 1000) `
                            -or $_.RecentlyCreatedDBAlert
                        )`
                        -and $_.State -ne "OFFLINE"
                    )             ##( !$_.IsBackupInstance) -or $_.AutoShrinkAlert -or ($_.LogFileSize -gt $_.DataFileSize -and $_.LogFileSize -gt 1000) -and $_.State -ne "OFFLINE") #>
                    {
                        #Write-Host "Warning for database : " $_.DatabaseName
                  	    $html += "<tr class='headsup'>"
                    }
                else{
			        $html += "<tr>"
		        }
                
		        $note = $notes | Where-Object {$_.Database -eq $db }
                $AutoShrinkAlertNote = $AutoShrinkAlertNotes | Where-Object {$_.Database -eq $db }  
                $LogFileSizeAlertNote = $LogFileSizeAlertNotes | Where-Object {$_.Database -eq $db }
                $RecentlyCreatedAlertDBNote = $RecentlyCreatedAlertDBNotes | Where-Object {$_.Database -eq $db }
                $DodgyDBAlertNote = $DodgyDBAlertNotes | Where-Object {$_.Database -eq $db }

                $lastBackupDuration = $_.LastBackupDuration
                if (!([string]::IsNullOrEmpty($lastBackupDuration))){
                    if ($lastBackupDuration -lt 60) {
                        $lastBackupDuration = "< 1"
                    } 
                    else {
                        $lastBackupDuration = ([math]::Round($lastBackupDuration / 60)).ToString()
                    }
                }

                ## Added bit to highlight the LastBackup and / or the LastLogBackup cells if they caused an alert
                ## Also added "No Backup", "Missing Log Backup" if those cells are empty but should not be. IanH 05/05/2022


     	        $html += "                		        
                <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.DatabaseName+`
                    $(if(!($note -eq $null)){"<sup> "+$note.ID+"</sup>"})+` 
                    $(if($_.AutoShrinkAlert){"<sup> "+[char]($AutoShrinkAlertNote.ID + 111) +"</sup>"})+` 
                    $(if($_.LogFileSize -gt $_.DataFileSize -and $_.LogFileSize -gt 1000 -and $_.State -ne "OFFLINE"){"<sup> "+[char]($LogFileSizeAlertNote.ID + 64) +"</sup>"})+` 
                    $(if($_.RecentlyCreatedDBAlert){"<sup> "+[char]($RecentlyCreatedAlertDBNote.ID + 96) +"</sup>"})+`
                "</td>
    		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.RecoveryModel+"</td>" +
## Last Backup
                "<td"+$(
                       if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"}
                       elseif (  
                                $highlightbackupissue -eq $true  `
                                -and ( 
                                        [string]::IsNullOrEmpty($_.LastBackup) -or $_.LastBackup -lt $_.LastBackupThreshold 
                                    ) `
                               -and (
                                       ( 
                                            ($_.LastBackupType -eq "FULL" -or [string]::IsNullOrEmpty($_.LastBackup) ) `
                                            -and $_.AGNotFullReplica -eq 0 
                                        )`
                                        -or 
                                        ( 
                                            ($_.LastBackupType -eq "DIFF" -or [string]::IsNullOrEmpty($_.LastBackup)) `
                                            -and $_.AGNotDiffReplica -eq 0
                                        ) 
                                    ) 
                              )   
                              {" class='backupsummaryhighlight'" } )  +
                    ">"+
                        $(if($highlightbackupissue -eq $true -and [string]::IsNullOrEmpty($_.LastBackup) ) 
                            { (" No backup ") }
                        else 
                            { (Format-DateTime $_.LastBackup) } ) + 
                "</td>" +
##
		        "<td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.LastBackupType+"</td>
                <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+ $lastBackupDuration +"</td>
		        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.DeviceType+"</td>" +
## Backup Threshold
		        "<td"+$(
                        if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+
                        $(if( ($_.LastBackupType -eq "FULL" -and $_.AGNotFullReplica -eq 1) -or ($_.LastBackupType -eq "DIFF" -and $_.AGNotDiffReplica -eq 1) `
                                -or ( ([DBNull]::Value).Equals($_.LastBackup) -and ( $_.AGNotFullReplica -eq 1 -or $_.AGNotDiffReplica -eq 1 )))
                            { "<small>(" + (Format-DateTime $_.LastBackupThreshold) + ")</small>"  }  ## Not relevant on this replica so write it small
                        else{ (Format-DateTime $_.LastBackupThreshold)  } ) +"</td>" +
## Last Log Backup
                "<td"+$(
                        if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"}
                            elseif (  
                                ($highlightbackupissue -eq $true ) `
                                -and ($_.RecoveryModel -eq "FULL") `
                                -and ( [string]::IsNullOrEmpty($_.LastLogBackup) -or $_.LastLogBackup -lt $_.LastLogBackupThreshold ) `
                                -and $_.SimpleOverride -eq $false  `
                                -and ( $_.AGNotLogReplica -eq 0)   
                                )
                                   {" class='backupsummaryhighlight'" } 
                             ) +">"+ 
                         $(if
                            ($highlightbackupissue -eq $true `
                                -and ($_.RecoveryModel -eq "FULL") `
                                -and [string]::IsNullOrEmpty($_.LastLogBackup) `
                                -and $_.SimpleOverride -eq $false   `
                                -and $_.AGNotLogReplica -eq 0
                            )
                            {  " Missing log backup " }  
                            elseif ( 
                            $_.RecoveryModel -eq "FULL"  
                        ) 
                            { (Format-DateTime $_.LastLogBackup) } ) + 
                  "</td>" +
## Log Backup Threshold
                  "<td"+
                        $(if(
                            !($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) )
                                {" class='disabled'"})  +">" +  
                        $(if( 
                            ($_.RecoveryModel -eq "SIMPLE") -or ($_.SimpleOverride -eq $true ) ) 
                                { " N/A "}
                        elseif(
                            $_.AGNotLogReplica -eq 1)
                                {"<small> (" + (Format-DateTime $_.LastLogBackupThreshold) + ")</small>"}  ## Not relevant on this replica so write it small
                        else 
                            {(Format-DateTime $_.LastLogBackupThreshold) } ) +
                  "</td>" +
##
                "<td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"} elseif($highlightDBissue){" class='backupsummaryhighlight'" } )+">"+$_.State+"</td> 
                <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.CompLevel+"</td> 
                <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">" + $(if($_.AutoShrink -eq 'True') {'ON'} else {''}) + "</td> 
		        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.MirrorRole+"</td>
		        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"}
                        elseif ($highlightMirrorissue -eq $true ){" class='backupsummaryhighlight'" } ) +">"+
                            $_.MirrorState+"</td>
        	    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.AGRole+"</td>
		        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"}  
                     elseif($_.AGAlert){" class='backupsummaryhighlight'" } ) +">" + $_.AGState+"</td>
		        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.PrefReplica+"</td>                 
		        "

                $html += "</tr>"
            }
        } 
    }

    $html += "</table>"	
 #############################################

    ## flag so we only put the heading in once at most
    $databasesummarynotesheading = $false



    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL") -and ($_.Alert -or !$_.IsBackupInstance)) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and $_.Alert)){
        if ($IsBackupInstanceFlag) {
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
	        $html += "<div class='notenormal'>"
		    $html+="Backup thresholds for databases which are mirror secondaries are ignored."
 		    $html += "</div>"
	        }

       if ($atLeastOneAGDatabase) {
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
	        $html += "<div class='notenormal'>"
            $html+="For databases which are members of an Availability Group, by default backups not on the Preferred Backup Replica are ignored. However different AG backup preferences can be specified in the config.xml.</br>"
		    $html += "</div>"
	        }

        if($backupsCheckedTodayFlag -eq 0){
	        if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
            $html += $backupsCheckedTodayMessage 
        }
    
        if ($atLeastOneDBDisabled -eq $true){
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
            $html += "<div class='notenormal'>"
            $html += "<i>Databases in italics</i> have been disabled in the config file and so backups for these can be ignored."
            $html += "</div>"
        }

        if($BackupHistoryCutoffNoteFlag -eq 1)
        {            
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
            $html += "<div class='notenormal'>"
            $html += "Database Backup History limited to " + $BackupHistoryCutoffNoteValue + " days to avoid query timeout"
            $html += "</div>"
  ##          $NoteSection = 1
        }

        
        if ($nonNativeBackups)
	    {
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Notes"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
            $html += "<div class='notenormal'>"
		    $html +="Pro DBA don't support Virtual/Third-party backups so we can't investigate any failures of such backup types."
            $html += "</div>"
	    }


        if ($atLeastOneAGDatabase)
        {   
            $AGNamesString = "("
            $AGNameTemp = $null 
            $AGNameList | Sort-Object -Property AGName | ForEach   {
                if([string]::IsNullOrEmpty($AGNameTemp)){ 
                    $AGNamesString += $_.AGName
                   }
                elseif($AGNameTemp -ne $_.AGName){
                    $AGNamesString += ", "
                    $AGNamesString += $_.AGName 
                   }
                $AGNameTemp = $_.AGName 
            }
            $AGNamesString += ")"
            $html += "<div class='headsupnote'>"
		    $html +="AG Backup Settings " 
            $html += $AGNamesString
            if ( ![string]::IsNullOrEmpty($AGFullDef) -or ![string]::IsNullOrEmpty($AGDiffDef) -or ![string]::IsNullOrEmpty($AGLogDef)){
                $html += " (Defined in config.xml)"
            }
            $html += "</div>"
            $html += "<div class='notenormal'>"
	        if ( ![string]::IsNullOrEmpty($AGFullDef)) {$html += "Fulls = " + $AGFullDef + " "} else {$html += "Fulls = Preferred "}
            if ( ![string]::IsNullOrEmpty($AGDiffDef)) {$html += "Diffs = " + $AGDiffDef + " "} else {$html += "Diffs = Preferred "}
            if ( ![string]::IsNullOrEmpty($AGLogDef)) {$html += "Log = " + $AGLogDef + " "} else {$html += "Log = Preferred "}
            $html += "</div>"
	    }
            


        if($notes.Count -gt 0){
	
		    $html += "<div class='headsupnote'>"
		    $html+="Database Notes (from config.xml)"
		    $html += "</div>"
            $html += "<table class='notenormal'>"
		
		    $notes | ForEach{
		
			    $html += "<tr>"
		
			    $html+="<td><sup>"+$_.ID+"</sup> "+$_.Note+"</td>"
			
			    $html += "</tr>"
		
		    }

              $html += "</table>"
        }

<# Let's not have all these AG database backup alert messages - too wordy
        if($AGBackupAlertNotes.Count -gt 0){
	
		    $html += "<table class='notenormal'>"
		
		    $AGBackupAlertNotes | ForEach{
		
			    $html += "<tr>"
		
			    $html+="<td><sup>"+[char]($_.ID + 33) +"</sup> "+$_.AGBackupAlertNote+"</td>"
			
			    $html += "</tr>"
		
		    }
             $html += "</table>"
        }
#>


        if($AutoShrinkAlertNotes.Count -gt 0){
            
		    $html += "<div class='headsupnote'>"
		    $html+="Databases with Auto Shrink enabled"
		    $html += "</div>"
            $html += "<table class='notenormal'>"
		
		    $AutoShrinkAlertNotes | ForEach{
		
		<#	    $html += "<tr class='headsup'>"#>
		
			    $html+="<td><sup>"+[char]($_.ID + 111) + "</sup>" + " (" + $_.Database + ")" + "</td>"
			
			    $html += "</tr>"
		
		    }

            $html += "</table>"
        }

        if($LogFileSizeAlertNotes.Count -gt 0){
            
		    $html += "<div class='headsupnote'>"
		    $html+="Log files larger than data files (and larger than 1 GB)"
		    $html += "</div>"
            $html += "<table class='notenormal'>"
		
		    $LogFileSizeAlertNotes | ForEach{
		
		<#	    $html += "<tr class='headsup'>" #>
		
			    $html+="<td><sup>"+[char]($_.ID + 64) +"</sup> Data file size: " + $_.DataFileSize +" GB, Log file size: "+$_.LogFileSize+" GB (" + $_.Database + ")" + "</td>"
			
			    $html += "</tr>"
		
		    }
            $html += "</table>"
        }

        if($RecentlyCreatedAlertDBNotes.Count -gt 0){
            $noteCount = $RecentlyCreatedAlertDBNotes.Count
            $html += "<div class='headsupnote'>"
		    $html+="Recently Created Database(s) - not had time to backup"
		    $html += "</div>"
            $html += "<table class='notenormal'>"
		
		    $RecentlyCreatedAlertDBNotes | ForEach{
		
		<#	    $html += "<tr class='headsup'>" #>
		
			    $html+="<td><sup>"+[char]($_.ID + 96) +"</sup> Creation Date : " + $_.CreateDate + " (" + $_.Database + ")" + "</td>"
			
			    $html += "</tr>"
		
		    }
            $html += "</table>"
        }

	
        if($DodgyDBAlertNotes.Count -gt 0){
            
		    $html += "<div class='headsupnote'>"
		    $html+="Databases with a status of SUSPECT or RECOVERY PENDING"
		    $html += "</div>"
            $html += "<table class='notenormal'>"
		
		    $DodgyDBAlertNotes | ForEach{
		
		<#	    $html += "<tr class='headsup'>" #>
		
			    $html+="<td><sup>"+[char]($_.ID + 96) +"</sup> Status : " + $_.Status + " (" + $_.Database + ")" + "</td>"
			
			    $html += "</tr>"
		
		    }
            $html += "</table>"
        }


    ## reuse this flag for this notes section
    $databasesummarynotesheading = $false

        if ($haAlert)
	    {
            if( $databasesummarynotesheading -eq $false){
                    $html += "<div class='headsupnote'>"
		            $html+="Warning(s)"
                    $html += "</div>"
                    $databasesummarynotesheading = $true 
            }
            $html += "<div class='notenormal'>"
		    $html +="One or more databases have an unhealthy Mirroring or Availability Group synchronisation status."
            $html += "</div>"
        }
		

		    
    }
	return $html

}


## Modified 30/10/18 - Gordon Feeney
## Replaced search for jobs run in the last day with jobs run since midnight on the previous day.

## Modified 27/11/18 - Gordon Feeney
## Search for jobs run since midnight on the previous day hadn't been implemented correctly.

## Modified 21/01/19 - Gordon Feeney
## Fixed bug in search for Reports where non-mirroring and non-AG systems weren't accounted for.

# GFF: 05/12/19
# Additional provision forj obs that have succeeded but with step failures

## GFF: 21/05/21
## Fixed bug whereby AG secondaries were causing the query below to fail but the report was still 
## being generated OK. Both mirroring and AG properties being checked should have been AND'd rather than OR'd

function Get-SQLFailedJobs($instance,$version,$config){

    #Exclude error messages for Pro-DBA jobs?
    if(![string]::IsNullOrEmpty($config.exclude.prodba_errors)) {
        $excludeprodba = $config.exclude.prodba_errors        
    }
    else {
        $excludeprodba = "n"
    }

    #Option to exclude SSIS error messages if SSISDB.internal.event_messages is too large

    if(![string]::IsNullOrEmpty($config.exclude.excludeSSISmessages)) {
        $excludeSSISmessages = $config.exclude.excludeSSISmessages        
    }
    else {
        $excludeSSISmessages = 0
    }

    $config = $config.exclude.job | ForEach {

		New-Object PSObject -Property @{
            Job = $_.name
			Enabled = $false
			Notes = $_.notes
		}
		
	}    

    ## Modified 27/11/18 - Gordon Feeney
    # Search for jobs run since midnight on the previous day hadn't been implemented correctly.

    ## Modified 21/01/19 - Gordon Feeney
    # Fixed bug in search for Reports where non-mirroring and non-AG systems weren't accounted for.

    ## Modified 30/07/20 - Ian Harris
    ## Fixed bug where job fails due to invalid owner (no steps are run so this was missed previously)

    $query = "
        SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

	    DECLARE @ExcludeProDBA char(1);
        SET @ExcludeProDBA = '$excludeprodba';

        DECLARE @ExcludeSSISMessages int;
        SET @ExcludeSSISMessages = '$excludeSSISmessages'
        
        DECLARE @SQL nvarchar(max);
		DECLARE @CompDate datetime;

		SET @CompDate = DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()) - 1, 0);
		
		DECLARE @ReportDBName sysname; -- database name 
		DECLARE @ReportJobsQuery nvarchar(MAX);
		DECLARE @ParamDefinition nvarchar(500);
		DECLARE @RowCount int;
		DECLARE @ErrorCount tinyint;
		DECLARE @ErrorMessage nvarchar(500);
		DECLARE @CrLf char(2);

		SET @ErrorCount = 0;
		SET @CrLf = char(13) + char(10);


		IF OBJECT_ID('tempdb..#Reports') IS NOT NULL
					DROP TABLE #Reports;

		CREATE TABLE #Reports
			(
			DatabaseName sysname, 
			ReportName nvarchar(500),
			ReportScheduleID sysname, 
			ReportStatus nvarchar(300)
			);


		--Reporting Services databases
		IF (SELECT CAST(LEFT(CAST(SERVERPROPERTY('productversion') as varchar), 4) AS DECIMAL(5, 3))) >= 11		
			BEGIN
				DECLARE db_cursor CURSOR READ_ONLY FORWARD_ONLY FOR 
				SELECT d.[name]
				FROM 
					sys.databases d
						LEFT OUTER JOIN 
					sys.database_mirroring m ON m.database_id = d.database_id 
						LEFT OUTER JOIN 
					(
						SELECT drs.database_id, drs.synchronization_health_desc, ars.role_desc
						FROM 
							sys.dm_hadr_database_replica_states drs 
								LEFT OUTER JOIN 
							sys.dm_hadr_availability_replica_states AS ars ON drs.replica_id = ars.replica_id AND ars.is_local = 1 
						WHERE ars.role_desc IS NOT NULL 
					) AG_info ON AG_info.database_id = d.database_id 
				WHERE 1 = 1
					AND d.database_id > 4 AND d.database_id <> 32767 
					--AND ((m.mirroring_guid IS NOT NULL AND m.mirroring_role_desc = 'PRINCIPAL' AND m.mirroring_state_desc = 'SYNCHRONIZED') OR (AG_info.role_desc = 'PRIMARY'))
					--Fixed bug in search for Reports where non-mirroring and non-AG systems weren't accounted for.
					AND state_desc = 'ONLINE'
                    --Fixed bug whereby AG secondaries were causing the query below to fail but the report was still being generated OK. Both mirroring and AG properties being 
                    --checked should have been AND'd rather than OR'd
					AND ((m.mirroring_guid IS NULL OR (m.mirroring_guid IS NOT NULL AND m.mirroring_role_desc = 'PRINCIPAL' AND m.mirroring_state_desc = 'SYNCHRONIZED')) AND (AG_info.role_desc IS NULL OR AG_info.role_desc = 'PRIMARY')) 					
				ORDER BY d.[name];
			END
		ELSE
			BEGIN
				DECLARE db_cursor CURSOR READ_ONLY FORWARD_ONLY FOR 
				SELECT d.[name]
				FROM 
					sys.databases d
						LEFT OUTER JOIN 
					sys.database_mirroring m ON m.database_id = d.database_id 					
				WHERE 1 = 1
					AND d.database_id > 4 AND d.database_id <> 32767				
					AND ((m.mirroring_guid IS NOT NULL AND m.mirroring_role_desc = 'PRINCIPAL' AND m.mirroring_state_desc = 'SYNCHRONIZED'))
				ORDER BY d.[name];
			END
		--END IF
					
		OPEN db_cursor  
		FETCH NEXT FROM db_cursor INTO @ReportDBName  
		WHILE @@FETCH_STATUS = 0  
		BEGIN  	
			BEGIN TRY
				SET @SQL = 'SELECT @RowCountOut = COUNT(*) FROM [' + @ReportDBName + '].dbo.sysobjects WHERE type = ''U'' AND name in (''Catalog'',  ''Subscriptions'', ''ReportSchedule'')';
				SET @ParamDefinition = N'@RowCountOut int OUTPUT';
				EXEC sp_executesql @SQL, @ParamDefinition, @RowCountOut = @RowCount OUTPUT
			END TRY

			BEGIN CATCH
				SET @ErrorMessage = ERROR_MESSAGE();
				SET @ErrorCount = @ErrorCount + 1;
				PRINT 'Error: ' + @ErrorMessage;
			END CATCH
			

			IF @RowCount = 3 AND @ErrorMessage IS NULL
				BEGIN
					BEGIN TRY
						SET @ReportJobsQuery  =        
									'INSERT INTO #Reports
									SELECT ' + '''' + 
										@ReportDBName + '''' + ', 
										c.Name AS ReportName
										, rs.ScheduleID AS ReportScheduleID
										, s.LastStatus AS ReportStatus
									FROM
										[' + @ReportDBName + ']..[Catalog] c
											INNER JOIN 
										[' + @ReportDBName + ']..Subscriptions s ON c.ItemID = s.Report_OID
											INNER JOIN 
										[' + @ReportDBName + ']..ReportSchedule rs ON c.ItemID = rs.ReportID
											AND rs.SubscriptionID = s.SubscriptionID;'
										
						EXEC sp_executesql @ReportJobsQuery 
					END TRY

					BEGIN CATCH
						SET @ErrorCount = @ErrorCount + 1;
						SET @ErrorMessage = ERROR_MESSAGE()
						PRINT 'Error: ' + @ErrorMessage 
					END CATCH
				END
			--END IF

			FETCH NEXT FROM db_cursor INTO @ReportDBName;		
		END

		CLOSE db_cursor  
		DEALLOCATE db_cursor;

			
		IF @ErrorCount = 0
			BEGIN
				SELECT j.name, j.[enabled], j.job_id, j.category_id, jh.instance_id, msdb.dbo.agent_datetime(jh.run_date,jh.run_time) AS run_time, jh.step_id, jh.run_status, STUFF(STUFF(REPLACE(STR(run_duration,6,0), ' ','0'),3,0,':'),6,0,':') AS duration, jh.message
				INTO #JobHistory
				FROM 
					msdb.dbo.sysjobs j 
						INNER JOIN 
					msdb.dbo.sysjobhistory jh ON jh.job_id = j.job_id
				WHERE 1 = 1
					AND jh.run_status IN (0, 1)
					AND msdb.dbo.agent_datetime(jh.run_date,jh.run_time) >= DATEADD(DAY, DATEDIFF(DAY, 0, GETDATE()) - 1, 0)
	--				AND ((j.[name] LIKE '(Pro-DBA)%' AND ISNULL('y', 'y') = 'n') OR j.[name] NOT LIKE '(Pro-DBA)%' )

				CREATE UNIQUE CLUSTERED INDEX IX_jobHistory_job_id_instance_id ON #JobHistory(job_id, instance_id);
				CREATE NONCLUSTERED INDEX IX_jobHistory_run_time ON #JobHistory(run_time);
				CREATE NONCLUSTERED INDEX IX_jobHistory_step_id_message ON #JobHistory(step_id) INCLUDE ([message]);

				--All jobs
				WITH Jobs As 
				(
					SELECT jh.*, js.job_steps, c.[name] AS category, jh2.run_time, jh2.duration
					FROM
					(
						SELECT 
							job_id, [name] AS job_name, 
							MAX(instance_id) AS instance_id,
							[enabled], category_id
						FROM 
							#JobHistory
						WHERE step_id = 0 AND run_status IN (0, 1)
						GROUP BY job_id, [name], [enabled], category_id
					) jh
						INNER JOIN 
					(
						SELECT job_id, COUNT(*) AS job_steps
						FROM msdb.dbo.sysjobsteps
						GROUP BY job_id
					) js ON js.job_id = jh.job_id
							INNER JOIN 
						msdb.dbo.syscategories c ON c.category_id = jh.category_id						
							INNER JOIN 
					#JobHistory jh2 ON jh2.job_id = jh.job_id AND jh2.instance_id = jh.instance_id					
				)
				,

				--Jobs that have succeeded
				Successes AS
				(
					SELECT jh.job_id, COUNT(*) AS success_count, MAX(jh.run_time) AS run_time
					FROM #JobHistory jh
					WHERE 1 = 1
						AND jh.run_status = 1
						AND jh.step_id = 0
					GROUP BY jh.name, jh.job_id
				)
				,

				--Jobs that have failed
				Failures AS
				(
					SELECT job_id, COUNT(*) AS failure_count, MAX(run_time) AS run_time
					FROM #JobHistory
					WHERE 1 = 1
						--AND   
						--	(step_id > 0 
						--	OR
						--	(step_id = 0 AND [message] like 'The job failed.  Unable to determine if the owner%')
						--	)
						AND step_id = 0 AND [message] like 'The job failed%'
						AND run_status = 0						
					GROUP BY job_id
				)
				,

				--Jobs that have succeeded but have step failures
				SuccessesWithStepFailures AS
				(
					SELECT jh.job_id, step_failure_count
					FROM #JobHistory jh
							LEFT OUTER JOIN 
						(
							SELECT job_id, msdb.dbo.agent_datetime(MAX(last_run_date), MAX(last_run_time)) AS step_failure_last_run_date, COUNT(*) AS step_failure_count
							FROM msdb.dbo.sysjobsteps
							WHERE last_run_outcome <> 1
								AND last_run_date <> 0 
							GROUP BY job_id
							HAVING msdb.dbo.agent_datetime(MAX(last_run_date), MAX(last_run_time)) > @CompDate
						) js ON js.job_id = jh.job_id
					WHERE 1 = 1
						AND jh.step_id = 0 
						AND jh.run_status = 1						
					GROUP BY jh.job_id, step_failure_count
				)
				,
		
				--Failed jobs' error messages
				Errors AS
				(
                    --We need to retrieve errors messages for step 0 and for the non-zero failed step if it exists.
					SELECT 
						jh.job_id, jh.run_time, 
						CASE 
							WHEN jh2.[message] IS NULL THEN jh.[message]
							ELSE jh2.[message]
					END AS [message], 
						js.subsystem AS job_subsystem, js.command AS job_command
					FROM 
						#JobHistory jh 
							INNER JOIN 
						msdb.dbo.sysjobsteps js ON js.job_id = jh.job_id AND js.step_id = 1
							LEFT OUTER JOIN 
						#JobHistory jh2 ON jh2.job_id = jh.job_id 
					WHERE 1 = 1
						AND jh.run_status = 0
						AND jh.instance_id = (
							SELECT MAX(instance_id) FROM #JobHistory
							WHERE job_id = jh.job_id 
								AND step_id = 0
								AND run_status = 0
							)
						AND jh2.run_status = 0
						AND jh2.instance_id = (
							SELECT MAX(instance_id) FROM #JobHistory
							WHERE job_id = jh2.job_id 
								AND step_id > 0
								AND run_status = 0
							)
						AND jh2.run_time = (
							SELECT MAX(run_time) FROM #JobHistory
							WHERE job_id = jh2.job_id 
								AND step_id > 0
								AND run_status = 0
							)
				)
				
				SELECT 
					CASE Jobs.category
						WHEN 'Report Server' THEN 'SSRS Report: ' + ISNULL(#Reports.ReportName, 'POTENTIAL ORPHANED SSRS SUBSCRIPTION') + ' (' + Jobs.job_name + ')'
						ELSE Jobs.job_name
					END AS job,		
					CASE Jobs.category
						WHEN 'Report Server' THEN 2
						ELSE 1
					END AS CategorySort,							
					Jobs.category, Jobs.[enabled], Jobs.job_steps, 
					Jobs.run_time AS start_time, 
					Successes.run_time AS latest_success, ISNULL(Successes.success_count, 0) AS success_count,
					Failures.run_time AS latest_failure, 
					--CASE 
					--	WHEN (ISNULL(Failures.failure_count, 0) = ISNULL(Successes.success_count, 0) AND ISNULL(SuccessesWithStepFailures.step_failure_count, 0) > 0) OR Jobs.job_name IN ('(Pro-DBA) Cycle Error Log') THEN 0
					--	ELSE ISNULL(Failures.failure_count, 0)
					--END AS failure_count,
					--CASE
					--	WHEN ISNULL(Successes.run_time, @CompDate) > ISNULL(Failures.run_time, @CompDate) OR Jobs.job_name IN ('(Pro-DBA) Cycle Error Log') THEN ISNULL(Failures.failure_count, 0) 
					--	ELSE 0
					--END AS warning_count, 
					ISNULL(Failures.failure_count, 0) AS failure_count,
					@CompDate AS CompDate, 
					Successes.run_time AS SuccessesRunTime, 
					Failures.run_time AS FailuresRunTime, 
					CASE
						--If the latest job run succeeded but there are step failures then it's a heads-up
						WHEN (ISNULL(Successes.run_time, @CompDate) >= ISNULL(Failures.run_time, @CompDate) AND ISNULL(SuccessesWithStepFailures.step_failure_count, 0) > 0) THEN 3
						--If a job has failed but its run time is earlier than the last success then it's a heads-up
						WHEN (ISNULL(Failures.failure_count, 0) > 0 AND ISNULL(Successes.run_time, @CompDate) > ISNULL(Failures.run_time, @CompDate)) THEN 2
						--If a job has failed and its run time is later than the last success then it's a failure/warning.
						WHEN (ISNULL(Failures.failure_count, 0) > 0 AND ISNULL(Failures.run_time, @CompDate) >= ISNULL(Successes.run_time, @CompDate)) THEN 1						
						ELSE 0
					END AS alert_type,					
					ISNULL(SuccessesWithStepFailures.step_failure_count, 0) AS step_failure_count,
                    CASE Errors.job_subsystem
                        WHEN 'SSIS' THEN Errors.[message] 
                        ELSE 
                            CASE 
								WHEN LEN(Errors.[message]) > 600 THEN
									CASE 
										WHEN CHARINDEX('Description: Executing the query', Errors.[message]) > 0 THEN SUBSTRING(Errors.[message], CHARINDEX('Description: Executing the query', Errors.[message]) + 33, 600) + ' .....'
										ELSE LEFT(Errors.[message], 300) + ' ..... ' + RIGHT(Errors.[message], 300)
									END
                                ELSE Errors.[message]
							END                            
                    END AS error_msg, 
					Errors.job_command, Errors.job_subsystem, 
					Jobs.duration
				INTO #TempJobs
				FROM 
					Jobs 
						LEFT OUTER JOIN 
					Successes ON Successes.job_id = Jobs.job_id 
						LEFT OUTER JOIN 
					Failures ON Failures.job_id = Jobs.job_id
						LEFT OUTER JOIN 
					SuccessesWithStepFailures ON SuccessesWithStepFailures.job_id = Jobs.job_id
						LEFT OUTER JOIN 
					Errors ON Errors.job_id = Jobs.job_id
						LEFT OUTER JOIN 
					#Reports ON #Reports.ReportScheduleID = Jobs.job_name
				WHERE [enabled] = 1
				--ORDER BY CategorySort, job, start_time DESC
				OPTION (RECOMPILE);

				IF EXISTS(SELECT * FROM sys.databases WHERE name = 'SSISDB')  AND @ExcludeSSISMessages = 0
                    BEGIN
				        ALTER TABLE #TempJobs 
				        ADD job_no int NOT NULL IDENTITY (1, 1);

				        --We need to update the error 
				        DECLARE @JobNo int;
						DECLARE	@JobName sysname;
				        DECLARE @JobCommand nvarchar(max);
				        DECLARE @SSISError nvarchar(max);
				        DECLARE @PosStart bigint, @PosEnd bigint;
                        DECLARE @Chars nvarchar(5);				
						
				        DECLARE jobs_cursor CURSOR READ_ONLY FORWARD_ONLY FOR 				
				        SELECT job_no, job, job_command
				        FROM #TempJobs
				        WHERE job_subsystem = 'SSIS'
						ORDER BY job_no;

				        OPEN jobs_cursor  
				        FETCH NEXT FROM jobs_cursor INTO @JobNo, @JobName, @JobCommand 
				        WHILE @@FETCH_STATUS = 0  
					        BEGIN  	
                                /*
                                Using ASCII codes as the actual characters screw up the PowerShell:

                                char(34) - perc
                                char(37) - double-quote
                                char(92) - backslash
                                */								
								SET @Chars = char(37) + char(34) + char(92) + char(34) + char(37);
								SET @PosStart = PATINDEX(@Chars, @JobCommand)+3;
										
								SET @Chars = char(37) + char(92) + char(34) + char(34) + char(37);
								SET @PosEnd = PATINDEX(@Chars, @JobCommand)
										
								IF @PosStart > 0 AND @PosEnd > 0
									BEGIN
										SET @JobCommand = SUBSTRING(@JobCommand, @PosStart, @PosEnd - @PosStart);
										SET @JobCommand = RIGHT(@JobCommand, CHARINDEX('\', REVERSE(@JobCommand)) - 1);

										SELECT TOP 1 @SSISError = [message]
										FROM SSISDB.[catalog].[event_messages] (NOLOCK)
										WHERE [package_name] COLLATE DATABASE_DEFAULT = @JobCommand COLLATE DATABASE_DEFAULT
											AND [event_name] IN ('OnError')
										ORDER BY [event_message_id] DESC;

										UPDATE #TempJobs
										SET error_msg = @SSISError
										WHERE job_no = @JobNo;
									END
								--END IF
						        
						        FETCH NEXT FROM jobs_cursor INTO @JobNo, @JobName, @JobCommand;
					        END
				        --END WHILE
		
				        CLOSE jobs_cursor  
				        DEALLOCATE jobs_cursor;
                    END
                --END WHILE

				SELECT *
				FROM #TempJobs
				ORDER BY CategorySort, job, start_time DESC;
			END
		--END IF

		DROP TABLE #Reports;
		DROP TABLE #TempJobs;
		DROP TABLE #JobHistory;
	" 


    ## Modified 13/09/18 - GFF
    ## Encapsulated query in try-catch block, created separate object in Catch block and added QueryError value to both Try and Catch objects
    try {
        #Write-host $instance
        ## Modified 10/01/19 - GFF
        ## Added global variable for script timeouts.

        $table =  Query-SQL $instance $query - ErrorAction Stop -Timeout 240

        $table | ForEach {
	
            $job = $_.job
            $category = $_.category
            $startDate = $_.start_time
		    $lastsuccess = $_.latest_success
		    $successes = $_.success_count
		    $lastfailure = $_.latest_failure
		    $failures = $_.failure_count
            $duration = $_.duration
            #$warnings = $_.warning_count
            ## GFF: 04/08/23
            ## Replaced 'warnings' with 'alertType'
            $alertType = $_.alert_type
            $stepFailures = $_.step_failure_count
            # We only want to include those error messages where alert_type is 1. Other non-zero alert types are for heads-ups rather than warnings.
            $errormsg = $(if(!([string]::IsNullOrEmpty($_.error_msg)) -and $alertType -eq 1){$_.error_msg})
            $enabled = $true
		    $notes = $null
		
            $config |  Where-Object {$_.Job -eq $job} | ForEach {
		
			    $enabled = $_.Enabled
			    $notes = $_.Notes
		
		    }
		
            New-Object PSObject -Property @{
			    QueryError = $false
			    Job = $job
			    Category = $category
                StartDate = $startDate
			    LastSuccess = $lastsuccess
			    Successes = $successes
			    LastFailure = $lastfailure
			    Failures = $failures
                Duration = $duration
                StepFailures = $stepFailures
                HeadsUp = $null
                ErrorMsg = $errormsg
                JobEnabled = $jobenabled
                Enabled = $enabled
			    Notes = $notes
		    } | Select-Object `
		    QueryError `
            ,Job `
		    ,Category `
		    ,StartDate `
            ,LastSuccess `
		    ,Successes `
		    ,LastFailure `
		    ,Failures `
            ,Duration `
            ,StepFailures `
            ,HeadsUp `
            ,ErrorMsg  `
            ,Enabled `
            ,Notes `
		    ,@{Name="Alert";Expression=			
			    { 	
                    <#
                    if ($stepFailures -gt 0)	{
                        #Jobs that have succeeded but with step failures
                        [int]3
                    }
                    elseif ($warnings -gt 0)	{
                        #Failed jobs that have subsequently succeeded
                        [int]2
                    }
                    elseif(($_.Failures -gt 0) -and ($_.Failures -ne $_.Successes)) {
                        #Enabled failed job
                        if ((($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled))){
                            [int]1
                        }
                        else {
                            #Disabled failed job
                            [int]2
                        }
			        } 
                    else {
        		        [int]0
			        }
                    #>

                    ## GFF: 04/08/23
                    ## Simplified different tests above with single 'alert type' variable.
			        
                    if($alertType -eq 1){
                        #Enabled failed job
                        if ((($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled))){
                            [int]1
                        }
                        else {
                            #Disabled failed job
                            [int]2
                        }
			        } 
                    elseif ($alertType -eq 2){
                        #Failed jobs that have subsequently succeeded
                        [int]2
                    }
                    elseif($alertType -eq 3){
                        #Jobs that have succeeded but with step failures
                        [int]3
                    }
                    else {
        		        [int]0
			        }
			    }
		    }
	    }
    }

    catch {
        $job =  $_.Exception.Message
        New-Object PSObject -Property @{
            QueryError = $true
			Job = $job
		} | Select-Object `
        QueryError `
		,Job `
        ,@{Name="Alert";Expression={[int]1}}
    }
        
}


#GFF: 18/12/19
#Additional provision forj obs that have succeeded but with step failures
function FormatHTML-SQLFailedJobs($obj, $agent_running){

	$notesCounter = [int]0
	$notes = @()
    $errorCounter = [int]0
	$errors = @()
	$notesBCounter = [int]0
	$notesB = @()    
	
    $obj | Where-Object {!($_.Notes -eq $null) -and ($_.Alert -ne 3)} | ForEach {
		$notesCounter += 1			
		$note = New-Object PSObject -Property @{
			ID = $notesCounter
			Job = $_.Job
			Note = $_.Notes
            Alert = $_.Alert
		}
        $notes += $note
	}
    
    $obj | Where-Object {!($_.ErrorMsg -eq $null) } | ForEach {
        $errorCounter += 1			
        $errormsg = New-Object PSObject -Property @{
			ID = $errorCounter
			Job = $_.Job
			ErrorMsg = $_.ErrorMsg
		}				
		$errors += $errorMsg
	}

    #GFF: 05/12/19
    #New Notes object for jobs that have succeeded but with step failures
	$obj | Where-Object {($_.Alert -eq 3) } | ForEach {
		$notesCounter += 1			
		$note = New-Object PSObject -Property @{
			ID = $notesCounter
			Job = $_.Job
            Note = $_.Notes
            Alert = $_.Alert
		}
        $notesB += $note
	}
    
    $html = "
	    <table class='summary'>
	    <tr>
		    <th>Job</th>
		    <th>Category</th>
		    <th>Last Success</th>
		    <th>Successes</th>
		    <th>Last Failure</th>
		    <th>Failures</th>
            <th>Duration (hh:mm:ss)</th>
	    </tr>
	"
	
    $headlinealert = $false
    $headsUp = $false

    if($agent_running){
        if( @($obj).Count -gt 0){
		    $obj | ForEach {
		
			    $job = $_.Job
                ## Modified 13/09/18 - GFF
                ## Test for new QueryError value created in Get-SQLFailedjobs.
                if ($_.QueryError -eq $true){
                    $html += "<tr class='warning'>"
                    $html += "<td colspan='6' class='disabled'" + ">" + $job + "</td>"
                }
                else{
                    if ($_.Alert -eq 1){
                        $headlinealert = $true
                        $html += "<tr class='warning'>"
			        } 
                    elseif ((($_.Alert -eq 2) -or ($_.Alert -eq 3)) -and (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL"))){
                        $html += "<tr class='headsup'>"	
                        $headsup = $true	
                    }
                    else {
                        if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
                            $html += "<tr>"
                        }				        
			        }

                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($_.Alert -eq 1) -and (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")))){

                        if ($_.Alert -eq 3){
                            $note = $notesB | Where-Object {$_.Job -eq $job }
                        }
                        else{
                            $note = $notes | Where-Object {$_.Job -eq $job }
                        }
                        
                        $error = $errors | Where-Object {$_.Job -eq $job }
                        
                        $html += "
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Job+$(if(!($note -eq $null)){"<sup> "+$note.ID+"</sup>"})+$(if(!($error -eq $null)){"<sup> "+[char]($error.ID+64)+"</sup>"})+"</td>
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Category+"</td>
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+(Format-DateTime $_.LastSuccess)+"</td>
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Successes+"</td>
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+(Format-DateTime $_.LastFailure)+"</td>
			            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+ $_.Failures + "</td>            
                        <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Duration+"</td>
			            "			

                        $html += "</tr>"
                    }
			        
                }			
			    
		    } #end of for-each
	    }
        else{
            $html += "<tr><td colspan='6'>No jobs run</td></tr>"
        }
	} 
    else {
		$html += "<tr class='warning'><td colspan='6'>Agent not  running</td></tr>"
	}

## Modified 29/12/17 - Gordon Feeney
## Amended to look for jobs to be highlighted rather than flagged as errors

    $html += "</table>"

        #flag to stop us including a notes heading more than once. 
    $jobnotesheading = $false

       if ($headlinealert -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL"))){
        if($errors.Count -gt 0){
	
		    $html += "<div class='headsupnote'>"
		    $html += "Job Failure Error Messages"
            $html += "</div><table class='notenormal'>"
		
		    $errors | ForEach{
                $job = $_.Job
                $alertObject = $obj | Where-Object {$_.Job -eq $job }                
                $alertvalue = $alertObject.Alert
                if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or ($alertvalue -eq 1)){

			        $html += "<tr>"
		
                    $html+="<td><sup>"+[char]($_.ID + 64)+"</sup> "+$_.ErrorMsg+"</td>"
			
			        $html += "</tr>"
                }	
		    }
		
		    $html += "</table>"
		
	    }


    if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL") ){
 ##       if ($headsup) {
                if( $jobnotesheading -eq $false){
                        $html += "<div class='headsupnote'>"
		                $html+="Notes"
                        $html += "</div>"
                        $jobnotesheading = $true 
                }
	        $html += "<div class='notenormal'><i>Jobs in italics</i> have been disabled in the config file and can be ignored.</div>"
 	        $html += "<div class='notenormal'>Jobs with failures that have subsequently succeeded can be ignored.</div>"
		    $html += "<div class='notenormal'>(Pro-DBA) Cycle Error Log job failures are treated as warnings as these sometimes fail due to locked log files.</div>"
  	        $html += "<div class='notenormal'>Disabled jobs with failures are omitted.</div>"
##	       }
        }

    
    if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
        $noteCount = $notes.Count

##Only put the header row in if we actually have some notes
$notesFromConfigFlag = 0

        if($notes.Count -gt 0){

            $jobnotes = $notes | Where-Object {($_.Alert -eq 1) -or ($_.Alert -eq 2) -or ($_.Alert -eq 3)} 
            $jobnotes | ForEach {
                if($notesFromConfigFlag -eq 0 ) {
                    $notesFromConfigFlag  = 1
                    $html += "<div class='headsupnote'>"
		            $html+="Job Exclusion Notes (from config.xml)"
		            $html += "</div>"
                    $html += "<table class='notenormal'>"		
                }
                $html += ""
                $html+="<sup>" + $_.ID + "</sup> " + $_.Note + " (" + $_.Job + ")" + "</div>"
            }
        } 

  <#        if($notesB.Count -gt 0){
            $notesB | ForEach {	
                if($notesFromConfigFlag -eq 0 ) {
                    $notesFromConfigFlag  = 1
                    $html += "<div class='headsupnote'>"
		            $html+="Job Exclusion Notes (from config.xml)"
		            $html += "</div>"
                    $html += "<table class='notenormal'>"		
                }	
                $html += ""
                if (![string]::IsNullOrEmpty($_.Note)){
                    $footNote = "$($_.Note)."
                    $footNote += " (" + $_.Job + ")"
                    $html+="<div class='notenormal'><sup>" + $_.ID + "</sup> " + $footNote  + "</div>"
                }
            }
                    $html+="<br/>"
        $html += "</table>"
    }
#>
   
        #GFF: 05/12/19
        #Jobs that have succeeded but with step failures. GFF
        if($notesB.Count -gt 0){
               $html += "<div class='headsupnote'>"
		       $html+="Jobs which succeeded but with step failures"
               $html += "</div>"
               
	
            $notesB | ForEach {		
                $html += ""
                $footNote = "(" + $_.Job + ")"
            <#    if (![string]::IsNullOrEmpty($_.Note)){
                    $footNote += ": $($_.Note)."
                }
            #>
                $html+="<div class='notenormal'><sup>" + $_.ID + "</sup> "  + $footNote  + "</div>"
            }
            $html+="<br/>"
	    }        
	}

 
	}

    return $html

}


## Modified /11/2017 - Gordon Feeney
## Added filter to exclude volumes prefixed with \\.
function Get-ServerDiskSpace($server,$config){
	
    ## Read in default values
    $defaultunittype = $config."default-unit-type"
    $defaultunitvalue = $config."default-unit-value"
    
	$config = $config.volume | ForEach {
		
        New-Object PSObject -Property @{
			Name = $_.name
			ThresholdUnitType = $_."unit-type"
			ThresholdUnitValue = $_."unit-value"
			Enabled = [bool]([int]$_.enabled)
			Notes = $_.notes
		}
		
	} 
		
        
	#Get-WmiObject Win32_Volume -Filter "DriveType='3'" -ComputerName $server | ForEach {
    #Exclude volumes prefixed with \\.
    Get-WmiObject Win32_Volume -ComputerName $server | Where { $_.drivetype -eq '3' -and $_.Name -notlike "\\?\Volume*"} | foreach{
		
        $name = $_.Name
		$label = $_.Label
		$freespace = [int64]$_.FreeSpace
		$capacity = [int64]$_.Capacity
		$freespacethreshold = [System.Nullable``1[[System.Int64]]] $null
		$enabled = $null
		$notes = $null
		
        #Populate the drive object with default values initially in case there isn't a config item for some of them
        $ThisThresholdUnitType = $defaultunittype
        $thisunitvalue = $defaultunitvalue

        if($ThisThresholdUnitType -eq "MB"){			
			$freespacethreshold = ([int64]$thisunitvalue)*1024*1024
		}			
		elseif($ThisThresholdUnitType -eq "GB"){
			$freespacethreshold = ([int64]$thisunitvalue)*1024*1024*1024
		}			
		elseif($ThisThresholdUnitType -eq "PERCENT"){
			$freespacethreshold = ([int64]$capacity)*([int64]$thisunitvalue/100)
		}

        $config |  Where-Object {$_.Name -eq $name} | ForEach {
            
            ## Modified on 13/10/17 by G Feeney
            ## Over-ride defaults with volumne-specific values

            if($_.ThresholdUnitType){
                $ThisThresholdUnitType = $_.ThresholdUnitType
                $thisunitvalue = [int64]$_.ThresholdUnitValue
            }
            else{
                $ThisThresholdUnitType = $defaultunittype
                $thisunitvalue = $defaultunitvalue
		    }         

            if($ThisThresholdUnitType -eq "MB"){			
				$freespacethreshold = ([int64]$thisunitvalue)*1024*1024
			}			
			elseif($ThisThresholdUnitType -eq "GB"){
				$freespacethreshold = ([int64]$thisunitvalue)*1024*1024*1024
			}			
			elseif($ThisThresholdUnitType -eq "PERCENT"){
				$freespacethreshold = ([int64]$capacity)*([int64]$thisunitvalue/100)
			}

            if(![string]::IsNullOrEmpty($_.Enabled) -And $_.Enabled){
				$enabled = [bool]1
			} elseif (![string]::IsNullOrEmpty($_.Enabled) -And !($_.Enabled)){
				$enabled = [bool]0
			}
			
            $notes = $_.Notes
            
		}
		
		New-Object PSObject -Property @{
			Name = $name
			Label = $label
			FreeSpace = $freespace
            ThresholdSpace = $freespacethreshold
			Capacity = $capacity
			Enabled = $enabled
			Notes = $notes
		}
	} | Select-Object `
	Name `
	,Label `
	,Enabled `
	,Notes `
	,@{Name="FreeSpacePercent";Expression={([Math]::Round(($_.FreeSpace/$_.Capacity)*100,2))}} `
	,@{Name="FreeSpaceGB";Expression={([Math]::Round($_.FreeSpace/1GB,2))}} `
    ,@{Name="ThresholdSpacePercent";Expression={([Math]::Round(($_.ThresholdSpace/$_.Capacity)*100,2))}} `
	,@{Name="ThresholdSpaceGB";Expression=
		{
			if(($_.ThresholdSpace)){([Math]::Round($_.ThresholdSpace/1GB,2))}
		}
	} `
	,@{Name="CapacityGB";Expression={([Math]::Round($_.Capacity/1GB,2))}} `
	,@{Name="Alert";Expression=
		{ 
			if(($_.FreeSpace -lt $_.ThresholdSpace -Or !($_.ThresholdSpace)) -And (($_.Enabled) -Or [string]::IsNullOrEmpty($_.Enabled))){
				[bool]1
			} else {
				[bool]0
			} 
			
		}
	} 

}


function FormatHTML-ServerDiskSpace($obj){

	$notesCounter = [int]0
	$notes = @()
		
	$obj | Where-Object {!($_.Notes -eq $null) } | ForEach {
		
		$notesCounter += 1
			
		$note = New-Object PSObject -Property @{
			ID = $notesCounter
			Name = $_.Name
			Note = $_.Notes
		}
				
		$notes += $note
				
	}
		
	$html = "
	<table class='summary'>
	<tr>
		<th>Name</th>
		<th>Label</th>
		<th>Free Space (%)</th>
		<th>Free Space (GB)</th>
		<th>Threshold (%)</th>
        <th>Threshold (GB)</th>
		<th>Capacity (GB)</th>
	</tr>
	"
	
    $alert = $false

    #Sort the drives by name to make the report more readable
    #IanH - 30/04/2020
    $obj =  $obj| Sort-Object Name 

    $obj | ForEach {
	
        $name = $_.Name
	
		if($_.Alert){
            $alert = $true
			$html += "<tr class='warning'>"
		} 
        else {
            if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
			    $html += "<tr>"
            }
		}

        #If it's an alert or it's a Full Health Check (as opposed to a Headline Health Check or a Disk or Services check) then display the row; otherwise don't disply it.
        if (($_.Alert) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL"))){
            $note = $notes | Where-Object {$_.Name -eq $name }
			
		    $html += "
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Name+$(if(!($note -eq $null)){"<sup> "+$note.ID+"</sup>"})+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.Label+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.FreeSpacePercent+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.FreeSpaceGB+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.ThresholdSpacePercent+"</td>
            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.ThresholdSpaceGB+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.CapacityGB+"</td>
		    "
		
		    $html += "</tr>"
        }		
	}
	
	
	$html += "</table>"
	
	if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL") -and ($notes.Count -gt 0)){
	`
		$html += "<table class='disabled'>"
		
		$notes | ForEach{
		
			$html += "<tr>"
		
			$html+="<td><sup>"+$_.ID+"</sup> "+$_.Note+"</td>"
			
			$html += "</tr>"
		
		}
		
		$html += "</table><p>"
		
	}
	    
	return $html
}


## Modified 14/11/2017 - David McDonald
## Added an event description column to Suspect Pages & suspect pages now only return event types 1-3.

function Get-SQLSuspectPages($instance)
{


    $table =  Query-SQL $instance "
    SELECT 
        d.name AS database_name, 
        mf.file_id, mf.name AS logical_filename, sp.event_type, mf.physical_name AS physical_filename, 
        sp.page_id, sp.error_count, sp.last_update_date,
		CASE 
        WHEN sp.event_type = 1
            THEN '823 error caused by an operating system CRC error or 824 error other than a bad checksum or a torn page'
        WHEN sp.event_type = 2
            THEN 'Bad checksum'
        WHEN sp.event_type = 3
            THEN 'Torn Page'
        WHEN sp.event_type = 4
            THEN 'Restored (The page was restored after it was marked bad)'
        WHEN sp.event_type = 5
            THEN 'Repaired (DBCC repaired the page)'
        WHEN sp.event_type = 7
            THEN 'Deallocated by DBCC'
    END AS event_description
    FROM 
	    msdb.dbo.suspect_pages sp 
		    INNER JOIN 
	    sys.master_files mf ON mf.database_id = sp.database_id AND mf.file_id = sp.file_id
		    INNER JOIN 
	    sys.databases d ON d.database_id = mf.database_id
    WHERE sp.event_type < 4
    "
    

    $table | ForEach-Object {
	       
		$dbname = $_.database_name
        $fileid = $_.file_id
        $logicalfile = $_.logical_filename
        $physicalfile = $_.physical_filename
		$pageid = $_.page_id
		$errorcount = $_.error_count
        $lastupdate = $_.last_update_date
        $eventdesc = $_.event_description
        $eventtype = $_.event_type
		
		New-Object PSObject -Property @{
			DatabaseName = $dbname
            FileID = $fileid
			LogicalFileName = $logicalfile
			PhysicalFileName = $physicalfile
			PageID = $pageid
			ErrorCount = $errorcount
			LastUpdateDate = $lastupdate
            EventDesc = $eventdesc
		} | Select-Object `
		DatabaseName `
        ,FileID `
		,LogicalFileName `
		,PhysicalFileName `
		,PageID `
		,ErrorCount `
		,LastUpdateDate `
        ,EventDesc `
	    ,@{Name="Alert";Expression=
			{ 			
			if(($_.errorcount -gt 0))
                {
				    [bool]1
			    } else 
                {
				    [bool]0
			    }	
					
			}		
		}	
	}    	
}


function FormatHTML-SQLSuspectPages($obj, $config){

    $html = "
	<table class='summary'>
	"
	
    if($obj.count -eq 0){
        $html += "<tr><td>No suspect pages</td></tr>"
    }
    else{   
        $html += "<tr>
		        <th>Database</th>
		        <th>Logical File</th>
		        <th>Physical File</th>
		        <th>Page ID</th>
		        <th>Error Count</th>
		        <th>Last Update</th>
                <th>Event Description</th>
	        </tr>	
            "        
		         
	    $obj | ForEach {
	
            if($_.Alert){
			    $html += "<tr class='warning'>"
		    } else {
			    $html += "<tr>"
		    }

		    #$note = $notes | Where-Object {$_.Database -eq $db }

        
            $html += "
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.DatabaseName+$(if(!($note -eq $null)){"<sup> "+$note.ID+"</sup>"})+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.LogicalFileName+"</td>
            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.PhysicalFileName+"</td>
            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.PageID+"</td>
            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.ErrorCount+"</td>
		    <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+(Format-DateTime $_.LastUpdateDate)+"</td>
            <td"+$(if(!($_.Enabled) -And !([string]::IsNullOrEmpty($_.Enabled)) ){" class='disabled'"})+">"+$_.EventDesc+"</td>
		    "
		
		    $html += "</tr>"
	    }
	
	    $html += "</table>"
	    
        if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
            $html += "<table class='disabled'>"

            $html += "<tr><td colspan='2'>Execute DBCC PAGE to identify corruption: " + $config.dbcc_page.command+"</td></tr>"

            $html += "<tr><td colspan='2'>&nbsp;</td></tr>"
			
	        $html += "<tr><td colspan='2'>Print options</td></tr>"


            $printopt = $config.print_opts.print_opt | ForEach {
		
              $html += "<tr>"
		
		        $html+="<td colspan='2'>"+$_+"</td>"
			
		        $html += "</tr>"
		
	        }
	

            $html += "<tr><td colspan='2'>&nbsp;</td></tr>"
    
            $html += "<tr><td colspan='2'>Example DBCC PAGE command: " + $config.dbcc_page_example.command+"</td></tr>"
        }
    }
        		
	$html += "</table>"

    return $html

}


function Get-SQLErrorLog($instance,$version,$config){

	$results = @()

    <#
	$exclude = $config.exclude.message | ForEach {
		
		New-Object PSObject -Property @{
			Message = $_
		}
		
	}
    #>

    $intervalMinutes = $config.intervalMinutes
    if([string]::IsNullOrEmpty($intervalMinutes)){
        $intervalMinutes = 1440
    }
    #Write-Host "intervalMinutes : $intervalMinutes "


	$query = "
	    /*
        Adapted from an original script by Pablo Echeverria at https://www.mssqltips.com/sqlservertip/5140/read-all-errors-and-warnings-in-the-sql-server-error-log-for-all-versions/ 
        */

        SET NOCOUNT ON

        SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;


        -- Load setup
        CREATE TABLE #InclusionList ([StringValue] VARCHAR(max));

        -- Declare variables
        DECLARE @IntervalMinutes INT, @DateStart DATETIME, @DateEnd DATETIME, @TraceFile VARCHAR(200)
        DECLARE @MinimumSeverity INT, @MaximumSeverity INT;

        SET @MinimumSeverity = 1;
        SET @MaximumSeverity = 25;


        -- Prepare severities list
        DECLARE @Severities AS TABLE(Severity INT);
        WHILE @MinimumSeverity <= @MaximumSeverity
        BEGIN
	        INSERT INTO @Severities
	        VALUES(@MinimumSeverity);
	        SET @MinimumSeverity = @MinimumSeverity + 1;
        END


        --Error log table
        CREATE TABLE #Info (ID INT NOT NULL IDENTITY (1, 1), [LogDate] DATETIME, [ProcessInfo] VARCHAR(1000), [Error] VARCHAR(7000))
        ALTER TABLE #Info ADD CONSTRAINT PK_Info PRIMARY KEY CLUSTERED (ID)
        CREATE NONCLUSTERED INDEX IX_LogDate ON #Info (LogDate);


        -- Set configuration values
        SET @IntervalMinutes = $intervalMinutes --1440 --1 day
        SET @DateEnd = GETDATE()
        SET @DateStart = DATEADD(mi, -@IntervalMinutes, @DateEnd)


        -- Read error log
        INSERT INTO #Info EXEC [xp_readerrorlog] 0, 1, NULL, NULL, @DateStart, @DateEnd;
        -- Get the second log in case we've just rolled over
        INSERT INTO #Info EXEC [xp_readerrorlog] 1, 1, NULL, NULL, @DateStart, @DateEnd;


        --Inclusions
        INSERT INTO #InclusionList VALUES (N'The client was unable to reuse a session with SPID %, which had been reset for connection pooling');
        INSERT INTO #InclusionList VALUES (N'BACKUP failed to complete the command BACKUP DATABASE');
        INSERT INTO #InclusionList VALUES (N'BackupVirtualDeviceFile::SendFileInfoBegin%failure on backup device%Operating system error 995');
        INSERT INTO #InclusionList VALUES (N'BackupVirtualDeviceFile::TakeSnapshot%failure on backup device%Operating system error 995');
        INSERT INTO #InclusionList VALUES (N'AppDomain % is marked for unload due to common language runtime');
        INSERT INTO #InclusionList VALUES (N'AppDomain % is marked for unload due to memory pressure');
        INSERT INTO #InclusionList VALUES (N'An error occurred in a Service Broker/Database Mirroring transport connection endpoint, Error: 8474, State:');
        INSERT INTO #InclusionList VALUES (N'FlushCache: cleaned up % bufs with % writes in % ms (avoided % new dirty bufs) for db');
        INSERT INTO #InclusionList VALUES (N'Warning Master Merge operation was not done for dbid %, objid %, so querying index will be slow. Please run alter fulltext catalog reorganize');
        INSERT INTO #InclusionList VALUES (N'SQL Server has encountered % occurrence(s) of I/O requests taking longer than 15 seconds to complete');
        INSERT INTO #InclusionList VALUES (N'Remote harden of transaction % in database % failed');
        INSERT INTO #InclusionList VALUES (N'The state of the local availability replica in availability group%has changed');
        INSERT INTO #InclusionList VALUES (N'Unsafe assembly');
        INSERT INTO #InclusionList VALUES (N'Replication-Replication Transaction-Log Reader Subsystem: agent % scheduled for retry. The process could not execute ''sp_replcmds''');
        INSERT INTO #InclusionList VALUES (N'Replication-Replication Distribution Subsystem: agent % scheduled for retry');
        INSERT INTO #InclusionList VALUES (N'A connection timeout has occurred on a previously established connection to availability replica % Either a networking or a firewall issue exists or the availability replica has transitioned to the resolving role');
	  INSERT INTO #InclusionList VALUES (N'Login failed for user%Password did not match%');
	  INSERT INTO #InclusionList VALUES (N'Login failed for user%Failed to open the explicitly specified database%');
	  INSERT INTO #InclusionList VALUES (N'Login failed for user%Reason: Could not find a login matching the name provided%');
        INSERT INTO #InclusionList VALUES (N'%Stack Dump being sent%');


        DELETE [i] FROM #Info [i] LEFT OUTER JOIN #InclusionList [li] ON [i].[Error] LIKE '%'+[li].[StringValue]+'%' WHERE [li].[StringValue] IS NULL;


        --Replacements
        UPDATE #Info
        SET Error = 'The client was unable to reuse a session with SPID ..., which had been reset for connection pooling. The failure ID is 8. This error may have been caused by an earlier operation failing. Check the error logs for failed operations immediately before this error message.'
        WHERE Error LIKE 'The client was unable to reuse a session with SPID %, which had been reset for connection pooling%';

        UPDATE #Info
        SET Error = 'BACKUP failed to complete .... Check the backup application log for detailed messages.'
        WHERE Error LIKE 'BACKUP failed to complete the command BACKUP DATABASE %';

        UPDATE #Info
        SET Error = 'BackupVirtualDeviceFile::SendFileInfoBegin: failure on backup device ..... Operating system error 995(The I/O operation has been aborted because of either a thread exit or an application request.).'
        WHERE Error LIKE 'BackupVirtualDeviceFile::SendFileInfoBegin%failure on backup device%Operating system error 995%';

        UPDATE #Info
        SET Error = 'BackupVirtualDeviceFile::TakeSnapshot: failure on backup device ..... Operating system error 995(The I/O operation has been aborted because of either a thread exit or an application request.).'
        WHERE Error LIKE 'BackupVirtualDeviceFile::TakeSnapshot%failure on backup device%Operating system error 995%';

        UPDATE #Info
        SET Error = 'AppDomain .... is marked for unload due to common language runtime (CLR) or security data definition language (DDL) operations.'
        WHERE Error LIKE 'AppDomain % is marked for unload due to common language runtime%';

        UPDATE #Info
        SET Error = 'AppDomain ..... is marked for unload due to memory pressure.'
        WHERE Error LIKE 'AppDomain % is marked for unload due to memory pressure%';

        UPDATE #Info
        SET Error = 'An error occurred in a Service Broker/Database Mirroring transport connection endpoint, Error: 8474, State: .... (Near endpoint role: Target, far endpoint address: '')'
        WHERE Error LIKE 'An error occurred in a Service Broker/Database Mirroring transport connection endpoint, Error: 8474, State: %';

        UPDATE #Info
        SET Error = 'FlushCache: cleaned up bufs'
        WHERE Error LIKE 'FlushCache: cleaned up % bufs with % writes in % ms (avoided % new dirty bufs) for db %';

        UPDATE #Info
        SET Error = 'Warning Master Merge operation was not done for dbid ..., objid ......., so querying index will be slow. Please run alter fulltext catalog reorganize'
        WHERE Error LIKE 'Warning Master Merge operation was not done for dbid %, objid %, so querying index will be slow. Please run alter fulltext catalog reorganize%';

        UPDATE #Info
        SET Error = 'SQL Server has encountered occurrences of I/O requests taking longer than 15 seconds to complete'
        WHERE Error LIKE 'SQL Server has encountered % occurrence(s) of I/O requests taking longer than 15 seconds to complete %';

        UPDATE #Info
        SET Error = 'Remote harden of transaction ..... in database ..... failed'
        WHERE Error LIKE 'Remote harden of transaction % in database % failed%';

        UPDATE #Info
        SET Error = 'The state of the local availability replica in availability group ..... has changed'
        WHERE Error LIKE 'The state of the local availability replica in availability group%has changed%';

        UPDATE #Info
        SET Error = 'Replication-Replication Distribution Subsystem: agent ..... scheduled for retry. Query timeout expired, Failed Command'
        WHERE Error LIKE 'Replication-Replication Distribution Subsystem: agent % scheduled for retry. Query timeout expired, Failed Command%';

        UPDATE #Info
        SET Error = 'Replication-Replication Distribution Subsystem: agent ..... scheduled for retry. The process could not execute ''sp_replcmds'''
        WHERE Error LIKE 'Replication-Replication Transaction-Log Reader Subsystem: agent % scheduled for retry. The process could not execute ''sp_replcmds''%';

        UPDATE #Info
        SET Error = 'A connection timeout has occurred on a previously established connection to availability replica .....  Either a networking or a firewall issue exists or the availability replica has transitioned to the resolving role'
        WHERE Error LIKE 'A connection timeout has occurred on a previously established connection to availability replica % Either a networking or a firewall issue exists or the availability replica has transitioned to the resolving role%';

        UPDATE #Info
        SET Error = LEFT(Error, CASE WHEN CHARINDEX('[CLIENT', Error) = 0 THEN LEN(Error) ELSE CHARINDEX('[CLIENT', Error) - 2 END)
        WHERE Error LIKE 'Login failed for user%Could not find a login%';

        UPDATE #Info
        SET Error = LEFT(Error, CASE WHEN CHARINDEX('[CLIENT', Error) = 0 THEN LEN(Error) ELSE CHARINDEX('[CLIENT', Error) - 2 END)
        WHERE Error LIKE 'Login failed for user%Password did not match%';


		--Finish
        WITH CTE_HiSeverity AS
        (
	        --Severity 16 and above
	        SELECT i2.ID AS ID2, i1.ID AS ID1, i2.LogDate, i2.ProcessInfo, sev.Severity, i2.[Error] AS Error
	        FROM 
		        #Info i1
			        INNER JOIN 
		        @Severities as sev ON i1.Error LIKE N'%Severity: ' + CAST(sev.Severity AS nvarchar(2)) + N'%'
			        INNER JOIN 
		        #Info i2 ON i2.ProcessInfo = i1.ProcessInfo
			        AND i2.LogDate = i1.LogDate 
			        AND i2.Error <> i1.Error
			        AND i2.ID = i1.ID + 1
	        WHERE sev.Severity > 15			
        ),
        --CTE_LoSeverity AS
        --(
        --	--Severity below 16
        --	SELECT i2.ID AS ID2, i1.ID AS ID1, i2.LogDate, i2.ProcessInfo, sev.Severity, i1.Error AS i1_Error, i2.[Error] AS Error
        --	FROM 
        --		#Info i1
        --			INNER JOIN 
        --		@Severities as sev ON i1.Error LIKE N'%Severity: ' + CAST(sev.Severity AS nvarchar(2)) + N'%'
        --			INNER JOIN 
        --		#Info i2 ON i2.ProcessInfo = i1.ProcessInfo
        --			AND i2.LogDate = i1.LogDate 
        --			AND i2.Error <> i1.Error
        --			AND i2.ID = i1.ID + 1
        --	WHERE sev.Severity < 16	
        --), 
        CTE_NoSeverity AS
        (
	        --Non-severity	
	        SELECT i.ID, i.LogDate, i.ProcessInfo, NULL AS Severity, i.Error
	        FROM 
		        #Info i
			        LEFT OUTER JOIN 
		        CTE_HiSeverity hs ON hs.ID1 = i.ID OR hs.ID2 = i.ID 		
	        WHERE hs.ID1 IS NULL AND hs.ID2 IS NULL 
		        AND i.Error NOT LIKE '%Severity%'
        )


        SELECT 
	        --ProcessInfo, 
	        Error AS [Text], 
			CASE 
				WHEN Error LIKE '%***Stack Dump%' THEN 1 
                ELSE 2
			END AS Alert, 
			COUNT(*) AS [Count], MAX(LogDate) AS [LastOccurred]
        FROM 
        (
	        SELECT 
		        LogDate,
		        --There are lots of errors with ProcessInfo spidnnn so convert them into SPID so they can be grouped.
		        --CASE 
		        --	WHEN ProcessInfo LIKE 'spid%' THEN 'SPID'
		        --	ELSE ProcessInfo
		        --END AS ProcessInfo, 
		        Severity, 'Severity: ' + CAST(Severity As nvarchar) + ' - ' + Error AS Error, ID1 as ID
	        FROM CTE_HiSeverity
	        UNION --ALL
	        SELECT 
		        LogDate, 
		        --CASE 
		        --	WHEN ProcessInfo LIKE 'spid%' THEN 'SPID'
		        --	ELSE ProcessInfo
		        --END AS ProcessInfo, 
		        NULL AS Severity, Error, ID
	        FROM CTE_NoSeverity
        ) ErrorLog
        GROUP BY 
	        --ProcessInfo, 
	        Error
        ORDER BY [LastOccurred];


        -- Unload setup
        DROP TABLE #Info;
        DROP TABLE #InclusionList;
	"


    ## Modified 13/09/18 - GFF
    ## Encapsulated query in try-catch block, created separate object in Catch block and added QueryError value to both Try and Catch objects
    try{
        $table =  Query-SQL $instance $query 120 "Get-SQLErrorLog"
        
	    if(($table)){
	
		    $table | ForEach {
                $error = $_.Text
                $alert = $_.Alert
                $results += New-Object PSObject -Property @{
                    QueryError = $false
				    ObjectType = "ErrorLog"
				    Text = $_.Text
                    Count = $_.Count
				    LastOccurred = $_.LastOccurred
                    Alert = $alert
				
			    }
		    }
	    }
    }

    catch{
        $results += New-Object PSObject -Property @{
				    QueryError = $true
				    ObjectType = "ErrorLog"
				    Text = $_.Exception.Message				    
        }
    }

    return $results
		
}


function FormatHTML-SQLErrorLog($obj){
		
	$html = "
	<table class='summary'>
	<tr>
		<th>Error</th>
		<th>Count</th>
		<th>Last Occurred</th>
	</tr>
	"
		
	$count = @($obj | Where-Object {$_.ObjectType -eq "ErrorLog"}).Count

	if($count -gt 0 ) {
	
		$obj | Where-Object {$_.ObjectType -eq "ErrorLog" } | ForEach {

            ## Modified 13/09/18 - GFF
            ## Test for new QueryError value created in Get-SQLErrorLog.
            if ($_.QueryError -eq $true){
                $html += "<tr class='warning'>"
			    $html += "
			    <td colspan='3' class='disabled'>" + $_.Text + "</td>			    
			    "					    
            }
            else{
                $error = $_.Text
                if ($_.Alert -eq 1){
                    $html += "<tr class='warning'>"
                }
                $html += "
			    <td>"+$_.Text+"</td>
			    <td>"+$_.Count+"</td>
			    <td>"+(Format-DateTime($_.LastOccurred))+"</td>
			    "
            }

            $html += "</tr>"
		}
	
	} else {
		
		$html += "
		<td colspan='3'>No errors found</td>
		"
	}
	
	$html += "</table>"
	
    $html += "<table class='disabled'><tr><td>Errors with severity levels less than 16 are excluded</td></tr><table>"
	
	return $html

}


function FormatHTML-ReportHeader($hostName,$smtpserver,$smtpPort,$username,$start,$end,$version,$to,$cc,$body,$serverSubject){

## 13/04/2020 IanH
## Modified mailto syntax 



    $html = "
	<table class='summary'>
	<tr>
		<th>Property</th>
		<th>Value</th>
	</tr>
	<tr>
		<td>Host</td>
		<td>$hostName</td>
	</tr>
	<tr>
		<td>Config File</td>
		<td>$config</td>
	</tr>
	<tr>
		<td>Mail Server</td>
		<td>$smtpserver`: $smtpPort</td>
	</tr>
	<tr>
		<td>User</td>
		<td>$username</td>
	</tr>
	<tr>
		<td>Start</td>
		<td>$(Format-DateTime $start)</td>
	</tr>
	<tr>
		<td>End</td>
		<td>$(Format-DateTime $end)</td>
	</tr>
	<tr>
		<td>Commit</td>
		<td>"+$version.Commit+"</td>
	</tr>
	<tr>
		<td>Tag</td>
		<td>"+$version.Tag+"</td>
	</tr>
	<tr>
		<td>Contact</td>
		<td><a href='mailto:$to&#63;cc=$cc&subject=SQL Daily Health Check - $serverSubject&body=$body'>Report Issue</a></td>
	</tr>
	</table>
	"
	
	return $html

}


function FormatHTML-ReportSummary($summary){
##Debug "Write-Host" - IanH
##Write-Host ($summary | Format-Table | Out-String)
		
    $html = "
	<table class='summary'>
	<tr>
		<th>Name</th>
		<th>Type</th>
		<th>Category</th>
		<th>Errors</th>
	</tr>
	"+$($summary | ForEach{

            ## Modified 29/12/17 - G Feeney
            ## Summary for Jobs category can be highlighted as well as flagged for errors
            ## Modified 25/07/18 Ian H
            ## Summary for instance can be highligted as well as flagged for errors
            ## (If restart during maintenance window - highlight, unless Agent also stopped in which case error) 

            #"<tr"+$(
            #    if(($_.InstanceHeadsUp -ge 1) -And ($_.Errors -le 1) -And !$_.Alert){" class='headsup'"}elseif ($_.Alert){" class='warning'"}elseif($_.HeadsUp ){" class='headsup'"})+">
            
            $class = ""

            #if ((($_.Category -eq 'ErrorLog') -and ($_.Errors -gt 0)) -or (($_.InstanceHeadsUp -ge 1) -And ($_.Errors -le 1) -And !$_.Alert) -or $_.HeadsUp){
            if ((($_.Category -eq 'ErrorLog') -and ($_.Errors -gt 0)) -or $_.Alert){
                $class = " class='warning'"
            }
            elseif ((($_.InstanceHeadsUp -ge 1) -And ($_.Errors -le 1) -And !$_.Alert) -or $_.HeadsUp){
                $class = " class='headsup'"
            }
            "<tr"+ $class +">
			<td>"+$_.Name+"</td>
			<td>"+$_.Type+"</td>
			<td>"+$_.Category+"</td>
			<td>"+$(
            if ($_.Category -ne 'Job') {
                if ($_.Category -eq 'ErrorLog') {
                    if ($_.Errors -eq 0 -and $_.HeadsUp -eq 0) {
                        "OK"
                    }
                    else {
                        if ([string]::IsNullOrEmpty($_.Errors)) {
                            "_"
                        }
                        elseif ([string]::IsNullOrEmpty($_.HeadsUp)) {
                            "_"
                        }
                        else {
                            $_.Errors + $_.HeadsUp
                        }
                    }
                }
                elseif ($_.Category -ne 'Instance') {
                    if ($_.Errors -eq 0) {
                        "OK"
                    }
                    else {
                        if ([string]::IsNullOrEmpty($_.Errors)) {
                            "_"
                        }
                        else {
                            $_.Errors
                        }
                    }
                 }
                 else   ## Category = Instance, check for InstanceHeadsUp
                 {
                    if ($_.Errors -ge 1)
                    {
                        $_.Errors
                    }
                    elseif ($_.Errors -eq 0) 
                    {
                            "OK"
                    }
                    elseif ($_.InstanceHeadsUp -eq 1)
                    {
                            "OK"
                    }
                 }
            }
            else {
                if ([string]::IsNullOrEmpty($_.Errors)) {
                    "_"
                }
                elseif ($_.Errors -gt 0) {
                    $_.Errors                   
                }
                elseif  ($_.HeadsUp -gt 0)  {
                    #$_.HeadsUp
                    "OK"
                }
                elseif ($_.Errors -eq 0) {
                    "OK"
                }
            }

			)+"</td>
			</tr>"
		
	})

	$html += "</table><p>"
	
	return $html

}


function FormatHTML-InstanceErrorsSummary($summary){
##Debug "Write-Host" - IanH
##Write-Host ($summary | Format-Table | Out-String)
	
    $summary_errors = $summary | Where-Object {($_.Category -eq 'Connect') -or ($_.Category -eq 'Permissions')}

    if (@($summary_errors).count -gt 0){

        $html = "
	    <table class='summary'>
	    <tr>
		    <th>Name</th>
		    <th>Type</th>
		    <th>Category</th>
		    <th>Errors</th>
	    </tr>
	    "+$($summary | Where-Object {($_.Category -eq 'Connect') -or ($_.Category -eq 'Permissions')} | ForEach{
            
                if ((($_.Category -eq 'Connect') -or ($_.Category -eq 'Permissions')) -and (($_.Errors -gt 0) -or $_.Alert)){

                    "<tr class='warning'>

			        <td>"+$_.Name+"</td>
			        <td>"+$_.Type+"</td>
			        <td>"+$_.Category+"</td>
			        <td>"+$(
                    if ($_.Category -ne 'Job') {

                        if ($_.Category -ne 'Instance') {
                                if ($_.Errors -eq 0) {
                                    "OK"
                                }
                                else {
                                    if ([string]::IsNullOrEmpty($_.Errors)) {
                                        "_"
                                    }
                                    else {
                                        $_.Errors
                                    }
                                }
                         }
                         else   ## Category = Instance, check for InstanceHeadsUp
                         {
                            if ($_.Errors -ge 1)
                            {
                                $_.Errors
                            }
                            elseif ($_.Errors -eq 0) 
                            {
                                    "OK"
                            }
                            elseif ($_.InstanceHeadsUp -eq 1)
                            {
                                    "OK"
                            }
                         }
                    }
                    else {
                        if ([string]::IsNullOrEmpty($_.Errors)) {
                            "_"
                        }
                        elseif ($_.Errors -gt 0) {
                            $_.Errors                   
                        }
                        elseif  ($_.HeadsUp -gt 0)  {
                            #$_.HeadsUp
                            "OK"
                        }
                        elseif ($_.Errors -eq 0) {
                            "OK"
                        }
                    }

			        )+"</td>
			        </tr>"
		    }
	    })
    }

    $html += "</table><p>"
	
	return $html

}


function FormatHTML-ServerErrorsSummary($summary){
##Debug "Write-Host" - IanH
##Write-Host ($summary | Format-Table | Out-String)
	
    $summary_errors = $summary | Where-Object {($_.Category -eq 'WMI')}    
    if (@($summary_errors).count -gt 0){
        $html = "
	    <table class='summary'>
	    <tr>
		    <th>Name</th>
		    <th>Type</th>
		    <th>Category</th>
		    <th>Errors</th>
	    </tr>
	    "+$($summary | Where-Object {($_.Category -eq 'WMI')} | ForEach{
            
                if ((($_.Category -eq 'WMI')) -and (($_.Errors -gt 0) -or $_.Alert)){

                    "<tr class='warning'>

			        <td>"+$_.Name+"</td>
			        <td>"+$_.Type+"</td>
			        <td>"+$_.Category+"</td>
			        <td>"+$(
                    if ($_.Category -ne 'Job') {

                        if ($_.Category -ne 'Instance') {
                                if ($_.Errors -eq 0) {
                                    "OK"
                                }
                                else {
                                    if ([string]::IsNullOrEmpty($_.Errors)) {
                                        "_"
                                    }
                                    else {
                                        $_.Errors
                                    }
                                }
                         }
                         else   ## Category = Instance, check for InstanceHeadsUp
                         {
                            if ($_.Errors -ge 1)
                            {
                                $_.Errors
                            }
                            elseif ($_.Errors -eq 0) 
                            {
                                    "OK"
                            }
                            elseif ($_.InstanceHeadsUp -eq 1)
                            {
                                    "OK"
                            }
                         }
                    }
                    else {
                        if ([string]::IsNullOrEmpty($_.Errors)) {
                            "_"
                        }
                        elseif ($_.Errors -gt 0) {
                            $_.Errors                   
                        }
                        elseif  ($_.HeadsUp -gt 0)  {
                            #$_.HeadsUp
                            "OK"
                        }
                        elseif ($_.Errors -eq 0) {
                            "OK"
                        }
                    }

			        )+"</td>
			        </tr>"
		    }
	    })
    }
    
    $html += "</table><p>"
	
	return $html

}


## Modified 29/12/17 - G Feeney
## HeadsUp variable created to allow summary to highlight some jobs rather than flag them as errors

## Modified 30/10/18 - G Feeney
## Fixed bug whereby duplicate servers had the same volume summary displayed for each one instead of just once.

function Get-SQLHealth ($directory, $config){

    writeTestMessage $globalTestFlag "Get-SQLHealth - START"
    
	$executingScriptDirectory = $directory

	$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	$FQhost = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain

	$reportVersion = Import-Csv -Delimiter "|" -Path "$executingScriptDirectory\version.rel"

	$css = Get-Content "$executingScriptDirectory\configuration\main.css"
	[xml]$xml = Get-Content "$executingScriptDirectory\configuration\$config"

	$client = $xml.configuration.metadata.client
	$subject = $xml.configuration.metadata.subject
	$clientcontact = $xml.configuration.metadata."client-contact"
	$clientcontactcc = $xml.configuration.metadata."client-contact-cc"
	$body = $xml.configuration.metadata.body
    #$use365 = $xml.configuration.smtp.use365
	$smtpServer = $xml.configuration.smtp.server
	$smtpPort = $xml.configuration.smtp.port
	$ssl = $(if($xml.configuration.smtp.ssl -eq 1){$true} else {$false})
	$smtpUser = $xml.configuration.smtp.user
	$smtpPassword = $xml.configuration.smtp.password
	$mailFrom = $xml.configuration.smtp.from
	$mailTo = $xml.configuration.smtp.to

    $buildAPI = $xml.configuration.buildAPI    
    $globalCommandTimeout = $xml.configuration.settings.commandtimeout
    if([string]::IsNullOrEmpty($globalCommandTimeout)){
        $globalCommandTimeout = 60
    }
    $globalScriptTimeout = $xml.configuration.settings.scripttimeout
    if([string]::IsNullOrEmpty($globalScriptTimeout)){
        $globalScriptTimeout = 50
    }


    $is_Sqlserver = 0
    
	$xml.configuration.server | ForEach {

		# Get Data

		$server = $_.name

        writeTestMessage $globalTestFlag "Get-SQLHealth - FIRST SERVER LOOP"
    

        if(!([string]::IsNullOrEmpty($_.servercontact))){
            $server_contact = $_.servercontact
		}
        else
        {
            $server_contact = $null 
        }
<#
        if($server_contact -eq $null)
        {
            Write-Host "server contact not defined"
        }
        else
        {
            Write-Host "Server contact is " $server_contact
        }
#>       
        New-Variable -Name "$($server)_contact" -Value $server_contact

##IanH - 14 Aug 18
## Allow for alias when using FQDN for server / instance

		if([string]::IsNullOrEmpty($_.alias)){
            $server_alias = $_.name
		} 
        else
        {
            $server_alias = $_.alias
        }

        $server_connect = $_.connect
		
		if([string]::IsNullOrEmpty($_.connect)){
            $server_connect = $server
		}
		
		New-Variable -Name "$($server)_start" -Value $(Get-Date)	
		
		# Check WMI Connectivity
		
		New-Variable -Name "$($server)_wmiconnect" -Value $(Test-WMIConnect $server_connect)
		
		if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)) {
		
			# Get Data
		
			New-Variable -Name "$($server)_diskSpace" -Value $(Get-ServerDiskSpace $server_connect $_.volumes)

			# Errors
			
			New-Variable -Name "$($server)_diskSpace_errors" -Value $(@($(Get-Variable -Name "$($server)_diskSpace" -ValueOnly) | Where-Object {$_.Alert}).Count)
		
		}
		
        	
	    $_.instances.instance | ForEach {
			            
			$instance = $_.name
            
            ## If not a production instance we don't have our monitoring installed
			if(!([string]::IsNullOrEmpty($_.production))){
				$production = $_.production
			}
            else {
                $production = 1
            }

          
            writeTestMessage $globalTestFlag "Get-SQLHealth - FIRST SERVER LOOP - FIRST INSTANCE LOOP"
    

            $instance_connect = $_.connect		
			if([string]::IsNullOrEmpty($_.connect)){
				$instance_connect = $instance
			}

            
            if([string]::IsNullOrEmpty($_.connect)){
			    $server_connect = $server
		    }
			

            $is_Sqlserver = $_.SQLServer
            if ($is_Sqlserver -eq 0){
                #do nowt
            }
            else{
                # Check Connectivity
			    New-Variable -Name "$($instance)_sqlconnect" -Value $(Test-SQLConnect $instance_connect)
			
			    if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly)) {
                
				    # Check Permissions
			
				    New-Variable -Name "$($instance)_sqlpermissions" -Value $(Test-SQLPermissions $instance_connect)
				
				    if($(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)) {
				
					    # Get Data
					    New-Variable -Name "$($instance)_sqlversion" -Value $(Get-SQLVersion $instance_connect)
					    New-Variable -Name "$($instance)_sqlinstance" -Value $(Get-SQLInstance $instance_connect $(Get-Variable -Name "$($instance)_sqlversion" -ValueOnly) $_.maintenance $_.flags $production $buildAPI $directory)
                                                

    ## Modified 17/08/2017 - Ian Harris
    ## Added parameters for default backup / log thresholds
					    New-Variable -Name "$($instance)_sqldbsummary" -Value $(Get-SQLDatabaseSummary $instance_connect $(Get-Variable -Name "$($instance)_sqlversion" -ValueOnly) $_.backup)

					    New-Variable -Name "$($instance)_sqlfailedjobs" -Value $(Get-SQLFailedJobs $instance_connect $(Get-Variable -Name "$($instance)_sqlversion" -ValueOnly) $_.jobs)
                        New-Variable -Name "$($instance)_sqlsuspectpages" -Value $(Get-SQLSuspectPages $instance_connect)
					    New-Variable -Name "$($instance)_sqlerrors" -Value $(Get-SQLErrorLog $instance_connect $(Get-Variable -Name "$($instance)_sqlversion" -ValueOnly) $_.errors)
						
					    # Errors
					
    ## Modified 24/07/18 IanH
    ## Changed to UptimeAlert equal 1 or 2 as now UptimeAlert can = 0, 1 or 2

    ## Modified on 17/12/21 by GFF
    ## Added new MonitoringInstalled alert and made the variable calculation a bit more legible
                       #New-Variable -Name "$($instance)_sqlinstance_errors" -Value $(@(@($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.UptimeAlert -eq 1 -or $_.UptimeAlert -eq 2}) + @($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert -or $_.MailProfileAlert}) + @($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert -or !$_.MonitoringInstalled})).Count)                            

<# Original (for when it breaks)
                       New-Variable -Name "$($instance)_sqlinstance_errors" -Value $(@((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.UptimeAlert -eq 1 -or $_.UptimeAlert -eq 2}).Count + 
                                        @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert -or $_.MailProfileAlert}).Count + 
                                        @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {!$_.MonitoringInstalled}).Count)   
#>

                       New-Variable -Name "$($instance)_sqlinstance_errors" -Value $(@((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.UptimeAlert -eq 8 -or $_.UptimeAlert -eq 9}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.MailProfileAlert -eq 1}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.MonitoringInstalledAlert -eq 1}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.EmailErrors -eq 1}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.VersionAvailability -gt 1}).Count)

<#                       New-Variable -Name "$($instance)_sqlinstance_errors" -Value $(@((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert -or $_.MailProfileAlert}).Count + 
                                        @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {!$_.MonitoringInstalled}).Count)                 
#>

    ## Modified 09/10/18 GFF
    ## Created new '_sqlinstanceheadline_errors' variables for headline-only alerts
                        New-Variable -Name "$($instance)_sqlinstanceheadline_errors" -Value $(@(@($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.UptimeAlert -eq 1}) + @($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentAlert})).Count)                    

    ## Modified 24/07/2018 Ianh
    ## Set _sqlinstance_headsup to 1 if UptimeAlert is set to 2 
    ## (i.e. a restart has taken place during a maintenance window) 
                        New-Variable -Name "$($instance)_sqlinstance_headsup"  -Value $(@($(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.UptimeAlert -eq 1 -or $_.UptimeAlert -eq 2}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.MailProfileHeadsup -eq 1}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.AgentHeadsup}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.MaintenanceAlert -gt 0}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.VersionAvailability -lt 2}).Count + @((Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) | Where-Object {$_.CoreHeadsUp -eq 1}).Count)
					    New-Variable -Name "$($instance)_sqldbsummary_errors" -Value $(@($(Get-Variable -Name "$($instance)_sqldbsummary" -ValueOnly) | Where-Object {$_.RedBackupAlert}).Count) 
					    New-Variable -Name "$($instance)_sqlfailedjobs_errors" -Value $(@($(Get-Variable -Name "$($instance)_sqlfailedjobs" -ValueOnly) | Where-Object {$_.Alert -eq 1}).Count)
                        New-Variable -Name "$($instance)_sqlfailedjobs_headsup" -Value $(@($(Get-Variable -Name "$($instance)_sqlfailedjobs" -ValueOnly) | Where-Object {($_.Alert -eq 2) -or ($_.Alert -eq 3)}).Count)
                        New-Variable -Name "$($instance)_sqlsuspectpages_errors" -Value $(@($(Get-Variable -Name "$($instance)_sqlsuspectpages" -ValueOnly) | Where-Object {$_.Alert}).Count)                        
                        New-Variable -Name "$($instance)_sqlerrorlog_errors" -Value $(@($(Get-Variable -Name "$($instance)_sqlerrors" -ValueOnly) | Where-Object {$_.ObjectType -eq "ErrorLog" -and $_.Alert -eq 1}).Count)
                        New-Variable -Name "$($instance)_sqlerrorlog_headsup" -Value $(@($(Get-Variable -Name "$($instance)_sqlerrors" -ValueOnly) | Where-Object {$_.ObjectType -eq "ErrorLog" -and $_.Alert -eq 2}).Count)
				    }
			    }
            } #End of check for SQL Server

		} #END OF INSTANCE LOOP
		
		New-Variable -Name "$($server)_end" -Value $(Get-Date)

	} #END OF SERVER LOOP


    $html = $null

    $wmiAlertFlag = $false
    $serverAlertFlag = $false        
    #$serverAlertTotal = 0
    

    #Create array to store results of servers and instances processed
    $resultsServerArray = @()    
     

	$xml.configuration.server | ForEach {

        $server = $_.name
        $wmiconnect_errors = 0

        writeTestMessage $globalTestFlag "Get-SQLHealth - $server - SECOND SERVER LOOP"
    

        #create and populate new Server item
        $resultsServerItem = New-Object PSObject -Property @{
            Server = $server
            ServerAlerts = 0
            ServerCompleted = $false
            Instances = @()
        }

        
        
## IanH 14 Aug 18
## Allow alias for the instance name when using a FQDN

		if([string]::IsNullOrEmpty($_.alias)){
				$server_alias = $_.name
		} 
        else
        {
                $server_alias = $_.alias
        }

		$summary = @()	
		
        if ($html -eq $null){
            $html = "
		    <html>
		    <head>
		    <style>"+$css+"</style>
		    </head>
		    <body>
		    "

            $html += "<h1>SQL Health Check" + $(if($globalScriptType -eq "HEALTHCHECK" -and $globalScriptSubType -eq "HEADLINE"){" (Headline)"}else{""}) + "</h1>"


            ##Purpose of this loop is to determine if there is more than one instance in the
            ##server block of config.xml. Use to this to set the subject in the "mailto" link
            ##For one instance per server block - use server\instance 
            ##For multiple instances just user the server name
            ##IanH - 30/04/2020

            $instanceCounter = 0
            $_.instances.instance | ForEach {
			
			    $instance =  $_.name

                #We haven't got our instance alias yet, so had to paste this bit in here

                if([string]::IsNullOrEmpty($_.alias)){
				    $instance_alias = $_.name
		        } 
                else
                {
                    $instance_alias = $_.alias
                }

                $instanceCounter = $instanceCounter + 1 
                
            }

            if($instanceCounter -gt 1)  #Multiple instances so use server name
            {
                $serverSubject = $server
            }
            else                        #Only one instance, so use server\instance
            {
                $serverSubject = $instance_alias
            }

            ## New variable $clientcontact_all used to include any contacts defined at the server level v2.30 IanH
            if($(Get-Variable -Name "$($server)_contact" -ValueOnly) -eq $null){
                $clientcontactcc_all = $clientcontactcc
            }
            else
            {
                $clientcontactcc_all = $clientcontactcc + ";" + $(Get-Variable -Name "$($server)_contact" -ValueOnly)
            }
		    		
		    $html += FormatHTML-ReportHeader $FQhost $smtpserver $smtpPort $user $(Get-Variable -Name "$($server)_start" -ValueOnly) $(Get-Variable -Name "$($server)_end" -ValueOnly) $reportVersion $clientcontact $clientcontactcc_all $body $serverSubject 

		    		
		   
        }
        
        

## IanH 14 Aug 18
## Use $server_alias when specifying server name in the HTML
		
		$summary += New-Object PSObject -Property @{
			Name = $server_alias
			Type = "Host"
			Category = "WMI"
			Errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){
		
					0

				} else { 1 })
		}
				
		$summary += New-Object PSObject -Property @{
			Name = $server_alias
			Type = "Host"
			Category = "Disk"
			Errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){
		
				$(Get-Variable -Name "$($server)_diskSpace_errors" -ValueOnly)

			} else { $null })
		}		
		
		$_.instances.instance | ForEach {
			
			$instance =  $_.name

            writeTestMessage $globalTestFlag "Get-SQLHealth - $server - SECOND SERVER LOOP - $instance - FIRST INSTANCE LOOP"
    

            if([string]::IsNullOrEmpty($_.alias)){
				$instance_alias = $_.name
		    } 
            else
            {
                $instance_alias = $_.alias
            }

            $is_Sqlserver = $_.SQLServer
            if ($is_Sqlserver -eq 0){
                #do nowt
            }
            else{
         	    $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Connect"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly)){
		
					    0

				    } else { 1 })
			    }
			
			    $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Permissions"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly)){
					
					    if($(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
		
						    0

					    } else { 1 }
					

				    } else { $null })
			    }
			
                ## Modified 20/07/2018 IanH
                ## Instance Headsup property added
                ## Modified 14/08/18
                ## Using $instance_alias



			    $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Instance"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
						
					    $(Get-Variable -Name "$($instance)_sqlinstance_errors" -ValueOnly)
						
				    } else { $null })
                    InstanceHeadsUp = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){

                        $(Get-Variable -Name "$($instance)_sqlinstance_headsup" -ValueOnly)
                        				
				    } else { $null })
			    }
			
			    $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Backup"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
						
					    $(Get-Variable -Name "$($instance)_sqldbsummary_errors" -ValueOnly)
						
				    } else { $null })

			    }


			
                ## Modified 29/12/17 - G Feeney
                ##HeadsUp variable populated
			    $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Job"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
					
                        $(Get-Variable -Name "$($instance)_sqlfailedjobs_errors" -ValueOnly)
						
				    } else { $null })
                    HeadsUp = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){

                        $(Get-Variable -Name "$($instance)_sqlfailedjobs_headsup" -ValueOnly)
						
				    } else { $null })
			    }

                $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "Suspect"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
						
					    $(Get-Variable -Name "$($instance)_sqlsuspectpages_errors" -ValueOnly)
						
				    } else { $null })
			    }

                <#
                $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "ErrorLog"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
						
					    $(Get-Variable -Name "$($instance)_sqlerrorlog_headsup" -ValueOnly)
						
				    } else { $null })
			    }
                #>

                $summary += New-Object PSObject -Property @{
				    Name = $instance_alias
				    Type = "SQL Server"
				    Category = "ErrorLog"
				    Errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){
						
					    $(Get-Variable -Name "$($instance)_sqlerrorlog_errors" -ValueOnly)
						
				    } else { $null })
                    HeadsUp = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){

                        $(Get-Variable -Name "$($instance)_sqlerrorlog_headsup" -ValueOnly)
						
				    } else { $null })
			    }
   
            } #End of check for SQL Server

		} #END OF INSTANCE LOOP

		
        if ($globalScriptType -eq "HEALTHCHECK" -and $globalScriptSubType -eq "FULL"){

            writeTestMessage $globalTestFlag "Get-SQLHealth - FULL HEALTH CHECK - REPORT SUMMARY"
    
		    $html += "<h2>Report Summary</h2><br>"
                        
            $html += FormatHTML-ReportSummary $($summary | Select-Object `
			Name `
			,Type `
			,Category `
			,Errors `
            ,@{Name="Alert";Expression=				
				{ 				
				    if($_.Errors -gt 0)
                    {
					    $true
				    } else 
                        {
					        $false
				        }						
			        }			
		        } `
            ,HeadsUp `
            ,InstanceHeadsUp 
            )
        }

        
		
        $wmiconnect_errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){0}else{1})
        $diskSpace_errors = $(Get-Variable -Name "$($server)_diskSpace_errors" -ValueOnly)
        $resultsServerItem.ServerAlerts = $wmiconnect_errors + $diskSpace_errors
        #$resultsServerArray += $resultsServerItem
               

        ## IanH 14/08/18
        ## Using $server_alias & $instance_alias in place of $server & $instance
	
        $serverCompleted = $false

        $arrayCount = $resultsServerArray.Count
            
        if ($wmiconnect_errors -eq 0){
            $resultsServerSubArray = $resultsServerArray | where-object {$_.Server -eq $server}
            if ($resultsServerSubArray -ne $null){
                $subArrayCount = @($resultsServerSubArray).Count
            }
            else{
                $subArrayCount = 0                
            }
            
            if (!($resultsServerSubArray -eq $null)){
                $resultsServerSubArrayItem = $resultsServerSubArray[$resultsServerSubArray.Count - 1]
                
                $serverCompleted = $resultsServerSubArrayItem.ServerCompleted
                
                if ($serverCompleted -eq $null){
                    $serverCompleted = $false
                }
            }
                
            #We don't need to display the Volume Summary more than once if we're doing a Headline report.            
            if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and !$serverCompleted -and ($diskSpace_errors -gt 0))){
            #if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($diskSpace_errors -gt 0))){

                writeTestMessage $globalTestFlag "Get-SQLHealth - VOLUME SUMMARY"
    
                $html += "<h2>$server_alias - Host</h2><br>"

                $html += "<h3>Volume Summary</h3>"
                
                $html +=FormatHTML-ServerDiskSpace $(Get-Variable -Name "$($server)_diskSpace" -ValueOnly)
            
                $html += "<br>"
                
                $serverAlertFlag = $true
            }
        }
		elseif ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")) -and ($wmiconnect_errors -gt 0)){
            
            writeTestMessage $globalTestFlag "Get-SQLHealth - HEADLINE HEALTH CHECK - HOST STATUS SUMMARY"
    
                
            $html += "<h2>$server_alias - Host</h2><br>"

            $html += "<h3>Host Status summary</h3><br>"
                        
            $html += FormatHTML-ServerErrorsSummary $($summary | Select-Object `
			Name `
			,Type `
			,Category `
			,Errors `
            ,@{Name="Alert";Expression=				
				{ 				
				    if($_.Errors -gt 0)
                    {
					    $true
				    } else 
                        {
					        $false
				        }						
			        }			
		        } `
            ,HeadsUp `
            ,InstanceHeadsUp 
            )
            + "<br><br>"

            $wmiAlertFlag = $($wmiconnect_errors -gt 0)
        }
				
		
        
        #$instanceAlertFlag = $false
        $instanceAlertTotal = 0


		$_.instances.instance | ForEach {

            $instance = $_.name

            writeTestMessage $globalTestFlag "Get-SQLHealth - $server - SECOND SERVER LOOP - $instance - SECOND INSTANCE LOOP"
    

            $is_Sqlserver = $_.SQLServer
            if ($is_Sqlserver -eq 0){
                #do nowt
            }
            else{
                
                #create and populate new Instance item
                $resultsInstanceItem = New-Object PSObject -Property @{
                    Instance = $instance
                    InstanceAlerts = 0
                    ErrorLogAlerts = 0
                }


                $sqlconnect_errors = $(if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly)){0}else{1})
                if ($sqlconnect_errors -eq 0){
                    $sqlpermissions_errors = $(if($(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)){0}else{1})
                }
                else
                {
                    #If we can't connect then permissions will fail as well, which we don't really need to record.
                    $sqlpermissions_errors = 0
                }
            
            
                $sqlconnectalertflag = $false
                $sqlpermissionalertflag = $false
                #$instanceAlertFlag = $false              

                $sqlinstance_errors = $(Get-Variable -Name "$($instance)_sqlinstance_errors" -ValueOnly)
                $sqlinstanceheadline_errors = $(Get-Variable -Name "$($instance)_sqlinstanceheadline_errors" -ValueOnly)
                $sqldbsummary_errors = $(Get-Variable -Name "$($instance)_sqldbsummary_errors" -ValueOnly)
                $agent_running = $(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly).SQLAgent
                $sqlfailedjobs_errors = $(Get-Variable -Name "$($instance)_sqlfailedjobs_errors" -ValueOnly)
                $sqlsuspectpages_errors = $(Get-Variable -Name "$($instance)_sqlsuspectpages_errors" -ValueOnly)
                $sqlerrorlog_errors = $(Get-Variable -Name "$($instance)_sqlerrorlog_errors" -ValueOnly)
                
                
                if(($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")){
                    $instanceAlerts = $sqlinstanceheadline_errors + $sqldbsummary_errors + $sqlfailedjobs_errors + $sqlsuspectpages_errors
                }            
                else{
                    $instanceAlerts = $sqlinstance_errors + $sqldbsummary_errors + $sqlfailedjobs_errors + $sqlsuspectpages_errors
                }
                $resultsInstanceItem.InstanceAlerts = $instanceAlerts                        
                $resultsInstanceItem.ErrorLogAlerts = $sqlerrorlog_errors
                $resultsServerItem.Instances += $resultsInstanceItem             
            

                if([string]::IsNullOrEmpty($_.alias)){
				    $instance_alias = $_.name
		        } 
                else
                {
                    $instance_alias = $_.alias
                }
		
            
			    if($(Get-Variable -Name "$($instance)_sqlconnect" -ValueOnly) -And $(Get-Variable -Name "$($instance)_sqlpermissions" -ValueOnly)) {

                    writeTestMessage $globalTestFlag "Get-SQLHealth - SECOND SERVER LOOP - SECOND INSTANCE LOOP - BUILDING HTML FOR FULL HEALTHCHECK"
			
                    #if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")) -and ($instanceAlerts -gt 0 -or !$agent_running))){
                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")) -and ($sqlinstanceheadline_errors -gt 0 -or $sqldbsummary_errors -gt 0 -or !$agent_running -or $sqlfailedjobs_errors -gt 0 -or $sqlsuspectpages_errors -gt 0))){
				        $html += "<h2>"+$instance_alias+" - SQL Server</h2><br>"
                    }
                
                    #if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($instanceAlerts -gt 0))){
                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($sqlinstanceheadline_errors -gt 0))){
                        $html += "<h3>"+$instance_alias+" - Instance Summary</h3>"

                        $html += FormatHTML-SQLInstance $(Get-Variable -Name "$($instance)_sqlinstance" -ValueOnly) + "<br>"

                        #$instanceAlertFlag = $true
                        $instanceAlertTotal += 1
                    }
				
				
                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($sqldbsummary_errors -gt 0))){
                        $html += "<h3>"+$instance_alias+" - Backup Summary and Database Status</h3>"
				                
				        $html += FormatHTML-SQLDatabaseSummary $(Get-Variable -Name "$($instance)_sqldbsummary" -ValueOnly) + "<br>"

                        #$instanceAlertFlag = $true
                        $instanceAlertTotal += 1
				    }


                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and (!$agent_running -or ($sqlfailedjobs_errors -gt 0)))){
                        $html += "<h3>"+$instance_alias+" - SQL Agent Jobs</h3>"				                
				
                        $html += FormatHTML-SQLFailedJobs $(Get-Variable -Name "$($instance)_sqlfailedjobs" -ValueOnly) $agent_running

                        #$instanceAlertFlag = $true
                        $instanceAlertTotal += 1
				    }
				
                
                    if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($sqlsuspectpages_errors -gt 0))){
				        $html += "<h3>"+$instance_alias+" - Suspect Pages</h3>"
				
				        $html += FormatHTML-SQLSuspectPages $(Get-Variable -Name "$($instance)_sqlsuspectpages" -ValueOnly) $_.suspect_pages 

                        #$instanceAlertFlag = $true
                        $instanceAlertTotal += 1
                    }

                
				    #if ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")) -or (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($alertcount -gt 0))){
                    #Include error log for the Headline reoport once we've managed to filter it down a good bit.
                    if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
                        $html += "<h3>"+$instance_alias+" - SQL Error Log</h3>"

                        #Need to work out what Error log entries are flagable.
               	        $html += FormatHTML-SQLErrorLog $(Get-Variable -Name "$($instance)_sqlerrors" -ValueOnly)

                        #$instanceAlertFlag = $true
                        $instanceAlertTotal += 1
                    }

                    if ($instanceAlertFlag){
			            #to do
                        <#Write-Host "--------------------------------------------------------"#>
                    }
			    }
			    elseif ((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")) -and (($sqlpermissions_errors -gt 0) -or ($sqlconnect_errors -gt 0))){

                    writeTestMessage $globalTestFlag "Get-SQLHealth - $server - SECOND SERVER LOOP - $instance - SECOND INSTANCE LOOP - BUILDING HTML FOR HEADLINE"
                    
                    $html += "<h2>$server_alias - Host</h2><br>"

                    $html += "<h3>Instance Connection and Permissions Summary</h3>"
                        
                    $html += FormatHTML-InstanceErrorsSummary $($summary | Select-Object `
			        Name `
			        ,Type `
			        ,Category `
			        ,Errors `
                    ,@{Name="Alert";Expression=				
				        { 				
				            if($_.Errors -gt 0)
                            {
					            $true
				            } else 
                                {
					                $false
				                }						
			                }			
		                } `
                    ,HeadsUp `
                    ,InstanceHeadsUp 
                    )
                    + "<p><br>"

                    $wmiAlertFlag = $($wmiconnect_errors -gt 0)
                }

                <#
                if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($instancealerttotal -eq 0)){
                    $html += "There are no instance issues"
                }
                #>
                #to do
                #Write-Host "--------------------------------------------------------"
            
            } #End of check for SQL Server

		} #END OF INSTANCE LOOP
		

        writeTestMessage $globalTestFlag "Get-SQLHealth - $server - SECOND SERVER LOOP - COMPLETING HTML"

        if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and (!($serverCompleted)) -and ((@($resultsServerItem.ServerAlerts) -gt 0) -or @($resultsInstanceItem.InstanceAlerts -gt 0))){
		    $html += "<br/><h3>&nbsp;</h3><br/><p/>"
        }
		
		if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
		    $html += "</body></html>"
        }
        
        $end = Get-Date
		        
        #Gordon Feeney: added Server Instance to output file
        $i = $instance_alias.indexof("\")
        if ($i -ne -1){
            $serverinstance = "$($server_alias)_$($instance_alias.substring($i + 1))"
        }
        else{
            $serverinstance = $server_alias
        }

        if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "FULL")){
            $outfile = "$executingScriptDirectory\output\$($serverinstance)_$(Get-Date -format ""yyyyMMddHHmm"")_HealthCheck.html"        
		    $html | Out-File $outfile 
            writeTestMessage $globalTestFlag "Get-SQLHealth - SENDING EMAIL FOR FULL HEALTHCHECK"    
            
            #If email contains multiple instances, then just use the server name in the email subject
            #If email contains only one instance, then use server\instance
            
            $instanceCounter = 0
            $_.instances.instance | ForEach {
			   $instanceCounter = $instanceCounter + 1 
            }

            if($instanceCounter -gt 1)
            {
                $serverSubject = $server
            }
            else
            {
                $serverSubject = $instance_alias
            }

            Send-Mail $smtpServer $smtpPort $ssl $smtpUser $smtpPassword $mailFrom $mailTo $("$client - $serverSubject - $subject") $html
            $html = $null
            
        }

        $resultsServerItem.ServerCompleted = $true
        $resultsServerArray += $resultsServerItem
                
	} #END OF SERVER LOOP

                     
    $prevServer = $null    
    $alertTotal = 0

    if (((($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")))){
        
		foreach ($resultsServerItem in $resultsServerArray ){

            $currServer = $resultsServerItem.Server   
            
            if (($prevServer -eq $null) -or ($currServer -ne $prevServer)){
                
                $alerts = $resultsServerItem.ServerAlerts
                #First loop or next server
                if ($alerts  -gt 0){                    
                    $html += "<strong>$currServer Host issues (WMI, Volumes): $alerts</strong><br>"
                    $alertTotal  += 1
                }
            }

            foreach ($resultsInstanceItem in $resultsServerItem.Instances ){
                
                $instance = $resultsInstanceItem.Instance

                $alerts = $resultsInstanceItem.InstanceAlerts
                if ($alerts  -gt 0){
                    $html += "<strong>$instance Instance issues (Connections, Permissions, Service restarts, Backups, Jobs, Suspect Pages): $alerts</strong><br>"
                    $alertTotal  += 1
                }
            }

            $prevServer = $currserver

            $html += "<p>"
        }        
        
        $html += "<p>"
    }
        

    #if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE") -and ($alertTotal -gt 0)){
    if (($globalScriptType -eq "HEALTHCHECK") -and ($globalScriptSubType -eq "HEADLINE")){

        if ($alertTotal -eq 0){
           $html += "<h3>&nbsp;</h3>" 
           $html += "<strong>No issues</strong>"
        }

        $html += "</body></html>"
        $outfile = "$executingScriptDirectory\output\$(Get-Date -format ""yyyyMMddHHmm"")_HealthCheck_Headline.html"
	    $html | Out-File $outfile       
        $config_name = $config.substring(0, $config.length - 4)
        if ($config_name -eq "config"){
            $mailSubject = "$client - $subject (Headline)"
        }
        else{
            $mailSubject = "$client - $config_name - $subject (Headline)"
        }
        writeTestMessage $globalTestFlag "Get-SQLHealth - SENDING EMAIL FOR HEADLINE"     
        Send-Mail $smtpServer $smtpPort $ssl $smtpUser $smtpPassword $mailFrom $mailTo $($mailSubject) $html
        $html = $null
    }
}


$directory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
SetScriptType $scriptType
SetScriptSubType $scriptSubtype
SetScriptMessage $scriptMessage
$globalTestFlag = $testFlag

if ($globalTestFlag){
    Write-Host "Get-SQLHealth - START OF FILE"
    Write-Host "globalScriptType : $globalScriptType"                    
    Write-Host "globalScriptSubType : $globalScriptSubType"
}
        
Get-SQLHealth $directory $config 2> "$directory\log\$(Get-Date -format "yyyyMMddHHmmssff")_HealthCheck.log"


## Purge old files
DeleteFiles "$directory\log" ".log" 30
DeleteFiles "$directory\output" ".html" 30
