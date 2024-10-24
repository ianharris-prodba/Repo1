<#
File		:	Get-SQLDisks.ps1

Parameters	:
			IN 	:	-Config
            	OUT	:	

Summary	:	Monitors free space on disks entered into the config.xml file in the configuration folder
#>

param([string]$config="config.xml", [string]$configsubject)

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
		Get-Date -Date $datetime -Format "dd-MMM-yy HH:mm:ss"
	}
	
}

function Format-Boolean ($bool){
	if($bool){
		return "Yes"
	} else {
		return "No"
	}
	
}

function Send-Mail ($smtpServer,$smtpPort,$ssl,$user,$pwd,$from,$to,$subject,$subsubject,$body){

    $SMTPClient = New-Object Net.Mail.SmtpClient($smtpServer,$smtpPort) 
	$msg = New-Object Net.Mail.MailMessage
	
	$msg.From = New-Object Net.Mail.MailAddress($from)
    #$msg.To.Add("gordon.feeney@pro-dba.com"); 
    $msg.To.Add($to); 
	$msg.IsBodyHTML = $true
	$msg.Body = $body
    $msg.Subject = $subject + " - " + $subsubject
 	
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

function Get-SQLLocalDateTime ($instance) {
	
	$server_date = $null
	
	Query-SQL $instance "SELECT GETDATE() AS server_date" | ForEach-Object {
		$server_date = $_.server_date
	}
	
	return $server_date

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


function Get-ServerDiskSpace($server,$config){
	
    ## Read in default values
    $defaultunittype = $config."default-unit-type"
    $defaultunitvalue = $config."default-unit-value"

    if(([string]::IsNullOrEmpty($defaultunittype )) -or ([string]::IsNullOrEmpty($defaultunitvalue ))){
        $defaultunittype  = 'PERCENT'
        $defaultunitvalue = '20'
    }

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
    Get-WmiObject Win32_Volume -ComputerName $server | Where { $_.drivetype -eq '3' -and $_.Name -notlike "\\?\Volume*"} | ForEach {
		
        $name = $_.Name
		$label = $_.Label
		$freespace = [int64]$_.FreeSpace
		$capacity = [int64]$_.Capacity
		$freespacethreshold = [System.Nullable``1[[System.Int64]]] $null
		$enabled = $true
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
		
        <#
        Write-Host "Server: $server"
        Write-Host "Name: $name"
        Write-Host "Label: $label"
        Write-Host "Enabled: $enabled"
        Write-Host "FreeSpace: $freespace"
        Write-Host "ThresholdSpace: $freespacethreshold"
        #>      

        if(($freespace -lt $freespacethreshold -Or !($freespacethreshold)) -And (($enabled) -Or [string]::IsNullOrEmpty($enabled))){
            New-Object PSObject -Property @{
			Name = $name
			Label = $label
			FreeSpace = $freespace
			ThresholdSpace = $freespacethreshold
			Capacity = $capacity
			Enabled = $enabled
			Notes = $notes
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
	            } #End of New-Object | Select-Object
        }
        
		
    }#End of Get-WmiObject 
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
		<th>Free Space(%)</th>
		<th>Free Space(GB)</th>
		<th>Threshold (%)</th>
        <th>Threshold(GB)</th>
		<th>Capacity(GB)</th>
	</tr>
	"
	
	$obj | ForEach {
	
		$name = $_.Name
	
		if($_.Alert){
			$html += "<tr class='warning'>"
		} else {
			$html += "<tr>"
		}

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
	
	$html += "</table>"
	
	if($notes.Count -gt 0){
	
		$html += "<table class='disabled'>"
		
		$notes | ForEach{
		
			$html += "<tr>"
		
			$html+="<td><sup>"+$_.ID+"</sup> "+$_.Note+"</td>"
			
			$html += "</tr>"
		
		}
		
		$html += "</table>"
		
	}
	
	return $html
}

function FormatHTML-ServerDiskEmailSubjects($obj){

    $drivealert = ""
    #Write-Host "Subect: $subject"

    $obj | ForEach {
	
        #$drivealert += "`r`n"
        if ($drivealert  -ne ""){
            $drivealert += "; "    
        }
        
		$name = $_.Name
	
		if($_.Alert){
			#Write-Host "Name: $name"
            #Write-Host "Label: "$_.Label
            #Write-Host "FreeSpacePercent: "$_.FreeSpacePercent
            #Write-Host "FreeSpaceGB: "$_.FreeSpaceGB
            #Write-Host "ThresholdSpaceGB: "$_.ThresholdSpaceGB
            #Write-Host "CapacityGB: "$_.CapacityGB
            $drivealert += $_.Name + " (" + $_.FreeSpaceGB + "GB/" + $_.CapacityGB + "GB, " + $_.FreeSpacePercent + "%)"
            #Write-Host $drivealert
		}
        #Write-Host ""
    }
    
    return $drivealert	
}

function FormatHTML-ReportHeader($hostName,$username,$start,$end,$version,$to,$cc,$body,$server){
		
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
		<td><a href='mailto:$to&cc=$cc&subject=SQL Disk Status Alert - $server&body=$body'>Report Issue</a></td>
	</tr>
	</table>
	"
	
	return $html

}

function FormatHTML-ReportSummary($summary){
		
	$html = "
	<table class='summary'>
	<tr>
		<th>Name</th>
		<th>Type</th>
		<th>Category</th>
		<th>Errors</th>
	</tr>
	"+$($summary | ForEach{
            
            if ($_.Errors -gt 0){
                "<tr"+$(if($_.Alert){" class='warning'"})+">
			    <td>"+$_.Name+"</td>
			    <td>"+$_.Type+"</td>
			    <td>"+$_.Category+"</td>
			    <td>"+$(
			    if(($_.Errors -eq 0)){ 
			    "OK"
			    } else {
				    if ([string]::IsNullOrEmpty($_.Errors)){
					    "-"
				    }
				    else {
					    $_.Errors
				    }
			    }
			    )+"</td>
			    </tr>"
            }			
		
	})+"</table>"
	
	return $html

}

function Get-SQLDisks ($directory, $config, $configsubject){

	$executingScriptDirectory = $directory

	$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
	$FQhost = (Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain

	$reportVersion = Import-Csv -Delimiter "|" -Path "$executingScriptDirectory\version.rel"

	$css = Get-Content "$executingScriptDirectory\configuration\main.css"
	[xml]$xml = Get-Content "$executingScriptDirectory\configuration\$config"

	$client = $xml.configuration.metadata.client
	
    if([string]::IsNullOrEmpty($configsubject)){
        $subject = $xml.configuration.metadata.subject        
    }
    else{
        $subject = $configsubject
    }
    
	$clientcontact = $xml.configuration.metadata."client-contact"
	$clientcontactcc = $xml.configuration.metadata."client-contact-cc"
	$body = $xml.configuration.metadata.body
	$smtpServer = $xml.configuration.smtp.server
	$smtpPort = $xml.configuration.smtp.port
	$ssl = $(if($xml.configuration.smtp.ssl -eq 1){$true} else {$false})
	$smtpUser = $xml.configuration.smtp.user
	$smtpPassword = $xml.configuration.smtp.password
	$mailFrom = $xml.configuration.smtp.from
	$mailTo = $xml.configuration.smtp.to

    $xml.configuration.server | ForEach {

		# Get Data
        $server = $_.name
        
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
	}	
	
	#Write-Host "smtpServer $smtpServer"
    #Write-Host "mailTo $mailTo"

    $xml.configuration.server | ForEach {

        $server = $_.name

        if([string]::IsNullOrEmpty($_.alias)){
            $server_alias = $_.name
		} 
        else
        {
            $server_alias = $_.alias
        }

        $Errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){		
		    $(Get-Variable -Name "$($server)_diskSpace_errors" -ValueOnly)
        } else { 0 })

        if ($Errors -gt 0){
		    $server = $_.name
            #Write-Host "HERE #1"
            #Write-Host "server $server "

            $disks = $(Get-Variable -Name "$($server)_diskSpace" -ValueOnly)
            $disks | foreach{

                $volume = $_.Name
                
                <#
                Write-Host "HERE #2"
                Write-Host "Server: $server"
                Write-Host "Volume: $volume"
                Write-Host "Label: "$_.Label
                Write-Host "Freespace: "$_.FreespaceGB
                Write-Host ""
                #>

                $summary = @()	
		
		        $html = "
		        <html>
		        <head>
		        <style>"+$css+"</style>
		        </head>
		        <body>
		        <h1>SQL Disk Status Check</h1>
		        "

		        #$html += FormatHTML-ReportHeader $FQhost $user $(Get-Variable -Name "$($server)_start" -ValueOnly) $(Get-Variable -Name "$($server)_end" -ValueOnly) $reportVersion $clientcontact $clientcontactcc $body $server
		
		        $summary += New-Object PSObject -Property @{
			        Name = $server
			        Type = "Host"
			        Category = "WMI"
			        Errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){
		
					        0

				        } else { 1 })
		        }
				
		        $summary += New-Object PSObject -Property @{
			        Name = $server
			        Type = "Host"
			        Category = "Disk"
			        Errors = $(if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)){
		
				        $(Get-Variable -Name "$($server)_diskSpace_errors" -ValueOnly)

			        } else { $null })
		        }		
		

		        $html += "<h2>Report Summary</h2><br>"
		
		        $html += FormatHTML-ReportSummary $($summary | Select-Object `
			        Name `
			        ,Type `
			        ,Category `
			        ,Errors `
			        ,@{Name="Alert";Expression=
				
				        { 
				
				        if($_.Errors -gt 0){
					        $true
				        } else {
					        $false
				        }
				
						
				        }
			
			        })

			
		        if($(Get-Variable -Name "$($server)_wmiconnect" -ValueOnly)) {
				
			        #$html += "<h2>$server - Host</h2><br>"
			
			        $html += "<h3>Volume Summary</h3>"

                    $serverdisk = $(Get-Variable -Name "$($server)_diskSpace" -ValueOnly) | Where-Object {$_.Name -eq $Volume}
			        $html +=FormatHTML-ServerDiskSpace $serverdisk
			
		        }
		
		        $html += "</body></html>"
		
		        $end = Get-Date
		
		        $html | Out-File "$executingScriptDirectory\output\$($server)_$(Get-Date -format ""yyyyMMddHHmm"")_Disks.html"

                $subsubject = FormatHTML-ServerDiskEmailSubjects $serverdisk
                
		        Send-Mail $smtpServer $smtpPort $ssl $smtpUser $smtpPassword $mailFrom $mailTo $("$subject - $client - $server_alias") $subsubject $html
            }
        }        
	}

}

$directory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Get-SQLDisks $directory $config $configsubject 2> "$directory\log\$(Get-Date -format "yyyyMMddHHmmssff")_Disks.log"

DeleteFiles "$directory\log" "Disks.log" 7
DeleteFiles "$directory\output" "Disks.html" 7