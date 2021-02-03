# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
# 
# comet-reporter.ps1 - Operational integration (reporting) for Comet Backup and SyncroMSP
#
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
#
# Description - Monitor Comet Backup job history and ticket appropriatly in Syncro.  Ideally, will run indefinitely on
#               Comet server as a Scheduled Task.  Use settings in Definitions section to control behavior.
#
#               One of a group of scripts to manage Comet Backup operations from Syncro.
#
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #
#
# Date     Description
# -------- ------------------------------------------------------------------------------------------------------------
# 20201110 Intial version (software@tsmidwest.com)
# 20210107 Adaptated from comet-dispatcher to report on completed Comet jobs (software@tsmidwest.com)
#
# # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # # #


###
### Identification
###
$ID  = "comet-reporter"
$VER = "20210108"


###
### Definitions
###
$BrandedName          = ""  # Name to use if service is branded
$CometPWD             = ""  # Comet password for API access
$CometURL             = ""  # Comet server URL
$CometUser            = ""  # Comet user for API access
$SleepTimer           = 180 # Duration to sleep between cycles
$SyncroAPIKey         = ""  # Syncro API key
$SyncroIssue          = ""  # Syncro ticket issue
$SyncroSubjectFailure = ""  # Syncro ticket subject on failure
$SyncroSubjectSuccess = ""  # Syncro ticket subject on success
$WorkDir              = ""  # Work directory
$GracePeriod          = 3   # Number of days a backup can miss its job


###
### Constants
###
$CometJobEntrySeverity  = @{I = "Info "; W = "Warn "; E = "Error"}
$CometJobClassification = @{7000 = "Timeout";
                            7001 = "Warning";
                            7002 = "Error";
                            7003 = "Quota";
                            7004 = "Schedule Missed";
                            7005 = "Cancelled";
                            7006 = "Skipped - Already Running";
                            7007 = "Abandoned"}


###
### Modules
###


###
### Function:  Message
###
function Message ($Message) {
    $out = $Message
    for($i = $Message.Length; $i -le 80; $i++) { $out = "$($out)." }
    Write-Host -NoNewline "$($out)"
    $out | Out-File -FilePath "$($WorkDir)\comet-reporter-$(Get-Date -UFormat "%Y%m%d").log" -Append -NoNewLine
}


###
### Function:  Result
###
function Result ($Result) {
    Write-Host "$($Result)"
    $Result | Out-File -FilePath "$($WorkDir)\comet-reporter-$(Get-Date -UFormat "%Y%m%d").log" -Append
}

    
###
### Function:  Create-Syncro-Ticket
###
function Create-Syncro-Ticket ($IssueType, $Status, $Subject, $UUID) {
    $url = "https://$($(Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\RepairTech\Syncro).shop_subdomain).syncromsp.com/api/syncro_device/tickets"
    $body = @{
        subject = $Subject
        problem_type = $IssueType
        status = if($status) { $Status } else { "New" }
        uuid = $UUID # This is a change from the Syncro integrated API
    }
    $bodyjson = ConvertTo-Json -InputObject $body -Compress
    $ticket = try {
        Invoke-WebRequest -Uri $url -Method Post -Body $bodyjson -ContentType 'application/json'
    } catch {
        ###! unable to create Syncro ticket, this is a fatal failure, exit so that admin attention will be brought by comet-monitor.
        Exit 1
    }
    $(ConvertFrom-Json -InputObject $ticket)
}


###
### Function:  Create-Syncro-Ticket-Comment
###
function Create-Syncro-Ticket-Comment ($Body, $DoNotEmail, $Hidden, $Subject, $TicketId, $UUID) {
    $url = "https://$($(Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\RepairTech\Syncro).shop_subdomain).syncromsp.com/api/syncro_device/tickets/$($TicketId)/add_comment"
    $result = $null
    if([bool]::TryParse([string]$Hidden, [ref]$result)) { if($result) { $Hidden = "1" } else { $Hidden =  "0" } } else { $Hidden = "0" }
    $result = $null
    if([bool]::TryParse([string]$DoNotEmail, [ref]$result)) { if($result) { $DoNotEmail = "1" } else { $DoNotEmail =  "0" } } else { $DoNotEmail = "1" }
    $body = @{
        subject = $Subject
        body = $Body
        hidden = $Hidden
        do_not_email = $DoNotEmail
        uuid = $UUID # This is a change from the Syncro integrated API
    }
    $bodyjson = ConvertTo-Json -InputObject $body -Compress
    $comment = try {
        Invoke-WebRequest -Uri $url -Method Post -Body $bodyjson -ContentType 'application/json'
    } catch {
        ###! unable to create Syncro ticket comment; this is NOT fatal, as the ticket was presumable already created
    }
    #$(ConvertFrom-Json -InputObject $comment) # This can be uncommented if output is needed; however, poor handling causes stdout issues
}


###
### Startup banner
###
Message("$($ID) $($VER)")


###
### Ensure work directory exists
###
if(-not (Test-Path $WorkDir -PathType Container)) {
    Result("failed")
    Exit 1
} else {
    Result("starting")
}


###
### Concurrency management; get date of PID file and ensure that the last update is at least 2x $SleepTimer before
### continuing to prevent multiple instances of reporter from running simultaneously
###
### #! Maybe I'll add this later...?


###
### Start the reporter loop
###
while(1) {


    ###
    ### Just some output to indicate things are working
    ###
    $cycle = Get-Date
    Message("Cycle $($cycle)")
    Result("starting")


    ###
    ### Retrieve the last successful backup end time
    ###
    if(Test-Path "$($WorkDir)\comet-reporter.json" -PathType Leaf) {
        $reporter = Get-Content -Path "$($WorkDir)\comet-reporter.json" | ConvertFrom-Json
    }


    ###
    ### Use Comet get-jobs-for-date-range API endpoint to retrieve a list of jobs that have a start or end date later
    ### than the most recently recorded end date in comet-reporter.json
    ###
    ### #! Enhancement - Use get-jobs-for-date-range API endpoint?
    ###
    Message("Retrieve Comet Backup job list")
    $comet_jobs = $(Invoke-WebRequest -URI "$($CometURL)/api/v1/admin/get-jobs-all" -Method Post -Body @{Username="$($CometUser)"; AuthType="Password"; Password="$($CometPWD)"}).Content | ConvertFrom-Json
    Result("done")


    ###
    ### First pass through the backup jobs; this pass sets $latestjobendtime variable.  If this variable is not equal
    ### to $reporter.lastjobendtime then it means an unticketed job has been found.  This test will be used to first 
    ### build the hash that contains Comet Device ID's and Syncro Asset ID's, and then finally to retrieve job details
    ### and create Syncro tickets.
    ###
    $latestjobendtime = $reporter.lastestjobendtime
    foreach($job in $comet_jobs) {
        if($job.EndTime -gt $latestjobendtime) {
            $latestjobendtime = $job.EndTime
        }
    }


    ###
    ### If $latestjobendtime and $reporter.lastjobendtime are not equal, then unticketed Comet jobs have been found;
    ### first step is to build the Comet Device ID / Syncro Asset UUID hash that can be used when tickets are created
    ### in round two.
    ###
    if($reporter.latestjobendtime -ne $latestjobendtime) {
        $cometid_syncroid = @{}
        $syncropage = 1
        $syncropages = 1
        while($syncropage -le $syncropages) {
            $url = "https://$($(Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\RepairTech\Syncro).shop_subdomain).syncromsp.com/api/v1/customer_assets?api_key=$($SyncroAPIKey)&page=$($syncropage)"
            Message("Retrieve page $($syncropage) of Syncro assets")
            $syncro_assets = Invoke-RestMethod -ContentType 'application/json' -Headers @{'Accept' = 'application/json'} -Method Get -Uri $url
            Result("done")
            $syncropages = $syncro_assets.meta.total_pages

            foreach($asset in $syncro_assets.assets) {
                if(!([string]::IsNullOrEmpty($asset.properties.'Comet Device ID'))) {
                    $cometid_syncroid.Add($asset.properties.'Comet Device ID', $asset.properties.kabuto_live_uuid);
                }    
            }
            $syncropage++
        }
    }


    ###
    ### Second pass through the backup jobs; on this pass, retrieve job details and create Syncro ticket
    ###
    foreach($job in $comet_jobs) {
        if($job.EndTime -gt $reporter.latestjobendtime) {
            

            ###
            ### Retrieve user profile details
            ###
            Message("Retrieve Comet Backup user profile $($job.Username)")
            $comet_userprofile = $(Invoke-WebRequest -URI "$($CometURL)/api/v1/admin/get-user-profile" -Method Post -Body @{Username="$($CometUser)"; AuthType="Password"; Password="$($CometPWD)"; TargetUser="$($job.Username)"}).Content | ConvertFrom-Json
            Result("done")


            ###
            ### Retrieve Comet job log entries
            ###
            Message("Retrieve Comet Backup job details")
            $comet_joblog = $(Invoke-WebRequest -URI "$($CometURL)/api/v1/admin/get-job-log-entries" -Method Post -Body @{Username="$($CometUser)"; AuthType="Password"; Password="$($CometPWD)"; JobID="$($job.GUID)"}).Content | ConvertFrom-Json
            Result("done")


            ###
            ### Use backup job details Status to determine success / failure of job and create Syncro ticket, set status
            ### and and backup job log entries
            ###
            if($job.Status -ge 5000 -and $job.Status -le 5999) {
                
                
                ###
                ### Create Syncro ticket
                ###
                Message("Create Syncro ticket")
                $syncroticket = Create-Syncro-Ticket -Subject "$($SyncroSubjectSuccess)" -IssueType "$($SyncroIssue)" -Status "Resolved" -UUID $cometid_syncroid.$($job.DeviceID)
                Result("done")


                ###
                ### Add Comet Backup job detail to Syncro ticket
                ###
                $syncroticket_comment = ""
                foreach($entry in $comet_joblog) {
                    $syncroticket_comment += "$(Get-Date -Date $([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($entry.Time))) -Format "yyyyMMdd HHmmss") | $($CometJobEntrySeverity.$($entry.Severity)) | $($entry.Message)`n"
                }
                Message("Add Syncro ticket detail comment")
                Create-Syncro-Ticket-Comment -TicketId $syncroticket.ticket.id -Subject "$($BrandedName) Detail" -Body "$($syncroticket_comment)" -UUID $cometid_syncroid.$($job.DeviceID)
                Result("done")


                ###
                ### Add initial detail to Syncro ticket (header regarding job details); added last so it will appear at
                ### top of ticket comments while viewing in reverse chronological order
                ###
                Message("Add Syncro ticket summary comment")
                Create-Syncro-Ticket-Comment -TicketId $syncroticket.ticket.id -Subject "$($BrandedName) Summary" -Body "$($BrandedName) backup job $($comet_userprofile.Sources.$($job.SourceGUID).Description) for user $($job.Username) completed successfully." -UUID $cometid_syncroid.$($job.DeviceID)
                Result("done")
            } elseif($job.Status -ge 7000 -and $job.Status -le 7999) {

                # Skip ticket creation if grace period hasn’t been passed yet
                $GracePeriodUnix = $($comet_userprofile.Sources.$($job.SourceGUID).Statistics.LastBackupJob.EndTime) + ($GracePeriod * 86400)
                if($GracePeriodUnix -ge $job.StartTime) {
                    continue
                }

                ###
                ### Create Syncro ticket
                ###
                Message("Create Syncro ticket")
                $syncroticket = Create-Syncro-Ticket -Subject "$($SyncroSubjectFailure)" -IssueType "$($SyncroIssue)" -Status "New" -UUID $cometid_syncroid.$($job.DeviceID)
                Result("done")


                ###
                ### Add Comet Backup job detail to Syncro ticket
                ###
                $syncroticket_comment = ""
                foreach($entry in $comet_joblog) {
                    $syncroticket_comment += "$(Get-Date -Date $([TimeZone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($entry.Time))) -Format "yyyyMMdd HHmmss") | $($CometJobEntrySeverity.$($entry.Severity)) | $($entry.Message)`n"
                }
                Message("Add Syncro ticket detail comment")
                Create-Syncro-Ticket-Comment -TicketId $syncroticket.ticket.id -Subject "$($BrandedName) Detail" -Body "$($syncroticket_comment)" -UUID $cometid_syncroid.$($job.DeviceID)
                Result("done")


                ###
                ### Add summary detail to Syncro ticket (header regarding job details); added last so it will appear at
                ### top of ticket comments while viewing in reverse chronological order
                ###
                Message("Add Syncro ticket summary comment")
                if($(job.Status) -le 7009) {


                        ###
                        ### For job status between 7000 and 7009, add detail based on Comet's defined constants; it may be
                        ### necessary to update this range and the reference constant if / when Comet changes their API
                        ### defined constant
                        ###
                        Create-Syncro-Ticket-Comment -TicketId $syncroticket.ticket.id -Subject "$($BrandedName) Summary" -Body "$($BrandedName) backup job $($comet_userprofile.Sources.$($job.SourceGUID).Description) for user $($job.Username) completed with one or more errors($CometJobClassification.$($job.Status))." -UUID $cometid_syncroid.$($job.DeviceID)
                } else {
 
 
                        ###
                        ### Job status outside of defined bounds, use a non-specific summary message
                        ###
                        Create-Syncro-Ticket-Comment -TicketId $syncroticket.ticket.id -Subject "$($BrandedName) Summary" -Body "$($BrandedName) backup job $($comet_userprofile.Sources.$($job.SourceGUID).Description) for user $($job.Username) completed with one or more errors." -UUID $cometid_syncroid.$($job.DeviceID)
                }
                Result("done")
            } else {


                ###
                ### Something odd has happened, so log it; never should a situation occur where a job has a non-zero
                ### end time and a status of anything other than a value between 5000-5999 or 7000-7999
                ###
                Message("Create Syncro ticket")
                Result("failed")
            }
        }
    }


    ###
    ### Update reporter file
    ###
    if($lastjobendtime -ne 0) {
        $(@{pid=$PID; latestjobendtime=$latestjobendtime; cycle=$SleepTimer} | ConvertTo-Json) | Out-File -FilePath "$($WorkDir)\comet-reporter.json"
        #"{pid: `"$($PID)`", latestjobendtime: `"$($latestjobendtime)`", cycle: `"$($SleepTimer)`"}" | Out-File -FilePath "$($WorkDir)\comet-reporter.json"
    }

    
    ###
    ### Complete cycle
    ###
    Message("Cycle $($cycle)")
    Result("finished")


    ###
    ### Pause before next cycle
    ###
    Sleep -Seconds $SleepTimer
}
