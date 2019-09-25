# Set Zebra printing language script
# Joshua Woleben
# Written 9/17/2019

Param(  [switch]$all,
        [string[]]$PrinterList,
        [string]$File)
$functions = @'
function is_printer_done {
    Param($printers_done,$printer_name_check)
    foreach ($printer_name in $printers_done) {
        if ($printer_name_check -match $printer_name) {
            return $true
        }
    }
    return $false
}
function is_printer_online {
    Param($printer_name)
    if (Test-Connection -ComputerName $printer_name -Quiet -Count 1) {
        return $true
    }
    else {
        return $false
    }
}
function write_log {
    Param([string]$log_entry,
            [string]$TranscriptFile)

            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                Write-Output $log_entry
            }
            finally {
                $mutex.ReleaseMutex() | out-null

            }
            $mutex.Dispose()
}
'@
function write_log {
       Param([string]$log_entry,
            [string]$TranscriptFile)
            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                Write-Output $log_entry
            }
            finally {
                $mutex.ReleaseMutex() | out-null
            }
            $mutex.Dispose()
}


# List of EPS servers
$print_servers = @('print_server_host')
$printer_list = @()
$printer_done_file = "C:\Temp\zebra_printers_done.txt"


# Get credentials
$username = "admin_user"

if (Test-Path -Path "C:\Temp\admin_pwd.txt" -PathType Leaf) {
    $password = Get-Content -Path "C:\Temp\eps_pwd.txt" | ConvertTo-SecureString
}
else {
    $password = Read-Host -AsSecureString -Prompt "Service account password"
}

$eps_creds = New-Object System.Management.Automation.PSCredential ($username, $password)

$TranscriptFile = "C:\Temp\Zebra_Language_Set_$(get-date -f MMddyyyyHHmmss).txt"

if (Test-Path $printer_done_file -PathType Leaf) {
    $printers_done = Get-Content -Path $printer_done_file
}
else {
    $printers_done = @()
}
write_log "Initializing..." $TranscriptFile
$global:printer_count = 0

ForEach ($server in $print_servers) {
    write_log "Gathering printers on $server..." $TranscriptFile
    if ($all) {
        $printer_list = Get-Printer -ComputerName $server | Where { $_.DriverName -match "ZDesigner QLn220" }
    }
    elseif (-not [string]::IsNullOrEmpty($File)) {
        if (Test-Path $File -PathType Leaf) {
            $printer_list = Get-Content -Path $File
            ForEach ($printer_name in $printer_names) {
                $printer_list += Get-Printer -ComputerName $server -Name "$printer_name"
            }
        }
        else {
            write_log "$File not found!" $TranscriptFile
            exit
        }
    }
    else {
        $printer_names = $PrinterList
        ForEach ($printer_name in $printer_names) {
            $printer_list += Get-Printer -ComputerName $server -Name "$printer_name"
        }
    }
    Copy-Item -Path "C:\Temp\zpl1.txt" -Destination "\\$server\C`$\Temp\zpl1.zpl" -Force
    Copy-Item -Path "C:\Temp\zpl2.txt" -Destination "\\$server\C`$\Temp\zpl2.zpl" -Force


    $printer_list | ForEach-Object {
<#        if ($printer_count -gt 30) {
            write_log "Pausing at 30 jobs to allow spooler to catch up..."
            sleep 20
            $printer_count = 0
        }#>

        
        $printer=$_.Name
        $PrintJob = { 
        Invoke-Expression $using:functions
            
        write_log "Working on $using:printer..." $using:TranscriptFile 
        if ((is_printer_done $using:printers_done $using:printer) -eq $true) {
            write_log ($using:printer + " has already been done. Continuing...") $using:TranscriptFile 
        }
        elseif ((is_printer_online $using:printer) -eq $false) {
            write_log ($using:printer + " is offline, skipping...") $using:TranscriptFile 
        }
        else {
            write_log "Changing driver to generic / text..." $using:TranscriptFile 
            Set-Printer -ComputerName $using:server -Name $using:printer -DriverName "Generic / Text Only"

            sleep 1
        
            $batch_text = "C:\Windows\ssdal.exe /p `"$using:printer`" send `"C:\Temp\zpl1.zpl`"`nC:\Windows\ssdal.exe /p `"$using:printer`" send `"C:\Temp\zpl2.zpl`""
        
            $batch_text | Out-File "C:\Temp\printer-$using:printer.cmd" -Encoding ascii
            Copy-Item -Path "C:\Temp\printer-$using:printer.cmd" -Destination "\\$using:server\C$\Temp\printer-$using:printer.cmd" -Force

            write_log "Issuing zpl commands to printer..." $using:TranscriptFile
            $printer_com = $($using:printer)
           
            $command_output =  Invoke-Command -ArgumentList $printer_com -ComputerName $using:server -Credential $using:eps_creds  { Param($printer_com)

              &  "C:\Temp\printer-$printer_com.cmd"
            }
            write_log $command_output $using:TranscriptFile 
            sleep 1
            write_log "Switching back to Zebra driver..." $using:TranscriptFile 
            Set-Printer -ComputerName $using:server -Name $using:printer -DriverName "ZDesigner QLn220"
            $using:printer | Out-File -FilePath $using:printer_done_file -Append
            if (Test-Path -Path "\\$using:server\C`$\Temp\printer-$using:printer.cmd" -PathType Leaf) {
                   Remove-Item -Path "\\$using:server\C`$\Temp\printer-$using:printer.cmd" -Force
            }

        }
        } # End script block
        Start-Job -ScriptBlock $PrintJob # -Credential $eps_creds
        while (@(Get-Job).Count -gt 30) {
            write_log (Get-Job | Receive-Job) $TranscriptFile
            
            Remove-Job -State Completed
            sleep 5
        }

    }
    write_log "Waiting on jobs to finish..." $TranscriptFile
    Get-Job | Wait-Job | Remove-Job
    write_log "Cleaning up..." $TranscriptFile
    Remove-Item -Path "\\$server\C`$\Temp\zpl1.zpl" -Force
    Remove-Item -Path "\\$server\C`$\Temp\zpl2.zpl" -Force
     
}


# Generate email report
$email_list=@("email1@example.com","user2@example.com")
$subject = "Zebra Printer Report"

$body = "Zebra language set report and transcript attached."

if ($all) {
    $body += "`nScript ran against ALL online printers."
}
if (-not [string]::IsNullOrEmpty($File)) {
    $body += "`nPrinter list loaded from file $File"
}
if (-not [string]::IsNullOrEmpty($PrinterList)) {
    $body += "`nPrinter list supplied at command line."
}


$MailMessage = @{
    To = $email_list
    From = "ZebraLanguageReport<Donotreply@example.com>"
    Subject = $subject
    Body = $body
    SmtpServer = "smtp.example.com"
    ErrorAction = "Stop"
    Attachment = $TranscriptFile
}
Send-MailMessage @MailMessage
