# create an array that tracks each file
$global:files = @()

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = "V:\Software\Utilities Formatting Tool\Error Log"
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true  

$action = {
    $path = $Event.SourceEventArgs.FullPath
    $changeType = $Event.SourceEventArgs.ChangeType
    $logline = "$(Get-Date -f o), $changeType, $path"
    Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value "==============================================="
    Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value $logline
    Write-Host $logline
    
    
    if($changeType -eq "Changed") {
        Write-Host -ForegroundColor Green "Found something useful $path"
        Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value "Found something useful $path"
        $global:files += $path
    }

}    

Register-ObjectEvent $watcher "Created" -Action $action
Register-ObjectEvent $watcher "Changed" -Action $action
Register-ObjectEvent $watcher "Deleted" -Action $action
Register-ObjectEvent $watcher "Renamed" -Action $action

while ($true) {

    # if there are any files to process, check if they are locked
    if($global:files) {

        $global:files | % {

            $file = $_
            # assume the file is locked
            $fileFree = $false

            Write-Host -ForegroundColor Yellow "Checking if the file is locked... ($file)"
            Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value "Checking if the file is locked... ($_)"

            # true  = file is free
            # false = file is locked
            try {
                [IO.File]::OpenWrite($file).close();Write-Host -ForegroundColor Green "File is free! ($file)"
                Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value "File is free! ($file)"
                $fileFree = $true
		$filepath = "V:\Software\Utilities Formatting Tool\Error Log"
                $filename = (get-childitem -file $filepath | sort CreationTime -Descending | select -last 1).Name
                $Join = Join-Path -path $filepath -ChildPath $filename
                $content = Get-Content -Raw $Join | out-string
			
                Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value $content
                
                $EmailParams = @{
                    credential = Import-CliXml -Path 'C:\Users\Antonio\Desktop\cred.xml' # this needs to be stored prior with "Get-Credential | Export-CliXml -Path C:\Users\Antonio\Desktop\cred.xml"
                    From = "antoniosigala2@outlook.com"
                    To = "asigala@sdiconsulting.com"
                    Subject = $filename
                    Body = $content
                    SMTPServer = "smtp-mail.outlook.com"
                    Port = "587"
                }
                Send-MailMessage @EmailParams -UseSsl
               
            }
            catch {
                Write-Host $_.ScriptStackTrace
                Write-Host -ForegroundColor Red "File is Locked ($file)"
                Add-content "C:\Users\Antonio\Desktop\Errors.txt" -value "File is Locked ($file)"
            }

            if($fileFree) {

                # do what we want with the file, since it is free
                #Move-Item $file $destination
                Write-Host -ForegroundColor Green "not locked, send some mail ($file)"
                #Add-content "C:\log.txt" -value "Moving file ($file)"


                # remove the current file from the array
                Write-Host -ForegroundColor Green "Done processing this file. ($file)"
                Add-content "C:\log.txt" -value "Done processing this file. ($file)"
                $global:files = $global:files | ? { $_ -ne $file }
            }
        }
    }

    sleep 2
}