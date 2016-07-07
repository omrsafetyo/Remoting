Function Get-SchTasks {
    [OutputType([PSObject])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)] 
        $ScheduleCom,
        [Parameter(Mandatory=$false)]
        $User,
        [Parameter(Mandatory=$false)]
        $TaskNameFilter,
        [Parameter(Mandatory=$false)]
        $TaskCommandFilter,
        [Parameter(Mandatory=$false,ValueFromPipeline=$True)]
        $Computername = $ENV:COMPUTERNAME,
        [Parameter(Mandatory=$false)]
        [switch] $Recurse = $false
    )
    begin {
        if ( -not($ScheduleCom) ) {
            $ScheduleCom = New-Object -ComObject "Schedule.Service"
        }
    }
    process {
        $tasks = New-Object System.Collections.ArrayList
        $ScheduleCom.Connect($Computername)
        
        if ( $recurse ) {
            $SubDirs = Get-SchSubFolders -ScheduleCom $ScheduleCom -BaseDirectory "\"
            ForEach ( $SubDir in $SubDirs ) {
                $tasks += @($ScheduleCom.getfolder($SubDir).gettasks(0))
            }
        }
        else {
            $tasks += @($ScheduleCom.getfolder("\").gettasks(0))
        }

        foreach ( $task in $tasks ) {
            if ( $task -eq $null ) {
                # v2 bug
                continue
            }
            
            # Grab the Xml 
            $xmlTask = [xml]($task.Xml)
            
            # Create an empty object to populate
            $objTask = "" | Select-Object Computername, Name, Author, Command, Arguments, WorkingDirectory, UserId, Enabled, LastRunTime, NextRunTime, Schedule, Description
            
            #pull the attributes from the XMl
            $objTask.Computername = $Computername
            $objTask.Name = $task.Name
            $objTask.Author = $xmlTask.Task.RegistrationInfo.Author
            $objTask.UserId = $xmlTask.Task.Principals.Principal.UserId
            $objTask.Command = $xmlTask.Task.Actions.Exec.Command
            $objTask.Arguments = $xmlTask.Task.Actions.Exec.Arguments
            $objTask.WorkingDirectory = $xmlTask.Task.Actions.Exec.WorkingDirectory
            $objTask.Enabled = $xmlTask.Task.Settings.Enabled
            $objTask.LastRunTime = $task.LastRunTime
            $objTask.NextRunTime = $task.NextRunTime
            $objTask.Schedule = $xmltask.Task.Triggers.InnerXml
            $objTask.Description = $xmlTask.Task.RegistrationInfo.Description
            
            # If filters or a user was provided, only return tasks that match
            # Task Name filter
            if ( $TaskNameFilter ) {
                if ( $objTask.Name -match $TaskNameFilter ) {
                    $objTask
                }
                else {
                    continue
                }
            }
            
            # Command filter
            if ( $TaskCommandFilter ) {
                if ( $objTask.Command -match $TaskCommandFilter ) {
                    $objTask
                }
                else {
                    continue
                }
            }
            
            # Tasks only matching a specific user
            if ( $User ) {
                if ( $objTask.UserId -match $User ) {
                    $objTask
                }
                else {
                    continue
                }
            }
            # Otherwise, no filters, return all tasks
            $objTask
        }
    }
    end {
        # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ScheduleCom) | Out-Null
    }
}

Function Get-SchSubFolders {
    [OutputType([PSObject])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] 
        $ScheduleCom,
        [Parameter(Mandatory=$true)]
        $BaseDirectory
    )
    
    process {
        $SubFolders = $ScheduleCom.GetFolder($BaseDirectory).GetFolders(0)
        if ( $SubFolders.Count -eq 0 ) {
            return
        }
        if ( $BaseDirectory -eq "\" ) {
            # return
            $BaseDirectory
        }
        
        ForEach ( $SubFolder in $SubFolders ) {
            $SubFolderPath = $SubFolder.Path
            # return
            $SubFolderPath
            Get-SchSubFolders -ScheduleCom $ScheduleCom -BaseDirectory "$SubFolderPath"
        }
    }
}

Function Get-SchTasksFromAllServers {
    [OutputType([PSObject])]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        $User,
        [Parameter(Mandatory=$false)]
        $TaskNameFilter,
        [Parameter(Mandatory=$false)]
        $TaskCommandFilter,
        [Parameter(Mandatory=$false)]
        $OutputFile
    )
    
    # TODO: Add parameter for input file, or input list of servers
    
    process {
        $output = New-Object System.Collections.ArrayList
        $scheduleCom = New-Object -ComObject "Schedule.Service"

        # AD for computer list
        Import-Module ActiveDirectory
        # The computer account should have a password update in the last 30 days
        $limitDate = (Get-Date).AddDays(-30)
        $servers = Get-ADComputer -Prop PasswordLastSet, IPv4Address, OperatingSystem -Filter "PasswordLastSet -gt '$limitDate'" | 
            Where-Object { $_.IPv4Address -ne $null -and $_.OperatingSystem -match "Windows Server" }  | 
            select -expand Name
        # Only servers that respond to pint
        $servers = @( $servers | Where-Object { (Test-Connection $_ -quiet -count 2) -eq $True } )
        ForEach ( $server in $servers ) {
            $ServerTasks = @()
            $ServerTasks = Get-SchTasks -ScheduleCom $scheduleCom -User $User -Computername $Server -TaskNameFilter $TaskNameFilter -TaskCommandFilter $TaskCommandFilter
            if ( $OutputFile ) {
                $output += $ServerTasks
            } else {
                $ServerTasks
            }
        }
        
        if ( $output.Count -gt 0 ) {
            $output | export-csv -notypeinformation $OutputFile
        }
    }
}
