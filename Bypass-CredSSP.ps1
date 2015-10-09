# Credential Functions
# http://powershell.com/cs/blogs/tips/archive/2014/03/28/exporting-and-importing-credentials-in-powershell.aspx
# The scheduled task must be running under the user account that was used to generate the XML credential file
. \\networkresource\security\CredentialFunctions.ps1
$cred = Import-Credential \\networkresource\security\credentials\SchedTaskUser.xml
$username = $cred.UserName
$Password = $cred.getnetworkcredential().password

#Import CSV containing failing job - client identifier and Source Server for failing job
$FailedClients = Import-CSV "\\networkresource\cubes\failedjobs.csv"

# Loop through the CSV
foreach ($clients in $FailedClients) {
  # Get the site code and the cube server
	$SiteCode = $clients.siteCode
	$CubeServer = $clients.cubeServer
	
	# This is the script that we want to execute.  Need to use tick marks to prevent variable expansion
	$Script1 = @"
		`$SiteCode = '$SiteCode'
		`$CubeServer = '$CubeServer'
        . \\NetworkResouce\Cubes\CubeFunctions.ps1
        Get-CubeInfo -server `$CubeServer -client `$SiteCode | ForEach {Set-CubeConnectionStringInfo `$_}
        \\NetworkResource\Cubes\CubesSetup.ps1 `$SiteCode
"@

  # Now convert the script to base64
	$Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script1))
	
	# Write it out to a file - the encoded string is too long to pass to psexec.exe
	if ( test-path -path \\NetworkResource\Cubes\CUBEFAILURES\b64 ) {
		Remove-Item \\NetworkResource\Cubes\CUBEFAILURES\b64
	}
	$Encoded | out-file -encoding ascii \\NetworkResource\Cubes\CUBEFAILURES\b64
	
	# Second script - this one reads the file written above, and then runs powershell with the output
	# passed to the encodedcommand parameter
	$Script2 = '$a=gc \\NetworkResource\Cubes\CUBEFAILURES\b64
		powershell.exe -EncodedCommand $a'
	
	# Convert this to base64	
	$Base64 = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script2))
	
	# Execute PSExec against the remote server with the credentials from the file
	# accept the EULA so there is no interaction required, and pass the encoded command
	\\NetworkResource\tools\Standalone\PsExec.exe \\$CubeServer -u $username -p $Password -accepteula powershell.exe -EncodedCommand $Base64
} 
