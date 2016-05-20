Function Get-Folder($initialDirectory) {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.RootFolder = [System.Environment+SpecialFolder]'MyComputer'
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Choose a directory"

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
            $loop = $false
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Try again or exit script?", "Choose a directory", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                return
            }
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
}

Function getDeploymentInfo(){
	
	$output = @()
	#get deployment info for last night
	$deploymentHolder = Get-CMDeployment | Where-Object {$_.EnforcementDeadline -gt (Get-Date).AddDays(-1)} | Select SoftwareName,CollectionID,EnforcementDeadline,NumberErrors,NumberInProgress,NumberSuccess,NumberTargeted

	foreach($deployment in $deploymentHolder){
		$deploymentInfo = New-Object -TypeName PSObject -Property @{
			SoftwareName = $deployment.SoftwareName
			CollectionID = $deployment.CollectionID
			EnforcementDeadline = $deployment.EnforcementDeadline
			NumberErrors = $deployment.NumberErrors
			NumberInProgress = $deployment.NumberInProgress
			NumberSuccess = $deployment.NumberSuccess
			NumberTargeted = $deployment.NumberTargeted
			Percentage = try {
				if($deployment.NumberTargeted -eq "0"){
				"0.00 %"
				}
				else{
				($deployment.NumberSuccess/$deployment.NumberTargeted).toString("P")
				}
			}
			Catch {}
		}
			
		$output += $deploymentInfo
	}
	
	return $output
}

$folderLoc = Get-Folder


Import-Module ConfigurationManager

$yesterday = (Get-Date).AddDays(-1).toString('MM-dd-yyyy')


$DEV = ""
$PROD = ""

CD $DEV

$DEVOut = getDeploymentInfo
$DEVOut | Select SoftwareName,CollectionID,EnforcementDeadline,NumberErrors,NumberInProgress,NumberSuccess,NumberTargeted,Percentage | Export-Csv "$folderLoc\DEV_$yesterday.csv" -noTypeInformation -append

CD $PROD

$PRODOut = getDeploymentInfo
$PRODOut | Select SoftwareName,CollectionID,EnforcementDeadline,NumberErrors,NumberInProgress,NumberSuccess,NumberTargeted,Percentage | Export-Csv "$folderLoc\PROD_$yesterday.csv" -noTypeInformation -append


cd $folderLoc

