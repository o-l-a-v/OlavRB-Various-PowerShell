#######################################
#######  Toggle on/off devices  #######
#######################################

[String] $StrNameScript = 'Toggle on/off web camera and internal mic'
[bool] $GUI = $true

# Devices
[String[]] $Devices = @('Internal Mic','Camera')


# Functions
Function Write-Out {
    Param(
        [Parameter(Mandatory=$true)]
        [String] $In,

        [Parameter(Mandatory=$false)]
        [bool] $StrOutputToUser = $false
        
    )

    Write-Output -InputObject $In

    If ($StrOutputToUser -and $GUI) {
        If (-not($Script:StrGUIOut)) {
            $Script:StrGUIOut = [String]::Empty
        }
        $Script:StrGUIOut += ('{0}{1}' -f ($In,"`r`n"))
    }

}


# Exit if not admin
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Out -StrOutputToUser $true -In 'This must be run as admin.'
}



Else {
    # Delete previous variables (mostly usefull when testing)
    $null = Clear-Variable -Name 'FirstDeviceStatus','StrGUIOut','StrDevicesFound' -ErrorAction SilentlyContinue

    # Help variables for output
    $StrDevicesFound = [string[]] @()

    # Loop devices, enable/disable based on first device status
    foreach ($Device in $Devices) {
        $DeviceInfo = Get-PnpDevice | Select-Object -Property DeviceID,FriendlyName,Status | Where-Object {$_.FriendlyName -match $Device} 
        Write-Out -In ('{0} | Found {1} device(s).' -f ($Device,(Measure-Object -InputObject $DeviceInfo).Count.ToString()))
        If ($DeviceInfo -ne $null) {       
            foreach ($D in $DeviceInfo) {
                $StrDevicesFound += $D.FriendlyName

                # Device status is either "OK" or "Error". Making it more readable with "enabled" or "disabled" instead.
                $StrStatus = [String] $(If($D.Status -eq 'OK'){'enabled'}Else{'disabled'})
                Write-Out -In ('   "{0}" current status: "{1}"' -f ($D.FriendlyName,$StrStatus))
                
                # Enable or disable based on the status of the first object
                If (-not($FirstDeviceStatus)) {
                    [String[]] $FirstDeviceStatus = @($D.Status,$StrStatus,$(If($D.Status -eq 'OK'){'disabled'}Else{'enabled'}))                    
                }
                
                # Change status if current device status is not equal to first device status
                If ($D.Status -ne $FirstDeviceStatus[0]) {
                    Write-Out -In ('      Device is already {0}.' -f ($StrStatus))
                }
                Else {
                    If ($FirstDeviceStatus -eq 'OK') {
                        Write-Out -In '      Disabling device...'
                        Disable-PnpDevice $D.DeviceID -Confirm:$false
                    } 
                    Else {
                        Write-Out -In '     Enabling device...'
                        Enable-PnpDevice $D.DeviceID -Confirm:$false
                    }
                    Write-Out -In ('         Success? {0}' -f ($?))
                }
            }
        }
    }
}

Write-Out -StrOutputToUser $true -In ('Following devices are now {0}:' -f ($FirstDeviceStatus[2]))
$StrDevicesFound | ForEach-Object {
    Write-Out -StrOutputToUser $true -In ('  - ' + $_)
}


If ($GUI) {

    ### Method 1  
    #$null = [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.MessageBox')  
    #[System.Windows.MessageBox]::Show($StrGUIOut,$StrNameScript)


    ### Method 2
    $null = [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $null = [System.Windows.Forms.MessageBox]::Show($StrGUIOut,$StrNameScript,[System.Windows.Forms.MessageBoxButtons]::OK)


    ### Method 3
    <#Add-Type -AssemblyName System.Windows.Forms.MessageBox
    $Form = New-Object system.Windows.Forms.MessageBox
    $Form.AcceptButton = 'OK'
    $Form.Text = $StrNameScript
    $Form.AutoScroll = $True
    $Form.AutoSize = $True
    $Form.AutoSizeMode = "GrowAndShrink"  # or GrowOnly
    $Font = New-Object System.Drawing.Font('Calibri',24,[System.Drawing.FontStyle]::Regular)  # Font styles are: Regular, Bold, Italic, Underline, Strikeout
    $Form.Font = $Font
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $StrGUIOut
    $Label.AutoSize = $True
    $Form.Controls.Add($Label)
    $Form.ShowDialog()#>


}