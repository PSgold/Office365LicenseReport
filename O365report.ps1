############################################START FUNCTION DEFINITIONS###########################################################
<#
function test-listBox{
    param(
        [Parameter()]$selectedItems,
        [Parameter()]$Data,
        [Parameter()]$formMain
    )

    foreach($s in $selectedItems){
        Add-Content -Path 'C:\Users\Shlomo\EFN Synced Folder\SG\Powershell\o365report\selected.txt' -Value $s
    }

    $Data|Export-Csv -Path 'C:\Users\Shlomo\EFN Synced Folder\SG\Powershell\o365report\data.csv' -NoTypeInformation
    $formMain.close()
}
#>

function Start-ErrorWin {
    param ([Parameter(Mandatory=$true)]$txt,
            [Parameter()][switch]$mainW
        )

    if($mainW){$formMain.hide()}

    $form = New-Object -TypeName System.Windows.Forms.Form
    $form.Size = New-Object -TypeName System.Drawing.Size(300,200)
    $form.Text = 'O365 Reports'
    $form.Opacity = .9
    $form.MaximizeBox = $false
    $form.FormBorderStyle = 1 #Fixed Single
    $form.StartPosition = 'CenterScreen'

    $label = New-Object -TypeName System.Windows.Forms.Label
    $label.Location = New-Object -TypeName System.Drawing.Point(20,20)
    $label.Size = New-Object -TypeName System.Drawing.Size(300,13)
    $label.Text = $txt

    $label2 = New-Object -TypeName System.Windows.Forms.Label
    $label2.Location = New-Object -TypeName System.Drawing.Point(20,35)
    $label2.Size = New-Object -TypeName System.Drawing.Size(230,13)
    $label2.Text = 'Please see the README for more info...'

    $okButton = New-Object -TypeName System.Windows.Forms.Button
    $okButton.Location = New-Object -TypeName System.Drawing.Point(100, 100)
    $okButton.Size = New-Object -TypeName System.Drawing.Size(70,23)
    $okButton.Text = "OK"
    $okButton.Add_Click({$form.Close()})
    
    $form.Controls.Add($label)
    $form.Controls.Add($label2)
    $form.Controls.Add($okButton)
    $form.ShowDialog()
}

function Get-Report{
    param(
        #[Parameter(Mandatory=$true)]$selectedItems,
        [Parameter(Mandatory=$true)]$Data
        #[Parameter(Mandatory=$true)]$formMain
    )
    $listbox.SetItemChecked(0,$false)
    $selectedItems = $listbox.CheckedItems
    $formMain.Controls.Remove($label)
    $formMain.Controls.Remove($listbox)
    $formMain.Controls.Remove($generateButton)
    $formMain.Controls.Remove($cancelButton)
    $formMain.Controls.Add($txtBox)
    $txtBox.Text = "Starting..."

    $formMain.Update()

    <# try{$encrypted = Get-Content -Path "$cwd\EPS.txt" -ErrorAction Stop}
    catch{Start-ErrorWin -txt "Missing EPS.txt" -mainW;$close = $true}
    $secure = ConvertTo-SecureString -String $encrypted
    try {$user = Get-Content -Path "$cwd\ur.txt" -ErrorAction Stop}
    catch{Start-ErrorWin -txt "Missing UR.txt" -mainW;$close = $true}
    if($close -eq $true){$formMain.Close()}
    else{
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user,$secure #>

        $day = [datetime]::Now.ToString("dd.MM.yyyy")
        $time = [datetime]::Now.ToString("HH_mm")
        $date = $day+"_"+$time

        Update-MainW -textout "Creating file"
        $filename = "O365report_${date}.csv"
        $filepath = "$cwd\$filename"
        $fileObj = [System.IO.StreamWriter]("$filepath")
        $fileObj.WriteLine("Client,Name,Email,License,LicenseID,,,SkuPartNumber,ConsumedUnits,PrePaidUnits")
        
        $firstLoop = $true
        foreach($s in $selectedItems){
            foreach($d in $data){
                if ($s -eq $d.Tenant){
                    Update-MainW -textout "Connecting to AAD: $s"
                    try{Connect-AzureAD -TenantId $d.ID -Credential $cred -ErrorAction Stop|Out-Null}
                    catch{$connect = $false}
                    if ($connect -eq $false){Update-MainW -textout "Failed to connect to $s";break}
                    
                    Update-MainW -textout "Getting license objects"
                    $tenantSkus = Get-AzureADSubscribedSku|Select-Object -Property skuid,skupartnumber,consumedunits,prepaidunits|Sort-Object -Property ConsumedUnits -Descending
                    $userLicense = Get-AzureADUser |Select-Object -Property displayname,userprincipalname, assignedlicenses|Sort-Object -Property displayname
                    $counter = 0;

                    Update-MainW -textout "Parsing and writing objects to file"
                    foreach ($u in $userLicense){
                        $license =""
                        $licenseCount = ($u.AssignedLicenses.skuid).count
                        if ($licenseCount -eq 0){
                            $license = "Unlicensed"
                            $licenseID = "Unlicensed"
                        }elseif ($licenseCount -eq 1) {
                            $licenseID = $u.assignedlicenses.skuid
                            foreach($ts in $tenantSkus){
                                if ($u.assignedlicenses.skuid -eq $ts.skuid){
                                    $license = $ts.SkuPartNumber
                                }    
                            }
                        }else{
                            for($c=0;$c -lt $licenseCount;$c++){
                                if($c -eq 0){
                                    $licenseID = $u.assignedlicenses.skuid[$c]    
                                    foreach($ts in $tenantSkus){
                                        if ($u.assignedlicenses.skuid[$c] -eq $ts.skuid){
                                            $license = $ts.SkuPartNumber
                                        }
                                    }
                                }
                                else{
                                    $licenseID = $licenseID, $u.assignedlicenses.skuid[$c] -join ";"
                                    foreach($ts in $tenantSkus){
                                        if ($u.assignedlicenses.skuid[$c] -eq $ts.skuid){
                                            $license = $license, $ts.SkuPartNumber -join ";"
                                        }
                                    }
                                }
                            }
                        }
                        $name = $u.DisplayName
                        $email = $u.UserPrincipalName
                        if($counter -ne 0){
                            if($tenantSkus[$counter]){
                                $partnumber = $tenantSkus[$counter].SkuPartNumber
                                $units = $tenantSkus[$counter].ConsumedUnits
                                $prepaidUnits = $tenantSkus[$counter].PrepaidUnits.Enabled
                                $fileObj.WriteLine(",$name,$email,$license,$licenseID,,,$partnumber,$units,$prepaidUnits")
                            }else{
                                $fileObj.WriteLine(",$name,$email,$license,$licenseID")
                            }
                        }else{
                            if($tenantSkus[$counter]){
                                $partnumber = $tenantSkus[$counter].SkuPartNumber
                                $units = $tenantSkus[$counter].ConsumedUnits
                                $prepaidUnits = $tenantSkus[$counter].PrepaidUnits.Enabled
                                if($firstLoop -ne $true){$fileObj.WriteLine("");$fileObj.WriteLine("")}
                                $fileObj.WriteLine("$s,$name,$email,$license,$licenseID,,,$partnumber,$units,$prepaidUnits")
                            }else{
                                if($firstLoop -ne $true){$fileObj.WriteLine("");$fileObj.WriteLine("")}
                                $fileObj.WriteLine("$s,$name,$email,$license,$licenseID")
                            }
                            $firstLoop = $false
                        }
                        $counter++
                    }
                    if(($tenantSkus.prepaidunits.enabled |Measure-Object -Sum).Sum -ne ($tenantSkus.ConsumedUnits|Measure-Object -Sum).Sum){
                        Update-MainW -textout "Writing unassigned licenses"
                        foreach($tenantSku in $tenantSkus){
                            if(($tenantSku.PrepaidUnits.Enabled -gt 500) -or ($tenantSku.SkuId -eq '6470687e-a428-4b7a-bef2-8a291ad947c9')){continue}
                            if ($tenantSku.ConsumedUnits -ne $tenantSku.prepaidunits.enabled){
                                $license = $tenantSku.SkuPartNumber
                                $licenseID = $tenantSku.SkuId
                                $unassignedcount = ($tenantSku.prepaidunits.enabled) - ($tenantSku.ConsumedUnits)
                                for($c=0;$c -lt $unassignedcount;$c++){
                                    $fileObj.WriteLine(",UNASSIGNED LICENSE, UNASSIGNED LICENSE,$license,$licenseID")
                                }
                            }
                        }
                    }
                    Disconnect-AzureAD
                    Update-MainW -textout "Disconnected AAD: $s"
                }else{continue}
            }
        }
        Update-MainW -textout "Saving file"
        $fileObj.Close()
        Update-MainW -textout "Finished..."
        $formMain.Controls.Add($finishButton)
        $formMain.Update()
}


function Start-MainW{
    
    $formMain = New-Object -TypeName System.Windows.Forms.Form
    $formMain.Size = New-Object -TypeName System.Drawing.Size(325,400)
    $formMain.Text = 'O365 Reports'
    $formMain.Opacity = .9
    $formMain.BackColor = 'Khaki'
    $formMain.MaximizeBox = $false
    $formMain.FormBorderStyle = 1 #Fixed Single
    $formMain.StartPosition = 'CenterScreen'

    $title = New-Object -TypeName System.Windows.Forms.Label
    $title.Location = New-Object -TypeName System.Drawing.Point(60,10)
    $title.Size = New-Object -TypeName System.Drawing.Size(200,24)
    $title.font = New-Object -TypeName System.Drawing.Font('Arial',12,[System.Drawing.FontStyle]::Bold)
    $title.ForeColor = 'MediumBlue'
    $title.Text = 'EFN O365 REPORTS'

    $label = New-Object -TypeName System.Windows.Forms.Label
    $label.Location = New-Object -TypeName System.Drawing.Point(20,50)
    $label.Size = New-Object -TypeName System.Drawing.Size(125,20)
    $label.font = New-Object -TypeName System.Drawing.Font('Arial',10,[System.Drawing.FontStyle]::Bold)
    $label.Text = "Tenant List:"
    $label.ForeColor = 'MidnightBlue'
    
    $listbox = New-Object -TypeName System.Windows.Forms.CheckedListBox
    $listbox.Location = New-Object -TypeName System.Drawing.Point(20,80)
    $listbox.Size = New-Object -TypeName System.Drawing.Size(180,180)
    $listbox.BackColor = 'OldLace'
    $listbox.ForeColor = 'Navy'
    $listbox.CheckOnClick = $true
    $listbox.TabIndex = 0
    $listbox.Items.Add('All')|Out-Null
    try {$tenants = Import-Csv -Path "$cwd\TenantList.csv" -ErrorAction Stop}
    catch{Start-ErrorWin -txt "Missing TenantList.csv";$close = $true}
    foreach($t in $tenants.Tenant){
        $listbox.Items.Add($t)|Out-Null
    }
    $listbox.Add_Click({
        if($this.SelectedItem -eq 'All' -and $listbox.GetItemChecked(0) -eq $false){
            for($c=1;$c -lt ($listbox.Items).Count;$c++){
                $listbox.SetItemChecked($c,$true)
            }
        }elseif($this.SelectedItem -eq 'All' -and $listbox.GetItemChecked(0) -eq $true){
            for($c=1;$c -lt ($listbox.Items).Count;$c++){
                $listbox.SetItemChecked($c,$false)
            }
        }
    })
    
    $generateButton = New-Object -TypeName System.Windows.Forms.Button
    $generateButton.Location = New-Object -TypeName System.Drawing.Point(220, 108)
    $generateButton.Size = New-Object -TypeName System.Drawing.Size(80,80)
    $generateButton.Text = "Generate Report"
    $generateButton.BackColor = "Yellow"
    $generateButton.ForeColor ="MediumBlue"
    $generateButton.Add_Click({Get-Report -Data $tenants})
    $generateButton.TabIndex = 1

    $cancelButton = New-Object -TypeName System.Windows.Forms.Button
    $cancelButton.Location = New-Object -TypeName System.Drawing.Point(115, 300)
    $cancelButton.Size = New-Object -TypeName System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.BackColor = "Blue"
    $cancelButton.ForeColor ="Yellow"
    $cancelButton.Add_Click({$formMain.Close()})
    $cancelButton.TabIndex = 2

    $finishButton = New-Object -TypeName System.Windows.Forms.Button
    $finishButton.Location = New-Object -TypeName System.Drawing.Point(115, 275)
    $finishButton.Size = New-Object -TypeName System.Drawing.Size(75,23)
    $finishButton.Text = "Finish"
    $finishButton.BackColor = "Blue"
    $finishButton.ForeColor ="Yellow"
    $finishButton.Add_Click({$formMain.Close()})

    $txtBox = New-Object -TypeName System.Windows.Forms.RichTextBox
    $txtBox.Location = New-Object -TypeName System.Drawing.Point(55,75)
    $txtBox.Size = New-Object -TypeName System.Drawing.Size(200,150)
    $txtBox.ReadOnly = $true

    $formMain.Controls.Add($title)
    $formMain.Controls.Add($label)
    $formMain.Controls.Add($listbox)
    $formMain.Controls.Add($generateButton)
    $formMain.Controls.Add($groupBox)
    $formMain.Controls.Add($cancelButton)
    
    if ($close -eq $true){
        $formMain.Close()
    }else{$formMain.BringToFront();$formMain.ShowDialog()|Out-Null}
}

function Update-MainW {
    param ([Parameter()]$textout)
    $txtBox.AppendText("${textout}`n")
    $formMain.Update()
}
###############################################END FUNCTION DEFINITIONS##########################################################
$cwd = Get-Location
$abort = $false
try{Import-Module -Name AzureAD -ErrorAction Stop}
catch{
    Start-ErrorWin -txt "Azure AD Powershell module is not installed."
    $abort = $true
}
if($abort -eq $false){
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $cred = Get-Credential
    Start-MainW
}