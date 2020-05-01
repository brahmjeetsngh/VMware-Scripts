#############################################################################
# Take VM Snapshot for multiple servers in GUI and send email report
#
#############################################################################									

Function Snapshot()
{
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    
    # Create the Main form.
    $form = New-Object System.Windows.Forms.Form 
    $form.Text = "VM Snapshot"
    $form.Size = New-Object System.Drawing.Size(650,320)
    $form.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $form.AutoSizeMode = 'GrowAndShrink'
    $form.Topmost = $True
    $form.ShowInTaskbar = $true  

    # Create the Email form.
    $Emailform = New-Object System.Windows.Forms.Form 
    $Emailform.Text = "Email Report"
    $Emailform.Size = New-Object System.Drawing.Size(420,200)
    $Emailform.FormBorderStyle = 'FixedSingle'
    $form.StartPosition = "CenterScreen"
    $Emailform.AutoSizeMode = 'GrowAndShrink'
    $Emailform.Topmost = $True
    $Emailform.ShowInTaskbar = $true  
    
    #Select vCenter
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Location = New-Object System.Drawing.Size(10,20) 
    $groupBox.size = New-Object System.Drawing.Size(180,80) 
    $groupBox.text = "Select the vCenter:" 
    $Form.Controls.Add($groupBox) 

    $RadioButton1 = New-Object System.Windows.Forms.RadioButton 
    $RadioButton1.Location = new-object System.Drawing.Point(15,15) 
    $RadioButton1.size = New-Object System.Drawing.Size(160,25) 
    $RadioButton1.Checked = $true 
    $RadioButton1.Text = "PROD" 
    $groupBox.Controls.Add($RadioButton1) 

    $RadioButton2 = New-Object System.Windows.Forms.RadioButton
    $RadioButton2.Location = new-object System.Drawing.Point(15,40)
    $RadioButton2.size = New-Object System.Drawing.Size(160,25)
    $RadioButton2.Text = "LAB"
    $groupBox.Controls.Add($RadioButton2)

    #Select Snapshot memory
    $groupBox1 = New-Object System.Windows.Forms.GroupBox
    $groupBox1.Location = New-Object System.Drawing.Size(200,20) 
    $groupBox1.size = New-Object System.Drawing.Size(180,80) 
    $groupBox1.text = "Snapshot Memory:" 
    $Form.Controls.Add($groupBox1) 

    $RadioButton3 = New-Object System.Windows.Forms.RadioButton 
    $RadioButton3.Location = new-object System.Drawing.Point(15,15) 
    $RadioButton3.size = New-Object System.Drawing.Size(160,25) 
    $RadioButton3.Checked = $true 
    $RadioButton3.Text = "Snapshot Without Memory" 
    $groupBox1.Controls.Add($RadioButton3) 

    $RadioButton4 = New-Object System.Windows.Forms.RadioButton
    $RadioButton4.Location = new-object System.Drawing.Point(15,40)
    $RadioButton4.size = New-Object System.Drawing.Size(160,25)
    $RadioButton4.Text = "Snapshot With Memory"
    $groupBox1.Controls.Add($RadioButton4)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(390,20) 
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.AutoSize = $true
    $label.Text = "Enter Hostname Here..."

    # Create the TextBox used to capture the user's text.
    $textBox = New-Object System.Windows.Forms.TextBox 
    $textBox.Location = New-Object System.Drawing.Size(390,40) 
    $textBox.Size = New-Object System.Drawing.Size(130,230)
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $false
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textbox.CharacterCasing='Upper'

    # Getting Input for Snapshot SR#, Snapshot Name, Snapshot Description.
    $SR_label = New-Object System.Windows.Forms.Label
    $SR_label.Location = New-Object System.Drawing.Size(15,130) 
    $SR_label.Size = New-Object System.Drawing.Size(280,20)
    $SR_label.AutoSize = $true
    $SR_label.Text = "*SR#"

    $SR_textBox = New-Object System.Windows.Forms.TextBox 
    $SR_textBox.Location = New-Object System.Drawing.Size(150,130) 
    $SR_textBox.Size = New-Object System.Drawing.Size(130,20)
    $SR_textBox.AcceptsReturn = $true
    $SR_textBox.AcceptsTab = $false
    $SR_textbox.CharacterCasing='Upper'
    
    $SN_label = New-Object System.Windows.Forms.Label
    $SN_label.Location = New-Object System.Drawing.Size(15,160) 
    $SN_label.Size = New-Object System.Drawing.Size(280,20)
    $SN_label.AutoSize = $true
    $SN_label.Text = "*Snapshot Name"
    
    $SN_textBox = New-Object System.Windows.Forms.TextBox 
    $SN_textBox.Location = New-Object System.Drawing.Size(150,160) 
    $SN_textBox.Size = New-Object System.Drawing.Size(200,20)
    $SN_textBox.AcceptsReturn = $true
    $SN_textBox.AcceptsTab = $false
    #$SN_textbox.CharacterCasing='Upper'

    $SD_label = New-Object System.Windows.Forms.Label
    $SD_label.Location = New-Object System.Drawing.Size(15,190) 
    $SD_label.Size = New-Object System.Drawing.Size(280,20)
    $SD_label.AutoSize = $true
    $SD_label.Text = "Snapshot Description"

    $SD_textBox = New-Object System.Windows.Forms.TextBox 
    $SD_textBox.Location = New-Object System.Drawing.Size(150,190) 
    $SD_textBox.Size = New-Object System.Drawing.Size(200,50)
    $SD_textBox.AcceptsReturn = $true
    $SD_textBox.AcceptsTab = $false
    $SD_textBox.Multiline = $true
    $SD_textBox.ScrollBars = 'Both'
    #$SD_textbox.CharacterCasing='Upper'
        
    #Create the Hardening Button.
    $HButton = New-Object System.Windows.Forms.Button
    $HButton.Location = New-Object System.Drawing.Size(540,40)
    $HButton.Size = New-Object System.Drawing.Size(100,40)
    $HButton.Text = "Take Snapshot"

    #Create the Report Button.
    $RButton = New-Object System.Windows.Forms.Button
    $RButton.Location = New-Object System.Drawing.Size(540,90)
    $RButton.Size = New-Object System.Drawing.Size(100,40)
    $RButton.Text = "Generate Report"

    #Create the Report Button.
    $EmailButton = New-Object System.Windows.Forms.Button
    $EmailButton.Location = New-Object System.Drawing.Size(540,140)
    $EmailButton.Size = New-Object System.Drawing.Size(100,40)
    $EmailButton.Text = "Send Email"

    #Create the Progress-Bar.
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Size(20,230) 
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.AutoSize = $true
    $label1.Text = "Progress..."

    $PB = New-Object System.Windows.Forms.ProgressBar
	$PB.Name = "PowerShellProgressBar"
	$PB.Value = 0
	$PB.Style="Continuous"

    $System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Width = 200 - 40
	$System_Drawing_Size.Height = 20
	$PB.Size = $System_Drawing_Size
	$PB.Left = 20
	$PB.Top = 250
    
    #Initiate Snapshot
    $HButton.Add_Click(
    { 
        $report= @()
        $counter = 0
        If ($SR_textBox.TextLength -ne 0 -and $SN_textBox.TextLength -ne 0)
        {
                $S_Name= $SR_textBox.text+ " - " +$SN_textBox.text  
        }
        Else{
            [System.Windows.Forms.MessageBox]::Show("SR# or Snapshot Name cannot be blank", "Info")
            return
        }
	   
        If ($textbox.TextLength -eq 0)
        {
            [System.Windows.Forms.MessageBox]::Show("Server List is empty", "Info")
            return       
        }
        [System.Windows.Forms.MessageBox]::Show("Sit back and relax while Snapshot is taken !!!", "VM Snapshot")
        $ServerList=$textbox.Text.Split("`n")|%{$_.trim()}
        Foreach ($vm in $ServerList)
        {
            if($vm -eq "")
            {
                $counter++
                [Int]$Percentage = ($Counter/$ServerList.Count)*100
                $PB.Value = $Percentage
                continue
            }
                
            if ($RadioButton1.Checked -eq $True) 
            {
                Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue
                Connect-VIServer -Server "Enter first vCenter Server FQDN"
            }
            if ($RadioButton2.Checked -eq $True) 
            {
                Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue
                Connect-VIServer -Server "Enter Second vCenter Server FQDN"
            }
            $Exists = get-vm -name $vm -ErrorAction SilentlyContinue
            If ($Exists)
            {
                If ($RadioButton3.Checked -eq $True)
                {
                    Get-VM $vm |New-snapshot -Name $S_Name -Description $SD_textBox.Text
                    $rep = Get-VM $vm | Get-Snapshot | Select-Object @{Name='VM';Expression={$_.vm}},@{Name='Snapshot_Name';Expression={$_.name}},@{Name='Description';Expression={$_.Description}},@{Name='Created';Expression={$_.Created}},@{Name='Remarks';Expression={""}}
                    $report = $report + $rep
                }
                If ($RadioButton4.Checked -eq $True)
                {
                    Get-VM $vm |New-snapshot -Name $S_Name -Description $SD_textBox.Text -Memory
                    $rep = Get-VM $vm | Get-Snapshot | Select-Object @{Name='VM';Expression={$_.vm}},@{Name='Snapshot_Name';Expression={$_.name}},@{Name='Description';Expression={$_.Description}},@{Name='Created';Expression={$_.Created}},@{Name='Remarks';Expression={""}}
                    $report = $report + $rep
                }
            }
            If (!$Exists)
            {
                $row= New-Object PSObject -Property @{VM = $vm;Snapshot_Name = "";Description = "";Created = "";Remarks="Server not Found"}
                $report += $row
            }
            $counter++
            [Int]$Percentage = ($Counter/$ServerList.Count)*100
            $PB.Value = $Percentage
        }
        [System.Windows.Forms.MessageBox]::Show("Snapshot taken successfully" , "Report Generation")
        $report |Select-object @{Name="HOSTNAME"; Expression={$_.VM}},@{Name="Snapshot_Name"; Expression={$_.Snapshot_Name}},@{Name="Description"; Expression={$_.Description}},@{Name="Created: Date & Time"; Expression={$_.Created}},@{Name='Remarks';Expression={$_.Remarks}}| Export-Csv $file -NoTypeInformation
    })

    #Report Generation
    $RButton.Add_Click(
    {    
		ii $path
    })

    #Send Email
    $EmailButton.Add_Click(
    {   
        #Create Label
        $ToLabel = New-Object System.Windows.Forms.Label
        $ToLabel.Location = New-Object System.Drawing.Size(20,20) 
        $ToLabel.Size = New-Object System.Drawing.Size(100,20)
        $ToLabel.AutoSize = $true
        $ToLabelFont = New-Object Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Bold)
        $ToLabel.Font = $ToLabelFont
        $ToLabel.Text = "*To:"

        $CcLabel = New-Object System.Windows.Forms.Label
        $CcLabel.Location = New-Object System.Drawing.Size(20,50) 
        $CcLabel.Size = New-Object System.Drawing.Size(100,20)
        $CcLabel.AutoSize = $true
        $CcLabelFont = New-Object Drawing.Font("Times New Roman",12,[System.Drawing.FontStyle]::Bold)
        $CcLabel.Font = $CcLabelFont
        $CcLabel.Text = "Cc: (Optional)"

        #Create the TextBox Email Address.
        $ToBox = New-Object System.Windows.Forms.TextBox 
        $ToBox.Location = New-Object System.Drawing.Size(140,20) 
        $ToBox.Size = New-Object System.Drawing.Size(250,20)
        $ToBoxFont = New-Object Drawing.Font("Times New Roman",8)
        $ToBox.Font=$ToBoxFont
        $ToBox.AcceptsReturn = $true
        $ToBox.AcceptsTab = $false
        #$ToBox.text=""
        $ToBox.CharacterCasing='lower'

        $CcBox = New-Object System.Windows.Forms.TextBox 
        $CcBox.Location = New-Object System.Drawing.Size(140,50) 
        $CcBox.Size = New-Object System.Drawing.Size(250,20)
        $CcBoxFont = New-Object Drawing.Font("Times New Roman",8)
        $CcBox.Font=$CcBoxFont
        $CcBox.AcceptsReturn = $true
        $CcBox.AcceptsTab = $false
        $CcBox.text=""
        $CcBox.CharacterCasing='lower'
        
        #Create Email Send Button
        $SendButton = New-Object System.Windows.Forms.Button
        $SendButton.Location = New-Object System.Drawing.Size(20,100)
        $SendButton.Size = New-Object System.Drawing.Size(100,40)
        $ButtonFont = New-Object Drawing.Font("Times New Roman",8,[System.Drawing.FontStyle]::Bold)
        $SendButton.Font=$ButtonFont
        $SendButton.Text = "Send Email"

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Size(150,100)
        $CancelButton.Size = New-Object System.Drawing.Size(100,40)
        $CancelButton.Font=$ButtonFont
        $CancelButton.Text = "Cancel"

        $Emailform.Controls.Add($ToLabel)
        $Emailform.Controls.Add($CcLabel)
        $Emailform.Controls.Add($ToBox)
        $Emailform.Controls.Add($CCBox)
        $Emailform.Controls.Add($SendButton)
        $Emailform.Controls.Add($CancelButton)
        
        $SendButton.Add_Click(
        {
            If ($ToBox.Text.Length -eq 0)
            {
                [System.Windows.Forms.MessageBox]::Show("Email address cannot be empty" , "Email Info")
                return
            }
            #---------------------------------------------------------------------
            # Generate the HTML report and output to file
            #---------------------------------------------------------------------

            $head = "<style>"
            $head = $head + "BODY{background-color:white;}"
            $head = $head + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
            $head = $head + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#778899}"
            $head = $head + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black}"
            $head = $head + "</style>"

            # SMTP info
            $Toemail=$ToBox.Text
            $strTo=$Toemail
            $strCc=$CCBox.Text
            $strSubject = "Snapshot Taken : $S_Name"
            $StrMsg="Hi All, <br>Snapshot has been taken successfully for below list of servers <br><br>"
            $strBody = "Attached is the list of Snapshots"
            $strMail = $strmsg

            # Write the output to an HTML file
            $strOutFile = $Path+"email.html"
            $ComputerName=(Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name
            Get-Content -Path $File |ConvertFrom-CSV | ConvertTo-HTML  -Head $head -Body $StrMsg | Out-File $StrOutFile
            
	        # Mail the output file
	        $msg = new-object Net.Mail.MailMessage
	        $att = new-object Net.Mail.Attachment($File)
	        $smtp = new-object Net.Mail.SmtpClient("smtp.corp.ad.zalando.net")
            $msg.From ="vcenter@zalando.de"
	        $msg.To.Add($strTo)
            If ($strCc.Length -ne 0)
            {
                $msg.cc.Add($strcc)
            }
	        $msg.Subject = $strSubject
	        $msg.IsBodyHtml = 1
	        $msg.Body = Get-Content $strOutFile
	        $msg.Attachments.Add($att)
            $smtp.Send($msg)
                [System.Windows.Forms.MessageBox]::Show("Email Send !!!","Info")
                $Emailform.close()
        })
        $CancelButton.Add_Click(
        {
            $Emailform.close()
        })
        $Emailform.Add_Shown({$Emailform.Activate()})
        $Emailform.ShowDialog() > $null 
               
    })
        
    # Add all of the controls to the form.
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($groupBox)
    $form.Controls.Add($groupBox1)
    $form.Controls.Add($RButton)
    $form.Controls.Add($HButton)
    $form.Controls.Add($label1)
    $form.Controls.Add($SR_label)
    $form.Controls.Add($SR_textBox)
    $form.Controls.Add($SN_label)
    $form.Controls.Add($SN_textBox)
    $form.Controls.Add($SD_label)
    $form.Controls.Add($SD_textBox)
    $form.Controls.Add($PB)
    $form.Controls.Add($EmailButton)
    # Initialize and show the form.
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() > $null
}
Set-ExecutionPolicy unrestricted -Force
$date=get-Date -format "ddMMyy_HHmm"
$ComputerName=(Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name
$Path=[System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)+"\Report\"
If ((Test-Path $Path) -eq $false)
{
    New-Item $Path -type directory
}
New-Item -ErrorAction Ignore -ItemType directory -Path Report
$File=$Path + "SnapshotReport_$date.csv"
Snapshot