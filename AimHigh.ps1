#Load Windows Forms Assembly
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

#For debugging in ISE
remove-variable * -ea 0

#Serialized training data
$TrainingOptions = @(
    @{
        ShortName = "cidod"
        LongName = "Counterintelligence Awareness and Reporting for DoD"
    }
    @{
        ShortName = "ci-security-brief"
        LongName = "Counterintelligence Awareness and Security Brief"
    }
    @{
        ShortName = "integrating"
        LongName = "Integrating CI and Threat Awareness into Your Security Program"
    }
    @{
        ShortName = "thwarting"
        LongName = "Thwarting the Enemy: Providing Counterintelligence and Threat Awareness to the Defense Industrial Base"
    }
    @{
        ShortName = "cybersecurity"
        LongName = "Cybersecurity Awareness"
    }
    @{
        ShortName = "rmf"
        LongName = "Introduction to the Risk Management Framework (RMF)"
    }
    @{
        ShortName = "derivative"
        LongName = "Derivative Classification"
    }
    @{
        ShortName = "awarenessrefresher"
        LongName = "DoD Annual Security Awareness Refresher"
    }
    @{
        ShortName = "initialorientation"
        LongName = "DoD Initial Orientation and Awareness Training"
    }
    @{
        ShortName = "markings"
        LongName = "Marking Classified Information"
    }
    @{
        ShortName = "oca"
        LongName = "Original Classification"
    }
    @{
        ShortName = "disclosure"
        LongName = "Unauthorized Disclosure of Classified Information for DoD and Industry"
    }
    @{
        ShortName = "insiderthreatprgm"
        LongName = "Establishing an Insider Threat Program"
    }
    @{
        ShortName = "itawareness"
        LongName = "Insider Threat Awareness"
    }
    @{
        ShortName = "opsec"
        LongName = "OPSEC Awareness for Military Members, DoD Employees and Contractors"
    }
)

#Creates the form and sets its size and position
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Select trainings to complete"
$objForm.Size = New-Object System.Drawing.Size(300,615) 
$objForm.StartPosition = "CenterScreen"
$objForm.FormBorderStyle = 5

#This creates a label for the TextBox1
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(8,10) 
$objLabel1.Size = New-Object System.Drawing.Size(280,20) 
$objLabel1.Text = "Enter full name:"
$objForm.Controls.Add($objLabel1) 

#This creates the TextBox1
$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(10,30) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20)
$objTextBox1.TabIndex = 0 
$objForm.Controls.Add($objTextBox1)

#This creates a label for the trainings
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(8,60) 
$objLabel2.Size = New-Object System.Drawing.Size(280,20) 
$objLabel2.Text = "Select trainings:"
$objForm.Controls.Add($objLabel2) 

#Dynamically creates checkboxes from $TrainingOptions
$YLoc = 80
ForEach ($Option in $TrainingOptions) {
    $VarName = "objCheckbox$($Option.ShortName)"
    New-Variable -Name $VarName
    $VarName = New-Object System.Windows.Forms.Checkbox 
    $VarName.Location = New-Object System.Drawing.Size(10,$Yloc) 
    $VarName.Size = New-Object System.Drawing.Size($objForm.Size.Width,30)
    $VarName.MaximumSize = New-Object System.Drawing.Size($($objForm.Size.Width - 25),40)
    $VarName.Text = $Option.LongName
    $VarName.TabIndex = 4
    $objForm.Controls.Add($VarName)
    $Yloc = $YLoc + 30
}

#Draws form
$objForm.ShowDialog()

#Get learners name from form
$Learner = $objTextBox1.Text

#Turn data from checkbox back into serialized data from $TrainingOptions
$Selection = @()
$CheckedTraining = $objForm.Controls | Where {$_.Checked -eq $true}
ForEach ($item in $CheckedTraining) {
    $Selection += $TrainingOptions | Where {$_.LongName -eq $item.Text}
}

#IE Method
$Payload = "javascript:window.parent.GetPlayer().SetVar(`"learner`",`"$Learner`");savePDF()"
$ie = New-Object -com internetexplorer.application
$ie.visible = $true
ForEach ($item in $Selection) {
    $ie.navigate("https://securityawareness.usalearning.gov/$($item.ShortName)/quiz/story_html5.html")
    While($ie.busy){
        Start-Sleep -Seconds 1
    }
    $ie.navigate($Payload)
    While($ie.busy){
        Start-Sleep -Seconds 1
    }

    <#
    $wshell = New-Object -ComObject WScript.Shell
    $id = (Get-Process iex* | where {$_.MainWindowTitle -match "View Downloads - Internet Explorer" -or $_.MainWindowTitle -match "Derivative Classification - Internet Explorer"}).id
    $wshell.AppActivate($id)
    start-sleep 1
    $wshell.SendKeys("{S}")
    Start-Sleep 1
    #>
}
