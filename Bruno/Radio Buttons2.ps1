Clear-Host
$AllData = $false

# Function to create the Filter form
function Filter_Form{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 430
    $Form.height = 460
    $Form.Text = "Options To Filter By"
    $Form.BackColor = "lightblue"
    $Form.startposition = "centerscreen"

    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Calibri",12)
    $Form.Font = $Font

    # Create a group that will contain your radio buttons
    $MyGroupBox = New-Object System.Windows.Forms.GroupBox
    $MyGroupBox.Location = '50,30'
    $MyGroupBox.size = '300,210' #width,height
    $MyGroupbox.BackColor = "white"
    $MyGroupBox.text = "Filter By..."

    #Create the collection of radio buttons
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton
    $RadioButton1.Location = '20,40'
    $RadioButton1.size = '360,30'
    $RadioButton1.Checked = $true
    $RadioButton1.Text = "All Data"

    $RadioButton2 = New-Object System.Windows.Forms.RadioButton
    $RadioButton2.Location = '20,70'
    $RadioButton2.size = '360,30'
    $RadioButton2.Checked = $false
    $RadioButton2.Text = "Filename"

    $RadioButton3 = New-Object System.Windows.Forms.RadioButton
    $RadioButton3.Location = '20,100'
    $RadioButton3.size = '360,30'
    $RadioButton3.Checked = $false
    $RadioButton3.Text = "Song Title"

    $RadioButton4 = New-Object System.Windows.Forms.RadioButton
    $RadioButton4.Location = '20,130'
    $RadioButton4.size = '360,30'
    $RadioButton4.Checked = $false
    $RadioButton4.Text = "Album"

    $RadioButton5 = New-Object System.Windows.Forms.RadioButton
    $RadioButton5.Location = '20,160'
    $RadioButton5.size = '360,30'
    $RadioButton5.Checked = $false
    $RadioButton5.Text = "Album Artist"

    #Create Inputbox
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(50,260)
    $label.Size = New-Object System.Drawing.Size(300,30)
    $label.Text = 'Enter name of to filter by...'
    $label.Visible = $true
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(50,290)
    $textBox.Size = New-Object System.Drawing.Size(300,150)
    $textBox.BackColor = "white"
    $textBox.Visible = $true
    $form.Controls.Add($textBox)

    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '75,340' #Horizontal axis, Vertical axis
    $OKButton.Size = '100,40'
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK

    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '200,340' #Horizontal axis, Vertical axis
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel

    # Add all the GroupBox controls on one line
    $MyGroupBox.Controls.AddRange(@($Radiobutton1,$RadioButton2,$RadioButton3,$Radiobutton4,$Radiobutton5))

    # Add all the Form controls on one line
    $form.Controls.AddRange(@($MyGroupBox,$OKButton,$CancelButton,$label,$textBox))

    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton

    # Activate the form
    $form.Add_Shown({$form.Activate()})

    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    # If the OK button is selected
    if ($dialogResult -eq "OK"){

        # Check the current state of each radio button and respond accordingly
        if ($RadioButton1.Checked){
            $AllData = $true
            Write-Host AllData_inside = $AllData}
        elseif ($RadioButton2.Checked){
                $FileNameData = $true
                Write-Host FileName_inside = $textBox.Text
                If ($textBox.TextLength -eq 0)
                    { Write-Host You forgot to type a Filename -BackgroundColor Red -ForegroundColor Yellow }}
        elseif ($RadioButton3.Checked){
                If ($textBox.TextLength -eq 0)
                    { Write-Host You forgot to type a Song title -BackgroundColor Red -ForegroundColor Yellow }}
        elseif ($RadioButton4.Checked){
                If ($textBox.TextLength -eq 0)
                    { Write-Host You forgot to type a Album name -BackgroundColor Red -ForegroundColor Yellow }}
        elseif ($RadioButton5.Checked){
                If ($textBox.TextLength -eq 0)
                    { Write-Host You forgot to type an Album Artist -BackgroundColor Red -ForegroundColor Yellow }}
    }

    return [PsCustomObject]@{
        AllData = $AllData
        TextBoxText = $textBox.Text
    }
}

$Result = Filter_Form
Write-Host `n"AllData_Outside : " $Result.AllData  #displaying False from line 2
Write-Host "FileName_Outside: " $Result.TextBoxText #displaying no text although i enter text inbox

