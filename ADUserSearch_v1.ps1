# Import the ActiveDirectory module
Import-Module ActiveDirectory

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Intertek AD User Search v1.0"
$form.Size = New-Object System.Drawing.Size(350, 150)
$form.StartPosition = "CenterScreen"

# Create the domain label
$domainLabel = New-Object System.Windows.Forms.Label
$domainLabel.Location = New-Object System.Drawing.Point(10, 20)
$domainLabel.Size = New-Object System.Drawing.Size(100, 20)
$domainLabel.Text = "Domain:"
$form.Controls.Add($domainLabel)

# Create the domain text box
$domainTextBox = New-Object System.Windows.Forms.TextBox
$domainTextBox.Location = New-Object System.Drawing.Point(120, 20)
$domainTextBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($domainTextBox)

# Create the search label
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(10, 50)
$searchLabel.Size = New-Object System.Drawing.Size(100, 20)
$searchLabel.Text = "Search String:"
$form.Controls.Add($searchLabel)

# Create the search text box
$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(120, 50)
$searchTextBox.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($searchTextBox)

# Create the search button
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(10, 80)
$searchButton.Size = New-Object System.Drawing.Size(75, 23)
$searchButton.Text = "Search"
$searchButton.Add_Click({
    # Search for the users in the specified domain
    $users = Get-ADUser -Filter {SamAccountName -like $searchTextBox.Text} -Server $domainTextBox.Text -Properties SamAccountName, UserPrincipalName, EmailAddress, Manager

    # Check if any results were found
    if ($users.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No matches found!")
        return
    }

    # Display the results in a GUI format
    $results = New-Object System.Collections.ArrayList
    foreach ($user in $users) {
        $result = New-Object PSObject
        $result | Add-Member -MemberType NoteProperty -Name "SamAccountName" -Value $user.SamAccountName
        $result | Add-Member -MemberType NoteProperty -Name "UniversalPrincipalName" -Value $user.UserPrincipalName
        $result | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $user.EmailAddress
        $result | Add-Member -MemberType NoteProperty -Name "Manager" -Value $user.Manager
        $results.Add($result) | Out-Null
    }

    # Display the results in a grid view
    $results | Out-GridView -Title "Search Results"

    # Export the results to a CSV file
    $results | Export-Csv -Path "C:\TEMP\SearchResults.csv" -NoTypeInformation

    # Display a message box to inform the user that the search is complete
    #[System.Windows.Forms.MessageBox]::Show("Search complete!")
})
$form.Controls.Add($searchButton)

# Create the export button REMOVED as it produced a blank export
#$exportButton = New-Object System.Windows.Forms.Button
#$exportButton.Location = New-Object System.Drawing.Point(90, 80)
#$exportButton.Size = New-Object System.Drawing.Size(75, 23)
#$exportButton.Text = "Export"
#$exportButton.Add_Click({
    # Export the results to a CSV file
#    $results | Export-Csv -Path "C:\TEMP\SearchResults.csv" -NoTypeInformation

    # Display a message box to inform the user that the export is complete
   # [System.Windows.Forms.MessageBox]::Show("Export complete!")
#})
#$form.Controls.Add($exportButton)

# Create the close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(246, 80)
$closeButton.Size = New-Object System.Drawing.Size(75, 23)
$closeButton.Text = "Close"
$closeButton.Add_Click({
    # Close the form
    $form.Close()
})
$form.Controls.Add($closeButton)

# Show the form
$form.ShowDialog() | Out-Null
```
