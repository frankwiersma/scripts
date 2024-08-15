# Function to create or update calendar items based on sent emails
function Add-EmailToCalendar {
    param(
        [DateTime]$TargetDate
    )

    # Create Outlook COM Object
    $Outlook = New-Object -ComObject Outlook.Application

    # Get the Sent Items folder
    $Namespace = $Outlook.GetNamespace("MAPI")
    $SentItems = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail)
    $Today = Get-Date -Format "yyyy-MM-dd"
    # Create an empty array to store email information
    $EmailList = @()

    # Iterate through each email in the Sent Items folder
    foreach ($MailItem in $SentItems.Items) {
        if ($MailItem -is [Microsoft.Office.Interop.Outlook.MailItem]) {
            # Get the date of the email
            $EmailDate = $MailItem.SentOn.Date
            
            # Check if the email was sent on the target date
            if ($EmailDate -eq $TargetDate.Date) {
                # Get the time, subject, recipient, etc.
                $Time = $MailItem.SentOn.ToString("HH:mm:ss")
                $SentOnDateTime = $MailItem.SentOn
                $Subject = $MailItem.Subject
                $Recipients = $MailItem.Recipients

                # Collect all recipients' email addresses or names in a single string
                $RecipientList = @()
                foreach ($Recipient in $Recipients) {
                    # Get the AddressEntry object and extract the email or display name
                    $AddressEntry = $Recipient.AddressEntry
                    $RecipientName = if ($AddressEntry.Type -eq "EX") { $AddressEntry.GetExchangeUser().PrimarySmtpAddress } else { $AddressEntry.Address }
                    $RecipientList += $RecipientName
                }
                $RecipientString = $RecipientList -join ", "

                # Add the details to the email list
                $EmailList += [PSCustomObject]@{
                    Date       = $EmailDate
                    Time       = $Time
                    Subject    = $Subject
                    Recipients = $RecipientString
                }

                # Check if a calendar item already exists for this email
                $CalendarItems = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar).Items
                $CalendarItems.Sort("[Start]")
                $CalendarItems.IncludeRecurrences = $false

                $ExistingAppointment = $CalendarItems | Where-Object {
                    $_.Start -eq $SentOnDateTime -and $_.Subject -eq "Email - $Subject - $RecipientString"
                }

                # Create a new appointment if it doesn't already exist
                if (-not $ExistingAppointment) {
                    $Appointment = $Outlook.CreateItem([Microsoft.Office.Interop.Outlook.OlItemType]::olAppointmentItem)
                    $Appointment.Subject = "Email - $Subject - $RecipientString"
                    $Appointment.Start = $SentOnDateTime
                    $Appointment.Duration = 15  # Duration in minutes, you can adjust if needed
                    $Appointment.BusyStatus = [Microsoft.Office.Interop.Outlook.OlBusyStatus]::olFree
                    $Appointment.Sensitivity = [Microsoft.Office.Interop.Outlook.OlSensitivity]::olPrivate
                    $Appointment.Save()
                }
            }
        }
    }

    # Optionally, display the result
    $EmailList | Format-Table -AutoSize
}

# Prompt for the date (defaults to today)
$UserInput = Read-Host "Enter the date (yyyy-MM-dd) or 'today', 'yesterday'"
switch ($UserInput.ToLower()) {
    "today" { $TargetDate = Get-Date }
    "yesterday" { $TargetDate = (Get-Date).AddDays(-1) }
    default { $TargetDate = [DateTime]::ParseExact($UserInput, 'yyyy-MM-dd', $null) }
}

# Run the function for the specified date
Add-EmailToCalendar -TargetDate $TargetDate
