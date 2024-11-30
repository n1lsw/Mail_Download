Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Program {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

# Define the base folder to save attachments
$baseFolder = ""

$allowedSenders = @(
    ""
)

# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Get the Inbox folder
$inbox = $namespace.GetDefaultFolder(6) # 6 refers to the Inbox folder

# Infinite loop to run the script every 5 minutes
while ($true) {
    # Get the current date and format it as "dd-MM-yyyy"
    $currentDate = Get-Date -Format "dd-MM-yyyy"

    # Define the download folder for the current day
    $downloadFolder = Join-Path -Path $baseFolder -ChildPath $currentDate

    # Define the download folder for the previous day
    $previousDownloadFolderRaw = Get-ChildItem $baseFolder | 
    Where-Object { $_.PSIsContainer } | 
    Sort-Object CreationTime -Descending | 
    Select-Object -Skip 1 -First 1

    # Get the full path
    $previousDownloadFolder = $previousDownloadFolderRaw.FullName

    # Create the download folder if it doesn't exist
    if (-Not (Test-Path -Path $downloadFolder)) {
        New-Item -ItemType Directory -Path $downloadFolder
        "New folder created"
    }

    # Get the items in the Inbox
    $items = $inbox.Items

    # Loop through the items in the Inbox
    foreach ($item in $items) {
        # Check if the item is a mail item
        if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
            $isAllowedSender = $allowedSenders -contains $item.SenderEmailAddress
            $recipients = $item.Recipients

            foreach ($recipient in $recipients) {
                if ($($recipient.Name -eq "Wilharm Nils")) {
                    $recipient = $true
                }
            }
            
            if ($isAllowedSender -or $recipient) {
                # Write-Host $item.SenderEmailAddress
                # Check if the mail item has attachments
                if ($item.Attachments.Count -gt 0) {
                    # Check if Received date of mail is equal to current date
                    if ($item.ReceivedTime.Date.ToString("dd-MM-yyyy") -eq $((Get-Date).ToString("dd-MM-yyyy"))) {
                        Write-Host "Processing email received on: $($item.ReceivedTime))"
                        # Loop through the attachments
                        foreach ($attachment in $item.Attachments) {
                            $previousAttachmentFilePath = Join-Path -Path $previousDownloadFolder -ChildPath $attachment.FileName
                            $attachmentFilePath = Join-Path -Path $DownloadFolder -ChildPath $attachment.FileName
                            # Get the current date and time
                            $currentDateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    
                            # Check if the previous attachment exists
                            if (Test-Path -Path $previousAttachmentFilePath) {
                                Write-Host "$currentDateTime Previous attachment already exists: $($attachment.FileName). Skipping download."
                            }
                            elseif (Test-Path -Path $attachmentFilePath) {
                                Write-Host "$currentDateTime Attachment already exists: $($attachment.FileName) from $($item.SenderEmailAddress). Skipping download."
                            }
                            else {
                                try {
                                    # Save the attachment
                                    $attachment.SaveAsFile($attachmentFilePath)
                                    Write-Host "$currentDateTime Downloaded attachment: $($attachment.FileName) from $($item.SenderEmailAddress)."
                                }
                                catch {
                                    Write-Host "$currentDateTime Failed to download attachment: $($attachment.FileName) from $($item.SenderEmailAddress). Error: $_"
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    # Wait for 5 minutes (300 seconds) before checking again
    Start-Sleep -Seconds 300
}

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null