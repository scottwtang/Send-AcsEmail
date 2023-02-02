$configFile = ".\config.json"

function Get-AccessToken {
    <#
        .SYNOPSIS
            Obtain an access token for a service principal using the OAuth 2.0 client credentials flow with a client secret.
        
        .NOTES
            https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    #>
    
    $config        = Get-Content $configFile | ConvertFrom-Json
    $tenant_id     = $config.service_principal.tenant_id
    $client_id     = $config.service_principal.client_id
    $client_secret = $config.service_principal.client_secret

    $params = @{
        Uri    = "https://login.microsoftonline.com/$($tenant_id)/oauth2/v2.0/token"
        Method = "POST"
        Body   = @{
	        client_id     = $client_id
	        client_secret = $client_secret
            grant_type    = "client_credentials"
            scope         = "https://communication.azure.com/.default"
        }
    }

    $token = Invoke-RestMethod @params
    return $token.access_token
}

function Send-AcsEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Subject,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $To,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $Cc,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string] $Bcc,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Body,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        $Attachments,
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Html", "PlainText")]
        [string]$ContentType = "Html",
        
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("High", "Normal", "Low")]
        [string]$Importance = "Normal"
    )

    <#
        .SYNOPSIS
            Uses the Azure Communication Services Email REST API to queue an email message to be sent to one or more recipients.
        
        .NOTES
            https://learn.microsoft.com/en-us/rest/api/communication/email/send

        .PARAMETER Subject
            Subject of the email message.

        .PARAMETER To
            To recipients.

        .PARAMETER Cc
            CC recipients.

        .PARAMETER Bcc
            BCC recipients.

        .PARAMETER Body
            Content of the email.

        .PARAMETER Attachments
            Attachment(s) to the email, defined as an array of attachment objects, each with the key names "Name", "AttachmentType", "ContentBytesBase64".
            Example:
                @(
                    @{
                        Name               = "MyFile.txt"
                        AttachmentType     = "txt"
                        ContentBytesBase64 = "TWFueSBoYW5kcyBtYWtlIGxpZ2h0IHdvcmsu"
                    }
                )                

        .PARAMETER ContentType
            Whether email body should be plain text or HTML.

        .PARAMETER Importance
            The importance type for the email.

        .EXAMPLE
            # Sends a plain text email to 3 recipients.
            Send-AcsEmail -Subject "Email Subject" -To "John.Smith@corp.com;Jane.Smith@corp.com" -CC "Alice.Smith@corp.com" -Body "This is the email body." -ContentType PlainText
    #>
    
    $config     = Get-Content $configFile | ConvertFrom-Json
    $endpoint   = $config.azure_communication_service.endpoint
    $apiVersion = $config.azure_communication_service.api_version
    $sender     = $config.azure_communication_service.sender

    # Create a hashtable property for each recipient type
    $recipients = New-Object -TypeName PsObject -Property @{}
    foreach ($type in $To, $Cc, $Bcc) {
        if (-not [string]::IsNullOrEmpty($type)) {

            # Set the hashtable member property name
            switch ($type) {
                $To  {$memberName = "To"}
                $Cc  {$memberName = "Cc"}
                $Bcc {$memberName = "Bcc"}
            }

            # Split the recipients list
            $recipientArray = $type.Split(";")

            # Create the member property to be added
            $property = @(    
                foreach ($recipient in $recipientArray) {
                    @{
                        Email       = $recipient
                        DisplayName = $recipient
                    }
                }
            )

            $recipients | Add-Member -Name $memberName -Type NoteProperty -Value $property
        }
    }

    $params = @{
        Method  = "POST"
        URI     = "$($endpoint.TrimEnd("/"))/emails:send?api-version=$apiVersion"
        Headers = @{
            Authorization              = "Bearer $(Get-AccessToken)"
            "Content-Type"             = "application/json"
            "repeatability-request-id" = [guid]::NewGuid().ToString()
            "repeatability-first-sent" = [DateTime]::UtcNow.ToString("r")
        }
        Body    = @{
            Sender      = $sender
            Recipients  = $recipients
            Importance  = $Importance
            Content     = @{
                Subject      = $Subject
                $ContentType = $Body
            }
            Attachments = $Attachments
        } | ConvertTo-Json -Depth 10
    }

    try {
        $response = Invoke-WebRequest @params -UseBasicParsing
        $messageId = $response.Headers.'x-ms-request-id'

        # Poll the status of the message sent
        $index = 0
        do {
            $status = Get-AcsEmailStatus -MessageId $messageId

            Write-Host "Status code $($status.StatusCode) - $($status.StatusDescription)"

            # Wait between requests to prevent throttling
            if ($status.Status -eq "Queued") {
                Start-Sleep -Seconds 5
            }

            $index++

            # Assume the email has been dropped after 30 seconds
            if (($status.Status -ne "OutForDelivery") -and ($index -ge 6)) {
                Write-Host "Error checking status for message Id $messageId"
            }
        }
        until (($status.Status -ne "Queued") -or ($index -ge 6))

    }

    catch {
        Write-Host $error[0]
    }
}

function Get-AcsEmailStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $MessageId
    )

    <#
        .SYNOPSIS
            Uses the Azure Communication Services Email REST API to get the status of a message sent previously.
        
        .NOTES
            https://learn.microsoft.com/en-us/rest/api/communication/email/get-send-status

        .PARAMETER MessageId
            System generated message id (GUID) returned from a previous call to send email, in the response header x-ms-request-id.
    #>
    
    $config     = Get-Content $configFile | ConvertFrom-Json
    $endpoint   = $config.azure_communication_service.endpoint
    $apiVersion = $config.azure_communication_service.api_version

    $params = @{
        Method  = "GET"
        URI     = "$($endpoint.TrimEnd("/"))/emails/$MessageId/status?api-version=$apiVersion"
        Headers = @{
            Authorization = "Bearer $(Get-AccessToken)"
        }
    }

    try {
        $response = Invoke-WebRequest @params -UseBasicParsing

        $statusDescription = switch (($response | ConvertFrom-Json).Status) {
            "Dropped"        {"The message could not be processed and was dropped."}
            "Queued"         {"The message has passed basic validations and has been queued to be processed further."}
            "OutForDelivery" {"The message has been processed and is now out for delivery."}
        }

        $status = [PSCustomObject] @{
            MessageId         = ($response | ConvertFrom-Json).MessageId
            StatusCode        = $response.StatusCode
            Status            = ($response | ConvertFrom-Json).Status
            StatusDescription = $statusDescription
        }

        return $status
    }

    catch {
        Write-Host $error[0]
    }
}
