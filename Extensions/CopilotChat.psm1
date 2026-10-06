<#
.SYNOPSIS
Module for Copilot Chat

.DESCRIPTION
This module adds the Copilot Chat view. The view can be used to chat with Microsoft 365 Copilot
about Intune policies. Policies are loaded from the connected tenant or from exported json files
and added as context (json data) to the chat, so questions about the policies can be asked and
the policies can be analyzed.

The chat uses the Microsoft 365 Copilot Chat API (Graph beta endpoint /copilot/conversations).
The API requires:

* A Microsoft 365 Copilot license (add-on) for the signed in user
* A delegated user login. Application (app-only) logins are not supported.
* All of the following Graph permissions: Sites.Read.All, Mail.Read, People.Read.All,
  OnlineMeetingTranscript.Read.All, Chat.Read, ChannelMessage.Read.All, ExternalItem.Read.All.
  Note that some of these permissions require an administrator to consent.

The API is in preview and only supports text responses. The data is processed by Microsoft 365
Copilot and stays within the Microsoft 365 trust boundary.

Microsoft links:
* Chat API overview: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview
* Create conversations: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations
* Chat messages: https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotconversation-chat

.NOTES
  Author:         Simon
#>

$global:CopilotViewObject = $null

# Maximum number of policies that can be added as context to a single message
$script:copilotMaxContextPolicies = 10

# The Microsoft 365 Copilot Chat API requires ALL of the below permissions
# Note: People.Read.All, OnlineMeetingTranscript.Read.All, ChannelMessage.Read.All and
# ExternalItem.Read.All require an administrator to consent (one time per tenant)
$script:copilotScopes = @(
    "Sites.Read.All",
    "Mail.Read",
    "People.Read.All",
    "OnlineMeetingTranscript.Read.All",
    "Chat.Read",
    "ChannelMessage.Read.All",
    "ExternalItem.Read.All"
)

function Get-ModuleVersion
{
    '1.0.0'
}

function Invoke-InitializeModule
{
    #Add settings
    $global:appSettingSections += (New-Object PSObject -Property @{
        Title = "Copilot Chat"
        Id = "CopilotChat"
        Values = @()
        Priority = 15
    })

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Disable web search grounding"
        Key = "CopilotDisableWebGrounding"
        Type = "Boolean"
        DefaultValue = $true
        Description = "Disables web search grounding for each chat message. Only enterprise data (Intune) will be used as context."
        SubPath = "CopilotChat"
    }) "CopilotChat"

    Add-SettingsObject (New-Object PSObject -Property @{
        Title = "Max context size per policy (KB)"
        Key = "CopilotMaxContextKB"
        Type = "Int"
        DefaultValue = 64
        Description = "Maximum json size per policy that is added as context to a chat message. Larger json data is truncated."
        SubPath = "CopilotChat"
    }) "CopilotChat"

    Add-CopilotView
}

function Add-CopilotView
{
    $viewPanel = Get-XamlObject ($global:AppRootFolder + "\Xaml\CopilotChatPanel.xaml") -AddVariables

    if(-not $viewPanel) { return }

    Set-CopilotViewPanel $viewPanel

    #Add menu group and items
    $global:CopilotViewObject = (New-Object PSObject -Property @{
        Title = "Copilot Chat"
        Description = "Chat with Microsoft 365 Copilot about Intune policies. The chat uses a separate Copilot account that needs a Microsoft 365 Copilot license."
        ID = "CopilotChat"
        ViewPanel = $viewPanel
        AuthenticationID = "MSAL"
        Activating = { Invoke-CopilotActivatingView }
        Authentication = (Get-MSALAuthenticationObject)
        Authenticate = { Invoke-CopilotAuthenticateToMSAL @args }
        AppInfo = (Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp)

        # The Copilot Chat scopes are NOT requested here. They are only requested
        # with the separate Copilot account (Connect-CopilotAccount) since the
        # tenant login does not need them.
        Permissions = @()
        HideMenu = $true
    })

    Add-ViewObject $global:CopilotViewObject
}

function Invoke-CopilotActivatingView
{
    Update-CopilotAccountInfo
}

function Invoke-CopilotAuthenticateToMSAL
{
    $global:CopilotViewObject.AppInfo = Get-GraphAppInfo "EMAzureApp" $global:DefaultAzureApp
    Set-MSALCurrentApp $global:CopilotViewObject.AppInfo
    $usr = (?? $global:MSALToken.Account.UserName (Get-Setting "" "LastLoggedOnUser"))
    if($usr)
    {
        & $global:msalAuthenticator.Login -Account $usr
    }
}

function Connect-CopilotAccount
{
    # Interactive sign in with the Copilot account. The Copilot account needs a
    # Microsoft 365 Copilot license and can be a different account than the
    # tenant login used for fetching the policies.
    Write-Log "Connect Copilot account"

    $appInfo = $global:MSGraphGlobalApps | Where ClientId -eq $global:DefaultAzureApp
    $script:copilotApp = Get-MSALApp $appInfo $null

    $acquireTokenObj = $script:copilotApp.AcquireTokenInteractive([string[]]$script:copilotScopes)
    [void]$acquireTokenObj.WithPrompt([Microsoft.Identity.Client.Prompt]::SelectAccount)

    [IntPtr]$ParentWindow = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle
    if ($ParentWindow)
    {
        [void]$acquireTokenObj.WithParentActivityOrWindow($ParentWindow)
    }

    $authResult = Get-MsalAuthenticationToken $acquireTokenObj
    if($authResult)
    {
        $script:copilotToken = $authResult
        $script:copilotAccount = $authResult.Account
        Write-Log "Copilot account connected: $($script:copilotAccount.UserName)"
    }

    Update-CopilotAccountInfo
    $script:copilotToken
}

function Get-CopilotToken
{
    # Returns the cached Copilot token or silently refreshes an expired token
    if(-not $script:copilotToken) { return $null }

    if($script:copilotToken.ExpiresOn.LocalDateTime.Ticks -gt ((Get-Date).AddMinutes(-5)).Ticks)
    {
        return $script:copilotToken
    }

    $authResult = $null
    try
    {
        $acquireTokenObj = $script:copilotApp.AcquireTokenSilent([string[]]$script:copilotScopes, $script:copilotAccount)
        $authResult = Get-MsalAuthenticationToken $acquireTokenObj
    }
    catch
    {
        Write-LogDebug "Silent token refresh failed: $($_.Exception.Message)"
    }

    if($authResult)
    {
        $script:copilotToken = $authResult
    }
    else
    {
        # UI required e.g. consent removed or password changed. Disconnect so the user reconnects.
        Write-Log "The Copilot token needs to be refreshed. Please connect the Copilot account again" 2
        Disconnect-CopilotAccount
    }

    $script:copilotToken
}

function Disconnect-CopilotAccount
{
    $script:copilotToken = $null
    $script:copilotAccount = $null
    $script:copilotConversationId = $null

    Update-CopilotAccountInfo
}

function Update-CopilotAccountInfo
{
    if(-not $global:lblCopilotAccount) { return }

    if($script:copilotAccount)
    {
        $global:lblCopilotAccount.Content = $script:copilotAccount.UserName
    }
    else
    {
        $global:lblCopilotAccount.Content = "Not connected"
    }
}

function Set-CopilotViewPanel
{
    param($viewPanel)

    # Create the chat document
    $script:copilotChatDoc = [System.Windows.Documents.FlowDocument]::new()
    $script:copilotChatDoc.PagePadding = [System.Windows.Thickness]::new(10)
    $global:fdsCopilotChat.Document = $script:copilotChatDoc

    # Populate the policy type combo with the object types from the Intune Manager view
    $global:cbCopilotType.ItemsSource = @(Get-CopilotObjectTypes)
    $global:cbCopilotType.DisplayMemberPath = "Title"

    # DataGrid columns
    $column = Get-GridCheckboxColumn "Selected"
    $global:dgCopilotPolicies.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Name")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Name"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgCopilotPolicies.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Type")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Type"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgCopilotPolicies.Columns.Add($column)

    $binding = [System.Windows.Data.Binding]::new("Source")
    $column = [System.Windows.Controls.DataGridTextColumn]::new()
    $column.Header = "Source"
    $column.IsReadOnly = $true
    $column.Binding = $binding
    $global:dgCopilotPolicies.Columns.Add($column)

    # Update the context info when a checkbox is checked/unchecked (routed events)
    $handler = [System.Windows.RoutedEventHandler] {
        Update-CopilotContextInfo
    }
    $global:dgCopilotPolicies.AddHandler([System.Windows.Controls.CheckBox]::CheckedEvent, $handler)
    $global:dgCopilotPolicies.AddHandler([System.Windows.Controls.CheckBox]::UncheckedEvent, $handler)

    # Select/deselect all items in the context list
    $global:dgCopilotPolicies.Columns[0].Header.add_Click({
        foreach($item in $global:dgCopilotPolicies.ItemsSource)
        {
            $item.Selected = $this.IsChecked
        }
        $global:dgCopilotPolicies.Items.Refresh()
        Update-CopilotContextInfo
    })

    $global:btnCopilotLoad.Add_Click({ Start-CopilotLoadPolicies })
    $global:btnCopilotLoadFiles.Add_Click({ Start-CopilotLoadExportedFiles })
    $global:btnCopilotClearContext.Add_Click({ Clear-CopilotContext })
    $global:btnCopilotNew.Add_Click({ Start-CopilotNewConversation })

    $global:btnCopilotConnect.Add_Click({ Connect-CopilotAccount | Out-Null })
    $global:btnCopilotDisconnect.Add_Click({ Disconnect-CopilotAccount })

    $global:btnCopilotSend.Add_Click({ Invoke-CopilotSend })

    $global:txtCopilotMessage.Add_KeyDown({
        if($_.Key -eq [System.Windows.Input.Key]::Enter)
        {
            Invoke-CopilotSend
        }
    })

    $script:copilotPolicies = @()
    $script:copilotConversationId = $null
    $script:copilotToken = $null
    $script:copilotAccount = $null

    Start-CopilotNewConversation
    Update-CopilotContextInfo
}

function Get-CopilotObjectTypes
{
    $viewObject = $global:viewObjects | Where { $_.ViewInfo.Id -eq "IntuneGraphAPI" }
    if(-not $viewObject) { return @() }

    $types = @()
    foreach($item in $viewObject.ViewItems)
    {
        if($item.API -and $item.Title)
        {
            $types += $item
        }
    }

    $types | Sort-Object -Property Title
}

function Start-CopilotLoadPolicies
{
    $objType = $global:cbCopilotType.SelectedItem
    if(-not $objType)
    {
        [System.Windows.MessageBox]::Show("Select a policy type first", "Copilot Chat", "OK", "Information") | Out-Null
        return
    }

    if(-not $global:MSALToken)
    {
        [System.Windows.MessageBox]::Show("Not logged in. Please login first", "Copilot Chat", "OK", "Information") | Out-Null
        return
    }

    Write-Status "Loading $($objType.Title)..."
    $objects = Get-GraphObjects -objectType $objType

    # Replace previously loaded tenant policies of the same type
    $script:copilotPolicies = @($script:copilotPolicies | Where { $_.Source -ne "Tenant" -or $_.Type -ne $objType.Title })

    foreach($obj in $objects)
    {
        $script:copilotPolicies += New-Object PSObject -Property @{
            Selected = $false
            Name = (?? $obj.displayName $obj.Id)
            Type = $objType.Title
            Source = "Tenant"
            JSON = $null
            Object = $obj
            ObjectType = $objType
        }
    }

    Write-Status ""
    Update-CopilotPolicies
}

function Start-CopilotLoadExportedFiles
{
    $folder = Get-Folder -path (Get-Setting "CopilotChat" "LastExportFolder" $env:temp) -title "Select a folder with exported json files"
    if(-not $folder) { return }

    Save-Setting "CopilotChat" "LastExportFolder" $folder

    Write-Status "Loading exported files..."
    $files = Get-ChildItem -Path $folder -Filter *.json -Recurse -ErrorAction SilentlyContinue

    $count = 0
    foreach($file in $files)
    {
        if($script:copilotPolicies | Where { $_.FilePath -eq $file.FullName }) { continue }

        try
        {
            $json = Get-Content $file.FullName -Raw -ErrorAction Stop
        }
        catch
        {
            Write-LogError "Failed to open file $($file.FullName)" $_.Exception
            continue
        }

        try
        {
            $obj = ConvertFrom-Json $json -ErrorAction Stop
        }
        catch
        {
            Write-LogError "Failed to parse json in file $($file.FullName)" $_.Exception
            continue
        }

        # One file can contain one or multiple objects (bulk export)
        $objs = @()
        if($obj -is [System.Array]) { $objs = $obj }
        else { $objs = @($obj) }

        foreach($jsonObj in $objs)
        {
            $name = $jsonObj.displayName
            if(-not $name) { $name = [IO.Path]::GetFileNameWithoutExtension($file.Name) }

            $script:copilotPolicies += New-Object PSObject -Property @{
                Selected = $false
                Name = $name
                Type = $file.Directory.Name
                Source = "Export"
                JSON = $json
                FilePath = $file.FullName
                Object = $jsonObj
                ObjectType = $null
            }
            $count++
        }
    }

    Write-Status ""
    if($count -gt 0)
    {
        Update-CopilotPolicies
    }
    else
    {
        Write-Log "No json files found in $folder" 2
    }
}

function Get-CopilotPolicyJson
{
    param($policy)

    # Exported files already have the json data
    if($policy.JSON) { return $policy.JSON }

    if(-not $policy.Object) { return $null }

    Write-Status "Loading json data for $($policy.Name)..."
    $objInfo = Get-GraphObject $policy.Object $policy.ObjectType

    if(-not $objInfo) { return $null }

    # Remove odata properties that are not part of the policy
    $objInfo = $objInfo | Select-Object -Property * -ExcludeProperty "*@odata*"

    ConvertTo-Json -InputObject $objInfo -Depth 20 -Compress
}

function Get-CopilotContextText
{
    $maxKB = Get-SettingValue "CopilotMaxContextKB" 64
    $maxSize = 64 * 1024
    try { $maxSize = [int]$maxKB * 1024 } catch { }

    $selected = @($script:copilotPolicies | Where Selected -eq $true)
    $context = @()

    foreach($policy in $selected)
    {
        if($context.Count -ge $script:copilotMaxContextPolicies)
        {
            Write-Log "Maximum number of policies in context ($($script:copilotMaxContextPolicies)) reached. The remaining policies are not included" 2
            break
        }

        $json = Get-CopilotPolicyJson $policy
        if(-not $json) { continue }

        if($json.Length -gt $maxSize)
        {
            $json = $json.Substring(0, $maxSize) + " ...[TRUNCATED]"
        }

        $context += @{ text = "Intune policy '$($policy.Name)' (Type: $($policy.Type)):" + [System.Environment]::NewLine + $json }
    }

    $context
}

function Invoke-CopilotApiRequest
{
    # Sends a request to the Copilot Chat API with the Copilot account token.
    # A separate function is used since Invoke-GraphRequest is bound to the
    # tenant login token which can be a different account than the Copilot account.
    param($Url, $Content)

    $token = Get-CopilotToken
    if(-not $token)
    {
        throw "The Copilot account is not connected. Connect the Copilot account first (an account with a Microsoft 365 Copilot license is required)."
    }

    if($Url -notmatch "^https?://")
    {
        $Url = "https://graph.microsoft.com/beta/" + $Url.TrimStart('/')
    }

    $headers = @{
        'Authorization' = "Bearer $($token.AccessToken)"
        'Content-Type' = 'application/json; charset=utf-8'
    }

    $invokeParams = @{
        Uri = $Url
        Method = "POST"
        Headers = $headers
        ContentType = "application/json; charset=utf-8"
        UseBasicParsing = $true
    }
    if($Content) { $invokeParams.Add("Body", [System.Text.Encoding]::UTF8.GetBytes($Content)) }

    $proxyURI = Get-ProxyURI
    if($proxyURI) { $invokeParams.Add("Proxy", $proxyURI) }

    $ret = $null
    $retryCount = 0
    do
    {
        $retryRequest = $false
        try
        {
            Write-LogDebug "Invoke Copilot Chat API: $Url"
            $ret = Invoke-RestMethod @invokeParams
        }
        catch
        {
            $retryCount++

            # Retry on throttling
            if($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 429 -and $retryCount -le 3)
            {
                $retryRequest = $true
                Write-Log "429 - Too many requests received. Wait 5 s before retry" 2
                Start-Sleep -Seconds 5
            }
            else
            {
                # Extract the graph error message from the response body
                $extMessage = $_.Exception.Message
                try
                {
                    $errorBody = $null
                    if($_.ErrorDetails -and $_.ErrorDetails.Message)
                    {
                        $errorBody = $_.ErrorDetails.Message | ConvertFrom-Json
                    }
                    elseif($_.Exception.Response -and $_.Exception.Response.GetResponseStream())
                    {
                        $stream = $_.Exception.Response.GetResponseStream()
                        $stream.Position = 0
                        $reader = New-Object System.IO.StreamReader($stream)
                        $errorBody = $reader.ReadToEnd() | ConvertFrom-Json
                    }

                    if($errorBody -and $errorBody.error.message) { $extMessage = $errorBody.error.message }
                }
                catch { }

                if($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 403)
                {
                    $extMessage = "Access denied by the Copilot Chat API. Verify that the Copilot account has a Microsoft 365 Copilot license and that the required permissions are consented. Original error: $extMessage"
                }

                Write-LogError "Failed to invoke the Copilot Chat API with URL $Url. Error: $extMessage" $_.Exception
                throw $extMessage
            }
        }
    } while($retryRequest)

    $ret
}

function Start-CopilotConversation
{
    Write-Log "Create Copilot conversation"
    $conversation = Invoke-CopilotApiRequest -Url "copilot/conversations" -Content "{}"

    if($conversation -and $conversation.Id)
    {
        Write-Log "Copilot conversation created: $($conversation.Id)"
    }

    $conversation
}

function Get-CopilotChatBody
{
    param($Message, $AdditionalContext)

    $body = @{
        message = @{ text = $Message }
        locationHint = @{ timeZone = [TimeZoneInfo]::Local.Id }
    }

    if($AdditionalContext -and @($AdditionalContext).Count -gt 0)
    {
        $body.Add("additionalContext", @($AdditionalContext))
    }

    if((Get-SettingValue "CopilotDisableWebGrounding" $true) -eq $true)
    {
        $body.Add("contextualResources", @{ webContext = @{ isWebEnabled = $false } })
    }

    ConvertTo-Json -InputObject $body -Depth 10
}

function Invoke-CopilotChatRequest
{
    param($Message, $AdditionalContext)

    $content = Get-CopilotChatBody -Message $Message -AdditionalContext $AdditionalContext

    Write-LogDebug "Copilot chat message: $content"

    Invoke-CopilotApiRequest -Url "copilot/conversations/$($script:copilotConversationId)/chat" -Content $content
}

function Get-CopilotAnswerText
{
    param($Response, $Message)

    if(-not $Response -or -not $Response.messages) { return $null }

    # The answer is the last message that is not an echo of the sent message
    $answer = $null
    foreach($message in $Response.messages)
    {
        if($message.text -and $message.text -ne $Message)
        {
            $answer = $message.text
        }
    }

    Remove-CopilotTextArtifacts $answer
}

function Remove-CopilotTextArtifacts
{
    param($text)

    if(-not $text) { return $text }

    # Remove entity reference tags (keep the inner text) e.g. <Person>John Doe</Person>
    $text = $text -replace "</?(Person|Event|File|Site|DriveItem|Email|Message|Chat|Team|Channel|Meeting|MailFolder|Document|FileAttachment)>", ""

    # Remove attribution markers e.g. [^1^]
    $text = $text -replace "\[\^\d+\^\]", ""

    $text.Trim()
}

function Invoke-CopilotSend
{
    $message = ""
    if($global:txtCopilotMessage) { $message = $global:txtCopilotMessage.Text.Trim() }
    if(-not $message) { return }

    $global:btnCopilotSend.IsEnabled = $false
    Write-Status "Asking Copilot..."
    [System.Windows.Forms.Application]::DoEvents()

    Add-CopilotChatMessage "You" $message
    $global:txtCopilotMessage.Text = ""

    try
    {
        # The Copilot Chat API requires a delegated login with a Microsoft 365 Copilot
        # license. This is a separate account than the tenant login used for the policies.
        if(-not $script:copilotToken)
        {
            Add-CopilotChatMessage "Info" "Connecting the Copilot account. Sign in with an account that has a Microsoft 365 Copilot license..."
            [System.Windows.Forms.Application]::DoEvents()

            $connected = Connect-CopilotAccount
            if(-not $connected)
            {
                throw "Copilot account not connected. The Copilot Chat API requires an account with a Microsoft 365 Copilot license."
            }
        }

        if(-not $script:copilotConversationId)
        {
            $conversation = Start-CopilotConversation
            if(-not $conversation -or -not $conversation.Id)
            {
                throw "Failed to create the Copilot conversation. Verify that the Copilot account has a Microsoft 365 Copilot license and that the required permissions are consented."
            }
            $script:copilotConversationId = $conversation.Id
        }

        $context = Get-CopilotContextText
        $response = Invoke-CopilotChatRequest -Message $message -AdditionalContext $context

        $answer = Get-CopilotAnswerText -Response $response -Message $message
        if(-not $answer)
        {
            throw "Copilot did not return an answer"
        }

        Add-CopilotChatMessage "Copilot" $answer
    }
    catch
    {
        Write-LogError "Copilot chat failed" $_.Exception
        Add-CopilotChatMessage "Error" $_.Exception.Message
    }
    finally
    {
        $global:btnCopilotSend.IsEnabled = $true
        Write-Status ""
    }
}

function Add-CopilotChatMessage
{
    param($Sender, $Text)

    if(-not $Text) { return }
    if(-not $script:copilotChatDoc) { return }

    $colors = @{
        "You" = "#1a6ab5"
        "Copilot" = "#742774"
        "Error" = "#c4314b"
        "Info" = "#606060"
    }
    $color = ?? $colors[$Sender] "#404040"

    $paragraph = [System.Windows.Documents.Paragraph]::new()
    $paragraph.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)

    $header = [System.Windows.Documents.Run]::new("$Sender ($((Get-Date).ToString("HH:mm"))):")
    $header.FontWeight = [System.Windows.FontWeights]::Bold
    $header.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($color))
    $paragraph.Inlines.Add($header)
    $paragraph.Inlines.Add([System.Windows.Documents.LineBreak]::new())

    foreach($line in ($Text -split "`r?`n"))
    {
        if($paragraph.Inlines.Count -gt 2) { $paragraph.Inlines.Add([System.Windows.Documents.LineBreak]::new()) }
        $paragraph.Inlines.Add([System.Windows.Documents.Run]::new($line))
    }

    $script:copilotChatDoc.Blocks.Add($paragraph)
    if($global:fdsCopilotChat -and $global:fdsCopilotChat.ScrollViewer) { $global:fdsCopilotChat.ScrollViewer.ScrollToEnd() }
}

function Clear-CopilotContext
{
    $script:copilotPolicies = @()
    Update-CopilotPolicies
}

function Start-CopilotNewConversation
{
    $script:copilotConversationId = $null

    if($script:copilotChatDoc)
    {
        $script:copilotChatDoc.Blocks.Clear()
    }

    Add-CopilotChatMessage "Info" "Welcome to Copilot Chat! Load policies from the tenant or from exported files, select the policies to include in the context and ask your questions."
}

function Update-CopilotPolicies
{
    if(-not $global:dgCopilotPolicies) { return }

    $global:dgCopilotPolicies.ItemsSource = @($script:copilotPolicies)
    Update-CopilotContextInfo
}

function Update-CopilotContextInfo
{
    if(-not $global:txtCopilotContextInfo) { return }

    $total = @($script:copilotPolicies).Count
    $selected = @($script:copilotPolicies | Where Selected -eq $true).Count

    $global:txtCopilotContextInfo.Text = "$selected of $total policies in context"
}
