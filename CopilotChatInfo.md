# Copilot Chat

The **Copilot Chat** view can be used to chat with Microsoft 365 Copilot about Intune policies. Policies are added as json context to the chat so questions about the policies can be asked and the policies can be analyzed e.g. security risks, misconfigurations, best practices or dependencies.

The chat uses the [Microsoft 365 Copilot Chat API (preview)](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview). The API is based on Microsoft Graph and processes the data with Microsoft 365 Copilot, which means that the data stays within the Microsoft 365 trust boundary. No API token or key is needed, the tool uses delegated authentications with the assigned licenses.

Microsoft links:

* [Microsoft 365 Copilot Chat API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview)
* [Create conversations](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations)
* [Chat messages](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotconversation-chat)

## Two accounts

The view uses **two separate logins**:

| Login | Used for | Account |
| --- | --- | --- |
| **Tenant login** (standard login of the tool) | Fetching the policies from the connected tenant | Can be any account with read access to the Intune policies, for example a customer tenant account |
| **Copilot account** (*Connect* button in the view) | Calling the Copilot Chat API | Must be an account with a **Microsoft 365 Copilot license**, for example your own account |

This makes it possible to analyze the policies of a customer tenant with your own Copilot account. The policy json data is sent as text to the Microsoft 365 Copilot conversation of the Copilot account (see [Privacy](#privacy)).

The Copilot account can also be connected directly with the **Connect** button, and **Send** connects the account automatically if it is not connected yet.

## Requirements

* **Microsoft 365 Copilot license** - The Chat API is available at no extra cost to users with a Microsoft 365 Copilot add-on license. The license must be assigned to the **Copilot account**.
* **Delegated user login** - The API does not support application (app-only) logins. A user login is required.
* **Graph permissions** - All of the following permissions are required to call the API (requested with the Copilot account login):
  * Sites.Read.All
  * Mail.Read
  * People.Read.All
  * OnlineMeetingTranscript.Read.All
  * Chat.Read
  * ChannelMessage.Read.All
  * ExternalItem.Read.All

**Note:** The following permissions require an **administrator** to consent: `People.Read.All`, `OnlineMeetingTranscript.Read.All`, `ChannelMessage.Read.All` and `ExternalItem.Read.All`. If the Copilot account has admin rights in its own tenant, the consent can be done directly in the login prompt (*Consent on behalf of your organization*). Otherwise an admin of the tenant of the Copilot account has to consent once. The consent only has to be done once per tenant.

**Why all permissions?** The Chat API validates that the token contains **all** of the above permissions before it accepts a call, even though the tool itself only reads the Intune policies. This is because Microsoft 365 Copilot grounds its answers on all workloads of the Copilot account (mail, people, meeting transcripts, chats, channel messages, external items). The tool never uses these permissions to read data - they only enable the Copilot service to ground responses. Since they are requested with the Copilot account login, the access applies to the Copilot account's own tenant, not to the tenant the policies are fetched from.

## How to use

1. Open the **Copilot Chat** view in the *Views* menu and login with the tenant account to fetch the policies (for example the customer tenant).
2. Connect the **Copilot account** with the *Connect* button. Sign in with an account that has a Microsoft 365 Copilot license.
3. Load policies into the context list:
   * **Load** - Loads all policies of the selected policy type from the connected tenant. Previously loaded tenant policies of the same type are replaced.
   * **Load exported files...** - Loads all json files in a folder (including sub-folders). This uses the files created by the export feature. The files are added to the policies already in the list.
   * **Clear context** - Removes all policies from the context list.
4. Select the policies to include as context in the chat messages with the checkboxes in the context list. A maximum of 10 policies per message is supported.
5. Ask a question and press **Send** (or Enter). The json data of the selected policies is added as additional context to the message.
6. The conversation is multi-turn, Copilot remembers the previous messages in the conversation. Use **New conversation** to start over. Use **Disconnect** to sign out from the Copilot account.

Example questions:

* "Analyze this policy and point out security risks"
* "Are there any settings in this policy that conflicts with the CIS benchmark?"
* "Explain what this policy configures on the devices"
* "Which users or groups are affected by the assignments in this policy?"

## Settings

| Setting | Description |
| --- | --- |
| Disable web search grounding | Disables web search grounding for each chat message. Only enterprise data (Intune) will be used as context. Default is on. |
| Max context size per policy (KB) | Maximum json size per policy that is added as context to a chat message. Larger json data is truncated. Default is 64 KB. |

## Known issues and limitations

* The Chat API is in **preview** and based on the Graph beta endpoint. The API might change.
* The API only responds with **text responses** and does not support actions like creating files, sending emails or code interpreter.
* The API uses both enterprise search grounding and web search grounding by default. Disabling web search grounding is a single-turn action and is applied to each chat message when the setting is on.
* Chat messages that include long running tasks are prone to gateway timeouts.
* The tool sends the request synchronously, the UI is blocked until Copilot returns an answer.
* The responses are AI-generated and might be inaccurate, so they should be verified before use.
* The time zone in the location hint is based on the Windows time zone of the client.
* The Copilot Chat API requires a full Microsoft 365 Copilot add-on license. Pay-as-you-go cannot be used to unlock the API.
* **Copilot Studio agents are not supported** - Agents created with Copilot Studio require a Copilot Studio license or pay-as-you-go billing and cannot be called programmatically for free. The Chat API is the free programmatic way to chat with Microsoft 365 Copilot.

## Privacy

The policy json data is sent as text to the Microsoft 365 Copilot conversation of the **Copilot account** and is processed within the Microsoft 365 trust boundary. Copilot does not use the Microsoft 365 prompts, responses and Graph data to train the foundation LLMs. No data is sent to any third party AI service.

**Note for consultants:** When the tenant login is a customer tenant account, the customer policy data is processed by Microsoft 365 Copilot in the tenant of the Copilot account (for example your own company tenant). Verify that this is allowed by the customer's data processing agreements.

## Roadmap

Possible future extensions:

* **Anonymization pipeline** - Copilot anonymizes the policy data (removes user names, device names, group names etc.) before it is forwarded to another AI model.
* **Local models** - Support for models installed on the PC e.g. Claude Code (with a Claude subscription, no API key needed) or a Gemini CLI, invoked as local processes with the anonymized data. This avoids Azure costs of a hosted model endpoint.

**Note:** There is no official command line interface for Microsoft 365 Copilot. The Chat API (Microsoft Graph) is the only free programmatic interface to Microsoft 365 Copilot, so the Graph scope consent cannot be avoided when using Microsoft 365 Copilot. Microsoft 365 Copilot Chat in the browser or in the desktop app does not offer a scriptable interface.
