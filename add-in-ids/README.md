# Find Outlook add-ins that use legacy Exchange Online tokens

Legacy Exchange Online tokens are deprecated and will begin being turned off across Microsoft 365 tenants in February 2025. If you are a Microsoft 365 administrator, you should check if any Outlook add-ins deployed in your tenant are using legacy Exchange Online tokens. If they are, we recommend you reach out to publishers and developers of those add-ins to ensure they are migrating their add-ins away from legacy tokens. Add-ins that continue to use legacy tokens won't work correctly when the legacy tokens are turned off beginning February 2025.

For more information about deprecation of legacy tokens, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

The following procedures in this README show how to find Outlook add-ins deployed in your tenant that are potentially using legacy tokens.

## List of Outlook add-ins deployed from the Microsoft store

This folder contains the [add-ins-using-exchange-tokens.xlsx](/add-ins-using-exchange-tokens.xlsx) spreadsheet. The spreadsheet contains a list of all Outlook add-ins published in the Microsoft store that were using legacy tokens as of April 2025. You can use the spreadsheet to look up any add-in by its asset ID. All add-ins published to the store have a unique Asset ID.

**Note:** The list is also available as a CSV file in [add-ins-using-exchange-tokens.csv](add-ins-using-exchange-tokens.csv).

## Open PowerShell and sign in

To use the commands in this README you need to connect to Exchange Online PowerShell and Centralized Deployment PowerShell. This requires you to sign in twice with your **User Admin** credentials.

1. Run PowerShell as administrator.
1. Run the command `Install-Module -Name O365CentralizedAddInDeployment`. For more information about this module, see [O365CentralizedAddInDeployment 3.0.2](https://www.powershellgallery.com/packages/O365CentralizedAddInDeployment/3.0.2)
1. Run the command `Import-Module -Name O365CentralizedAddInDeployment`.  
1. Run the command `Connect-OrganizationAddInService`. Sign in with your **User Admin** credentials
1. Run the command `Import-Module ExchangeOnlineManagement`. For more information about this module, see [Connect to Exchange Online PowerShell](https://learn.microsoft.com/powershell/exchange/connect-to-exchange-online-powershell)
1. Run the command `Connect-ExchangeOnline`. Sign in with your **User Admin** credentials.

## Get a list of all add-ins deployed at the organization level

Admins can deploy add-ins for the entire organization from the Microsoft store or through central deployment. You’ll need to run two commands to get the complete list of add-ins deployed to the organization.

Run the following command to get a list of all add-ins deployed at the organization level in your tenant.

- `Get-OrganizationAddIn | Select-Object -Property DisplayName, AssetID, ProductID, ProviderName`

The command lists all add-ins you deployed on your tenant. If any add-ins have a value for the **AssetID** then they are deployed from the Microsoft store and you can search the `add-ins-using-exchange-tokens.xlsx` spreadsheet to see if they use legacy tokens. Add-ins without a value for the **AssetID** are centrally deployed. These can’t be found in the spreadsheet. For next steps, see [Determine if centrally deployed add-ins are using legacy tokens](#determine-if-centrally-deployed-add-ins-are-using-legacy-tokens).

Next, run the following command in PowerShell to get a list of organization add-ins. This list may include additional add-ins not found from the previous command.

- `Get-App -OrganizationApp | ft DisplayName, MarketplaceAssetID, AppId, ProviderName`

The command lists all organization add-ins deployed on your tenant. The **MarketplaceAssetID** lists the asset ID of any add-ins deployed from the Microsoft store. You can search the spreadsheet to see if these are using legacy tokens.

Add-ins without a value for the **MarketplaceAssetID** are centrally deployed. These can’t be found in the spreadsheet. 

For next steps, see [Determine if centrally deployed add-ins are using legacy tokens](#determine-if-centrally-deployed-add-ins-are-using-legacy-tokens).

**Note:** If the **ProviderName** is Microsoft, those add-ins are already being migrated and will be updated by February 2025.

## Get a list of all add-ins deployed by users

Users can also deploy add-ins from the Microsoft store, or sideload add-ins. Run the following command in PowerShell to get a list of all add-ins deployed by users in your tenant. The command loops through all user mailboxes so it may take some time to complete.

- `Get-EXOMailbox -ResultSize Unlimited | ForEach-Object {Get-App -Mailbox $_.identity | Select-Object -Property MailboxOwnerId, DisplayName, Enabled, AppVersion, AppId, MarketplaceAssetID} | Sort-Object -Property DisplayName | Get-Unique -AsString | Export-Csv add-in-list-export.csv -NoTypeInformation -Append`

The command exports an `add-in-list-export.csv` file with a list of all add-ins found for all users. For more information about this command including how to build reports around add-ins in your organization, see [Report Office Store Web Add-ins (apps) usage in your organization](https://www.howto-outlook.com/howto/report-office-store-app-usage.htm)

Open the `add-in-list-export.csv` file in Excel. If the list doesn't open with the correct columns, go to the **Data** tab, in the **Get & Transform Data group**, and choose **From Text/CSV**.

The **MarketplaceAssetID** column will list an asset ID if the add-in was deployed from the Microsoft store. You can search the spreadsheet to see if these are using legacy tokens. 

Add-ins that don't have a **MarketplaceAssetID** were deployed by the user from the organization or by sideloading. These aren’t listed in the spreadsheet and are either line-of-business add-ins or 3rd party add-ins. 

For next steps, see [Determine if centrally deployed add-ins are using legacy tokens](#determine-if-centrally-deployed-add-ins-are-using-legacy-tokens).

## Determine if Microsoft store add-ins are using legacy tokens

When you run the Get-OrganizationAddIn command it returns an **AssetID** column. The Get-App command returns a **MarketplaceAssetID** column which is equivalent to **AssetID**. The asset id is a unique value assigned to an add-in when published to the Microsoft store. Any add-ins with an asset id are deployed from the Microsoft store and can be looked up in the spreadsheet to see if they use legacy tokens.

1. Open the [add-ins-using-exchange-tokens.xlsx](add-ins-using-exchange-tokens.xlsx) spreadsheet. Or if you prefer you can open the [add-ins-using-exchange-tokens.csv](add-ins-using-exchange-tokens.csv) which contains the same list.
2. Use the asset id to search the **AssetID** column in the spreadsheet. If the add-in is listed in the spreadsheet then it was known to use legacy tokens as of April 2025.

For any add-ins deployed in your tenant that are also listed in the spreadsheet, we recommend you contact the publisher as soon as possible to confirm they have a plan and a timeline for moving off legacy tokens.

## Determine if centrally deployed add-ins are using legacy tokens

Add-ins that don’t have an asset id are deployed through central deployment or sideloaded for developer testing. You can’t look these up in the spreadsheet since they are not deployed from the Microsoft store. In many cases these add-ins are line-of-business add-ins created by your organization. It’s also possible that these add-ins are from a 3rd party. In either case we recommend reaching out to the responsible developers or publisher to ensure they are checking if they use legacy tokens and have a plan to move off legacy tokens. Developers can find more information on how to check your add-in for legacy token usage in [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).

## More resources

For more information about deprecation of legacy Exchange Online tokens, and administrator questions, see [Nested app authentication and Outlook legacy tokens deprecation FAQ](https://aka.ms/NAAFAQ).
