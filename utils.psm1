Using namespace System.Net
using namespace Microsoft.Xrm.Sdk
using namespace Microsoft.Xrm.Sdk.Query
using namespace Microsoft.Crm.Sdk.Messages
using namespace Microsoft.Xrm.Sdk.Messages





#test mailbox request
#$requestUrl = "https://graph.microsoft.com/v1.0/me/mailbox"
#$response = Invoke-RestMethod -Uri $requestUrl -Headers $authHeader -Method GET