# Azure Function Proxy for Document Intelligence

This folder contains a simple Azure Function that proxies SharePoint Framework requests to Azure Document Intelligence. It allows the SPFx web part to send PDF bytes to the function and receive classification results without calling Azure Cognitive Services directly from the browser.

## Files

- `host.json` - Azure Functions host configuration.
- `local.settings.json` - Local function settings with placeholders for the Document Intelligence endpoint and API key.
- `.funcignore` - Files ignored when publishing the function.
- `ClassifyDocument/function.json` - HTTP-triggered function binding.
- `ClassifyDocument/index.js` - Function implementation.

## Setup

1. Install the Azure Functions Core Tools.
2. Set the real values in `azure-function/local.settings.json`:
   - `DOCUMENT_INTELLIGENCE_ENDPOINT`
   - `DOCUMENT_INTELLIGENCE_API_KEY`
3. Run the function locally from the `azure-function` folder:
   - `func start`

## Request format

The function accepts POST requests to `/api/classify` with JSON:

```json
{
  "modelId": "<your-model-id>",
  "fileBase64": "<base64-encoded-pdf>"
}
```

The function returns the Document Intelligence analysis result JSON.

## Deploy to Azure

1. Install the Azure CLI and Azure Functions Core Tools.
2. Sign in:
   - `az login`
3. Create a resource group (if you do not already have one):
   - `az group create --name <resource-group> --location <region>`
4. Create a storage account:
   - `az storage account create --name <storage-account-name> --location <region> --resource-group <resource-group> --sku Standard_LRS`
5. Create the Function App:
   - `az functionapp create --resource-group <resource-group> --consumption-plan-location <region> --name <function-app-name> --storage-account <storage-account-name> --runtime node --runtime-version 18 --functions-version 4`
6. Configure app settings for Azure Document Intelligence:
   - `az functionapp config appsettings set --name <function-app-name> --resource-group <resource-group> --settings DOCUMENT_INTELLIGENCE_ENDPOINT="https://<your-document-intelligence-resource>.cognitiveservices.azure.com" DOCUMENT_INTELLIGENCE_API_KEY="<your-api-key>"`
7. Publish from the `azure-function` folder:
   - `func azure functionapp publish <function-app-name> --javascript`

After publish, use the function URL:

`https://<function-app-name>.azurewebsites.net/api/classify`

> Do not publish `local.settings.json`. It is only for local development.

## Notes

- The function is configured as anonymous for easy SPFx integration. If you want to secure it, change `authLevel` in `function.json` and pass the function key from the client.
- Do not commit real keys into source control.
