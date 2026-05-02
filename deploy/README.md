# Deploy to your own Azure account

This folder contains a generic ARM template that provisions everything needed to run the Donation Automation app in **your** Azure subscription:

- Log Analytics workspace
- Azure Container Registry (ACR)
- Storage account + `templates` blob container
- Container Apps managed environment
- Container App (with all env vars / secrets wired up)

## One-click deploy

Replace `<OWNER>/<REPO>/<BRANCH>` with this repo's location, then click:

> [![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2F<OWNER>%2F<REPO>%2F<BRANCH>%2Fdeploy%2Fazuredeploy.json)

(For this repo: `afn7081/autmoation/main`.)

## CLI deploy

```bash
# 1. Login + pick subscription
az login
az account set --subscription "<your-subscription-id>"

# 2. Create a resource group
az group create -n donations-rg -l eastus

# 3. Deploy (you'll be prompted for any secrets you leave blank in parameters)
az deployment group create \
  -g donations-rg \
  --template-file azuredeploy.json \
  --parameters @azuredeploy.parameters.example.json \
  --parameters stripeSecretKey='sk_test_...' \
               stripeWebhookSecret='whsec_...' \
               sendgridApiKey='SG....'
```

## After the first deploy

The template starts with a placeholder hello-world image so the Container App
boots successfully. To run the real app:

```bash
# Get the ACR name from the deployment output
ACR=$(az deployment group show -g donations-rg -n azuredeploy \
      --query properties.outputs.acrLoginServer.value -o tsv | cut -d. -f1)

# Build & push your image directly in Azure (no local Docker needed)
az acr build -r $ACR -t donations-app:latest .

# Point the Container App at the new image
az containerapp update -g donations-rg -n donations-app \
  --image ${ACR}.azurecr.io/donations-app:latest
```

## Parameters

| Parameter | Required | Notes |
|---|---|---|
| `appName` | no | Base name for resources. Default `donations-app`. |
| `location` | no | Defaults to the resource group's region. |
| `acrName` / `storageAccountName` | no | Auto-generated unique names if omitted. |
| `containerImage` | no | Replace with your ACR image after first push. |
| `stripeSecretKey`, `stripeWebhookSecret`, `sendgridApiKey` | **yes** | App will not function without these. |
| `cloudconvertApiKey`, `convertapiSecret` | optional | Only needed for PDF/PPTX conversion features. |
| `fromEmail`, `fromName` | yes | Sender identity used by SendGrid. |
| `cpu`, `memory`, `minReplicas`, `maxReplicas` | no | Container App sizing. |

## Outputs

After deployment, run:

```bash
az deployment group show -g donations-rg -n azuredeploy --query properties.outputs
```

You'll get:

- `containerAppFqdn` — public URL of your app
- `acrLoginServer` — push images here
- `storageAccountName` / `templatesContainer` — upload your `cert_template.pptx`,
  email header image, etc., into the `templates` container.

## Security notes

- All secrets are passed as `securestring` parameters and stored as **Container
  App secrets** — never baked into the template or image.
- The ACR password is fetched at deploy time via `listCredentials()` and stored
  as a Container App secret; admin user is enabled so the Container App can
  pull without a managed identity. For production, switch to a user-assigned
  managed identity with `AcrPull` and disable admin user.
- The storage account uses key-based auth via a connection-string secret. For
  production, prefer managed identity + `Storage Blob Data Contributor`.
