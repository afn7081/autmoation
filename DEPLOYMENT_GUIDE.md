# Azure Portal Deployment Guide — GWFBOB Donation Automation

## Prerequisites
- Azure account at [portal.azure.com](https://portal.azure.com)
- Docker Desktop installed locally
- Your Stripe & SendGrid API keys ready

---

## Step 1 — Create a Resource Group

1. Go to **portal.azure.com** → search **"Resource groups"** in the top search bar
2. Click **+ Create**
3. Fill in:
   - **Subscription**: Select yours
   - **Resource group**: `gwfbob-rg`
   - **Region**: `East US` (or closest to your users)
4. Click **Review + Create** → **Create**

---

## Step 2 — Create a Container Registry (ACR)

1. Search **"Container registries"** in the top bar → click **+ Create**
2. Fill in:
   - **Resource group**: `gwfbob-rg`
   - **Registry name**: `gwfbobacr` (must be globally unique)
   - **Location**: `East US`
   - **SKU**: `Basic`
3. Click **Review + Create** → **Create**
4. Once created, go to the registry → **Settings** → **Access keys**
5. Toggle **Admin user** to **Enabled**
6. Note down:
   - **Login server**: `gwfbobacr.azurecr.io`
   - **Username**: `gwfbobacr`
   - **Password**: (copy one of the passwords)

### Push your Docker image (run locally in terminal):
```bash
cd C:\Users\afanankhan\automation
docker login gwfbobacr.azurecr.io -u gwfbobacr -p <PASSWORD_FROM_ABOVE>
docker build -t gwfbobacr.azurecr.io/donation-func:latest .
docker push gwfbobacr.azurecr.io/donation-func:latest
```

---

## Step 3 — Create a Storage Account

1. Search **"Storage accounts"** → click **+ Create**
2. Fill in:
   - **Resource group**: `gwfbob-rg`
   - **Storage account name**: `gwfbobstorage` (must be globally unique, lowercase)
   - **Region**: `East US`
   - **Performance**: `Standard`
   - **Redundancy**: `LRS` (cheapest)
3. Click **Review + Create** → **Create**

---

## Step 4 — Create the Function App

1. Search **"Function App"** → click **+ Create**
2. Select **"Container App"** hosting (not the default)
3. **Basics** tab:
   - **Resource group**: `gwfbob-rg`
   - **Function App name**: `gwfbob-donations` (this becomes your URL)
   - **Region**: `East US`
4. **Container** tab:
   - **Image source**: `Azure Container Registry`
   - **Registry**: `gwfbobacr`
   - **Image**: `donation-func`
   - **Tag**: `latest`
5. **Storage** tab:
   - **Storage account**: Select `gwfbobstorage`
6. Click **Review + Create** → **Create**
7. Wait for deployment to complete (~2-3 minutes)

---

## Step 5 — Set Environment Variables

1. Go to your Function App → **Settings** → **Environment variables**
2. Click **+ Add** for each of these:

   | Name                     | Value                                      |
   |--------------------------|---------------------------------------------|
   | `STRIPE_SECRET_KEY`      | `sk_live_...` (from Stripe Dashboard)       |
   | `STRIPE_WEBHOOK_SECRET`  | `whsec_...` (you'll get this in Step 6)     |
   | `SENDGRID_API_KEY`       | `SG...` (from SendGrid Dashboard)           |
   | `FROM_EMAIL`             | `info@gwfbob.org`                           |
   | `FROM_NAME`              | `Global Women Foundation & Band of Brothers`|

3. Click **Apply** → **Confirm**

---

## Step 6 — Configure Stripe Webhook

1. Go to **[dashboard.stripe.com/webhooks](https://dashboard.stripe.com/webhooks)**
2. Click **+ Add endpoint**
3. Fill in:
   - **Endpoint URL**: `https://gwfbob-donations.azurewebsites.net/api/webhook`
   - Under **Select events** → search and check **`payment_intent.succeeded`**
4. Click **Add endpoint**
5. On the endpoint page, click **Reveal** next to **Signing secret**
6. Copy the `whsec_...` value
7. Go back to **Azure Portal → Function App → Environment variables**
8. Update `STRIPE_WEBHOOK_SECRET` with the signing secret → **Apply**

---

## Step 7 — Verify Everything Works

### Health Check
Open your browser and go to:
```
https://gwfbob-donations.azurewebsites.net/api/health
```
You should see: `{"status": "healthy"}`

### Check Functions are Registered
1. In Azure Portal → your Function App → **Functions** (left sidebar)
2. You should see two functions listed:
   - `stripe_webhook`
   - `health`

### Test with Stripe CLI (optional)
```bash
stripe trigger payment_intent.succeeded
```
Then check **Function App → Monitor → Invocations** to see the log.

---

## Monitoring & Logs

### View live logs:
1. Function App → **Monitor** → **Log stream**
2. This shows real-time output when webhooks fire

### View invocation history:
1. Function App → **Functions** → click `stripe_webhook` → **Monitor**
2. Shows each invocation with status (Success/Failure) and logs

### Set up alerts (recommended):
1. Function App → **Monitoring** → **Alerts** → **+ Create alert rule**
2. Add condition: **Function Execution Count** where **Result = Failure**
3. Add action: email notification to your team

---

## Updating the App

When you make code changes, rebuild and push:

```bash
cd C:\Users\afanankhan\automation
docker build -t gwfbobacr.azurecr.io/donation-func:latest .
docker push gwfbobacr.azurecr.io/donation-func:latest
```

Then in Azure Portal:
1. Function App → **Deployment Center** → click **Sync** (or restart the app)

---

## Cost Estimate

| Resource            | Tier     | ~Cost/month |
|---------------------|----------|-------------|
| Function App (B1)   | Basic    | ~$13        |
| Container Registry  | Basic    | ~$5         |
| Storage Account     | LRS      | ~$1         |
| **Total**           |          | **~$19/mo** |

For lower cost, you can use the **Consumption plan** instead of B1 (pay-per-execution, free for up to 1M requests/month), but cold starts may add a few seconds delay to the first webhook after idle time.
