#!/bin/bash
set -e
source "$(dirname "$0")/setup-deploy-env.sh"

echo "👤 Active Account: $(gcloud config get-value account)"

# Dynamically retrieve the Project Number to construct the Service Account email
PROJECT_NUMBER=$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)' 2>/dev/null)
if [ -z "$PROJECT_NUMBER" ]; then
    echo "❌ Error: Could not retrieve Project Number. Please ensure you are authenticated."
    exit 1
fi
COMPUTE_SVC_ACCOUNT="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"

# 1. Check and Add IAM policy binding. 
echo "🔑 Checking if IAM binding for Cloud Run invoker exists..."
# add-iam-policy-binding is idempotent; we can run it directly to ensure the state.
echo "🔑 Ensuring Cloud Run invoker role is assigned to the compute service account..."
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
    --member="serviceAccount:$COMPUTE_SVC_ACCOUNT" \
    --role="roles/run.invoker" \
    --quiet >/dev/null || echo "⚠️ Warning: Could not set IAM policy. Ensure your account has 'Project IAM Admin' permissions."
echo "✅ IAM Binding complete"

# 2. Deploy the Cloud Function
gcloud functions deploy kpi-automation-job \
    --gen2 \
    --runtime=python311 \
    --region="$REGION" \
    --service-account="$COMPUTE_SVC_ACCOUNT" \
    --memory=1Gi \
    --timeout=540s \
    --trigger-http \
    --no-allow-unauthenticated \
    --entry-point=run_my_script

echo "✅ Deployment complete"

if [ "$SKIP_SCHEDULER" = "true" ]; then
    echo "⏭️ Skipping Cloud Scheduler update as requested."
    exit 0
fi


source "$(dirname "$0")/deploy_scheduler.sh"
