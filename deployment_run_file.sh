#!/bin/bash
set -e
source "$(dirname "$0")/setup-deploy-env.sh"

# Dynamically retrieve the Project Number to construct the Service Account email
PROJECT_NUMBER=$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)')
COMPUTE_SVC_ACCOUNT="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"

# 1. Add IAM policy binding. 
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
    --member="serviceAccount:$COMPUTE_SVC_ACCOUNT" \
    --role="roles/run.invoker" \
    --account="kpi@cityteam.org" \
    --quiet

echo "✅ IAM Binding complete"

# 2. Deploy the Cloud Function
gcloud functions deploy nightly-program-utilization-job \
    --gen2 \
    --runtime=python311 \
    --region="$REGION" \
    --service-account="$COMPUTE_SVC_ACCOUNT" \
    --trigger-http \
    --no-allow-unauthenticated \
    --entry-point=run_my_script

echo "✅ Deploment complete"

# Capture the function URL to use in the scheduler
FUNCTION_URL=$(gcloud functions describe nightly-program-utilization-job --region="$REGION" --format='value(serviceConfig.uri)')

# Create or update the Cloud Scheduler job to run at 6:00 AM nightly
COMMON_ARGS=(
    --location="$REGION"
    --schedule="0 6 * * *"
    --time-zone="America/Los_Angeles"
    --uri="${FUNCTION_URL}?task=utilization"
    --http-method=POST
    --oidc-service-account-email="$COMPUTE_SVC_ACCOUNT"
    --message-body='{"task": "utilization"}'
)

gcloud scheduler jobs update http nightly-utilization-trigger "${COMMON_ARGS[@]}" --update-headers="Content-Type=application/json" || \
gcloud scheduler jobs create http nightly-utilization-trigger "${COMMON_ARGS[@]}" --headers="Content-Type=application/json"
