#!/bin/bash

# Capture the function URL to use in the scheduler
FUNCTION_URL=$(gcloud functions describe kpi-automation-job --region="$REGION" --format='value(serviceConfig.uri)')

if [ -z "$FUNCTION_URL" ]; then
    echo "❌ Error: Could not retrieve Function URL. Skipping Scheduler update."
    return 1 2>/dev/null || exit 1
fi

# Create or update the Cloud Scheduler job to run at 6:00 AM PST
COMMON_ARGS=(
    --location="$REGION"
    --schedule="0 6 * * *"
    --time-zone="America/Los_Angeles"
    --uri="${FUNCTION_URL}?task=all"
    --http-method=POST
    --oidc-service-account-email="$COMPUTE_SVC_ACCOUNT"
    --message-body='{"task": "all"}'
)

#gcloud scheduler jobs update http kpi-automation-job-trigger "${COMMON_ARGS[@]}" --update-headers="Content-Type=application/json" 2>/dev/null || \
#gcloud scheduler jobs create http kpi-automation-job-trigger "${COMMON_ARGS[@]}" --headers="Content-Type=application/json"

if gcloud scheduler jobs describe kpi-automation-job-trigger --location="$REGION" &>/dev/null; then
    gcloud scheduler jobs update http kpi-automation-job-trigger "${COMMON_ARGS[@]}" --update-headers="Content-Type=application/json"
else
    gcloud scheduler jobs create http kpi-automation-job-trigger "${COMMON_ARGS[@]}" --headers="Content-Type=application/json"
fi