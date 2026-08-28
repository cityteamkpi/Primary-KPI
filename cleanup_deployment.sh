#!/bin/bash
# Source environment variables (Project ID, Region, Credentials)
source "$(dirname "$0")/setup-deploy-env.sh"

FUNCTION_NAME="kpi-automation-job"
SCHEDULER_JOB="kpi-automation-job-trigger"

echo "Checking status of Cloud Function: $FUNCTION_NAME in region: $REGION..."
if gcloud functions describe "$FUNCTION_NAME" --region="$REGION" --format="value(status)" &>/dev/null; then
    echo "Deleting Cloud Function: $FUNCTION_NAME..."
    gcloud functions delete "$FUNCTION_NAME" --region="$REGION" --quiet
else
    echo "⚠️ Function $FUNCTION_NAME not found or already deleted."
fi

echo "Attempting to delete Cloud Scheduler job: $SCHEDULER_JOB..."
gcloud scheduler jobs delete "$SCHEDULER_JOB" --location="$REGION" --quiet 2>/dev/null || echo "⚠️ Scheduler job not found."