#!/bin/bash
# Source environment variables (Project ID, Region, Credentials)
source "$(dirname "$0")/setup-deploy-env.sh"

echo "Attempting to delete Cloud Function: nightly-cleanup-job in region: $REGION..."
gcloud functions delete nightly-cleanup-job --region="$REGION" --quiet

# echo "Attempting to delete Cloud Scheduler job: nightly-utilization-trigger..."
# gcloud scheduler jobs delete nightly-utilization-trigger --location="$REGION" --quiet