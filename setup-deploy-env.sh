#!/bin/bash

# 1. Define Paths and Constants
# Get the absolute path of the directory containing this script
SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" &> /dev/null && pwd)
export SA_KEY_PATH="$SCRIPT_DIR/ct-kpi-automation-d56fab25ab61.json"
export PROJECT_ID="ct-kpi-automation"
export REGION="us-central1"
export GOOGLE_APPLICATION_CREDENTIALS="$SA_KEY_PATH"
export ACCOUNT_EMAIL=kpi@cityteam.org


# Toggle this to "true" to skip Cloud Scheduler updates by default
export SKIP_SCHEDULER="true"

# 2. Configure gcloud CLI (Claude says not to do this)
#echo "🔐 Authenticating with Service Account..."
#gcloud auth activate-service-account --key-file="$SA_KEY_PATH" --quiet || { echo "❌ Auth failed"; exit 1; }

gcloud config set account "$ACCOUNT_EMAIL" --quiet
gcloud config set project "$PROJECT_ID" --quiet
gcloud config set functions/region "$REGION" --quiet

# Dynamically retrieve the Project Number to construct the Service Account email
PROJECT_NUMBER=$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)' 2>/dev/null)
if [ -z "$PROJECT_NUMBER" ]; then
    echo "❌ Error: Could not retrieve Project Number. Please ensure you are authenticated."
    exit 1
fi
COMPUTE_SVC_ACCOUNT="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"


echo "✅ Environment initialized for $PROJECT_ID."
