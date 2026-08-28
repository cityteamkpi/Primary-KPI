#!/bin/bash

# 1. Define Paths and Constants
# Get the absolute path of the directory containing this script
SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" &> /dev/null && pwd)
export SA_KEY_PATH="$SCRIPT_DIR/ct-kpi-automation-d56fab25ab61.json"
export PROJECT_ID="ct-kpi-automation"
export REGION="us-central1"
export GOOGLE_APPLICATION_CREDENTIALS="$SA_KEY_PATH"

# Toggle this to "true" to skip Cloud Scheduler updates by default
export SKIP_SCHEDULER="true"

# 2. Configure gcloud CLI
echo "🔐 Authenticating with Service Account..."
gcloud auth activate-service-account --key-file="$SA_KEY_PATH" --quiet || { echo "❌ Auth failed"; exit 1; }

SA_EMAIL=$(gcloud auth list --filter=status:ACTIVE --format='value(account)')
gcloud config set account "$SA_EMAIL" --quiet
gcloud config set project "$PROJECT_ID" --quiet
gcloud config set functions/region "$REGION" --quiet

echo "✅ Environment initialized for $PROJECT_ID."
