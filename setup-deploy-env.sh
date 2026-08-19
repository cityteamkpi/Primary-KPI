#!/bin/bash

# 1. Define Paths and Constants
# Get the absolute path of the directory containing this script
SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" &> /dev/null && pwd)
export SA_KEY_PATH="$SCRIPT_DIR/ct-kpi-automation-d56fab25ab61.json"
export PROJECT_ID="ct-kpi-automation"
export REGION="us-central1"

# 2. Configure gcloud CLI
gcloud auth activate-service-account --key-file="$SA_KEY_PATH" --quiet
gcloud config set project "$PROJECT_ID" --quiet
gcloud config set functions/region "$REGION" --quiet

echo "✅ Environment initialized for $PROJECT_ID."
