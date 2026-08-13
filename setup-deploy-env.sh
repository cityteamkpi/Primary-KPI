#!/bin/bash

# 1. Define Paths and Constants
export SA_KEY_PATH="/home/kpi/ct-kpi-automation/ct-kpi-automation-d56fab25ab61.json"
export PROJECT_ID="ct-kpi-automation"
export REGION="us-central1"

# 2. Configure gcloud CLI
gcloud auth activate-service-account --key-file="$SA_KEY_PATH" --quiet
gcloud config set project "$PROJECT_ID" --quiet
gcloud config set functions/region "$REGION" --quiet

echo "✅ Environment initialized for $PROJECT_ID."

