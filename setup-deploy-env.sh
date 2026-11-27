#!/bin/bash

# 1. Define Paths and Constants
SCRIPT_DIR="$( cd -- "$( dirname -- "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
export PROJECT_ID="ct-kpi-automation"
export REGION="us-central1"
export SA_KEY_PATH="$SCRIPT_DIR/ct-kpi-automation-d56fab25ab61.json"
export SKIP_SCHEDULER="true"

# 2. Parse Arguments
# Arg 1: Account Email (Defaults to gdejager@cityteam.org)
# Arg 2: Environment Mode "dev" or "deploy" (Defaults to dev)
export ACCOUNT_EMAIL="${1:-gdejager@cityteam.org}"
ENV_MODE="${2:-dev}"

# 3. Configure Authentication based on Mode
if [ "$ENV_MODE" = "dev" ]; then
    echo "🛠️ Initializing DEVELOPMENT mode for account: $ACCOUNT_EMAIL"
    
    # Set active user account & project in gcloud
    gcloud config set account "$ACCOUNT_EMAIL" --quiet || { echo "❌ Failed to set account"; exit 1; }
    gcloud config set project "$PROJECT_ID" --quiet

elif [ "$ENV_MODE" = "deploy" ]; then
    echo "🚀 Initializing DEPLOYMENT mode..."
    
    # Authenticate gcloud directly using the Service Account key
    gcloud auth activate-service-account --key-file="$SA_KEY_PATH" --quiet || { echo "❌ SA Auth failed"; exit 1; }
    gcloud config set project "$PROJECT_ID" --quiet
    
    # Point libraries directly to the key file
    export GOOGLE_APPLICATION_CREDENTIALS="$SA_KEY_PATH"
else
    echo "❌ Unknown mode: $ENV_MODE. Usage: source setup-deploy-env.sh [ACCOUNT_EMAIL] [dev|deploy]"
    exit 1
fi

gcloud config set functions/region "$REGION" --quiet

# 4. Dynamically retrieve the Project Number
PROJECT_NUMBER=$(gcloud projects describe "$PROJECT_ID" --format='value(projectNumber)' 2>/dev/null)
if [ -z "$PROJECT_NUMBER" ]; then
    echo "❌ Error: Could not retrieve Project Number. Please ensure [$ACCOUNT_EMAIL] is authenticated via 'gcloud auth login'."
    exit 1
fi

COMPUTE_SVC_ACCOUNT="${PROJECT_NUMBER}-compute@developer.gserviceaccount.com"
echo "✅ Environment initialized in [$ENV_MODE] mode for $PROJECT_ID."