#!/bin/bash
if [[ "${BASH_SOURCE[0]}" != "${0}" ]]; then
    echo "❌ Error: This script must be executed, not sourced. Please run it as an executable: ./deployment_run_file.sh"
    return 1 2>/dev/null || exit 1
fi

set -e
source "$(dirname "$0")/setup-deploy-env.sh"

echo "👤 Active Account: $(gcloud config get-value account)"

# 1. Check and Add IAM policy binding. 
echo "🔑 Ensuring Cloud Run invoker role is assigned to $COMPUTE_SVC_ACCOUNT..."

# We attempt to add the binding. If it fails due to permissions, we warn the user 
# but continue, as the role might already be granted at the project or resource level.
if ! gcloud projects add-iam-policy-binding "$PROJECT_ID" \
     --member="serviceAccount:$COMPUTE_SVC_ACCOUNT" \
     --role="roles/run.invoker" \
     --quiet >/dev/null 2>&1; then
    echo "⚠️  Note: Could not update project IAM policy. This is expected if the role is already assigned or if this identity lacks 'Project IAM Admin' permissions."
else
    echo "✅ IAM Binding verified/updated."
fi

# 2. Deploy the Cloud Function (intentionally omitting    --no-allow-unauthenticated \)
# so this deploy doesn't require run.services.setIamPolicy
echo "🚀 Deploying Cloud Function..."
gcloud functions deploy kpi-automation-job \
    --gen2 \
    --runtime=python311 \
    --region="$REGION" \
    --service-account="$COMPUTE_SVC_ACCOUNT" \
    --memory=1Gi \
    --timeout=540s \
    --trigger-http \
    --entry-point=run_my_script

echo "✅ Deployment complete!"
echo "🌐 Service URL: $(gcloud functions describe kpi-automation-job --region="$REGION" --format='value(serviceConfig.uri)')"
echo "📊 Status: $(gcloud functions describe kpi-automation-job --region="$REGION" --format='value(state)')"

if [ "$SKIP_SCHEDULER" = "true" ]; then
    echo "⏭️ Skipping Cloud Scheduler update as requested."
    exit 0
fi


source "$(dirname "$0")/deploy_scheduler.sh"
