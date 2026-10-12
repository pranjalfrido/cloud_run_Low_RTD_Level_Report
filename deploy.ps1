# ── Low RTD Email — Deploy to Google Cloud Run + Scheduler ───────────────────
# Run this once from the Low_RTD_Email_Automation folder:
#   cd "Low_RTD_Email_Automation"
#   .\deploy.ps1

$PROJECT_ID   = "frido-429506"
$REGION       = "asia-south1"          # Mumbai — closest to IST timezone
$IMAGE        = "gcr.io/$PROJECT_ID/low-rtd-email"
$JOB_NAME     = "low-rtd-email-job"
$SA_EMAIL     = "for-quickcomm-automation@$PROJECT_ID.iam.gserviceaccount.com"

Write-Host "=== Step 1: Build & push Docker image ===" -ForegroundColor Cyan
gcloud builds submit --tag $IMAGE --project $PROJECT_ID

Write-Host "`n=== Step 2: Deploy Cloud Run Job ===" -ForegroundColor Cyan
gcloud run jobs deploy $JOB_NAME `
  --image $IMAGE `
  --region $REGION `
  --project $PROJECT_ID `
  --service-account $SA_EMAIL `
  --task-timeout 600 `
  --max-retries 1 `
  --memory 512Mi `
  --cpu 1

Write-Host "`n=== Step 3: Create Cloud Scheduler jobs (IST = UTC+5:30) ===" -ForegroundColor Cyan

# 7:15 AM IST  = 01:45 UTC
gcloud scheduler jobs create http "low-rtd-email-morning" `
  --location $REGION `
  --project $PROJECT_ID `
  --schedule "45 1 * * *" `
  --time-zone "Asia/Kolkata" `
  --uri "https://$REGION-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/$PROJECT_ID/jobs/$JOB_NAME`:run" `
  --http-method POST `
  --oauth-service-account-email $SA_EMAIL `
  --message-body "" `
  --attempt-deadline 600s

# 4:15 PM IST  = 10:45 UTC
gcloud scheduler jobs create http "low-rtd-email-evening" `
  --location $REGION `
  --project $PROJECT_ID `
  --schedule "45 10 * * *" `
  --time-zone "Asia/Kolkata" `
  --uri "https://$REGION-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/$PROJECT_ID/jobs/$JOB_NAME`:run" `
  --http-method POST `
  --oauth-service-account-email $SA_EMAIL `
  --message-body "" `
  --attempt-deadline 600s

Write-Host "`n=== Done! ===" -ForegroundColor Green
Write-Host "Morning  job: 7:15 AM IST  (cron: 45 1 * * *  Asia/Kolkata)"
Write-Host "Evening  job: 4:15 PM IST  (cron: 45 10 * * * Asia/Kolkata)"
Write-Host "Cloud Run Job: $JOB_NAME in $REGION"
