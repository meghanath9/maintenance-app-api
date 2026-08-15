param(
    [Parameter(Mandatory = $true)]
    [string]$ProjectId,

    [string]$Region = "asia-south1",

    [string]$ServiceName = "maintenance-app"
)

$ErrorActionPreference = "Stop"

Write-Host "Setting active GCP project to $ProjectId..."
gcloud config set project $ProjectId | Out-Host

Write-Host "Building container image with Cloud Build..."
$Image = "gcr.io/$ProjectId/$ServiceName"
gcloud builds submit --tag $Image | Out-Host

Write-Host "Deploying to Cloud Run..."
gcloud run deploy $ServiceName `
    --image $Image `
    --platform managed `
    --region $Region `
    --allow-unauthenticated | Out-Host

Write-Host "Deployment completed."
Write-Host "Get URL with: gcloud run services describe $ServiceName --region $Region --format='value(status.url)'"
