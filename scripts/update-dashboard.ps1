$ErrorActionPreference = "Stop"

Set-Location "E:\personal\Sin Dashboard"

python scripts\build-data.py --input data\lot-details.xlsx --output data\lots.json

git add data\lot-details.xlsx data\lots.json

git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    exit 0
}

git commit -m "Auto update lot details"
git push
