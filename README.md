# confluence-scanner
#### docker build
```
docker build -t confluence-scanner .
```
#### Run with auth
```
docker run --rm \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://company.atlassian.net/ \
  --token TOKEN \
  --username user@company.com \
  --regex-file /app/regex.txt \
  -m both
```
#### Run without auth
```
docker run --rm \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://company.atlassian.net/ \
  --regex-file /app/regex.txt \
  --public-only \
  -m both
```
#### Output
```
$(pwd)/output/
├── confluence_results_pages.csv
├── confluence_results_pages.xlsx
├── confluence_results_files.csv
└── confluence_results_files.xlsx
```
