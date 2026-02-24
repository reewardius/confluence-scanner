# Confluence Secret Scanner

A security tool for scanning Confluence spaces for secrets, credentials, and sensitive data using regex patterns and keywords. Supports page content scanning, file attachment scanning, OCR on images, and automated email alerts via AWS SES.

---

## Features

- 🔍 **Scan page content** — searches the body of all Confluence pages
- 📎 **Scan attachments** — scans files: PDF, DOCX, XLSX, JSON, TXT, images, archives
- 🖼️ **OCR support** — extracts text from images using Tesseract
- 📦 **Archive unpacking** — scans contents of ZIP and TAR archives
- 📊 **XLSX + CSV reports** — formatted Excel report with color-coded findings
- 📧 **Email notifications** — sends scan results via AWS SES
- 🔔 **Author alerts** — notifies page editors who exposed secrets
- 🔓 **Public-only mode** — scan without credentials (anonymous access)
- 🔁 **Resume support** — resume interrupted scans from a specific space
- 🐳 **Docker-ready** — runs fully containerized

---

## Requirements

### Python (local run)

Install dependencies:
```bash
pip install -r requirements.txt
```

System dependency — Tesseract OCR:
```bash
# Ubuntu/Debian
apt-get install tesseract-ocr tesseract-ocr-eng tesseract-ocr-rus

# macOS
brew install tesseract
```

### Docker

Just Docker installed — no other dependencies needed.

---

## Quick Start

### Docker (recommended)

**1. Build the image:**
```bash
docker build -t confluence-scanner .
```

**2. Run a basic scan (pages only):**
```bash
docker run --rm \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://your-org.atlassian.net \
  --username user@example.com \
  --token YOUR_API_TOKEN \
  --regex-file /app/regex.txt \
  --output /output/results.csv
```

**3. Scan pages + attachments (creates 2 separate reports):**
```bash
docker run --rm \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://your-org.atlassian.net \
  --username user@example.com \
  --token YOUR_API_TOKEN \
  --regex-file /app/regex.txt \
  --output /output/results.csv \
  -m both \
  --no-duplicates
```

**4. Scan without credentials (public spaces only):**
```bash
docker run --rm \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://your-org.atlassian.net \
  --public-only \
  --regex-file /app/regex.txt \
  --m both \
  --output /output/results.csv
```

### Local run

```bash
python3 confluence.py \
  --base-url https://your-org.atlassian.net \
  --username user@example.com \
  --token YOUR_API_TOKEN \
  --regex-file regex.txt \
  --output results.csv
```

---

## Regex File Format

Each line in the regex file must follow the format:

```
Name:::Regex:::GroupIndex
```

Example `regex.txt`:
```
AWS Access Key:::AKIA[0-9A-Z]{16}:::0
Password:::(?i)password\s*[:=]\s*(\S+):::1
API Key:::(?i)api[_-]?key\s*[:=]\s*(\S+):::1
Private Key:::-----BEGIN (RSA |EC )?PRIVATE KEY-----:::0
Slack Token:::xox[baprs]-[0-9A-Za-z]{10,48}:::0
```

- `Name` — label shown in the report
- `Regex` — Python-compatible regular expression
- `GroupIndex` — capture group index (`0` = full match, `1` = first group, etc.)

---

## Output Files

Results are saved to the directory specified in `--output`.

| Mode | Files created |
|------|--------------|
| Default (pages only) | `results.csv`, `confluence_secrets.xlsx` |
| `--mode files` | `results.csv`, `confluence_secrets_in_files.xlsx` |
| `--mode both` | `results_pages.csv`, `confluence_secrets.xlsx`, `results_files.csv`, `confluence_secrets_in_files.xlsx` |

---

## All Arguments

### Connection

| Argument | Description |
|----------|-------------|
| `--base-url` | Confluence base URL (e.g. `https://your-org.atlassian.net`) |
| `--username` | Atlassian account email |
| `--token` | Atlassian API token |
| `--public-only` | Scan only public spaces without authentication |

### Scan patterns

| Argument | Description |
|----------|-------------|
| `--regex-file` | Path to regex patterns file (`Name:::Regex:::GroupIndex`) |
| `--regex` | Single regex pattern (legacy, simple use) |
| `--keywords` | Path to keywords file (one keyword per line) |
| `--trufflehog-patterns`, `-tp` | Path to a TruffleHog YAML file with detectors |
| `--trufflehog-keywords`, `-tk` | Include only TruffleHog detectors whose keywords field matches any of the specified values. **Example:** `aws,api,internal` |
| `--trufflehog-exclude-keywords`, `-tek` | Exclude TruffleHog detectors whose keywords field matches any of the specified values. **Example:** `gateway,arn` |

### Scan mode

| Argument | Description |
|----------|-------------|
| `-m`, `--mode` | `files` — attachments only; `both` — pages and attachments separately. Default: pages only |
| `--filetype` | File extensions to scan, e.g. `docx,pdf,json` (only with `-m files` or `-m both`) |
| `--exclude-filetype` | File extensions to exclude, e.g. `pdf` |
| `--max-size` | Maximum file size to scan, e.g. `2mb`, `500kb` |
| `--scan-images-only` | Scan only image attachments using OCR |
| `--archive-support` | Unpack and scan ZIP/TAR archives |

### Filtering

| Argument | Description |
|----------|-------------|
| `--space-keys` | Comma-separated space keys to scan, or path to a file with keys |
| `--exclude-space-keys` | Comma-separated space keys to exclude |
| `--modified-after` | Scan pages modified after this date (`D.M.Y` or `D/M/Y`) |
| `--modified-before` | Scan pages modified before this date |
| `--created-in-year` | Scan pages/files created in specific year(s), e.g. `2025` or `2025,2026` |
| `--no-duplicates` | Remove duplicate findings from the report |

### Output

| Argument | Description |
|----------|-------------|
| `-o`, `--output` | Output CSV file path (default: `confluence_results.csv`). XLSX is auto-generated alongside. |
| `--secret-max-length` | Max characters shown in "Matched Value" column (default: `20`) |
| `--config` | Path to JSON config file with arguments |
| `--debug` | Stop after first 5 findings and generate report |
| `--resume-from` | Resume scan from a specific space key |

### Email notifications (AWS SES)

| Argument | Description |
|----------|-------------|
| `--email-sender` | Sender email (must be verified in AWS SES) |
| `--email-recipient` | Recipient(s) for scan results, comma-separated |
| `--aws-region` | AWS region for SES (default: `eu-central-1`) |

### Author alerts

| Argument | Description |
|----------|-------------|
| `--alert` | Send individual alerts to page editors who exposed secrets |
| `--security-contact` | Security team email included in author alerts (default: `security@company.com`) |
| `--security-wiki` | URL to security documentation included in author alerts |

---

## Email Notifications via AWS SES

The scanner can send results by email using AWS SES.

**Required IAM permissions:**
```json
{
  "Effect": "Allow",
  "Action": ["ses:SendEmail", "ses:SendRawEmail"],
  "Resource": "*"
}
```

**When running on EC2 — assign an IAM Role with SES permissions to the instance:**

- Local run (`python3 confluence.py ...`) — no extra configuration needed, credentials are picked up automatically

- Docker run — add `--network host` so the container can reach the instance credentials:

```bash
docker run --rm \
  --network host \
  -v $(pwd)/output:/output \
  -v $(pwd)/regex.txt:/app/regex.txt:ro \
  confluence-scanner \
  --base-url https://your-org.atlassian.net \
  --username user@example.com \
  --token YOUR_API_TOKEN \
  --regex-file /app/regex.txt \
  --output /output/results.csv
```

You can also pass credentials via environment variables:
```bash
docker run --rm \
  -e AWS_ACCESS_KEY_ID=YOUR_KEY \
  -e AWS_SECRET_ACCESS_KEY=YOUR_SECRET \
  -e AWS_DEFAULT_REGION=eu-central-1 \
  ...
```

---

## Config File

Instead of passing all arguments on the command line, you can use a JSON config file:

```json
{
  "base_url": "https://your-org.atlassian.net",
  "username": "user@example.com",
  "token": "YOUR_API_TOKEN",
  "regex_file": "/app/regex.txt",
  "output": "/output/results.csv",
  "mode": "both",
  "no_duplicates": true,
  "email_sender": "security@company.com",
  "email_recipient": "appsec@company.com"
}
```

Use with:
```bash
python3 confluence.py --config config.json
```

---

## Security Notice

> ⚠️ This tool is intended for use by authorized security personnel only. Always ensure you have permission to scan the target Confluence instance before running.
