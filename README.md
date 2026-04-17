# Resume Tailor

Automatically tailor your resume to any job description using GitHub Models AI.

## How It Works

1. **Upload a JD** — Drop a job description file (`.txt`, `.pdf`, or `.docx`) into the `jd/` folder
2. **GitHub Actions triggers** — The workflow detects the new JD and runs automatically
3. **AI tailors your resume** — GitHub Models API rewrites bullet points and reorders skills to match the JD
4. **Get your files** — A tailored `.docx` and `.pdf` appear in the `output/` folder, committed back to the repo

The original resume structure, formatting, fonts, and layout are **fully preserved**. Only the text content of experience bullets and skills ordering are adjusted.

## Setup

### 1. Add your GitHub Models token as a repository secret

Go to **Settings → Secrets and variables → Actions → New repository secret**

- Name: `GH_MODELS_TOKEN`
- Value: Your GitHub personal access token (with Models API access)

### 2. Upload a job description

Add any `.txt`, `.pdf`, or `.docx` file to the `jd/` folder and push:

```bash
cp my_job_description.txt jd/
git add jd/
git commit -m "Add JD for Google SWE position"
git push
```

### 3. Get results

After the workflow completes (~2-3 minutes), check the `output/` folder for:
- `Rehan_Malik_Resume_<jd_name>.docx`
- `Rehan_Malik_Resume_<jd_name>.pdf`

You can also download them from the **Actions → Artifacts** tab.

### Manual trigger

You can also trigger the workflow manually from the Actions tab using "Run workflow" and specifying the JD filename.

## What Gets Tailored

| Section | What Changes |
|---------|-------------|
| **Experience bullets** | Rewritten to emphasize relevant skills and achievements matching the JD |
| **Skills & Certifications** | Reordered to prioritize skills mentioned in the JD |
| **Name, Contact, Education** | Never changed |
| **Formatting** | Fully preserved (fonts, sizes, colors, alignment, styles) |

## Local Usage

```bash
pip install -r requirements.txt

python scripts/tailor_resume.py \
  --jd jd/my_job.txt \
  --template template/resume_template.docx \
  --output-dir output \
  --token YOUR_GITHUB_TOKEN
```

## Project Structure

```
├── .github/workflows/
│   └── tailor_resume.yml    # GitHub Actions workflow
├── jd/                      # Drop JD files here
├── output/                  # Tailored resumes appear here
├── template/
│   └── resume_template.docx # Your master resume
├── scripts/
│   └── tailor_resume.py     # Main tailoring script
├── requirements.txt
└── README.md
```
