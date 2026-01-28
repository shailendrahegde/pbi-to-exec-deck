# GitHub Repository Setup Instructions

Follow these steps to create and publish your Claude PowerPoint Constitution repository.

## Repository Contents

Your `github-repo` directory contains:

```
github-repo/
├── .gitignore                              # Excludes temporary files
├── README.md                               # Main documentation (purpose & steps)
├── Claude PowerPoint Constitution.md       # The governance framework
├── QUICK-START.md                          # 5-minute setup guide
├── EXAMPLE-SETUP.md                        # Directory structure examples
├── LICENSE                                 # MIT License
└── GITHUB-SETUP-INSTRUCTIONS.md           # This file
```

## Step 1: Create GitHub Repository

### Option A: Via GitHub Web UI

1. Go to https://github.com/new
2. Enter repository details:
   - **Name:** `pbi-to-exec-deck`
   - **Description:** `Transform Power BI dashboards into executive-ready presentations using Claude Code and governance framework`
   - **Visibility:** Public (or Private if preferred)
   - **Initialize:** ❌ Do NOT initialize with README (we have one)
3. Click "Create repository"

### Option B: Via GitHub CLI

```bash
gh repo create pbi-to-exec-deck --public --description "Transform Power BI dashboards into executive-ready presentations using Claude Code"
```

## Step 2: Initialize Local Git Repository

Open terminal/command prompt in your github-repo directory:

```bash
cd C:\Users\shahegde\claudex\github-repo

# Initialize git
git init

# Add all files
git add .

# Create initial commit
git commit -m "Initial commit: Claude PowerPoint Constitution v1.4"
```

## Step 3: Link to GitHub and Push

Replace `[your-username]` with your GitHub username:

```bash
# Add remote repository
git remote add origin https://github.com/[your-username]/pbi-to-exec-deck.git

# Push to GitHub
git branch -M main
git push -u origin main
```

### Using SSH (Alternative)

If you prefer SSH authentication:

```bash
git remote add origin git@github.com:[your-username]/pbi-to-exec-deck.git
git branch -M main
git push -u origin main
```

## Step 4: Verify Repository

1. Go to `https://github.com/[your-username]/pbi-to-exec-deck`
2. Verify all files are visible:
   - README.md displays as the main page
   - Constitution file is accessible
   - Quick Start and Example Setup are present

## Step 5: Add Topics and Description (Optional)

On your GitHub repository page:

1. Click "⚙️ Settings" (if you have admin access) or "About" gear icon
2. Add topics (helps with discoverability):
   - `claude-code`
   - `powerpoint`
   - `power-bi`
   - `analytics`
   - `executive-presentations`
   - `dashboard`
   - `constitution`
3. Add website URL if you have documentation hosted elsewhere

## Step 6: Create Release (Optional)

Tag the current version:

```bash
git tag -a v1.4 -m "Version 1.4: Blank template requirement, 2-3 insights per slide"
git push origin v1.4
```

On GitHub:
1. Go to "Releases" → "Create a new release"
2. Choose tag `v1.4`
3. Release title: `Claude PowerPoint Constitution v1.4`
4. Description:
```markdown
## What's New in v1.4
- Blank template requirement (Section 8.1)
- 2-3 insights per slide for executive focus (Section 3.5)
- Image-left, text-right layout (Section 7.7)

## Getting Started
See [QUICK-START.md](QUICK-START.md) for setup instructions.
```

## Updating the Repository

When you make changes to the constitution:

```bash
cd C:\Users\shahegde\claudex\github-repo

# Make your changes to files

# Stage changes
git add .

# Commit with descriptive message
git commit -m "Update constitution: Add new rule for [description]"

# Push to GitHub
git push origin main
```

## Repository Structure Best Practices

### Recommended README Sections (Already Included)

✅ Purpose statement
✅ How it works
✅ Getting started steps
✅ Troubleshooting
✅ Version history

### Consider Adding (Future)

- **CHANGELOG.md**: Detailed version history
- **CONTRIBUTING.md**: Guidelines for contributors
- **Examples/**: Folder with before/after examples
- **Templates/**: Sample corporate templates (if shareable)
- **Issues**: Enable GitHub Issues for community questions

## Sample Repository URLs

```
Repository: https://github.com/[your-username]/pbi-to-exec-deck
Raw Constitution: https://raw.githubusercontent.com/[your-username]/pbi-to-exec-deck/main/Claude%20PowerPoint%20Constitution.md
```

Users can download the constitution directly:

```bash
curl -O https://raw.githubusercontent.com/[your-username]/pbi-to-exec-deck/main/Claude%20PowerPoint%20Constitution.md
```

## Sharing the Repository

### Social Media Description

"New: Claude PowerPoint Constitution - A governance framework to transform Power BI dashboards into executive-ready presentations with Claude Code. Source-grounded insights, compelling headlines, and consistent formatting. https://github.com/[your-username]/pbi-to-exec-deck"

### Internal Distribution

Share with your team:
```
🎯 New Tool Available: Claude PowerPoint Constitution

Automatically convert Power BI exports into executive-ready presentations using Claude Code.

✅ Compelling insights
✅ Source fidelity
✅ Corporate template styling

Get started: https://github.com/[your-username]/pbi-to-exec-deck
```

## License Notes

This repository uses MIT License, which allows:
- ✅ Commercial use
- ✅ Modification
- ✅ Distribution
- ✅ Private use

Users must:
- Include copyright notice
- Include license text

## Next Steps

1. ✅ Create GitHub repository
2. ✅ Push files
3. ✅ Verify all content visible
4. Share with team
5. Gather feedback
6. Iterate on constitution based on usage

---

**Questions?** Open an issue on GitHub or reach out to the repository maintainer.
