# -*- coding: utf-8 -*-
import os
import sys
import subprocess
from pathlib import Path

PROJECT_DIR = Path(__file__).parent.resolve()
os.chdir(PROJECT_DIR)

print("=" * 60)
print("  Coach Studio - Automated 1-Click GitHub Pages Deployer")
print("=" * 60)

# 1. Ensure latest index.html and builder.html are generated
print("\n[1/4] Rebuilding latest Coach Studio HTML...")
subprocess.run([sys.executable, "build_studio.py"], check=True)

# 2. Check Git repository
print("\n[2/4] Checking Git repository...")
if not (PROJECT_DIR / ".git").exists():
    print("Initializing Git repository...")
    subprocess.run(["git", "init", "-b", "main"], check=True)
else:
    print("Git repository already initialized.")

# Check remotes
res = subprocess.run(["git", "remote", "-v"], capture_output=True, text=True)
remote_output = res.stdout.strip()

repo_url = ""
if "origin" not in remote_output:
    print("\nNo GitHub remote 'origin' found.")
    print("Please create a new repository on GitHub (e.g. https://github.com/new)")
    print("Name it: coach-studio (or whatever you prefer)")
    repo_url = input("Enter your GitHub Repository URL (e.g. https://github.com/USERNAME/coach-studio.git): ").strip()
    if not repo_url:
        print("ERROR: GitHub repository URL is required to push. Aborting.")
        sys.exit(1)
    subprocess.run(["git", "remote", "add", "origin", repo_url], check=True)
    print(f"Added remote origin: {repo_url}")
else:
    for line in remote_output.splitlines():
        if line.startswith("origin") and "(push)" in line:
            repo_url = line.split()[1]
            break
    print(f"Using existing remote: {repo_url}")

# 3. Stage & Commit
print("\n[3/4] Staging and committing files...")
subprocess.run(["git", "add", ".gitignore", "index.html", "builder.html", "build_studio.py", "make_simplified_studio.py", "deploy_to_github.py"], check=True)

# Ensure Git author identity is set
name_check = subprocess.run(["git", "config", "user.name"], capture_output=True, text=True)
if not name_check.stdout.strip():
    subprocess.run(["git", "config", "user.name", "winoftime"], check=True)
    subprocess.run(["git", "config", "user.email", "winoftime@users.noreply.github.com"], check=True)

status_res = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
if status_res.stdout.strip():
    subprocess.run(["git", "commit", "-m", "Deploy Coach Studio PWA to GitHub Pages"], check=True)
else:
    print("No file changes to commit.")

# 4. Push to GitHub
print("\n[4/4] Pushing to GitHub...")
push_res = subprocess.run(["git", "push", "-u", "origin", "main"])
if push_res.returncode != 0:
    print("\nNOTE: If push was rejected because remote has files, you can run: git pull --rebase origin main")
    sys.exit(push_res.returncode)

# Parse username and repo name from URL
pages_url = ""
username = ""
repo_name = ""
if "github.com/" in repo_url or "github.com:" in repo_url:
    clean_url = repo_url.replace("git@github.com:", "").replace("https://github.com/", "").replace(".git", "")
    parts = clean_url.split("/")
    if len(parts) >= 2:
        username, repo_name = parts[0], parts[1]
        pages_url = f"https://{username}.github.io/{repo_name}/"

print("\n" + "=" * 60)
print("  DEPLOYMENT COMPLETE!")
print("=" * 60)
if pages_url:
    print(f"\nYour Permanent Coach Studio App URL: {pages_url}")
    print("\nTo activate GitHub Pages (One-time step):")
    print(f"1. Open https://github.com/{username}/{repo_name}/settings/pages")
    print("2. Under 'Branch', select 'main' and '/ (root)', then click 'Save'.")
    print(f"3. Open {pages_url} on your phone in Chrome or Edge.")
    print("4. Tap 'Install app' or 'Add to Home screen'.")
    print("5. You will NEVER have to send or transfer the HTML file again!")
print("=" * 60)
