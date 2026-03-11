Link this folder to your GitHub repository
=========================================

Option 1 — Run the script (e.g. in PowerShell or a terminal where Git is on PATH)

  cd "C:\Users\ryanc\OneDrive\repos\Trading_Algo"
  .\Link_GitHub.ps1

  When prompted, paste your GitHub repo URL (e.g. https://github.com/YourUsername/Trading-Algo.git).

  Or pass the URL as an argument:
  .\Link_GitHub.ps1 https://github.com/YourUsername/Trading-Algo.git

Option 2 — Run the commands yourself

  cd "C:\Users\ryanc\OneDrive\repos\Trading_Algo"
  git init
  git remote add origin YOUR_GITHUB_REPO_URL
  git add .
  git commit -m "Initial commit"
  git branch -M main
  git push -u origin main

  If the remote already has commits (e.g. README), run this before pushing:
  git pull origin main --allow-unrelated-histories -m "Merge remote with local project"
  then: git push -u origin main

A .gitignore has been added so Python cache, venv, and common IDE/OS files are not committed.
