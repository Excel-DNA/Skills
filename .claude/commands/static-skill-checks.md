Run the Excel-DNA skill static validator and capture the result.

```powershell
New-Item -ItemType Directory -Force runs\static-skill-checks | Out-Null
python evals\excel-dna-addins\scripts\run_static_skill_checks.py skills\excel-dna-addins *>&1 | Tee-Object -FilePath runs\static-skill-checks\static-validation.txt
```

Summarize whether the command passed, the warning/error count, and the output path.
