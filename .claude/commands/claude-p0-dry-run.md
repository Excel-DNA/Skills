Check that the Claude Code P0 explicit campaign command is runnable without starting the long campaign.

```powershell
python evals\excel-dna-addins\scripts\run_claude_campaign.py `
  --skill-dir skills\excel-dna-addins `
  --scenarios evals\excel-dna-addins\evals\scenarios.jsonl `
  --out runs\claude-p0-explicit `
  --priority P0 `
  --activation explicit `
  --dry-run
```

Report selected scenario count, runner command, and any path or CLI prerequisite issues.
