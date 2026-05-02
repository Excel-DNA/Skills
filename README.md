# Excel-DNA Skills

Codex skill-development workspace for Excel-DNA add-in guidance, examples, and evaluation material.

This repository currently contains:

- `skills/excel-dna-addins/` - the Excel-DNA skill package.
- `evals/excel-dna-addins/` - the testing and evaluation campaign.
- `.agents/skills/excel-dna-addins/` - local Codex activation junction or mirror; ignored by Git.
- `AGENTS.md` - repository operating guidance and target local development layout.

The project is licensed under the MIT License.

Run the static skill baseline with:

```powershell
python evals\excel-dna-addins\scripts\run_static_skill_checks.py skills\excel-dna-addins
```

Run the explicit P0 campaign with:

```powershell
python evals\excel-dna-addins\scripts\run_codex_campaign.py `
  --skill-dir skills\excel-dna-addins `
  --scenarios evals\excel-dna-addins\evals\scenarios.jsonl `
  --out runs\p0-explicit `
  --priority P0 `
  --activation explicit
```
