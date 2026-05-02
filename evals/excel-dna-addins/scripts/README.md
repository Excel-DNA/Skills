# Scripts

- `run_static_skill_checks.py`: validates skill package shape, frontmatter, direct links, stale phrases, and script syntax.
- `run_codex_campaign.py`: runs filtered scenarios through the local Codex CLI.
- `run_campaign.py`: skeleton for invoking a selected agent runner over all scenarios.
- `score_campaign.py`: aggregates deterministic checks for a captured campaign run.
- `score_agent_run.py`: deterministic scorer for captured answer/workspace output.

The runner integration should be adapted to the chosen framework so `trace.json` records skill activation, reference loads, commands, token/cost data, and generated artifacts.
