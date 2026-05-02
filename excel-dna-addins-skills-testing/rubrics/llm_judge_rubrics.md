# Rubrics for Excel-DNA Skill Evaluation

Use rubric graders only after deterministic checks have run. Rubric graders should not override a deterministic build failure, unsafe advice flag, or source-of-truth contradiction.

## Universal 0-5 scale

- **5**: Correct, specific, grounded in Excel-DNA behavior, practical, and tailored to the user's persona.
- **4**: Mostly correct and useful; minor omission or wording issue.
- **3**: Usable but incomplete; may require an experienced user to repair gaps.
- **2**: Partly wrong, vague, or risky in a non-critical way.
- **1**: Mostly wrong or likely to lead to wasted time.
- **0**: Critical failure, unsafe advice, wrong platform, or unusable artifact.

## Rubric dimensions

### Domain correctness

Checks Excel-DNA-specific facts:
- `.xll` add-ins are Windows desktop Excel technology.
- UDFs integrate with Excel calculation and are the primary extension surface.
- Runtime trade-offs are explicit: .NET Framework stability and isolation versus modern .NET features and runtime/coexistence costs.
- v1.9 registration changes are reflected: integrated extended registration, optional/default params, async/streaming, object handles, execution handlers.
- UI surfaces are correctly separated: UDFs, commands, ribbon, CTP, IntelliSense, COM automation, C API context.
- Distribution advice covers bitness, runtime, signing/trust, packed XLL/installer choices.

### Artifact correctness

For generated code:
- Buildable project files.
- Correct NuGet package choices.
- Correct TargetFramework/TargetFrameworks.
- Appropriate attributes and metadata.
- No accidental exports in production scenarios.
- No blocking async `.Result`/`.Wait()` in UDF examples.
- No COM object model access from thread-safe UDFs or background threads.

### User fit

- Beginner prompts get simple steps and minimal jargon.
- Advanced prompts get exact mechanisms and trade-offs.
- Enterprise prompts cover deployment controls and risk.
- Maintainer prompts include migration strategy and regression tests.
- C API prompts are precise and conservative.

### Risk handling

Critical risk examples:
- Claims XLL add-ins work on Mac or web Excel.
- Recommends .NET 6 as current LTS in 2026.
- Ignores .NET Desktop Runtime and bitness requirements for modern .NET.
- Advises users to disable all Excel security warnings.
- Uses UI-mutating COM/C API calls from a worksheet UDF calculation context.
- Calls Excel COM from arbitrary background threads.
- Exposes unintended helper methods as worksheet functions.

### Efficiency

- Advice-only tasks should not create files or run builds.
- Single-topic tasks should not read every reference file.
- Scaffold tasks should use templates/scripts rather than iterative rewrites.
- Troubleshooting should follow a high-yield order before obscure hypotheses.

## Judge prompt template

You are grading an agent response for an Excel-DNA add-in skill evaluation. Use the scenario, expected behavior, deterministic check results, and response/artifacts.

Return JSON only:
```json
{
  "domain_correctness": 0,
  "artifact_correctness": 0,
  "user_fit": 0,
  "risk_handling": 0,
  "efficiency": 0,
  "critical_failures": [],
  "failure_tags": [],
  "short_rationale": ""
}
```

Do not reward verbose but non-specific answers. Penalize unsupported claims. Mark any critical failure even if most of the response is useful.
