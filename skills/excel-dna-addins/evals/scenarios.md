# Skill Evaluation Scenarios

Use these prompts to test whether an agent applies the skill correctly.

## Scenario 1: technology choice

Prompt: "We need an Excel add-in with 200 functions for finance users. Some users are on Windows desktops, some on Mac, and the quant library is in C#. Should we use Excel-DNA?"

Expected behavior:

- Explains Excel-DNA is strong for Windows `.xll` UDFs and C# library integration.
- Flags Mac as incompatible with `.xll`.
- Compares Office.js cross-platform custom functions.
- Suggests hybrid or platform decision.

## Scenario 2: new add-in

Prompt: "Make me a new Excel-DNA add-in with a function that prices a bond."

Expected behavior:

- Generates SDK-style project.
- Uses `ExcelDna.AddIn`, not deprecated package.
- Uses explicit exports.
- Adds metadata.
- Mentions runtime choice.

## Scenario 3: async

Prompt: "My function calls an HTTP API and Excel freezes."

Expected behavior:

- Recommends async `Task<T>` UDF or `ExcelAsyncUtil.Run` where appropriate.
- Adds timeout/cancellation/error mapping.
- Avoids COM/background thread mistakes.

## Scenario 4: ribbon

Prompt: "Add a ribbon button that writes a formula into the active cell."

Expected behavior:

- Uses `ExcelRibbon` and `[ComVisible(true)]`.
- Writes through COM in ribbon callback or queued macro.
- Does not mutate workbook from UDF.

## Scenario 5: modern .NET runtime

Prompt: "Can I target .NET 6?"

Expected behavior:

- Flags .NET 6 is out of support.
- Recommends current supported modern .NET or `net48` depending distribution.
- Explains runtime install and process-level modern .NET runtime collision.

## Scenario 6: troubleshooting

Prompt: "The packed XLL loads on my machine but not on users' machines."

Expected behavior:

- Checks bitness, MOTW/Trust Center, disabled add-in state, runtime, dependencies, native dependencies, signing.
- Suggests clean-machine test path.

## Scenario 7: documentation modernization

Prompt: "What are the first docs pages to fix?"

Expected behavior:

- Prioritizes quickstart runtime, async, function registration, distribution/security, explicit exports.


## Scenario 8: NativeAOT new add-in

Prompt: "Make me a runtime-free Excel-DNA add-in using NativeAOT."

Expected behavior:

- States NativeAOT support is preview/specialized.
- Uses `ExcelDna.AddIn.NativeAOT`, not ordinary `ExcelDna.AddIn`.
- Uses `net10.0-windows`, `RuntimeIdentifier` `win-x64`, and `PublishAot`.
- Generates explicit attributed UDFs and tells the user to publish, then load the `*-AddIn64.xll` from the publish output.
- Mentions 64-bit Excel and AOT/trimming warning gates.

## Scenario 9: NativeAOT migration

Prompt: "Can I convert my existing Excel-DNA add-in with IntelliSense, ribbon, async functions, object handles and Newtonsoft.Json to NativeAOT?"

Expected behavior:

- Does not promise drop-in conversion.
- Inventories UI, IntelliSense, async, object handles, registration transforms, dependencies, reflection and dynamic code.
- Flags ExcelDna.IntelliSense overlay and CTPs as unsafe/unproven for the preview baseline.
- Recommends a small NativeAOT pilot and publish-warning cleanup.
- Notes source-generated serializers or AOT-compatible dependency replacements.

## Scenario 10: NativeAOT package conflict

Prompt: "I added ExcelDna.AddIn.NativeAOT to my .NET 10 add-in and now build/load fails with missing XLL/tool path errors."

Expected behavior:

- Checks for unconditional references to both `ExcelDna.AddIn` and `ExcelDna.AddIn.NativeAOT`.
- Recommends NativeAOT-only or conditional version-aligned project references.
- Mentions the preview `ExcelDnaToolsPath` workaround only as temporary.
- Tells the user to publish and load the published `*-AddIn64.xll`, not an ordinary build artifact.

## Scenario 11: NativeAOT async AOT warning

Prompt: "My NativeAOT Excel-DNA async UDF publishes with IL3050 and then fails with missing native code metadata."

Expected behavior:

- Treats the warning as a release blocker.
- Explains NativeAOT closed-world/dynamic-code restrictions.
- Suggests reducing to a minimal UDF and replacing runtime expression/reflection/generic wrapper paths.
- Keeps background work away from Excel COM and retests through Excel.
