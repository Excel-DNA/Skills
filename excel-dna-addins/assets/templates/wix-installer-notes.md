# WiX Installer Notes

Use the Excel-DNA WiX installer template as a starting point for MSI distribution.

Checklist:

- Decide per-user vs per-machine.
- Add product/upgrade codes.
- Copy packed `.xll` outputs to stable install path.
- Register Excel add-in for user or configure enterprise deployment.
- Detect/install required .NET Desktop Runtime for modern .NET targets.
- Sign `.xll` and MSI.
- Test install, repair, upgrade, uninstall.
