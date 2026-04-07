---
description: "Use when the user mentions Access databases (.accdb/.mdb), Access-POSH, Access-POSH.ps1, PowerShell COM automation for Access, or asks to open/query/modify an Access database."
---
# Access-POSH Script

When working with Microsoft Access databases, use the PowerShell script at:

```
k:\Workgrp\PERSONAL SHARE\Colozzi\Access POSH\Access-POSH.ps1
```

Dot-source it in a PowerShell 7 terminal before calling any Access functions:

```powershell
. "C:\PATH\TO\Access POSH\Access-POSH.ps1"
```

This script provides 54 PowerShell functions for full Access database automation via COM. Use `-AsJson` on any function for structured output. The `@access-dev` agent has the complete function reference.
