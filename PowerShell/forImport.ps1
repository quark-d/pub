{
    "version": "2.0.0",
    "tasks": [
        {
            "label": "Import VBA Module",
            "type": "shell",
            "command": "powershell",
            "args": [
                "-ExecutionPolicy", "Bypass",
                "-File", "${workspaceFolder}/RunMacro.ps1",
                "-excelFilePath", "C:/path/to/Sample.xlsm",
                "-vbaModulePath", "${file}"
            ],
            "problemMatcher": [],
            "group": {
                "kind": "build",
                "isDefault": true
            }
        }
    ]
}
