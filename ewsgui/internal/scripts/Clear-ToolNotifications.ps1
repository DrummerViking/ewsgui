﻿powershell.exe -command "Set-PSFConfig -Module 'EWSGui' -Name 'ExoGraphGuiNotification' -Value $false -Initialize -Validation 'bool' -Description 'Whether the module should display notifications about ExoGraphGUI tool.' -PassThru | Register-PSFConfig"