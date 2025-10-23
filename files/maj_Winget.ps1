$nom_module = "Microsoft.WinGet.Client"
        $installed_module = Get-Module -ListAvailable -Name $nom_module

        if ($null -eq $installed_module) {
            Install-Module -Name Microsoft.WinGet.Client -Force -AllowClobber
        } else {
            $module_version_instal = ($installed_module | Select-Object -ExpandProperty Version)
            $available_Module = Find-Module -Name $nom_module -ErrorAction SilentlyContinue
        if ($available_Module) {
            $source_microsoft = $available_Module.Version
            if ($module_version_instal -lt $source_microsoft) {
                Update-Module -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue    
            }
        }
        }

        winget upgrade --id Microsoft.Winget --silent --accept-source-agreements --accept-package-agreements