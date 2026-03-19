#
# decoster.tech.config.base.psd1
# Shared configuration file for DecosterOutput scripts (DecomVMWithGui, DHCPreservation_AND_VMCreation_GUI, ...).
# Place this file in the same directory as the scripts that use it.
#
@{
    # -------------------------------------------------------------------------
    # Logging
    # Each script appends its own subfolder: Serverdecom, Servercreatie, ...
    # Log files older than LogRetentionDays are deleted automatically at startup.
    # -------------------------------------------------------------------------
    Logging = @{
        BasePath         = "c:\programdata\decoster.tech\scripting\logs"
        LogRetentionDays = 365
    }

    # -------------------------------------------------------------------------
    # Paths
    # JsonBasePath   : default folder for saving / opening JSON files,
    #                  used by the ManagedBookmarksCreator script.
    # -------------------------------------------------------------------------
    Paths = @{
        ScriptBasePath = "c:\programdata\decoster.tech\scripting\Powershell"
        JsonBasePath   = "c:\programdata\decoster.tech\scripting\Json"
    }

    # -------------------------------------------------------------------------
    # Company
    # Name         : "company" name used in generated script headers and registry paths.
    # Author       : default author name inserted in generated script headers.
    # -------------------------------------------------------------------------
    Company = @{
        Name         = "Decoster.tech"
        Author       = "Decoster Hans"
        RegistryBase = "HKLM:\Software\Decoster.tech\Scripting"
    }
}
