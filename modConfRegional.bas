Attribute VB_Name = "modConfRegional"
Option Explicit

Public Sub ponerConfiguracionRegional()

    SaveSettingString HKEY_CURRENT_USER, "Control Panel\International", "sDecimal", "."
    SaveSettingString HKEY_CURRENT_USER, "Control Panel\International", "sMonDecimalSep", "."
    SaveSettingString HKEY_CURRENT_USER, "Control Panel\International", "sMonThousandSep", ","
    SaveSettingString HKEY_CURRENT_USER, "Control Panel\International", "sThousand", ","
    SaveSettingString HKEY_CURRENT_USER, "Control Panel\International", "sNegativeSign", "-"
    
End Sub
