Outfile "DBxcel_Setup.exe"
InstallDir "$PROGRAMFILES\DBxcel"
RequestExecutionLevel admin

Section "Install"
    SetOutPath $INSTDIR
    File "dist\DBxcel.exe"
    CreateShortcut "$DESKTOP\DBxcel.lnk" "$INSTDIR\DBxcel.exe"
    CreateDirectory "$SMPROGRAMS\DBxcel"
    CreateShortcut "$SMPROGRAMS\DBxcel\DBxcel.lnk" "$INSTDIR\DBxcel.exe"

SectionEnd

Section "Uninstall"
    Delete "$INSTDIR\DBxcel.exe"
    Delete "$INSTDIR\Uninstall.exe"
    Delete "$INSTDIR\params.json"
    Delete "$INSTDIR\"
    RMDir "$INSTDIR"

    Delete "$DESKTOP\DBxcel.lnk"
    Delete "$SMPROGRAMS\DBxcel\DBxcel.lnk"
SectionEnd

Section -"Write Uninstaller"
    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd