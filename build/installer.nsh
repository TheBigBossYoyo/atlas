!macro preInit
  ; Atlas branding fix: legacy MD Reader installs registered an InstallLocation
  ; under HKLM\Software\<appId> pointing at "C:\Program Files\MD Reader\Atlas\".
  ; Wipe those legacy install-location values so the NSIS InstallDirRegKey
  ; lookup falls through and $INSTDIR defaults to "$PROGRAMFILES64\Atlas".
  SetRegView 64
  DeleteRegValue HKLM "Software\${APP_GUID}" "InstallLocation"
  DeleteRegValue HKLM "Software\${APP_64_NAME}" "InstallLocation"
  DeleteRegKey HKLM "Software\MD Reader"
  DeleteRegKey HKCU "Software\MD Reader"
  SetRegView 32
  DeleteRegValue HKLM "Software\${APP_GUID}" "InstallLocation"
  DeleteRegKey HKLM "Software\MD Reader"
  DeleteRegKey HKCU "Software\MD Reader"
  SetRegView default

  ; Force the default install directory for fresh installs and for legacy
  ; upgrades whose old InstallLocation no longer exists on disk.
  StrCpy $INSTDIR "$PROGRAMFILES64\Atlas"
!macroend

!macro customInstall
  ; Best-effort cleanup for legacy MD Reader registration during Atlas upgrades
  DeleteRegKey HKLM "SOFTWARE\MDReader"
  DeleteRegValue HKLM "SOFTWARE\RegisteredApplications" "MD Reader"
  DeleteRegKey HKLM "SOFTWARE\Classes\MDReader.md"
  DeleteRegValue HKLM "SOFTWARE\Classes\.md\OpenWithProgids" "MDReader.md"
  DeleteRegValue HKLM "SOFTWARE\Classes\.markdown\OpenWithProgids" "MDReader.md"
  DeleteRegValue HKLM "SOFTWARE\Classes\.mdown\OpenWithProgids" "MDReader.md"
  DeleteRegValue HKLM "SOFTWARE\Classes\.mkd\OpenWithProgids" "MDReader.md"

  ; Register Atlas application capabilities
  WriteRegStr HKLM "SOFTWARE\RegisteredApplications" "Atlas" "SOFTWARE\Atlas\Capabilities"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities" "ApplicationName" "Atlas"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities" "ApplicationDescription" "Universal document viewer for Windows"

  ; Capabilities file associations
  ; --- Atlas.Document ---
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".docx" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".odt" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".rtf" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".md" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".markdown" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".mdown" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".mkd" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".txt" "Atlas.Document"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".log" "Atlas.Document"

  ; --- Atlas.Spreadsheet ---
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".xlsx" "Atlas.Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".xlsm" "Atlas.Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".ods" "Atlas.Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".csv" "Atlas.Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".tsv" "Atlas.Spreadsheet"

  ; --- Atlas.Presentation ---
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".pptx" "Atlas.Presentation"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".odp" "Atlas.Presentation"

  ; --- Atlas.Pdf ---
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".pdf" "Atlas.Pdf"

  ; --- Atlas.Code ---
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".json" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".yaml" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".yml" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".toml" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".ini" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".js" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".ts" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".tsx" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".jsx" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".py" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".rs" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".go" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".java" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".c" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".cpp" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".h" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".hpp" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".cs" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".rb" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".php" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".sh" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".ps1" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".sql" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".html" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".css" "Atlas.Code"
  WriteRegStr HKLM "SOFTWARE\Atlas\Capabilities\FileAssociations" ".scss" "Atlas.Code"

  ; ProgID definitions
  ; --- Atlas.Document ---
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Document" "" "Atlas Document"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Document" "FriendlyTypeName" "Atlas Document"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Document\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Document\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'

  ; --- Atlas.Spreadsheet ---
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Spreadsheet" "" "Atlas Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Spreadsheet" "FriendlyTypeName" "Atlas Spreadsheet"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Spreadsheet\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Spreadsheet\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'

  ; --- Atlas.Presentation ---
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Presentation" "" "Atlas Presentation"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Presentation" "FriendlyTypeName" "Atlas Presentation"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Presentation\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Presentation\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'

  ; --- Atlas.Pdf ---
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Pdf" "" "Atlas PDF"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Pdf" "FriendlyTypeName" "Atlas PDF"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Pdf\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Pdf\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'

  ; --- Atlas.Code ---
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Code" "" "Atlas Source Code"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Code" "FriendlyTypeName" "Atlas Source Code"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Code\DefaultIcon" "" "$INSTDIR\${APP_EXECUTABLE_FILENAME},0"
  WriteRegStr HKLM "SOFTWARE\Classes\Atlas.Code\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'

  ; Register Atlas under Applications for Open With
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\shell\open\command" "" '"$INSTDIR\${APP_EXECUTABLE_FILENAME}" "%1"'
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}" "FriendlyAppName" "Atlas"

  ; SupportedTypes and OpenWithProgids
  ; --- Atlas.Document ---
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".docx" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.docx\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".odt" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.odt\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".rtf" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.rtf\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".md" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.md\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".markdown" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.markdown\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".mdown" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.mdown\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".mkd" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.mkd\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".txt" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.txt\OpenWithProgids" "Atlas.Document" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".log" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.log\OpenWithProgids" "Atlas.Document" ""

  ; --- Atlas.Spreadsheet ---
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".xlsx" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.xlsx\OpenWithProgids" "Atlas.Spreadsheet" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".xlsm" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.xlsm\OpenWithProgids" "Atlas.Spreadsheet" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".ods" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.ods\OpenWithProgids" "Atlas.Spreadsheet" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".csv" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.csv\OpenWithProgids" "Atlas.Spreadsheet" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".tsv" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.tsv\OpenWithProgids" "Atlas.Spreadsheet" ""

  ; --- Atlas.Presentation ---
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".pptx" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.pptx\OpenWithProgids" "Atlas.Presentation" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".odp" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.odp\OpenWithProgids" "Atlas.Presentation" ""

  ; --- Atlas.Pdf ---
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".pdf" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.pdf\OpenWithProgids" "Atlas.Pdf" ""

  ; --- Atlas.Code ---
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".json" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.json\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".yaml" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.yaml\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".yml" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.yml\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".toml" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.toml\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".ini" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.ini\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".js" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.js\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".ts" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.ts\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".tsx" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.tsx\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".jsx" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.jsx\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".py" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.py\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".rs" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.rs\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".go" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.go\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".java" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.java\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".c" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.c\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".cpp" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.cpp\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".h" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.h\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".hpp" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.hpp\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".cs" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.cs\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".rb" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.rb\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".php" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.php\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".sh" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.sh\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".ps1" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.ps1\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".sql" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.sql\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".html" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.html\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".css" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.css\OpenWithProgids" "Atlas.Code" ""
  WriteRegStr HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}\SupportedTypes" ".scss" ""
  WriteRegStr HKLM "SOFTWARE\Classes\.scss\OpenWithProgids" "Atlas.Code" ""

  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)'
!macroend

!macro customUnInstall
  DeleteRegKey HKLM "SOFTWARE\Atlas"
  DeleteRegValue HKLM "SOFTWARE\RegisteredApplications" "Atlas"
  DeleteRegKey HKLM "SOFTWARE\Classes\Atlas.Document"
  DeleteRegKey HKLM "SOFTWARE\Classes\Atlas.Spreadsheet"
  DeleteRegKey HKLM "SOFTWARE\Classes\Atlas.Presentation"
  DeleteRegKey HKLM "SOFTWARE\Classes\Atlas.Pdf"
  DeleteRegKey HKLM "SOFTWARE\Classes\Atlas.Code"
  DeleteRegKey HKLM "SOFTWARE\Classes\Applications\${APP_EXECUTABLE_FILENAME}"

  ; Remove OpenWithProgids registrations
  ; --- Atlas.Document ---
  DeleteRegValue HKLM "SOFTWARE\Classes\.docx\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.odt\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.rtf\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.md\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.markdown\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.mdown\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.mkd\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.txt\OpenWithProgids" "Atlas.Document"
  DeleteRegValue HKLM "SOFTWARE\Classes\.log\OpenWithProgids" "Atlas.Document"

  ; --- Atlas.Spreadsheet ---
  DeleteRegValue HKLM "SOFTWARE\Classes\.xlsx\OpenWithProgids" "Atlas.Spreadsheet"
  DeleteRegValue HKLM "SOFTWARE\Classes\.xlsm\OpenWithProgids" "Atlas.Spreadsheet"
  DeleteRegValue HKLM "SOFTWARE\Classes\.ods\OpenWithProgids" "Atlas.Spreadsheet"
  DeleteRegValue HKLM "SOFTWARE\Classes\.csv\OpenWithProgids" "Atlas.Spreadsheet"
  DeleteRegValue HKLM "SOFTWARE\Classes\.tsv\OpenWithProgids" "Atlas.Spreadsheet"

  ; --- Atlas.Presentation ---
  DeleteRegValue HKLM "SOFTWARE\Classes\.pptx\OpenWithProgids" "Atlas.Presentation"
  DeleteRegValue HKLM "SOFTWARE\Classes\.odp\OpenWithProgids" "Atlas.Presentation"

  ; --- Atlas.Pdf ---
  DeleteRegValue HKLM "SOFTWARE\Classes\.pdf\OpenWithProgids" "Atlas.Pdf"

  ; --- Atlas.Code ---
  DeleteRegValue HKLM "SOFTWARE\Classes\.json\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.yaml\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.yml\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.toml\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.ini\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.js\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.ts\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.tsx\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.jsx\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.py\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.rs\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.go\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.java\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.c\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.cpp\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.h\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.hpp\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.cs\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.rb\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.php\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.sh\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.ps1\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.sql\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.html\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.css\OpenWithProgids" "Atlas.Code"
  DeleteRegValue HKLM "SOFTWARE\Classes\.scss\OpenWithProgids" "Atlas.Code"

  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, i 0, i 0)'
!macroend
