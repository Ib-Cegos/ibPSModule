# Manifeste de module pour le module « ib »
# Généré par : Renaud WANGLER
# Généré le : 30/10/2024

@{

RootModule = 'ib.psm1'
ModuleVersion = '1.2.6'
GUID = '8afa264f-71b6-4f7c-b16b-36463452660c'
Author = 'Renaud WANGLER'
CompanyName = 'ib'
Copyright = '(c) 2024 ib Cegos. Tous droits réservés.'
Description = 'Simplification des actions techniques pour les installations des machines en salle'
PowerShellVersion = '5.0'
ScriptsToProcess = @('.\moduleImport.ps1')
FunctionsToExport = @(
'get-ibComputers',
'invoke-ibNetCommand',
'invoke-ibMute',
'stop-ibNet',
'new-ibTeamsShortcut',
'get-ibComputerInfo',
'optimize-ibComputer',
'get-ibPassword',
'wait-ibNetwork',
'write-ibLog',
'get-ibLog',
'install-ibScreenPaint',
'install-ibZoomit',
'Reset-Office365',
'Reset-Ib',
'Invoke-InstaConfig',
'Write-Log',
'Ensure-Directory',
'Get-InstalledModuleVersion',
'Ensure-IbModuleUpToDate',
'Get-TodayYmd',
'Is-DateInRange',
'Get-ServiceTag',
'Load-JsonFile',
'Convert-SessionsToList',
'Get-RoomInfoFromRef',
'Get-StageActionsFromRef',
'Invoke-ResetAction',
'New-UrlShortcut',
'Write-ShortcutsToPublicDesktop',
'Ensure-MsalPs',
'Get-GraphAccessToken',
'Invoke-Graph',
'Download-GraphFileToLocal',
'Test-GraphFileExists',
'Put-GraphTextFile',
'Wait-Internet',
'Get-MarqueurSession',
'Remove-MarqueurOutDated',
'Remove-ExistingMarqueurs'
)
CmdletsToExport = @()
VariablesToExport = '*'
AliasesToExport = @('oic','optib','ibPaint','ResetIb','Reset365','InstaConfig')
FileList = @('svcl.exe')
PrivateData = @{
    PSData = @{
         Tags = @()
         LicenseUri = 'https://www.powershellgallery.com/policies/Terms'
         ProjectUri = 'https://github.com/ib-cegos/ibPSModule'
    }}}