# Manifeste de module pour le module « ib »
# Généré par : Renaud WANGLER
# Généré le : 30/10/2024

@{

RootModule = 'ib.psm1'
ModuleVersion = '1.1.9'
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
'Reset365',
'ResetIb')
CmdletsToExport = @()
VariablesToExport = '*'
AliasesToExport = @('oic','optib','ibPaint')
FileList = @('svcl.exe')
PrivateData = @{
    PSData = @{
         Tags = @()
         LicenseUri = 'https://www.powershellgallery.com/policies/Terms'
         ProjectUri = 'https://github.com/ib-cegos/ibPSModule'
    }}}