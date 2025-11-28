#URL des fichiers json de référence ib partagés sur OneDrive
$infoUrl = 'https://ibgroupecegos-my.sharepoint.com/:u:/g/personal/distanciel_ib_cegos_fr/'
$computersInfoUrl = "$($infoUrl)EZu4bAqgln5PlEwkMPtryEcB8UL-RJvUxig2GfHESWQjeQ?e=UMd3jn"
$sessionsInfoUrl = "$($infoUrl)EYrPnfJ16fFLp4QJsD9cwF4BvfayFcnqbnpVn8DZhghOOQ?e=NhPWEk"
$ibPassKey = (83,124,0,8,91,12,213,127,158,123,148,248,53,200,192,219,165,223,105,253,73,86,183,226,187,204,21,4,115,230,153,114)
$eventSource = 'ibPowershellModule'

function write-ibLog {
  param (
    [int32]$id = 0, [parameter(Position = 0)][string]$message, [string]$session = '', [string]$command, [switch]$warning, [switch]$error, [switch]$out, [switch]$stop )
  if ($warning) { $eventId = New-Object System.Diagnostics.EventInstance($id,1,2) }
  elseif ($error) { $eventId = New-Object System.Diagnostics.EventInstance($id,1,1) }
  else { $eventId = New-Object System.Diagnostics.EventInstance($id,1) }
  $eventObject = New-Object -TypeName System.Diagnostics.EventLog -Property @{Log = 'Application'; Source = $eventSource}
  if ($ibCommandLaunch -ne $PSCmdlet.MyInvocation.HistoryId) {
    $commandTrace = Get-PSCallStack|Where-Object {$_.Command -notlike '*ScriptBlock*' -and $_.InvocationInfo.CommandOrigin -like 'Runspace'}
    $commandTable=@("Lancement de la commande '$($commandTrace.Command)'","Arguments : $($commandTrace.Arguments)","Version du module ib : $((Get-Module -Name ib).Version)","Version Powershell : $($PSVersionTable.PSVersion)","Edition Powershell : $($PSVersionTable.PSEdition)")
    $eventObject.WriteEvent((New-Object System.Diagnostics.EventInstance(4,1)),$commandTable)
    $global:ibCommandLaunch = $PSCmdlet.MyInvocation.HistoryId }
  switch ($id) {
    1 { #La cible du raccourci est passée dans $message, son nom dans $command
        $eventContent = @{title = 'Création d''un raccourci sur le bureau.'; name = $command; shortcut = $message } }
    2 { #La commande passée est dans $command, son résultat (ou erreur) dans $message
        $eventContent = @{title = 'Lancement d''une commande.'; command = $command; result = $message } }
    3 { $eventContent = @{title = 'Installation automatique du client Teams.' } }
    default {
      if ($message -eq '') {$eventContent = @{title = 'Message d''évènement manquant.'} }
      else {$eventContent = @{title = $message } }}}
    if ($session -ne '') { $eventContent.add('session',$session) }
    $eventObject.WriteEvent($eventId,@($eventContent.title,(ConvertTo-Json -InputObject $eventContent)))
    if ($error) {
      if ($id -eq 2) { write-error -Message $message } else { Write-Error -Message $eventContent.title -ErrorId $id }}
    elseif ($warning) { Write-Warning -Message $eventContent.title }
    elseif ($out) { Write-Output -Message $eventContent.title }
    else { Write-debug $eventContent.title }
    if ($stop) { break } }

function get-ibLog {
  param (
  [int32]$id = -1,
  [string]$session = '')
  $eventObject = New-Object System.Diagnostics.EventLog
  $eventObject.Source = $eventSource
  $eventResult = @()
  foreach ($event in $eventObject.Entries) {
    if ($event.Source -eq $eventSource -and $event.replacementStrings.count -gt 1) {
      if ($event.eventId -eq 4) { $eventContent=New-Object -TypeName System.Object 
        $eventContent|Add-Member -NotePropertyName Title -NotePropertyValue $event.replacementStrings[0] }
      else { $eventContent = ($event.replacementStrings[1]|ConvertFrom-Json) }
      $eventContent|Add-Member -NotePropertyName Date -NotePropertyValue $event.TimeGenerated.Tostring('yyyyMMdd')
      if (($id -eq -1 -or ($event.eventId -eq $id)) -and ($session -eq '' -or $eventContent.session -eq $session)) {
        $eventResult += $eventContent } } }
  return ($eventResult) }

function optimize-ibComputer {
  <#
  .DESCRIPTION
  Cette commande est faite pour etre lancée au démarrage de la machine de formation et optimiser son fonctionnement pour la formation en cours.
    .PARAMETER force
  Si ce paramètre est utilisé, les commande de type 'oneTimeCommand' seront jouées, même si déjà trouvées dans les logs pour la session de formation de la machine
  #>
  param ([switch]$Force)
  wait-ibNetwork
  Write-Debug 'Vérification de la version du module.'
  $oldDebug = $global:DebugPreference
  $global:DebugPreference = 'silentlyContinue'
  if ( [version](Find-Module -Name ib).Version -gt (Get-Module -Name ib -ListAvailable|sort-object -property Version | select-object -Last 1).Version ) {
    $global:DebugPreference = $oldDebug
    write-ibLog 'Mise à jour du module.' -warning
    try { Remove-Module -Name ib -Force -ErrorAction stop}
    catch { write-ibLog -id 2 -command 'Remove-Module -Name ib -Force' -message $_.Exception.Message -error }
    try { Update-Module -Name ib -Force -ErrorAction stop}
    catch { write-ibLog -id 2 -command 'Update-Module -Name ib -Force' -message $_.Exception.Message -error }
    Import-Module -Name ib}
  $global:DebugPreference = $oldDebug
  get-ibComputerInfo -force
  if ($ibComputerInfo.currentSession) { 
    $todoMessage = "Optimisation de la machine pour la session en cours $($ibComputerInfo.currentSession)."
    $sessionToSetup = $ibComputerInfo.currentSession }
  elseif ($ibComputerInfo.nextSession) {
    $todoMessage = "Optimisation de la machine pour la prochaine session $($ibComputerInfo.nextSession)."
    $sessionToSetup = $ibComputerInfo.nextSession }
    else { write-ibLog 'Aucune session trouvée pour cette machine.' -warning -stop}
  if ($global:ibComputerInfo) {
    write-ibLog $todoMessage -session $sessionToSetup
    if ($global:ibComputerInfo.teamsMeeting -ne $null) { new-ibTeamsShortcut -meetingUrl $global:ibComputerInfo.teamsMeeting }
    else { new-ibTeamsShortcut }
    ForEach ($shortcut in $global:ibComputerInfo.shortcuts) {
      write-ibLog -id 1 -command $shortcut.name -message $shortcut.url -session $sessionToSetup
      New-Item -Path "$env:PUBLIC\Desktop" -Name "$($shortcut.name).url" -ItemType File -Value "[InternetShortcut]`nURL=$($shortcut.url)" -Force|Out-Null}
    ForEach ($command in $global:ibComputerInfo.commands) {
      try { 
        Invoke-Expression $command -OutVariable commandResult | Out-Null
        write-ibLog -id 2 -command $command -message (out-string -InputObject $commandResult) -session $sessionToSetup }
      catch { write-ibLog -id 2 -command $command -message $_.Exception.Message -error }}
    ForEach ($command in $global:ibComputerInfo.oneTimeCommands) {
      $runCommand = $true
      foreach ($oldCommand in (get-ibLog -id 2 -session $sessionToSetup)) {
        if ($oldCommand.command -eq $command) {
          write-ibLog "La commande '$command' a déjà été lançée (date de référence : $($oldCommand.Date))." -warning
          $runCommand = $false }}
      if ($runCommand -or $Force) {
        try {
          Invoke-Expression $command -OutVariable commandResult | Out-Null
          write-ibLog -id 2 -command $command -message (out-string -InputObject $commandResult) -session $sessionToSetup }
        catch { write-ibLog -id 2 -command $command -message $_.Exception.Message -error }}}
      }}

function install-ibScreenPaint {
  <#
  .DESCRIPTION
  Cette commande installe silencieusement l'outil "Screen Marker and recorder" depuis le store Windows.
    .PARAMETER autoStart
  Si ce paramètre est mentionné, un raccourci de démarrage automatique sera ajouté pour que ce programme se lance à l'ouverture de session utilisateur.
  #>
  param([switch]$autoStart)
  #Installer "Screen Marker and Recorder"
  write-ibLog -message 'Installation du logiciel "Screen Marker and Recorder" depuis le Microsoft Store.'
  winget install 9n0fw68w0dfw --accept-source-agreements --accept-package-agreements --silent --force --nowarn --disable-interactivity >> out-null
  if ($autoStart) {
    #Ajouter "Screen Marker and Recorder" au démarrage de Windows
    $ibShortcut = (New-Object -ComObject WScript.Shell).CreateShortcut("$([System.Environment]::GetFolderPath('startup'))\Screen Marker and Recorder.lnk")
    $ibShortcut.Arguments = $ibShortcut.TargetPath = "shell:AppsFolder\$((Get-StartApps "*screen Marker*").appId)"
    $ibShortcut.Save() }}

function install-ibZoomit {
  <#
  .DESCRIPTION
  Cette commande va utomatiquement télécharger, installer et configurer l'outil "ZoomIt" qui permet largement de dynamiser les démonstrations.
  .PARAMETER noAutostart
  Si ce paramètre est mentionné, ZoomIt ne sera pas configuré pour démarrer automatiquement.
  #>
  param([switch]$noAutoStart)
  function Set-RegistryValue {
    param ( [string]$Path='HKCU:\Software\Sysinternals\ZoomIt', [string]$Name, [string]$Value, [string]$PropertyType='Dword' )
  if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $PropertyType -Force | Out-Null }
  #Stop current ZoomIt running
  foreach ($Process in (Get-Process -Name "ZoomIt*" -ErrorAction silentlyContinue)) {
    Stop-Process -Id $Process.Id -Force
    Start-Sleep -Seconds 1}
  $DownloadURL = 'https://live.sysinternals.com/ZoomIt64.exe'
  $DestinationFile = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath ($DownloadURL.Split('/')[-1])
  Invoke-WebRequest -Uri $DownloadURL -OutFile $DestinationFile -ErrorAction Stop #Download ZoomIt in current user Documents folder
  Set-RegistryValue -Name 'EulaAccepted' -Value 1
  Set-RegistryValue -Name 'OptionsShown' -Value 1
  Set-RegistryValue -Name 'ToggleKey' -value 625          #Toggle Key to [Ctrl+F2]
  Set-RegistryValue -Name 'DrawToggleKey' -value 624      #Draw Toggle Key to [Ctrl+F1]
  Set-RegistryValue -Name 'RecordToggleKey' -value 626    #Record Toggle Key to [Ctrl+F3]
  if (!$noAutoStart) { Set-RegistryValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'ZoomIt' -Value $DestinationFile -PropertyType String } # Run On Startup via Registry
  Start-Process -FilePath $DestinationFile -NoNewWindow } #Start ZoomIt

function get-ibPassword {
  param ([parameter(Mandatory=$true)][string]$password)
  $key = 0..255 | Get-Random -Count 32 |ForEach-Object {[byte]$_}
  write-ibLog 'Utilisation de la commande get-ibPassword pour générer de nouvelles informations de sécurité' -warning
  Write-Output "En tête du module :`n`$ibPassKey = ($($key -join ','))"
  $SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force
  $encryptedPass = ($SecurePassword | ConvertFrom-SecureString -Key $key)
  Write-Output "Dans le fichier JSON :`n""password"" : ""$encryptedPass""" }

function wait-ibNetwork {
  if (!$global:ibNetOk) {
    Write-Debug 'Vérification du réseau.'
    $netCount = 0
    do {
      $netCount ++
      if ($netCount -eq 10) { write-ibLog 'Attente longue (10 essais) du réseau.' -warning}
      elseif ($netCount -eq 100) {write-ibLog 'Attente du réseau en vain.' -error -stop}
      write-debug 'Attente d''une réponse réseau.'
      $netTest = Test-NetConnection -InformationLevel Quiet }
    until ($netTest)
  $global:ibNetOk = $true}}

function get-ibComputersInfo {
  <#
  .DESCRIPTION
  Cette commande récupère les informations techniques/d'installation depuis les fichiers de réference ib (sur oneDrive).
  .PARAMETER force
  Si ce paramètre n'est pas mentionné, la machine pourra conserver les informations déjà récupérées depuis le réseau
  #>
  param ([switch]$force)
  if (!$global:ibComputersInfo -or !$global:ibSessionsInfo -or $force) {
    wait-ibNetwork
    if (!($global:ibComputersInfo = ((invoke-WebRequest -Uri "$computersInfoUrl&download=1" -UseBasicParsing).content|ConvertFrom-Json))) { write-ibLog 'Impossible de récuperer les informations des machines ib depuis le partage oneDrive' -error -stop}
    if (($global:ibSessionsInfo = ((invoke-WebRequest -Uri "$sessionsInfoUrl&download=1" -UseBasicParsing).content|ConvertFrom-Json))) {
      write-ibLog 'Récupération des informations de référence et de sessions.'
      foreach ($session in $global:ibSessionsInfo.Sessions.psObject.Properties) {
        if ($session.Value.salle -ne $null -and $global:ibComputersInfo.salles.($session.Value.salle) -eq $null) {
          write-debug "Ajout de la salle '$($session.value.salle)' pour la session '$($session.Name)'."
          $global:ibComputersInfo.Salles|add-member -NotePropertyName $session.Value.salle -NotePropertyValue @{sessions=@($session.Name)}}
        elseif ($session.Value.salle -ne $null) {
          write-debug "  Ajout de la session '$($session.Name)' à la salle '$($session.value.salle)'."
          if ($global:ibComputersInfo.Salles.($session.Value.salle).sessions -eq $null) { $global:ibComputersInfo.Salles.($session.Value.salle)|Add-Member -NotePropertyName sessions -NotePropertyValue @($session.Name)}
          else { $global:ibComputersInfo.Salles.($session.Value.salle) += $session.Name }}}}
    else { write-ibLog 'Impossible de récuperer les informations des sessions ib depuis le partage oneDrive' -error -stop}
    Write-Debug 'Stockage des informations de connexion'
    $global:ibAdminAccount = New-Object pscredential ($global:ibComputersInfo.Account.name, ($global:ibComputersInfo.Account.password|ConvertTo-SecureString -Key $ibPassKey))}}

function add-ibComputerInfo {
#Fonction facilitant le peuplement de la variable $global:ibComputerInfo (utilisation interne)
  param ($Names,$Value,[switch]$Add)
  foreach ($Name in $Names) {
    if ($Value.($Name) -ne $null) {
      if ($Add) {
        Write-Debug "  Ajout de valeur(s) '$Name' aux informations de la machine."
        if ($global:ibComputerInfo.($name) -eq $null) {$global:ibComputerInfo|Add-Member -NotePropertyName $Name -NotePropertyValue $Value.($Name)}
        else {$global:ibComputerInfo.($Name) += $value.($Name) }}
      elseif ($global:ibComputerInfo.($Name) -eq $null) {
        Write-Debug "  Ajout de la valeur '$Name' aux informations de la machine."
        $global:ibComputerInfo|Add-Member -NotePropertyName $Name -NotePropertyValue $Value.($Name) }}}}

function get-ibComputerInfo {
  <#
  .DESCRIPTION
  Cette commande récupere les informations techniques/d'installation sur la machine en cours depuis la réference ib.
  .PARAMETER force
  Si ce paramètre n'est pas mentionné, la machine pourra conserver les informations déjà recupérées depuis le réseau
  #>
  param ([switch]$force)
  if (!$global:ibComputersInfo -or $force) { get-ibComputersInfo -force}
  $ibComputersInfo = $global:ibComputersInfo
  $serialNumber = (Get-CimInstance Win32_BIOS).SerialNumber
  if ($global:ibComputerInfo = $ibComputersInfo.($serialNumber)) {
    Write-Debug 'Numéro de série de la machine trouvé.'
    if ($salleNumber = $global:ibComputerInfo.salle) {
      Write-Debug '  Référence de la salle trouvée dans les informations de machine.'
      if ($salle=$global:ibComputersInfo.Salles.$salleNumber) {
        write-debug '  Salle trouvée dans la référence ib.'
        add-ibComputerInfo -Names teamsMeeting -Value $salle
        add-ibComputerInfo -Names shortcuts,commands,oneTimeCommands,sessions -Value $salle -Add }}
      else { write-ibLog "  Salle '($($global:ibComputerInfo.salle))' trouvée sur la machine mais pas dans la référence." -warning}
    $currentDate = (Get-Date).ToString('yyyyMMdd')
    $nextSession = @{'name'='init';'date'='99999999'}
    foreach ($session in $global:ibComputerInfo.sessions) {
      $sessionMessage = "  La machine est inscrite pour la session '$session'"
      if ($ibSessionsInfo.Sessions.$session.debut -gt $currentDate) {
        if ($ibSessionsInfo.Sessions.$session.debut -lt $nextSession.date) { $nextSession = @{name=$session ; date = $ibSessionsInfo.Sessions.$session.debut}}
        $sessionMessage += ' qui n''a pas démarré.'}
      elseif ($ibSessionsInfo.Sessions.$session.fin -lt $currentDate) {$sessionMessage += ' qui est terminée.'}
      else {
        $sessionMessage += ' qui est en cours.'
        $currentSession = $session }
      write-debug $sessionMessage }
    if ($currentSession -ne $null) {
      write-ibLog "Inscription des informations de la session en cours ($currentSession)." -session $currentSession
      add-ibComputerInfo -Names stage,salle,formateur,debut,fin,teamsMeeting -Value $ibSessionsInfo.Sessions.$currentSession
      add-ibComputerInfo -Names shortcuts,commands,oneTimeCommands -Value $ibSessionsInfo.Sessions.$currentSession -Add
      $global:ibComputerInfo|Add-Member -NotePropertyName currentSession -NotePropertyValue $currentSession }
    elseif ($nextSession.name -ne 'init') {
      Write-ibLog "Inscription des informations de la prochaine session ($($nextSession.name))." -session $nextSession.name
      add-ibComputerInfo -Names stage,salle,formateur,debut,fin,teamsMeeting -Value $ibSessionsInfo.Sessions.($nextSession.name)
      add-ibComputerInfo -Names shortcuts,commands,oneTimeCommands -Value $ibSessionsInfo.Sessions.($nextSession.name) -Add
      $global:ibComputerInfo|Add-Member -NotePropertyName nextSession -NotePropertyValue $nextSession.name}
    if ($formateurTRG = $global:ibComputerInfo.formateur) {
      Write-Debug '  Formateur trouvé sur la machine.'
      if ($formateur = $global:ibComputersInfo.Formateurs.$formateurTRG) {
        write-debug '  Formateur trouvé dans la référence ib.'
        add-ibComputerInfo -Names shortcuts,commands,oneTimeCommands -value $formateur -Add }
      else { write-ibLog "  Formateur '($($global:ibComputerInfo.formateur))' trouvé sur la machine mais pas dans la référence." -warning}}
    if ($stageRef = $global:ibComputerInfo.stage) {
      Write-Debug '  Stage trouvé sur les informations de machine.'
      if ($stage = $global:ibComputersInfo.Stages.$stageRef) {
        write-debug '  Stage trouvé dans la référence ib.'
        add-ibComputerInfo -Names shortcuts,commands,oneTimeCommands -value $stage -Add }
      else { write-ibLog "  Stage '($($global:ibComputerInfo.stage))' trouvé sur la machine mais pas dans la référence." -warning}
        }}
else { write-ibLog "Numéro de série '$serialNumber' introuvable dans le fichier de références." -error -stop}}

function new-ibTeamsShortcut {
  <#
  .DESCRIPTION
  Cette commande Installe le nouveau client Teams et pose un raccourci pour la réunion sur le bureau le cas échéant.
  .PARAMETER meetingUrl
  Si ce paramètre est renseigné, un raccourci sera posé sur le bureau (de tous les utilisateurs de la machine) qui pointera sur l'adresse fournie et s'appelera 'Réunion Teams'.
  #>
  param( $meetingUrl = 'noUrl')
  # URL vers Teamsbootstrapper.exe depuis https://learn.microsoft.com/en-us/microsoftteams/new-teams-bulk-install-client
  write-ibLog -id 3
  $DownloadExeURL='https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
  $WebClient=New-Object -TypeName System.Net.WebClient
  write-debug '  Téléchargement du client Teams.'
  $WebClient.DownloadFile($DownloadExeURL,(Join-Path -Path $env:TEMP -ChildPath 'Teamsbootstrapper.exe'))
  $WebClient.Dispose()
  Write-Debug '  Installation du client Teams.'
  & "$($Env:TEMP)\Teamsbootstrapper.exe" -p >> $null
  if ($meetingUrl -ne 'noUrl') {
    write-ibLog -id 1 -command 'Réunion Teams' -message $meetingUrl
    New-Item -Path "$env:PUBLIC\Desktop" -Name 'Réunion Teams.url' -ItemType File -Value "[InternetShortcut]`nURL=$meetingUrl" -Force|out-null}}

function set-ibRemoteManagement {
  <#
  .DESCRIPTION
  Cette commande vérifie et/ou met en place la configuration nécessaire pour utiliser le service WinRM en local.
  #>
  Write-Debug 'Passage du profil des cartes réseau en "Privé" si elles sont en "Public".'
  Get-NetConnectionProfile|where-object {$_.NetworkCategory -notlike '*Domain*'}|Set-NetConnectionProfile -NetworkCategory Private
  Write-Debug 'Activation du Powershell Remoting.'
  enable-PSRemoting -Force|out-null
  try {$saveTrustedHosts=(Get-Item WSMan:\localhost\Client\TrustedHosts).value}
  catch {$saveTrustedHosts=''}
  Set-Item WSMan:\localhost\Client\TrustedHosts -value * -Force
  Set-ItemProperty -Path HKLM:\System\CurrentControlSet\Control\Lsa –Name ForceGuest –Value 0 -Force
  return $saveTrustedHosts }

function get-ibSubNet {
    #retourne un tableau des addresses IP du sous-réseau correspondant à l'adresse fournie (mais excluant celle-ci)
    param (
        [ipaddress]$ip,
        [ValidateRange(1,31)][int]$MaskBits)
    Write-Debug "Création d'un tableau de toutes les adresses IP du sous-réseau pour l'adresse $ip"
    $mask = ([Math]::Pow(2,$MaskBits)-1)*[Math]::Pow(2,(32-$MaskBits))
    $maskbytes = [BitConverter]::GetBytes([UInt32] $mask)
    $DottedMask = [IPAddress]((3..0 | ForEach-Object { [String] $maskbytes[$_] }) -join '.')
    write-debug "  Utilisation du masque de sous-réseau $DottedMask."
    [ipaddress]$subnetId = $ip.Address -band $DottedMask.Address
    $LowerBytes = [BitConverter]::GetBytes([UInt32] $subnetId.Address)
    [IPAddress]$broadcast = (0..3 | ForEach-object{$LowerBytes[$_] + ($maskbytes[(3-$_)] -bxor 255)}) -join '.'
    $subList = @()
    $current=$subnetId
    do {
        $curBytes = [BitConverter]::GetBytes([UInt32] $current.Address)
        [Array]::Reverse($curBytes)
        $nextBytes = [BitConverter]::GetBytes([UInt32]([bitconverter]::ToUInt32($curBytes,0) +1))
        [Array]::Reverse($nextBytes)
        $current = [ipaddress]$nextBytes
        if (($current -ne $broadcast) -and ($current -ne $ip)) { $subList+=$current.IPAddressToString }}  while ($current -ne $broadcast)
    return ($subList)}

function get-ibComputers {
  <#
  .DESCRIPTION
  Cette commande renvoit un tableau contenant les adresses IP de toutes les machines du sous-réseau de la machine depuis laquelle elle est lancée.
  #>

  #prérequis
  if (!(Get-Command Start-ThreadJob -errorAction silentlyContinue)) {
    Write-Debug 'Installation du module "ThreadJob".'
    Install-Module -Name Microsoft.Powershell.ThreadJob -Force -scope allUsers
    import-module -Name Microsoft.Powershell.ThreadJob}
  #Récuperation des informations sur le subnet
  $netIPConfig = get-NetIPConfiguration|Where-Object {$_.netAdapter.status -like 'up' -and $_.InterfaceDescription -notlike '*VirtualBox*' -and $_.InterfaceDescription -notlike '*vmware*' -and $_.InterfaceDescription -notlike '*virtual*'}
  $netIpAddress = $netIPConfig|Get-NetIPAddress -AddressFamily ipv4
  [System.Collections.ArrayList]$ipList = (get-ibSubNet -ip $netIpAddress.IPAddress -MaskBits $netIpAddress.PrefixLength)
  #Enlever le routeur de la liste !
  $ipList.Remove([ipaddress]($netIPConfig.ipv4defaultGateway.nextHop))
  write-debug 'Lancement des Ping des machines du sous-réseau.'
  $ipLoop = 0
  $ipLength = $ipList.Count
  ForEach ($ip in $ipList) {
    $ipLoop ++
    Write-Progress -Activity "Tentatives de connexion" -Status "Machine $ip." -PercentComplete (($ipLoop/$ipLength)*100)
    Start-ThreadJob -ScriptBlock {Test-Connection -ComputerName $using:ip -count 1 -buffersize 8 -Quiet} -ThrottleLimit 50 -Name $ip|Out-Null }
    Write-Progress -Activity "Tentatives de connexion" -Completed
  $ipLoop = 0
  $pingJobs = Get-Job
  $ipLength = $pingJobs.count
  foreach ($pingJob in $pingJobs) {
    $ipLoop ++
    Write-Progress -Activity "Attente des résultats" -Status "Adresse $($pingJob.name)." -PercentComplete (($ipLoop/$ipLength)*100)
    $pingResult = Receive-Job $pingJob -Wait -AutoRemoveJob
    #Enlever l'adresse de la liste si pas de réponse au ping
    if (!$pingResult) {$ipList.Remove($pingJob.name)}}
    Write-Progress -Activity "Attente des résultats" -Completed
  return($ipList)}

function invoke-ibNetCommand {
<#
.DESCRIPTION
Cette commande permet de lancer une commande sur toutes les machines accessibles sur le subnet.
.PARAMETER Command
Syntaxe complète de la commande à lancer (dans une chaine de caracteres)
.PARAMETER getCred
Ce switch permet de demander le nom et mot de passe de l'utilisateur à utiliser sur les machines distantes.
S'il est omis, l'utilisateur actuellement connecté sera utilisé.
.PARAMETER autoCred
Ce switch permet de spécifier automatiquement le nom et mot de passe de l'utilisateur à utiliser sur les machines distantes depuis la référence ib.
S'il est omis, l'utilisateur actuellement connecté sera utilisé.
.EXAMPLE
invoke-ibNetCommand -Command {$env:computername}
Va se connecter à chaque machine du réseau pour récupérer son nom d'ordinateur et l'afficher
#>
    param([parameter(Mandatory=$true,HelpMessage='Commande à lancer sur toutes les machines du sous-réseau')][string]$command,[switch]$getCred,[switch]$autoCred)
    if ($getCred) {
        $cred=Get-Credential -Message "Merci de saisir le nom et mot de passe du compte administrateur WinRM à utiliser pour éxecuter la commande '$Command'"
        if (-not $cred) {
            Write-Error "Arrêt suite à interruption utilisateur lors de la saisie du Nom/Mot de passe"
            break}}
    elseif ($autoCred) {
      if (!$global:ibComputersInfo) { get-ibComputersInfo }
      $cred = $global:ibAdminAccount
      $getCred = $true }
    $savedTrustedHosts = set-ibRemoteManagement
    foreach ($computer in get-ibComputers) {
        try {
            if ($getCred) {$commandOutput=(invoke-command -ComputerName $computer -ScriptBlock ([scriptBlock]::create($command)) -Credential $cred -ErrorAction Stop)}
            else {$commandOutput=(invoke-command -ComputerName $computer -ScriptBlock ([scriptBlock]::create($command)) -ErrorAction Stop)}
            if ($commandOutput) {
                Write-Host "[$computer] Résultat de la commande:" -ForegroundColor Green
                Write-host $commandOutput -ForegroundColor Gray }
            else { Write-Host "[$computer] Commande executée." -ForegroundColor Green}}
        catch {
            if ($_.Exception.message -ilike '*Access is denied*' -or $_.Exception.message -like '*Accès refusé*') { Write-Host "[$computer] Accès refusé." -ForegroundColor Red}
            else { Write-Host "[$computer] Erreur: $($_.Exception.message)" -ForegroundColor Red }}}
    Set-Item WSMan:\localhost\Client\TrustedHosts -value $savedTrustedHosts -Force}

function invoke-ibMute {
  <#
  .DESCRIPTION
  Cette commande permet de désactiver le son sur toutes les machines accessibles sur le subnet (dans la salle).
  Pour ce faire, elle utilise, un freeware (svcl.exe https://www.nirsoft.net/utils/sound_volume_command_line.html) qui sera uploadé dans le répertoire temporaire de chaque machine.
  Ne fonctionnera, à priori, que si un utilisateur est deja connecté sur la machine...
  .PARAMETER getCred
  Ce switch permet de demander le nom et mot de passe de l'utilisateur à utiliser sur les machines distantes.
  S'il est omis, l'utilisateur actuellement connecté sera utilisé.
  .PARAMETER autoCred
  Ce switch permet de spécifier automatiquement le nom et mot de passe de l'utilisateur à utiliser sur les machines distantes depuis la référence ib.
  S'il est omis, l'utilisateur actuellement connecté sera utilisé.
  #>
    param([switch]$getCred,[switch]$autoCred)
    if ($getCred) {
        $cred=Get-Credential -Message 'Merci de saisir le nom et mot de passe du compte administrateur WinRM à utiliser pour couper le son'
        if (-not $cred) {
            Write-Error "Arrêt suite à interruption utilisateur lors de la saisie du Nom/Mot de passe"
            break}}
    elseif ($autoCred) {
      if (!$global:ibComputersInfo) { get-ibComputersInfo }
      $cred = $global:ibAdminAccount
      $getCred = $true }
    $savedTrustedHosts = set-ibRemoteManagement
    Write-Debug 'Récupération de l''executable svcl.exe'
    $svclFile = (get-module -listAvailable ib).path
    $svclFile = $svclFile.substring(0,$svclFile.LastIndexOf('\')) + '\svcl.exe'
    Write-Debug 'Dépot de l''outil svcl et lancement sur les machines du sous-réseau.'
    foreach ($computer in get-ibComputers) {
        try {
            if ($getCred) {$session = New-PSSession -ComputerName $computer -Credential $cred -errorAction Stop}
            else {$session = New-PSSession -ComputerName $computer -errorAction Stop}
            if ($session) {
                $remoteTemp = (Invoke-Command -Session $session -ScriptBlock {$env:Temp})
                Copy-Item $svclFile -Destination "$remoteTemp\svcl.exe" -ToSession $session
                invoke-command -session $session -scriptBlock {set-location $using:remoteTemp;.\svcl.exe /mute (.\svcl.exe /scomma|ConvertFrom-Csv|where-object Default -eq render).name}
                Write-Host "[$computer] OK" -ForegroundColor Green}
            }
        catch {
            if ($_.Exception.message -ilike '*Access is denied*' -or $_.Exception.message -like '*Accès refusé*') { Write-Host "[$computer] Accès refusé." -ForegroundColor Red}
            else { Write-Host "[$computer] Erreur: $($_.Exception.message)" -ForegroundColor Red }}}
    Set-Item WSMan:\localhost\Client\TrustedHosts -value $savedTrustedHosts -Force}


function Reset-Ib {

$resultats = Get-BCDentry
$command2019 = "bcdedit /set {default} description Ib"
$command2019_1 = "bcdedit /set {default} device vhd=[D:]\Ib.vhdx"
$command2019_2 = "bcdedit /set {default} osdevice vhd=[D:]\Ib.vhdx"
$command_2019 = "bcdedit /set {default} description Ib2"
$command_2019_1 = "bcdedit /set {default} device vhd=[D:]\Ib2.vhdx"
$command_2019_2 = "bcdedit /set {default} osdevice vhd=[D:]\Ib2.vhdx"

# on recherche la descritpion du boot
foreach ($resultat in $resultats) {
    # Convertir l'objet en chaîne de caractères
    $resultat_str = $resultat | Out-String

    # Diviser le résultat en lignes
    $lignes = $resultat_str -split "`n"

    # Parcourir les lignes
    foreach ($ligne in $lignes) {
        # Vérifier si la ligne contient la description
        if ($ligne -match "description\s+(.+)") {
            # Extraire le contenu de la description
            $description = $matches[1].Trim()
            # Sortir de la boucle intérieure une fois la description trouvée
            break
        }
    }

    # Sortir de la boucle externe si la description a été trouvée
    if ($description) {
        break
    }
}


    if ($description -eq "Ib") {
        #Supprime le vhd
        Remove-Item -Path "D:\Ib2.vhdx" -Force
        #Re-créer le vhd
        diskpart /s C:\VHD\Office_2019.txt
        #passe en premier option de boot le vhd Office_2019
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_2019" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_2019_1" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_2019_2" -Verb RunAs -Wait
        #Affiche la pop-up de redémarrage puis redémarre au bout de 10sec
        #10 secondes d'affichage
        $Temps = 10
        #contenant de la pop-up
        $Message = "Votre machine doit redémarrer pour terminer sa configuration, merci de patienter..."
        #Titre de la pop_up
        $Titre = " Information importante "
        #Création d'un objet de type pop-up
        $Prompt = New-Object -ComObject WScript.Shell
        #Action d'afficher la pop-up
        $AffichageMessage = $Prompt.popup($Message, $Temps, $Titre, 16+0)
        #Redémarrer l'ordinateur
        Restart-Computer -force
   } 
	#Si la description de l'entry du boot actuel est "Office_2019"
	else {
        #Supprime le vhd 
        Remove-Item -Path "D:\Ib.vhdx" -Force
        #Re-créer le vhd
        diskpart /s C:\VHD\Office2019.txt
        #Passe en premier option de boot le vhd Office2019
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command2019" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command2019_1" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command2019_2" -Verb RunAs -Wait
        #Affiche la pop-up de redémarrage puis redémarre au bout de 10sec
        Show-MessageAndRestart
    }
} 

function Reset-Office365 {

$storeFile = "D:\Reset.txt"
$resultats = Get-BCDentry
$command365 = "bcdedit /set {default} description Office365"
$command365_1 = "bcdedit /set {default} device vhd=[D:]\Office365.vhdx"
$command365_2 = "bcdedit /set {default} osdevice vhd=[D:]\Office365.vhdx"
$command_365 = "bcdedit /set {default} description Office_365"
$command_365_1 = "bcdedit /set {default} device vhd=[D:]\Office_365.vhdx"
$command_365_2 = "bcdedit /set {default} osdevice vhd=[D:]\Office_365.vhdx"

foreach ($resultat in $resultats) {
    # Convertir l'objet en chaîne de caractères
    $resultat_str = $resultat | Out-String

    # Diviser le résultat en lignes
    $lignes = $resultat_str -split "`n"

    # Parcourir les lignes
    foreach ($ligne in $lignes) {
        # Vérifier si la ligne contient la description
        if ($ligne -match "description\s+(.+)") {
            # Extraire le contenu de la description
            $description = $matches[1].Trim()
            # Sortir de la boucle intérieure une fois la description trouvée
            break
        }
    }

    # Sortir de la boucle externe si la description a été trouvée
    if ($description) {
        break
    }
}
    
    if ($description -eq "Office365") {
        #Supprime le vhd
        Remove-Item -Path "D:\Office_365.vhdx" -Force
        #Re-créer le vhd
        diskpart /s C:\VHD\Office_365.txt
        #passe en premier option de boot le vhd Office_365
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_365" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_365_1" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command_365_2" -Verb RunAs -Wait
        #Affiche la pop-up de redémarrage puis redémarre au bout de 10sec
        #10 secondes d'affichage
        $Temps = 10
        #contenant de la pop-up
        $Message = "Votre machine doit redémarrer pour terminer sa configuration, merci de patienter..."
        #Titre de la pop_up
        $Titre = " Information importante "
        #Création d'un objet de type pop-up
        $Prompt = New-Object -ComObject WScript.Shell
        #Action d'afficher la pop-up
        $AffichageMessage = $Prompt.popup($Message, $Temps, $Titre, 16+0)
        #Redémarrer l'ordinateur
        Restart-Computer -force
   } 
	else {
        #Supprime le vhd 
        Remove-Item -Path "D:\Office365.vhdx" -Force
        #Re-créer le vhd
        diskpart /s C:\VHD\Office365.txt
        #Passe en premier option de boot le vhd Office365
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command365" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command365_1" -Verb RunAs -Wait
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $command365_2" -Verb RunAs -Wait
        #Affiche la pop-up de redémarrage puis redémarre au bout de 10sec
        #10 secondes d'affichage
        $Temps = 10
        #contenant de la pop-up
        $Message = "Votre machine doit redémarrer pour terminer sa configuration, merci de patienter..."
        #Titre de la pop_up
        $Titre = " Information importante "
        #Création d'un objet de type pop-up
        $Prompt = New-Object -ComObject WScript.Shell
        #Action d'afficher la pop-up
        $AffichageMessage = $Prompt.popup($Message, $Temps, $Titre, 16+0)
        #Redémarrer l'ordinateur
        Restart-Computer -force
    }
} 

function stop-ibNet {
<#
.DESCRIPTION
Cette commande permet d'arrêter toutes les machines du réseau local, en terminant par la machine sur laquelle est lançée la commande
.PARAMETER GetCred
Si ce switch n'est pas spécifié, l'identité de l'utilisateur actuellement connecté sera utilisée pour stopper les machines.
#>
param(
[switch]$GetCred)
if ($GetCred) {invoke-ibNetCommand -Command 'stop-Computer -Force' -GetCred}
elseif ($autoCred) { invoke-ibNetCommand 'stop-computer -Force' -autoCred }
else {invoke-ibNetCommand 'Stop-Computer -Force'}
  Stop-Computer -Force}

#######################
#  Gestion du module  #
#######################
New-Alias -Name oic -Value optimize-ibComputer -ErrorAction SilentlyContinue
New-Alias -Name optib -Value optimize-ibComputer -ErrorAction SilentlyContinue
New-Alias -Name ibPaint -value install-ibScreenPaint -errorAction SilentlyContinue
New-Alias -Name Resetib -Value Reset-Ib -ErrorAction SilentlyContinue
New-Alias -Name Reset365 -Value Reset-Office365 -ErrorAction SilentlyContinue
Export-moduleMember -Function invoke-ibMute,get-ibComputers,invoke-ibNetCommand,stop-ibNet,new-ibTeamsShortcut,get-ibComputerInfo,optimize-ibComputer,get-ibPassword,wait-ibNetwork,write-ibLog,get-ibLog,install-ibScreenPaint,install-ibZoomit,Reset-Office365,Reset-Ib -Alias oic,optib,ibPaint,ResetIb,Reset365
