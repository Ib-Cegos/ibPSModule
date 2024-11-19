#########################################
# Script lançé à l'import du module ib #
#########################################
$eventSource = 'ibPowershellModule'
Write-Warning "Les commandes du module 'ib' ne sont pas prévues pour être utilisées en dehors de notre environnement de formation..."
if (-not ([System.Diagnostics.EventLog]::SourceExists($eventSource))) {
    Write-Warning "Création de la source d'évènements '$eventSource' et attente de sa disponibilité (possible indisponibilité des logs avant prochaine utilisation du module)."
    [System.Diagnostics.EventLog]::CreateEventSource($eventSource,'Application') }
  while (-not [System.Diagnostics.EventLog]::SourceExists($eventSource)) { Start-Sleep -Seconds 10 }