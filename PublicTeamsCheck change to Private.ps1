
# Lister ut alle public teams i tenant
# Genererer to filer. En med oversikt over alle team, og en med eiere.
# Laget av John Riis 2024

#Kople opp til MSTeams
Connect-MicrosoftTeams

#Hent alle public Team
Write-Host "Getting public teams - It will take a while 30-60 minutes" -ForegroundColor Cyan
$publicteams = Get-Team | Where-Object {$_.Visibility -eq 'Public'}

#Export alle public team til Excel 
$publicteams | Export-Excel C:\temp\PublicTeamTFK.xlsx -TableStyle Medium2

#Export alle public team med eier til Excel
$counter = 0
$noteams=$publicteams.Count
foreach ($publicteam in $publicteams){
    $counter ++
    $whatever = Get-TeamUser -GroupId $publicteam.GroupId | Where-Object {$_.Role -eq 'Owner'} | select {$publicteam.DisplayName}, User | sort {$publicteam.DisplayName}
    Write-host "$counter / $noteams"
    $whatever | Export-Excel -Path "c:\temp\PublicTeamsWithOwnersTFK.xlsx" -TableStyle Medium2 -Append
}


## Set teams private

$pubteams = Import-Excel -Path "c:\temp\PublicTeamTFK.xlsx"
Write-Host -ForegroundColor Yellow "Found $($pubteams.Count) teams"
$i=0 
foreach($p in $pubteams){
$i++
    Write-Host -ForegroundColor Cyan "$i/$($pubteams.Count) -> $($p.groupId)"
    Set-Team -GroupId $($p.GroupId) -Visibility Private
}
