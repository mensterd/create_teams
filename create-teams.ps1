#
# https://docs.microsoft.com/nl-nl/MicrosoftTeams/teams-powershell-overview
# https://docs.microsoft.com/en-us/powershell/module/teams/add-teamuser?view=teams-ps
# https://blog.jijitechnologies.com/create-teams-microsoft-teams-powershell
#
# https://365tips.be/installatie-van-de-microsoft-teams-powershell-module/
#
# Lay-out CSV file: TeamsName, Description, TeamsType, ChannelName[;], Template, Owners[;], Members[;]
#


function Create-Channel
{
    param
    (
        $channelName, $groupId
    )
    process{
        try
        {
            $teamChannels = $channelName -split ";"
            if($teamChannels)
            {
                for($i=0; $i -le ($teamChannels.count -1 ); $i++)
                {
                    New-TeamChannel -GroupId $groupId -DisplayName $teamChannel[$i]
                }
            }
        }
        catch
        {
            Write-Host "Fout bij creeren Teams-kanalen."
        }
    }
}



function Add-Users{
    param
    (
        $Users, $GroupId, $CurrentUserName, $Role
    )
    process
    {
        try
        {
            $team_users = $Users -Split ";"    
            if($team_users)
            {
                for($j=0; $j -le $team_users.count - 1; $j++)
                {
                    if($team_users[$j] -ne $CurrentUsername)
                    {
                        Add-TeamUser -GroupId $GroupId -User $team_user[$j] -Role $Role
                    }
                }
            }
        }
        catch
        {
            Out-Host "Fout bij toevoegen Users"    
        }
    }




}




function Create-NewTeam
{
    param(
        $ImportPath
    )
    Process
        {
            Import-Module MicrosoftTeams
            $cred = Get-Credential
            $username = $cred.UserName
            Connect-MicrosoftTeams -Credential $cred
            $teams = Import-Csv -Path $ImportPath

            foreach($team in $teams)
            {
                # Controleer of het Teams al bestaat. Gaat niet goed, te veel Teams, of rechten..
                # beter Exception handler schrijven
                # $getteam = Get-Team | Where-Object {$_.displayname -eq $team.TeamsName}
                try
                {
                    Write-Host "Team wordt gemaakt: " $team.TeamsName
                    $group = New-Team -DisplayName $team.TeamsName -Description $team.Description -Visibility $team.TeamsType
                }
                catch
                {
                    Write-Host "Fout bij het maken van het Team" $team.TeamsName
                }
                
                Write-Host "Kanalen worden gemaakt: " $team.ChannelName
                Create-Channel -ChannelName $team.ChannelName -GroupId $group.Id

                Write-Host "Teamleden toevoegen... "
                #Add-User -Users $team.Members -GroupId $group.Id -CurrentUserName $username -Role Member

                Write-Host "Team-eigenaren toevoegen..."
                #Add-User -Users $team.Owners -GroupId $group.Id -CurrentUserName $username -Role Owner

                Write-Host "Team creeren gereed: " $team.TeamsName
                $team = $null
            }

        Disconnect-MicrosoftTeams
        }

}