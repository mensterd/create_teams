#
# https://docs.microsoft.com/nl-nl/MicrosoftTeams/teams-powershell-overview
# https://docs.microsoft.com/en-us/powershell/module/teams/add-teamuser?view=teams-ps
# https://blog.jijitechnologies.com/create-teams-microsoft-teams-powershell
#
# https://365tips.be/installatie-van-de-microsoft-teams-powershell-module/
#
# Lay-out CSV file: TeamsName, Description, TeamsType, ChannelName[;], Template, Owners[;], Members[;]
#

param(
    [String]$ImportPath
)


$outPath = "./log.txt"


function Create-Channel
{
    param
    (
        $channelName, $groupId
    )
    process{
        try
        {
            $teamChannels = $channelName.Split(";")
            if($teamChannels)
            {
                for($i=0; $i -le ($teamChannels.Count -1 ); $i++)
                {
                    Write-Host $teamChannels[$i] "wordt ingesteld..."
                    Start-Sleep -Seconds 5

                    New-TeamChannel -GroupId $groupId -DisplayName $teamChannels[$i] -Description $teamChannels[$i] -MembershipType "Standard"
                    
                    Write-Host "Kanaal:" $teamChannels[$i] "Gereed."
                }
            }
        }
        catch
        {
            Write-Host "Fout bij creeren Teams-kanalen."
        }
    }
}



function Add-User{
    param
    (
        $Users, $GroupId, $CurrentUserName, $Role
    )
    process
    {
        try
        {
            $team_users = $Users.Split(";")
            if($team_users)
            {
                for($j=0; $j -le $team_users.count - 1; $j++)
                {
                    if($team_users[$j] -ne $CurrentUsername)
                    {
                        Add-TeamUser -GroupId $GroupId -User $team_users[$j] -Role $Role
                        #$team_users[$j], $Role| Out-File -FilePath $outPath
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
            Clear-Host
            Write-Host "Microsoft Teams module importeren"
            Import-Module MicrosoftTeams
            #$cred = Get-Credential
            #$username = $cred.UserName
            Write-Host "Verbinden met Office 365"
            $context = Connect-MicrosoftTeams
            $username = $context.Account.Id

            $teams = Import-Csv -Path $ImportPath

           
            foreach($team in $teams)
            {
                # Controleer of het Teams al bestaat. Gaat niet goed, te veel Teams, of rechten..
                # beter Exception handler schrijven
                # $getteam = Get-Team | Where-Object {$_.displayname -eq $team.TeamsName}
                try
                {
                    Write-Host "Team wordt gemaakt: " $team.TeamsName
                    $group = New-Team -DisplayName $team.TeamsName -Description $team.TeamsDescription -Visibility $team.TeamsType
                    Write-Host $team.TeamsName 'Gecreeerd.'
                }
                catch
                {
                    Write-Host "Fout bij het maken van het Team" $team.TeamsName
                }
                
                Start-Sleep -Seconds 5
                
                Write-Host "Kanalen worden gemaakt: " $team.ChannelName
                Create-Channel -ChannelName $team.ChannelName -GroupId $group.GroupId

                Write-Host "Teamleden toevoegen... "
                Add-User -Users $team.Members -GroupId $group.GroupId -CurrentUserName $username -Role Member

                #Write-Host "Team-eigenaren toevoegen..."
                #Add-User -Users $team.Owners -GroupId $group.Id -CurrentUserName $username -Role Owner

                Write-Host "Team creeren gereed: " $team.TeamsName
                $team = $null
            }

        Disconnect-MicrosoftTeams
        }

}


# Alternatief
# $import | foreach-object -Process{$group = new-team -displayname $_.TeamsName -Description $_.TeamsDescription -Visibility $_.Teamstype}

Create-NewTeam -ImportPath $ImportPath
