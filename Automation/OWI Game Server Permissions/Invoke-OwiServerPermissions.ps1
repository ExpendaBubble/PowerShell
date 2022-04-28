<#
    .SYNOPSIS
        Use this script to OWI server (Squad, Post Scriptum, etc.)
        permissions based on data from a Google Docs sheet
        and a json configuration file.

    .DESCRIPTION
        This script ingests user data from a Google Docs sheet
        and loops through server admin permission configurations
        in a json file.

        Its result can be manipulated by changing
        the Google Docs column names and their respective values,
        along with the corresponding property names and values
        of the json file.

        You can schedule this script to run periodically with
        Windows' Task Scheduler or a similar tool.

    .LINK
        https://expendablog.nl

    .LINK
        https://developers.google.com/identity/protocols/oauth2

    .NOTES
        Author  : ExpendaBubble
        Version : 2.2
        Date    : 4/28/2022

        Requirements:
          Servers.json
          Google OAuth 2.0 client credentials:
            ClientId
            ClientSecret
            RefreshToken
          Google Docs sheet with the following columns:
            Username
            SteamID
            .. and 1 column per game server type
#>

Begin {
    function Get-GAuthToken {
        [CmdletBinding()]
        param (
            [string]
            $ClientId,

            [string]
            $ClientSecret,

            [string]
            $RefreshToken
        )
        $GrantType = "refresh_token"
        $RequestUri = "https://accounts.google.com/o/oauth2/token"
        $gAuthBody = "refresh_token=$RefreshToken&client_id=$ClientId&client_secret=$ClientSecret&grant_type=$GrantType"
        $GAuthResponse = Invoke-RestMethod -Method Post -Uri $RequestUri -ContentType "application/x-www-form-urlencoded" -Body $gAuthBody -UseBasicParsing
        Return $GAuthResponse.access_token
    }

    function Get-GSheet {
        [CmdletBinding()]
        param (
            [string]
            $AuthToken,

            [string]
            $DocumentId
        )
        begin {
            $headers = @{
                "Authorization" = "Bearer $AuthToken"
                "Content-type"  = "application/json"
            }
        }
        process {
            Invoke-RestMethod -Uri "https://www.googleapis.com/drive/v3/files/$DocumentId/export?mimeType=text/csv" -Method Get -Headers $headers -UseBasicParsing
        }
    }

    ## Variables
    $ClientId = 'xxxxxxxxxxxxxxxxxxxxxx.apps.googleusercontent.com'
    $ClientSecret = 'xxxxxxxxxxxxxxxxxxxxx'
    $DocumentId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
    $RefreshToken = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
    $ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
    $JsonFile = 'Servers.json'

    ## Get Google Sheet data
    $GAuthToken = Get-GAuthToken -ClientId $ClientId -ClientSecret $ClientSecret -RefreshToken $RefreshToken
    $GSheet = Get-GSheet -AuthToken $GAuthToken -DocumentId $DocumentId | ConvertFrom-Csv

    ## Load configuration
    $json = Get-Content -Path $(Join-Path -Path $ScriptDir -ChildPath $JsonFile)
    $servers = $JSON | ConvertFrom-Json
}
Process {
    ForEach ($server in $servers) {
        $AdminsCfg = Join-Path -Path $server.ServerConfig -ChildPath 'Admins.cfg'
        If (Test-Path -Path $AdminsCfg) {
            ## Get group definitions
            $groups = New-Object -TypeName System.Collections.ArrayList
            ForEach ($group in $server.Groups) {
                $permissions = $group.permissions -join ','
                $entry = "Group=$($group.Name):$permissions"
                $groups.Add($entry) | Out-Null
            }
            ## Get admin definitions
            $AdminList = New-Object -TypeName System.Collections.ArrayList
            ForEach ($group in $server.Groups.Name) {
                ForEach ($user in $GSheet) {
                    If ($user.$($server.Game) -eq $group) {
                        $entry = "Admin=$($user.'SteamID'):$group // $($user.'Username')"
                        $AdminList.Add($entry) | Out-Null
                    }
                }
            }
            ## Get header sections
            $content = Get-Content -Path $AdminsCfg
            $LineBreaks = $content | Select-String -Pattern "^(/){10}"
            $Header1 = Get-Content -Path $AdminsCfg -TotalCount $LineBreaks.LineNumber[1] | Select-Object -Skip ($LineBreaks.LineNumber[0]-1)
            $Header2 = Get-Content -Path $AdminsCfg -TotalCount $LineBreaks.LineNumber[3] | Select-Object -Skip ($LineBreaks.LineNumber[2]-1)
            $Header3 = Get-Content -Path $AdminsCfg -TotalCount $LineBreaks.LineNumber[5] | Select-Object -Skip ($LineBreaks.LineNumber[4]-1)
            ## Construct new Admins.cfg
            $Header1 | Out-File -FilePath $AdminsCfg
            "" | Add-Content -Path $AdminsCfg
            $Header2 | Add-Content -Path $AdminsCfg
            "" | Add-Content -Path $AdminsCfg
            ForEach ($entry in $groups) {
                $entry | Add-Content -Path $AdminsCfg
            }
            "" | Add-Content -Path $AdminsCfg
            $Header3 | Add-Content -Path $AdminsCfg
            "" | Add-Content -Path $AdminsCfg
            ForEach ($entry in $AdminList) {
                $entry | Add-Content -Path $AdminsCfg
            }
        }
    }
}
