<#
.NOTES
    Created on:     6/6/2017
    Created by:     Andy Simmons (slick AD logic stolen from Rami Harasimowicz)
    Organization:   St. Luke's Health System
    Filename:       Get-OSInstalls.ps1
.SYNOPSIS
    Searches a domain for OS installs in a given date range.

.DESCRIPTION
    Initially written to track progress on our Windows 10 pilot.

.PARAMETER Domain
    Active Directory fully qualified domain name

.PARAMETER OperatingSystem
    Name of the Operating System tied to the AD computer object

.PARAMETER PilotStartDate
    Start date filter

.PARAMETER PilotEndDate
    End date filter

.PARAMETER OutFile
    File path to HTML report

.PARAMETER MailRecipient
    List of email addresses to receive the report

.PARAMETER MailServer
    SMTP smarthost

.PARAMETER MailSender
    Email address of the sender

.PARAMETER MailSubject
    Email message subject text

.PARAMETER SuppressEmail
    Do not send an email, just launch the report in a browser.
#>  
[CmdletBinding()]
param (
	[string]
	$Domain = 'sl1.stlukes-int.org',

	[string] 
	$OperatingSystem = 'Windows 10 Enterprise',

	[datetime]
	$PilotStartDate = [datetime]::Parse('05/01/2017'),

    [datetime]
    $PilotEndDate = [datetime]::Now,

    [IO.FileInfo]
    $OutFile = "${env:TEMP}\OSInstallReport.html",

    [string[]]
    $MailRecipient = @('simmonsa@slhs.org'),

    [string]
    $MailServer = 'mailgate.slhs.org',

    [string]
    $MailSender = 'spamgenerator@slhs.org',

    [string]
    $MailSubject = 'Windows 10 Pilot - Daily Report',

    [string]
    $MailBody = 'Report attached.',

    [switch]
    $SuppressEmail
)

$htmlHeader = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
    <head>
        <title>Windows 10 Pilot Machine Report</title>
        <style type="text/css">
        body {
            font-family: Calibri, Candara, Segoe, 'Segoe UI', Optima, Arial, sans-serif;
        }

        #report { width: 835px; }

        table {
            border-collapse: collapse;
            border: none;
            font: 11pt Calibri, Candara, Segoe, 'Segoe UI', Optima, Arial, sans-serif;
            color: black;
            margin-bottom: 10px;
        }

        table td {
            font-size: 12px;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
        }

        table th {
            font-size: 13px;
            font-weight: bold;
            padding-left: 0px;
            padding-right: 20px;
            text-align: left;
        }

        h2 { clear: both; font-size: 130%; color: #0077b3; }

        h3 {
            clear: both;
            font-size: 115%;
            color: #0077b3;
            margin-left: 20px;
            margin-top: 30px;
        }

        p { margin-left: 20px; font-size: 12px; }

        table.list { float: left; }

        table.list td:nth-child(1) {
            font-weight: bold;
            border-right: 1px grey solid;
            text-align: right;
        }

        table.list td:nth-child(2) { padding-left: 7px; }
        table tr:nth-child(even) td:nth-child(n) { background: #F2F2FF; }
        table tr:nth-child(odd) td:nth-child(n) { background: #FFFFFF; }
        table { margin-left: 20px; }
        </style>
    </head>
<body>
"@

$htmlFooter = @"
<hr/>
<p>Generated from: $($MyInvocation.MyCommand.Source)<br/>
Date: $(Get-Date)<br/>
User: $($env:UserDomain + "\" + $env:UserName)<br/>
Computer: $($env:COMPUTERNAME)</p>
</body>
</html>
"@

Write-Verbose "Retrieving all '$OperatingSystem' machines on domain: $Domain ..."

#Build a regEx to get the container from their DN
$dnParser = [regex]'(?<!\\),(.*)$'

#objects for getting meta data about AD properties (to know when they last changed)
$adContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain',$Domain)
$dc        = [DirectoryServices.ActiveDirectory.DomainController]::findOne($adContext)

$osMatches = Get-ADComputer -filter {OperatingSystem -eq $OperatingSystem} -properties * |
    Select-Object -Property Name, OperatingSystemVersion,
    @{ 
        Name       = "OperatingSystemModDate"
        Expression = {$dc.GetReplicationMetadata($_.distinguishedName).operatingsystem.LastOriginatingChangeTime}
    }, WhenCreated, LastLogonDate, 
    @{
        Name       = "Container"
        Expression = {$dnParser.Matches($_.DistinguishedName)[0].Groups[1].Value.ToString()}
    }, Enabled

Write-Verbose "Analyzing $($osMatches.Count) results ..."

$pilotFilter = [scriptblock] { 
    ($_.OperatingSystemModDate -ge $PilotStartDate) -and
    ($_.OperatingSystemModDate -le $PilotEndDate) -and
    ($_.Enabled)
}
$pilotMachines = $osMatches | Where-Object $pilotFilter | Sort-Object -Property OperatingSystemModDate -Descending

$installsByDate = $pilotMachines.OperatingSystemModDate.ToShortDateString() | 
    Group-Object -NoElement | Select-Object -Property @{
        Name       = 'Date'
        Expression = {$_.Name}
    },
    @{
        Name       = 'Installs'
        Expression = {$_.Count}
    }

$installsByOU = $pilotMachines | Group-Object -Property Container -NoElement | 
    Sort-Object -Property Count -Descending | Select-Object -Property Count, Name

Write-Verbose "Generating HTML report ..."

$shortStartDate = $PilotStartDate.ToShortDateString()
$shortEndDate   = $PilotEndDate.ToShortDateString()
$htmlBody       = "<h2>$OperatingSystem installations on '$Domain' ($shortStartDate - $shortEndDate)</h2>"

$htmlBody += $pilotMachines  | ConvertTo-Html -Fragment -PreContent "<h3>$($pilotMachines.Count) New $OperatingSystem Installs</h3>" | Out-String
$htmlBody += $installsByDate | ConvertTo-Html -Fragment -PreContent "<h3>$OperatingSystem Installations By Date</h3>" | Out-String
$htmlBody += $installsByOU   | ConvertTo-Html -Fragment -PreContent "<h3>$OperatingSystem Installations by OU</h3>" | Out-String

$htmlDocument = $htmlHeader + $htmlBody + $htmlFooter

Write-Verbose "Saving report file: $OutFile"
$htmlDocument | Out-File -FilePath $OutFile

if (Test-Path $OutFile)
{
    if ($SuppressEmail)
    {
        Write-Verbose 'Email disabled. Launching report in browser.'
        Invoke-Item $OutFile
    }
    else 
    {
        Write-Verbose "Emailing report: $($MailRecipient -join ', ')"
        $smmParams = @{
            To          = $MailRecipient
            From        = $MailSender
            SmtpServer  = $MailServer
            Subject     = $MailSubject
            Body        = $htmlDocument
            Attachments = $OutFile
            BodyAsHtml  = $true
        }
        Send-MailMessage @smmParams

        Write-Verbose "Deleting report file: $OutFile"
        Remove-Item $OutFile -Force
    }
}
