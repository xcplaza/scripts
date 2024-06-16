﻿<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 24.06.13.2304

#Get-AuthenticodeSignature -FilePath "C:\Scripts\HealthChecker.ps1" | ft -AutoSize
#.\HealthChecker.ps1 -Server "EX01-2016"
#.\HealthChecker.ps1 -BuildHtmlServersReport
#Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | .\HealthChecker.ps1; .\HealthChecker.ps1 -BuildHtmlServersReport -HtmlReportFile "ExchangeAllServersReport.html"; .\ExchangeAllServersReport.html[PS] C:\scripts>Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | .\HealthChecker.ps1; .\HealthChecker.ps1 -BuildHtmlServersReport -HtmlReportFile "ExchangeAllServersReport.html"; .\ExchangeAllServersReport.html
#.\HealthChecker.ps1 -BuildHtmlServersReport -HtmlReportFile "EX01-2016Report.html"


<#
.NOTES
	Name: HealthChecker.ps1
	Requires: Exchange Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
    Major Release History:
        4/20/2021  - Initial Public Release on CSS-Exchange.
        11/10/2020 - Initial Public Release of version 3.
        1/18/2017 - Initial Public Release of version 2.
        3/30/2015 - Initial Public Release.

.SYNOPSIS
	Checks the target Exchange server for various configuration recommendations from the Exchange product group.
.DESCRIPTION
	This script checks the Exchange server for various configuration recommendations outlined in the
	"Exchange 2013 Performance Recommendations" section on Microsoft Docs, found here:

	https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help

	Informational items are reported in Grey.  Settings found to match the recommendations are
	reported in Green.  Warnings are reported in yellow.  Settings that can cause performance
	problems are reported in red.  Please note that most of these recommendations only apply to latest Support Exchange versions.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER MailboxReport
	This optional parameter gives a report of the number of active and passive databases and
	mailboxes on the server.
.PARAMETER LoadBalancingReport
    This optional parameter will check the connection count of the Default Web Site for every server
    running Exchange 2013+ with the role in the org.  It then breaks down servers by percentage to
    give you an idea of how well the load is being balanced.
.PARAMETER ServerList
    Used with -LoadBalancingReport. A comma separated list of servers to operate against. Without
    this switch the report will use all 2013+ servers in the organization.
.PARAMETER SiteName
	Used with -LoadBalancingReport.  Specifies a site to pull  servers from instead of querying every server
    in the organization.
.PARAMETER XMLDirectoryPath
    Used in combination with BuildHtmlServersReport switch for the location of the HealthChecker XML files for servers
    which you want to be included in the report. Default location is the current directory.
.PARAMETER BuildHtmlServersReport
    Switch to enable the script to build the HTML report for all the servers XML results in the XMLDirectoryPath location.
.PARAMETER HtmlReportFile
    Name of the HTML output file from the BuildHtmlServersReport. Default is ExchangeAllServersReport.html
.PARAMETER DCCoreRatio
    Gathers the Exchange to DC/GC Core ratio and displays the results in the current site that the script is running in.
.PARAMETER AnalyzeDataOnly
    Switch to analyze the existing HealthChecker XML files. The results are displayed on the screen and an HTML report is generated.
.PARAMETER SkipVersionCheck
    No version check is performed when this switch is used.
.PARAMETER SaveDebugLog
    The debug log is kept even if the script is executed successfully.
.PARAMETER ScriptUpdateOnly
    Switch to check for the latest version of the script and perform an auto update. No elevated permissions or EMS are required.
.PARAMETER Verbose
	This optional parameter enables verbose logging.
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME1,SERVERNAME2
	Run against a list of servers
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME -MailboxReport -Verbose
	Run against a single remote Exchange server with verbose logging and mailbox report enabled.
.EXAMPLE
	Get-ExchangeServer | .\HealthChecker.ps1
	Run against all the Exchange servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport
    Run a load balancing report comparing all Exchange 2013+ servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport -ServerList EX01,EX02,EXS03
    Run a load balancing report comparing servers named EX01, EX02, and EX03.
.LINK
    https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help
    https://docs.microsoft.com/en-us/exchange/exchange-2013-virtualization-exchange-2013-help#requirements-for-hardware-virtualization
    https://docs.microsoft.com/en-us/exchange/plan-and-deploy/virtualization?view=exchserver-2019#requirements-for-hardware-virtualization
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "HealthChecker", SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ParameterSetName = "HealthChecker", HelpMessage = "Enter the list of servers names on which the script should execute against.")]
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, ParameterSetName = "MailboxReport", HelpMessage = "Enter the list of servers names on which the script should execute against.")]
    [string[]]$Server = ($env:COMPUTERNAME),

    [Parameter(Mandatory = $false, HelpMessage = "Provide the location of where the output files should go.")]
    [ValidateScript( {
            -not $_.ToString().EndsWith('\') -and (Test-Path $_)
        })]
    [string]$OutputFilePath = ".",

    [Parameter(Mandatory = $true, ParameterSetName = "MailboxReport", HelpMessage = "Enable the MailboxReport feature data collection against the server.")]
    [switch]$MailboxReport,

    [Parameter(Mandatory = $true, ParameterSetName = "LoadBalancingReport", HelpMessage = "Enable the LoadBalancingReport feature data collection.")]
    [Parameter(Mandatory = $true, ParameterSetName = "LoadBalancingReportBySite", HelpMessage = "Enable the LoadBalancingReport feature data collection.")]
    [switch]$LoadBalancingReport,

    [Alias("CASServerList")]
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport", HelpMessage = "Provide a list of servers to run against for the LoadBalancingReport.")]
    [string[]]$ServerList = $null,

    [Parameter(Mandatory = $true, ParameterSetName = "LoadBalancingReportBySite", HelpMessage = "Provide the AD SiteName to run the LoadBalancingReport against.")]
    [string]$SiteName = ([string]::Empty),

    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport", HelpMessage = "Provide the directory where the XML files are located at from previous runs of the Health Checker to Import the data from.")]
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly", HelpMessage = "Provide the directory where the XML files are located at from previous runs of the Health Checker to Import the data from.")]
    [Parameter(Mandatory = $false, ParameterSetName = "VulnerabilityReport", HelpMessage = "Provide the directory where the XML files are located at from previous runs of the Health Checker to Import the data from.")]
    [ValidateScript( {
            -not $_.ToString().EndsWith('\')
        })]
    [string]$XMLDirectoryPath = ".",

    [Parameter(Mandatory = $true, ParameterSetName = "HTMLReport", HelpMessage = "Enable the HTMLReport feature to run against the XML files from previous runs of the Health Checker script.")]
    [switch]$BuildHtmlServersReport,

    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport", HelpMessage = "Provide the name of the Report to be created.")]
    [string]$HtmlReportFile = "ExchangeAllServersReport-$((Get-Date).ToString("yyyyMMddHHmmss")).html",

    [Parameter(Mandatory = $true, ParameterSetName = "DCCoreReport", HelpMessage = "Enable the DCCoreReport feature data collection against the current server's AD Site.")]
    [switch]$DCCoreRatio,

    [Parameter(Mandatory = $true, ParameterSetName = "AnalyzeDataOnly", HelpMessage = "Enable to reprocess the data that was previously collected and display to the screen")]
    [switch]$AnalyzeDataOnly,

    [Parameter(Mandatory = $true, ParameterSetName = "VulnerabilityReport", HelpMessage = "Enable to collect data on the entire environment and report only the security vulnerabilities.")]
    [switch]$VulnerabilityReport,

    [Parameter(Mandatory = $false, ParameterSetName = "HealthChecker", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "DCCoreReport", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [Parameter(Mandatory = $false, ParameterSetName = "VulnerabilityReport", HelpMessage = "Skip over checking for a new updated version of the script.")]
    [switch]$SkipVersionCheck,

    [Parameter(Mandatory = $false, HelpMessage = "Always keep the debug log output at the end of the script.")]
    [switch]$SaveDebugLog,

    [Parameter(Mandatory = $true, ParameterSetName = "ScriptUpdateOnly", HelpMessage = "Only attempt to update the script.")]
    [switch]$ScriptUpdateOnly
)

begin {



function Add-AnalyzedResultInformation {
    [CmdletBinding()]
    param(
        # Main object that we are manipulating and adding entries to
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$AnalyzedInformation,

        # The value of the display entry
        [object]$Details,

        [object]$DisplayGroupingKey,

        [int]$DisplayCustomTabNumber = -1,

        [string]$DisplayWriteType = "Grey",

        # The name of the display entry
        [string]$Name,

        # Used for when the name might have a duplicate and we want it to be unique for logic outside of display
        [string]$CustomName,

        # Used for when the value might have a duplicate and we want it to be unique for logic outside of display
        [object]$CustomValue,

        # Used to display an Object in a table
        [object]$OutColumns,

        [ScriptBlock[]]$OutColumnsColorTests,

        [string]$TestingName,

        [object]$DisplayTestingValue,

        [string]$HtmlName,

        [string]$HtmlDetailsCustomValue = "",

        [bool]$AddDisplayResultsLineInfo = $true,

        [bool]$AddHtmlDetailRow = $true,

        [bool]$AddHtmlOverviewValues = $false,

        [bool]$AddHtmlActionRow = $false
        #[string]$ActionSettingClass = "",
        #[string]$ActionSettingValue,
        #[string]$ActionRecommendedDetailsClass = "",
        #[string]$ActionRecommendedDetailsValue,
        #[string]$ActionMoreInformationClass = "",
        #[string]$ActionMoreInformationValue,
    )
    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand): $name"
        function GetOutColumnsColorObject {
            param(
                [object[]]$OutColumns,
                [ScriptBlock[]]$OutColumnsColorTests,
                [string]$DefaultDisplayColor = ""
            )

            $returnValue = New-Object System.Collections.Generic.List[object]

            foreach ($obj in $OutColumns) {
                $objectValue = New-Object PSCustomObject
                foreach ($property in $obj.PSObject.Properties.Name) {
                    $displayColor = $DefaultDisplayColor
                    foreach ($func in $OutColumnsColorTests) {
                        $result = $func.Invoke($obj, $property)
                        if (-not [string]::IsNullOrEmpty($result)) {
                            $displayColor = $result[0]
                            break
                        }
                    }

                    $objectValue | Add-Member -MemberType NoteProperty -Name $property -Value ([PSCustomObject]@{
                            Value        = $obj.$property
                            DisplayColor = $displayColor
                        })
                }
                $returnValue.Add($objectValue)
            }
            return $returnValue
        }
    }
    process {
        Write-Verbose "Calling $($MyInvocation.MyCommand): $name"

        if ($AddDisplayResultsLineInfo) {
            if (!($AnalyzedInformation.DisplayResults.ContainsKey($DisplayGroupingKey))) {
                Write-Verbose "Adding Display Grouping Key: $($DisplayGroupingKey.Name)"
                [System.Collections.Generic.List[object]]$list = New-Object System.Collections.Generic.List[object]
                $AnalyzedInformation.DisplayResults.Add($DisplayGroupingKey, $list)
            }

            $lineInfo = [PSCustomObject]@{
                DisplayValue = [string]::Empty
                Name         = [string]::Empty
                TestingName  = [string]::Empty       # Used for pestering testing
                CustomName   = [string]::Empty       # Used for security vulnerability
                TabNumber    = 0
                TestingValue = $null                 # Used for pester testing down the road
                CustomValue  = $null                 # Used for security vulnerability
                OutColumns   = $null                 # Used for colorized format table option
                WriteType    = [string]::Empty
            }

            if ($null -ne $OutColumns) {
                $lineInfo.OutColumns = $OutColumns
                $lineInfo.WriteType = "OutColumns"
                $lineInfo.TestingValue = (GetOutColumnsColorObject -OutColumns $OutColumns.DisplayObject -OutColumnsColorTests $OutColumnsColorTests -DefaultDisplayColor "Grey")
                $lineInfo.TestingName = $TestingName
            } else {

                $lineInfo.DisplayValue = $Details
                $lineInfo.Name = $Name

                if ($DisplayCustomTabNumber -ne -1) {
                    $lineInfo.TabNumber = $DisplayCustomTabNumber
                } else {
                    $lineInfo.TabNumber = $DisplayGroupingKey.DefaultTabNumber
                }

                if ($null -ne $DisplayTestingValue) {
                    $lineInfo.TestingValue = $DisplayTestingValue
                } else {
                    $lineInfo.TestingValue = $Details
                }

                if ($null -ne $CustomValue) {
                    $lineInfo.CustomValue = $CustomValue
                } elseif ($null -ne $DisplayTestingValue) {
                    $lineInfo.CustomValue = $DisplayTestingValue
                } else {
                    $lineInfo.CustomValue = $Details
                }

                if (-not ([string]::IsNullOrEmpty($TestingName))) {
                    $lineInfo.TestingName = $TestingName
                } else {
                    $lineInfo.TestingName = $Name
                }

                if (-not ([string]::IsNullOrEmpty($CustomName))) {
                    $lineInfo.CustomName = $CustomName
                } elseif (-not ([string]::IsNullOrEmpty($TestingName))) {
                    $lineInfo.CustomName = $TestingName
                } else {
                    $lineInfo.CustomName = $Name
                }

                $lineInfo.WriteType = $DisplayWriteType
            }

            $AnalyzedInformation.DisplayResults[$DisplayGroupingKey].Add($lineInfo)
        }

        $htmlDetailRow = [PSCustomObject]@{
            Name        = [string]::Empty
            DetailValue = [string]::Empty
            TableValue  = $null
            Class       = [string]::Empty
        }

        if ($AddHtmlDetailRow) {
            if (!($analyzedResults.HtmlServerValues.ContainsKey("ServerDetails"))) {
                [System.Collections.Generic.List[object]]$list = New-Object System.Collections.Generic.List[object]
                $AnalyzedInformation.HtmlServerValues.Add("ServerDetails", $list)
            }

            $detailRow = $htmlDetailRow

            if ($displayWriteType -ne "Grey") {
                $detailRow.Class = $displayWriteType
            }

            if ([string]::IsNullOrEmpty($HtmlName)) {
                $detailRow.Name = $Name
            } else {
                $detailRow.Name = $HtmlName
            }

            if ($null -ne $OutColumns) {
                $detailRow.TableValue = (GetOutColumnsColorObject -OutColumns $OutColumns.DisplayObject -OutColumnsColorTests $OutColumnsColorTests)
            } elseif ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
                $detailRow.DetailValue = $Details
            } else {
                $detailRow.DetailValue = $HtmlDetailsCustomValue
            }

            $AnalyzedInformation.HtmlServerValues["ServerDetails"].Add($detailRow)
        }

        if ($AddHtmlOverviewValues) {
            if (!($analyzedResults.HtmlServerValues.ContainsKey("OverviewValues"))) {
                [System.Collections.Generic.List[object]]$list = New-Object System.Collections.Generic.List[object]
                $AnalyzedInformation.HtmlServerValues.Add("OverviewValues", $list)
            }

            $overviewValue = $htmlDetailRow

            if ($displayWriteType -ne "Grey") {
                $overviewValue.Class = $displayWriteType
            }

            if ([string]::IsNullOrEmpty($HtmlName)) {
                $overviewValue.Name = $Name
            } else {
                $overviewValue.Name = $HtmlName
            }

            if ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
                $overviewValue.DetailValue = $Details
            } else {
                $overviewValue.DetailValue = $HtmlDetailsCustomValue
            }

            $AnalyzedInformation.HtmlServerValues["OverviewValues"].Add($overviewValue)
        }

        if ($AddHtmlActionRow) {
            #TODO
        }
    }
}

function Get-DisplayResultsGroupingKey {
    param(
        [string]$Name,
        [bool]$DisplayGroupName = $true,
        [int]$DisplayOrder,
        [int]$DefaultTabNumber = 1
    )
    return [PSCustomObject]@{
        Name             = $Name
        DisplayGroupName = $DisplayGroupName
        DisplayOrder     = $DisplayOrder
        DefaultTabNumber = $DefaultTabNumber
    }
}



function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [ScriptBlock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [ScriptBlock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}

# Common method used to handle Invoke-Command within a script.
# Avoids using Invoke-Command when running locally on a server.
function Invoke-ScriptBlockHandler {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,

        [Parameter(Mandatory = $true)]
        [ScriptBlock]
        $ScriptBlock,

        [string]
        $ScriptBlockDescription,

        [object]
        $ArgumentList,

        [bool]
        $IncludeNoProxyServerOption,

        [ScriptBlock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $returnValue = $null
        $currentErrors = $null
    }
    process {

        if (-not([string]::IsNullOrEmpty($ScriptBlockDescription))) {
            Write-Verbose "Description: $ScriptBlockDescription"
        }

        try {

            if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {

                $params = @{
                    ComputerName = $ComputerName
                    ScriptBlock  = $ScriptBlock
                    ErrorAction  = "Stop"
                }

                if ($IncludeNoProxyServerOption) {
                    Write-Verbose "Including SessionOption"
                    $params.Add("SessionOption", (New-PSSessionOption -ProxyAccessType NoProxyServer))
                }

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Invoke-Command with argument list"
                    $params.Add("ArgumentList", $ArgumentList)
                } else {
                    Write-Verbose "Running Invoke-Command without argument list"
                }

                $returnValue = Invoke-Command @params
            } else {
                # Handle possible errors when executed locally.
                $currentErrors = $Error.Count

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Script Block Locally with argument list"

                    # if an object array type expect the result to be multiple parameters
                    if ($ArgumentList.GetType().Name -eq "Object[]") {
                        $returnValue = & $ScriptBlock @ArgumentList
                    } else {
                        $returnValue = & $ScriptBlock $ArgumentList
                    }
                } else {
                    Write-Verbose "Running Script Block Locally without argument list"
                    $returnValue = & $ScriptBlock
                }

                Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
            }
        } catch {
            Write-Verbose "Failed to run $($MyInvocation.MyCommand) - $ScriptBlockDescription"

            # Possible that locally we hit multiple errors prior to bailing out.
            if ($null -ne $currentErrors) {
                Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
            } else {
                Invoke-CatchActionError $CatchActionFunction
            }
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $returnValue
    }
}

function Get-VisualCRedistributableInstalledVersion {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $softwareList = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        $installedSoftware = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
            -ScriptBlock { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* } `
            -ScriptBlockDescription "Querying for software" `
            -CatchActionFunction $CatchActionFunction

        foreach ($software in $installedSoftware) {

            if ($software.PSObject.Properties.Name -contains "DisplayName" -and $software.DisplayName -like "Microsoft Visual C++ *") {
                Write-Verbose "Microsoft Visual C++ Found: $($software.DisplayName)"
                $softwareList.Add([PSCustomObject]@{
                        DisplayName       = $software.DisplayName
                        DisplayVersion    = $software.DisplayVersion
                        InstallDate       = $software.InstallDate
                        VersionIdentifier = $software.Version
                    })
            }
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $softwareList
    }
}

function Get-VisualCRedistributableInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year
    )

    if ($Year -eq 2012) {
        return [PSCustomObject]@{
            VersionNumber = 184610406
            DownloadUrl   = "https://www.microsoft.com/en-us/download/details.aspx?id=30679"
            DisplayName   = "Microsoft Visual C++ 2012*"
        }
    } else {
        return [PSCustomObject]@{
            VersionNumber = 201367256
            DownloadUrl   = "https://support.microsoft.com/en-us/topic/update-for-visual-c-2013-redistributable-package-d8ccd6a5-4e26-c290-517b-8da6cfdf4f10"
            DisplayName   = "Microsoft Visual C++ 2013*"
        }
    }
}

function Test-VisualCRedistributableInstalled {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object { $_.DisplayName -like $desired.DisplayName }))
}

function Test-VisualCRedistributableUpToDate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object {
                $_.DisplayName -like $desired.DisplayName -and $_.VersionIdentifier -eq $desired.VersionNumber
            }))
}

function Get-VisualCRedistributableLatest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return $Installed |
        Sort-Object VersionIdentifier -Descending |
        Where-Object { $_.DisplayName -like $desired.DisplayName } |
        Select-Object -First 1
}




function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-CatchActions {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $script:ErrorsExcluded += $CurrentError
    Write-Verbose "Error Excluded Count: $($Script:ErrorsExcluded.Count)"
    Write-Verbose "Error Count: $($Error.Count)"
    Write-VerboseErrorInformation $CurrentError
}

function Get-UnhandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -eq $handledError) {
                    [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index
                    }
                }
                $index++
            }
}

function Get-HandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -ne $handledError) {
                    [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index
                    }
                }
                $index++
            }
}

function Test-UnhandledErrorsOccurred {
    return $Error.Count -ne $Script:ErrorsExcluded.Count
}

function Invoke-ErrorCatchActionLoopFromIndex {
    [CmdletBinding()]
    param(
        [int]$StartIndex
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Start Index: $StartIndex Error Count: $($Error.Count)"

    if ($StartIndex -ne $Error.Count) {
        # Write the errors out in reverse in the order that they came in.
        $index = $Error.Count - $StartIndex - 1
        do {
            Invoke-CatchActions $Error[$index]
            $index--
        } while ($index -ge 0)
    }
}

function Invoke-ErrorMonitoring {
    # Always clear out the errors
    # setup variable to monitor errors that occurred
    $Error.Clear()
    $Script:ErrorsExcluded = @()
}

function Invoke-WriteDebugErrorsThatOccurred {

    function WriteErrorInformation {
        [CmdletBinding()]
        param(
            [object]$CurrentError
        )
        Write-VerboseErrorInformation $CurrentError
        Write-Verbose "-----------------------------------`r`n`r`n"
    }

    if ($Error.Count -gt 0) {
        Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

        Get-UnhandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }

        Write-Verbose "`r`n`r`nErrors that were handled"
        Get-HandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

function Invoke-AnalyzerKnownBuildIssues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [string]$CurrentBuild,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Extract for Pester Testing - Start
    function GetVersionFromString {
        param(
            [object]$VersionString
        )
        try {
            return New-Object System.Version $VersionString -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to convert '$VersionString' in $($MyInvocation.MyCommand)"
            Invoke-CatchActions
        }
    }

    function GetKnownIssueInformation {
        param(
            [string]$Name,
            [string]$Url
        )

        return [PSCustomObject]@{
            Name = $Name
            Url  = $Url
        }
    }

    function GetKnownIssueBuildInformation {
        param(
            [string]$BuildNumber,
            [string]$FixBuildNumber,
            [bool]$BuildBound = $true
        )

        return [PSCustomObject]@{
            BuildNumber    = $BuildNumber
            FixBuildNumber = $FixBuildNumber
            BuildBound     = $BuildBound
        }
    }

    function TestOnKnownBuildIssue {
        [CmdletBinding()]
        [OutputType("System.Boolean")]
        param(
            [object]$IssueBuildInformation,
            [version]$CurrentBuild
        )
        $knownIssue = GetVersionFromString $IssueBuildInformation.BuildNumber
        Write-Verbose "Testing Known Issue Build $knownIssue"

        if ($null -eq $knownIssue -or
            $CurrentBuild.Minor -ne $knownIssue.Minor) { return $false }

        $fixValueNull = [string]::IsNullOrEmpty($IssueBuildInformation.FixBuildNumber)
        if ($fixValueNull) {
            $resolvedBuild = GetVersionFromString "0.0.0.0"
        } else {
            $resolvedBuild = GetVersionFromString $IssueBuildInformation.FixBuildNumber
        }

        Write-Verbose "Testing against possible resolved build number $resolvedBuild"
        $buildBound = $IssueBuildInformation.BuildBound
        $withinBuildBoundRange = $CurrentBuild.Build -eq $knownIssue.Build
        $fixNeeded = $fixValueNull -or $CurrentBuild -lt $resolvedBuild
        Write-Verbose "BuildBound: $buildBound | WithinBuildBoundRage: $withinBuildBoundRange | FixNeeded: $fixNeeded"
        if ($CurrentBuild -ge $knownIssue) {
            if ($buildBound) {
                return $withinBuildBoundRange -and $fixNeeded
            } else {
                return $fixNeeded
            }
        }

        return $false
    }

    # Extract for Pester Testing - End

    function TestForKnownBuildIssues {
        param(
            [version]$CurrentVersion,
            [object[]]$KnownBuildIssuesToFixes,
            [object]$InformationUrl
        )
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Testing CurrentVersion $CurrentVersion"

        if ($null -eq $Script:CachedKnownIssues) {
            $Script:CachedKnownIssues = @()
        }

        foreach ($issue in $KnownBuildIssuesToFixes) {

            if ((TestOnKnownBuildIssue $issue $CurrentVersion) -and
                    (-not($Script:CachedKnownIssues.Contains($InformationUrl)))) {
                Write-Verbose "Known issue Match detected"
                if (-not ($Script:DisplayKnownIssueHeader)) {
                    $Script:DisplayKnownIssueHeader = $true

                    $params = $baseParams + @{
                        Name             = "Known Issue Detected"
                        Details          = "True"
                        DisplayWriteType = "Yellow"
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Details                = "This build has a known issue(s) which may or may not have been addressed. See the below link(s) for more information.`r`n"
                        DisplayWriteType       = "Yellow"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                $params = $baseParams + @{
                    Details                = "$($InformationUrl.Name):`r`n`t`t`t$($InformationUrl.Url)"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params

                if (-not ($Script:CachedKnownIssues.Contains($InformationUrl))) {
                    $Script:CachedKnownIssues += $InformationUrl
                    Write-Verbose "Added known issue to cache"
                }
            }
        }
    }

    try {
        $currentVersion = New-Object System.Version $CurrentBuild -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to set the current build to a version type object. $CurrentBuild"
        Invoke-CatchActions
    }

    try {
        Write-Verbose "Working on November 2021 Security Updates - OWA redirection"
        $infoParams = @{
            Name = "OWA redirection doesn't work after installing November 2021 security updates for Exchange Server 2019, 2016, or 2013"
            Url  = "https://support.microsoft.com/help/5008997"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.2.986.14" "15.2.986.15"),
                (GetKnownIssueBuildInformation "15.2.922.19" "15.2.922.20"),
                (GetKnownIssueBuildInformation "15.1.2375.17" "15.1.2375.18"),
                (GetKnownIssueBuildInformation "15.1.2308.20" "15.1.2308.21"),
                (GetKnownIssueBuildInformation "15.0.1497.26" "15.0.1497.28"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params

        Write-Verbose "Working on March 2022 Security Updates - MSExchangeServiceHost service may crash"
        $infoParams = @{
            Name = "Exchange Service Host service fails after installing March 2022 security update (KB5013118)"
            Url  = "https://support.microsoft.com/kb/5013118"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.2.1118.7" "15.2.1118.9"),
                (GetKnownIssueBuildInformation "15.2.986.22" "15.2.986.26"),
                (GetKnownIssueBuildInformation "15.2.922.27" $null),
                (GetKnownIssueBuildInformation "15.1.2507.6" "15.1.2507.9"),
                (GetKnownIssueBuildInformation "15.1.2375.24" "15.1.2375.28"),
                (GetKnownIssueBuildInformation "15.1.2308.27" $null),
                (GetKnownIssueBuildInformation "15.0.1497.33" "15.0.1497.36"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params

        Write-Verbose "Working on January 2023 Security Updates - Management issues after SerializedDataSigning is enabled on Exchange Server 2013"
        $infoParams = @{
            Name = "Management issues after SerializedDataSigning is enabled on Exchange Server 2013"
            Url  = "https://techcommunity.microsoft.com/t5/exchange-team-blog/released-january-2023-exchange-server-security-updates/ba-p/3711808"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.0.1497.45" "15.0.1497.47"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params

        Write-Verbose "Working on January 2023 Security Updates - Other known issues"
        $infoParams = @{
            Name = "Known Issues with Jan 2023 Security for Exchange 2016 and 2019"
            Url  = "https://techcommunity.microsoft.com/t5/exchange-team-blog/released-january-2023-exchange-server-security-updates/ba-p/3711808"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.1.2507.17" "15.1.2507.21"),
                (GetKnownIssueBuildInformation "15.2.986.37" "15.2.986.41"),
                (GetKnownIssueBuildInformation "15.2.1118.21" "15.2.1118.25"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params

        Write-Verbose "Working on February 2023 Security Updates"
        $infoParams = @{
            Name = "Known Issues with Feb 2023 Security Updates"
            Url  = "https://techcommunity.microsoft.com/t5/exchange-team-blog/released-february-2023-exchange-server-security-updates/ba-p/3741058"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.2.1118.25" "15.2.1118.26"),
                (GetKnownIssueBuildInformation "15.2.986.41" "15.2.986.42"),
                (GetKnownIssueBuildInformation "15.1.2507.21" "15.1.2507.23"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params

        Write-Verbose "Work on March 2024 Security Updates"
        $infoParams = @{
            Name = "Known Issues with Mar 2024 Security Updates"
            Url  = "https://support.microsoft.com/help/5037171"
        }
        $params = @{
            CurrentVersion          = $currentVersion
            KnownBuildIssuesToFixes = @((GetKnownIssueBuildInformation "15.2.1544.9" "15.2.1544.11"),
                (GetKnownIssueBuildInformation "15.2.1258.32" "15.2.1258.34"),
                (GetKnownIssueBuildInformation "15.1.2507.37", "15.1.2507.39"))
            InformationUrl          = (GetKnownIssueInformation @infoParams)
        }
        TestForKnownBuildIssues @params
    } catch {
        Write-Verbose "Failed to run TestForKnownBuildIssues"
        Invoke-CatchActions
    }
}
function Invoke-AnalyzerExchangeInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $keyExchangeInformation = Get-DisplayResultsGroupingKey -Name "Exchange Information"  -DisplayOrder $Order
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation
    $getWebServicesVirtualDirectory = $exchangeInformation.VirtualDirectories.GetWebServicesVirtualDirectory |
        Where-Object { $_.Name -eq "EWS (Default Web Site)" }
    $getWebServicesVirtualDirectoryBE = $exchangeInformation.VirtualDirectories.GetWebServicesVirtualDirectory |
        Where-Object { $_.Name -eq "EWS (Exchange Back End)" }

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keyExchangeInformation
    }

    $params = $baseParams + @{
        Name                  = "Name"
        Details               = $HealthServerObject.ServerName
        AddHtmlOverviewValues = $true
        HtmlName              = "Server Name"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Generation Time"
        Details               = $HealthServerObject.GenerationTime
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Version"
        Details               = $exchangeInformation.BuildInformation.VersionInformation.FriendlyName
        AddHtmlOverviewValues = $true
        HtmlName              = "Exchange Version"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "Build Number"
        Details = $exchangeInformation.BuildInformation.ExchangeSetup.FileVersion
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.BuildInformation.VersionInformation.Supported -eq $false) {
        $daysOld = ($date - $exchangeInformation.BuildInformation.VersionInformation.ReleaseDate).Days

        $params = $baseParams + @{
            Name                   = "Error"
            Details                = "Out of date Cumulative Update. Please upgrade to one of the two most recently released Cumulative Updates. Currently running on a build that is $daysOld days old."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
            TestingName            = "Out of Date"
            DisplayTestingValue    = $true
            HtmlName               = "Out of date"
        }
        Add-AnalyzedResultInformation @params
    }

    $extendedSupportDate = $exchangeInformation.BuildInformation.VersionInformation.ExtendedSupportDate
    if ($extendedSupportDate -le ([DateTime]::Now.AddYears(1))) {
        $displayWriteType = "Yellow"

        if ($extendedSupportDate -le ([DateTime]::Now.AddDays(178))) {
            $displayWriteType = "Red"
        }

        $displayValue = "$($exchangeInformation.BuildInformation.VersionInformation.ExtendedSupportDate.ToString("MMM dd, yyyy",
            [System.Globalization.CultureInfo]::CreateSpecificCulture("en-US"))) - Please note of the End Of Life date and plan to migrate soon."

        if ($extendedSupportDate -le ([DateTime]::Now)) {
            $displayValue = "Error: Your Exchange server reached end of life on " +
            "$($exchangeInformation.BuildInformation.VersionInformation.ExtendedSupportDate.ToString("MMM dd, yyyy",
                [System.Globalization.CultureInfo]::CreateSpecificCulture("en-US"))), and is no longer supported."
        }

        $params = $baseParams + @{
            Name                   = "End Of Life"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $exchangeInformation.BuildInformation.LocalBuildNumber) {
        $local = $exchangeInformation.BuildInformation.LocalBuildNumber
        $remote = [system.version]$exchangeInformation.BuildInformation.ExchangeSetup.FileVersion

        if ($local -ne $remote) {
            $params = $baseParams + @{
                Name                   = "Warning"
                Details                = "Running commands from a different version box can cause issues. Local Tools Server Version: $local"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    # If the ExSetup wasn't found, we need to report that.
    if ($exchangeInformation.BuildInformation.ExchangeSetup.FailedGetExSetup -eq $true) {
        $params = $baseParams + @{
            Name                   = "Warning"
            Details                = "Didn't detect ExSetup.exe on the server. This might mean that setup didn't complete correctly the last time it was run."
            DisplayCustomTabNumber = 2
            DisplayWriteType       = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $exchangeInformation.BuildInformation.KBsInstalled) {
        Add-AnalyzedResultInformation -Name "Exchange IU or Security Hotfix Detected" @baseParams
        $problemKbFound = $false
        $problemKbName = "KB5029388"

        foreach ($kb in $exchangeInformation.BuildInformation.KBsInstalled) {
            $params = $baseParams + @{
                Details                = $kb
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            if ($kb.Contains($problemKbName)) {
                $problemKbFound = $true
            }
        }

        if ($problemKbFound) {
            Write-Verbose "Found problem $problemKbName"
            if ($null -ne $HealthServerObject.OSInformation.BuildInformation.OperatingSystem.OSLanguage) {
                [int]$OSLanguageID = [int]($HealthServerObject.OSInformation.BuildInformation.OperatingSystem.OSLanguage)
                # https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-operatingsystem
                $englishLanguageIDs = @(9, 1033, 2057, 3081, 4105, 5129, 6153, 7177, 8201, 10249, 11273)
                if ($englishLanguageIDs.Contains($OSLanguageID)) {
                    Write-Verbose "OS is english language. No action required"
                } else {
                    Write-Verbose "Non english language code: $OSLanguageID"
                    $params = $baseParams + @{
                        Details                = "Error: August 2023 SU 1 Problem Detected. More Information: https://aka.ms/HC-Aug23SUIssue"
                        DisplayWriteType       = "Red"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            } else {
                Write-Verbose "Language Code is null"
            }
        }
    }

    # Both must be true. We need to be out of extended support AND no longer consider the latest SU the latest SU for this version to be secure.
    if ($extendedSupportDate -le ([DateTime]::Now) -and
        $exchangeInformation.BuildInformation.VersionInformation.LatestSU -eq $false) {
        $params = $baseParams + @{
            Details                = "Error: Your Exchange server is out of support and no longer receives SUs." +
            "`n`t`tIt is now considered persistently vulnerable and it should be decommissioned ASAP."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    } elseif ($exchangeInformation.BuildInformation.VersionInformation.LatestSU -eq $false) {
        $params = $baseParams + @{
            Details                = "Not on the latest SU. More Information: https://aka.ms/HC-ExBuilds"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $params = @{
        AnalyzeResults     = $AnalyzeResults
        DisplayGroupingKey = $keyExchangeInformation
        CurrentBuild       = $exchangeInformation.BuildInformation.ExchangeSetup.FileVersion
    }
    Invoke-AnalyzerKnownBuildIssues @params

    $params = $baseParams + @{
        Name                  = "Server Role"
        Details               = $exchangeInformation.BuildInformation.ServerRole
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.GetExchangeServer.IsMailboxServer -eq $true) {
        $dagName = [System.Convert]::ToString($exchangeInformation.GetMailboxServer.DatabaseAvailabilityGroup)
        if ([System.String]::IsNullOrWhiteSpace($dagName)) {
            $dagName = "Standalone Server"
        }
        $params = $baseParams + @{
            Name    = "DAG Name"
            Details = $dagName
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "AD Site"
        Details = ([System.Convert]::ToString(($exchangeInformation.GetExchangeServer.Site)).Split("/")[-1])
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) {

        Write-Verbose "Working on MRS Proxy Settings"
        $mrsProxyDetails = $getWebServicesVirtualDirectory.MRSProxyEnabled
        if ($getWebServicesVirtualDirectory.MRSProxyEnabled) {
            $mrsProxyDetails = "$mrsProxyDetails`n`r`t`tKeep MRS Proxy disabled if you do not plan to move mailboxes cross-forest or remote"
            $mrsProxyWriteType = "Yellow"
        } else {
            $mrsProxyWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name             = "MRS Proxy Enabled"
            Details          = $mrsProxyDetails
            DisplayWriteType = $mrsProxyWriteType
        }
        Add-AnalyzedResultInformation @params
    }

    if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) {
        Write-Verbose "Determining Server Group Membership"

        $params = $baseParams + @{
            Name             = "Exchange Server Membership"
            Details          = "Passed"
            DisplayWriteType = "Grey"
        }

        if ($null -ne $exchangeInformation.ComputerMembership -and
            $null -ne $HealthServerObject.OrganizationInformation.WellKnownSecurityGroups) {
            $localGroupList = $HealthServerObject.OrganizationInformation.WellKnownSecurityGroups |
                Where-Object { $_.WellKnownName -eq "Exchange Trusted Subsystem" }
            # By Default, I also have Managed Availability Servers and Exchange Install Domain Servers.
            # But not sure what issue they would cause if we don't have the server as a member, leaving out for now
            $adGroupList = $HealthServerObject.OrganizationInformation.WellKnownSecurityGroups |
                Where-Object { $_.WellKnownName -in @("Exchange Trusted Subsystem", "Exchange Servers") }
            $displayMissingGroups = New-Object System.Collections.Generic.List[string]

            foreach ($localGroup in $localGroupList) {
                if (($null -eq ($exchangeInformation.ComputerMembership.LocalGroupMember.SID | Where-Object { $_.ToString() -eq $localGroup.SID } ))) {
                    $displayMissingGroups.Add("$($localGroup.WellKnownName) - Local System Membership")
                }
            }

            foreach ($adGroup in $adGroupList) {
                if (($null -eq ($exchangeInformation.ComputerMembership.ADGroupMembership.SID | Where-Object { $_.ToString() -eq $adGroup.SID }))) {
                    $displayMissingGroups.Add("$($adGroup.WellKnownName) - AD Group Membership")
                }
            }

            if ($displayMissingGroups.Count -ge 1) {
                $params.DisplayWriteType = "Red"
                $params.Details = "Failed"
                Add-AnalyzedResultInformation @params

                foreach ($group in $displayMissingGroups) {
                    $params = $baseParams + @{
                        Details                = $group
                        TestingName            = $group
                        DisplayWriteType       = "Red"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                $params = $baseParams + @{
                    Details                = "More Information: https://aka.ms/HC-ServerMembership"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            } else {
                Add-AnalyzedResultInformation @params
            }
        } else {
            $params.DisplayWriteType = "Yellow"
            $params.Details = "Unknown - Wasn't able to get the Computer Membership information"
            Add-AnalyzedResultInformation @params
        }
    }

    if ($exchangeInformation.BuildInformation.MajorVersion -eq "Exchange2013" -and
        $exchangeInformation.GetExchangeServer.IsClientAccessServer -eq $true) {

        if ($null -ne $exchangeInformation.ApplicationPools -and
            $exchangeInformation.ApplicationPools.Count -gt 0) {
            $mapiFEAppPool = $exchangeInformation.ApplicationPools["MSExchangeMapiFrontEndAppPool"]
            [bool]$enabled = $mapiFEAppPool.GCServerEnabled
            [bool]$unknown = $mapiFEAppPool.GCUnknown
            $warning = [string]::Empty
            $displayWriteType = "Green"
            $displayValue = "Server"

            if ($hardwareInformation.TotalMemory -ge 21474836480 -and
                $enabled -eq $false) {
                $displayWriteType = "Red"
                $displayValue = "Workstation --- Error"
                $warning = "To Fix this issue go into the file MSExchangeMapiFrontEndAppPool_CLRConfig.config in the Exchange Bin directory and change the GCServer to true and recycle the MAPI Front End App Pool"
            } elseif ($unknown) {
                $displayValue = "Unknown --- Warning"
                $displayWriteType = "Yellow"
            } elseif (!($enabled)) {
                $displayWriteType = "Yellow"
                $displayValue = "Workstation --- Warning"
                $warning = "You could be seeing some GC issues within the Mapi Front End App Pool. However, you don't have enough memory installed on the system to recommend switching the GC mode by default without consulting a support professional."
            }

            $params = $baseParams + @{
                Name                   = "MAPI Front End App Pool GC Mode"
                Details                = $displayValue
                DisplayCustomTabNumber = 2
                DisplayWriteType       = $displayWriteType
            }
            Add-AnalyzedResultInformation @params
        } else {
            $warning = "Unable to determine MAPI Front End App Pool GC Mode status. This may be a temporary issue. You should try to re-run the script"
        }

        if ($warning -ne [string]::Empty) {
            $params = $baseParams + @{
                Details                = $warning
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    $internetProxy = $exchangeInformation.GetExchangeServer.InternetWebProxy

    $params = $baseParams + @{
        Name    = "Internet Web Proxy"
        Details = $internetProxy
    }

    if ([string]::IsNullOrEmpty($internetProxy)) {
        $params.Details = "Not Set"
    } elseif ($internetProxy.Scheme -ne "http") {
        <#
        We use the WebProxy class WebProxy(Uri, Boolean, String[]) constructor when running Set-ExchangeServer -InternetWebProxy,
        which throws an UriFormatException if the URI provided cannot be parsed.
        This is the case if it doesn't follow the scheme as per RFC 2396 (https://datatracker.ietf.org/doc/html/rfc2396#section-3.1).
        However, we sometimes see cases where customers have set an invalid proxy url that cannot be used by Exchange Server
        (e.g., https://proxy.contoso.local, ftp://proxy.contoso.local or even proxy.contoso.local).
        #>
        $params.Details = "$internetProxy is invalid as it must start with http://"
        $params.Add("DisplayWriteType", "Red")
    }
    Add-AnalyzedResultInformation @params

    if (-not ([string]::IsNullOrWhiteSpace($getWebServicesVirtualDirectory.InternalNLBBypassUrl))) {
        $params = $baseParams + @{
            Name             = "EWS Internal Bypass URL Set"
            Details          = "$($getWebServicesVirtualDirectory.InternalNLBBypassUrl) - Can cause issues after KB 5001779" +
            "`r`n`t`tThe Web Services Virtual Directory has a value set for InternalNLBBypassUrl which can cause problems with Exchange." +
            "`r`n`t`tSet the InternalNLBBypassUrl to NULL to correct this."
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $getWebServicesVirtualDirectoryBE -and
        $null -ne $getWebServicesVirtualDirectoryBE.InternalNLBBypassUrl) {
        Write-Verbose "Checking EWS Internal NLB Bypass URL for the BE"
        $expectedValue = "https://$($exchangeInformation.GetExchangeServer.Fqdn.ToString()):444/ews/exchange.asmx"

        if ($getWebServicesVirtualDirectoryBE.InternalNLBBypassUrl.ToString() -ne $expectedValue) {
            $params = $baseParams + @{
                Name             = "EWS Internal Bypass URL Incorrectly Set on BE"
                Details          = "Error: '$expectedValue' is the expected value for this." +
                "`r`n`t`tAnything other than the expected value, will result in connectivity issues."
                DisplayWriteType = "Red"
            }

            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Working on results from Test-ServiceHealth"
    $servicesNotRunning = $exchangeInformation.ExchangeServicesNotRunning

    if ($null -ne $servicesNotRunning -and
        $servicesNotRunning.Count -gt 0 ) {
        Add-AnalyzedResultInformation -Name "Services Not Running" @baseParams

        foreach ($stoppedService in $servicesNotRunning) {
            $params = $baseParams + @{
                Details                = $stoppedService
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Working on Exchange Dependent Services"
    if ($null -ne $exchangeInformation.DependentServices) {

        if ($exchangeInformation.DependentServices.Critical.Count -gt 0) {
            Write-Verbose "Critical Services found to be not running."
            Add-AnalyzedResultInformation -Name "Critical Services Not Running" @baseParams

            foreach ($service in $exchangeInformation.DependentServices.Critical) {
                $params = $baseParams + @{
                    Details                = "$($service.Name) - Status: $($service.Status) - StartType: $($service.StartType)"
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Red"
                    TestingName            = "Critical $($service.Name)"
                }
                Add-AnalyzedResultInformation @params
            }
        }
        if ($exchangeInformation.DependentServices.Common.Count -gt 0) {
            Write-Verbose "Common Services found to be not running."
            Add-AnalyzedResultInformation -Name "Common Services Not Running" @baseParams

            foreach ($service in $exchangeInformation.DependentServices.Common) {
                $params = $baseParams + @{
                    Details                = "$($service.Name) - Status: $($service.Status) - StartType: $($service.StartType)"
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Yellow"
                    TestingName            = "Common $($service.Name)"
                }
                Add-AnalyzedResultInformation @params
            }
        }

        if ($exchangeInformation.DependentServices.Misconfigured.Count -gt 0) {
            Write-Verbose "Misconfigured Services found."
            Add-AnalyzedResultInformation -Name "Misconfigured Services" @baseParams

            foreach ($service in $exchangeInformation.DependentServices.Misconfigured) {
                $params = $baseParams + @{
                    Details                = "$($service.Name) - Status: $($service.Status) - StartType: $($service.StartType) - CorrectStartType: $($service.CorrectStartType)"
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Yellow"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false -and
        $null -ne $exchangeInformation.ExtendedProtectionConfig) {
        $params = $baseParams + @{
            Name    = "Extended Protection Enabled (Any VDir)"
            Details = $exchangeInformation.ExtendedProtectionConfig.ExtendedProtectionConfigured
        }
        Add-AnalyzedResultInformation @params

        # If any directory has a higher than expected configuration, we need to throw a warning
        # This will be detected by SupportedExtendedProtection being set to false, as we are set higher than expected/recommended value you will likely run into issues of some kind
        # Skip over Default Web Site/Powershell if RequireSsl is not set.
        $notSupportedExtendedProtectionDirectories = $exchangeInformation.ExtendedProtectionConfig.ExtendedProtectionConfiguration |
            Where-Object { ($_.SupportedExtendedProtection -eq $false -and
                    $_.VirtualDirectoryName -ne "Default Web Site/Powershell") -or
                ($_.SupportedExtendedProtection -eq $false -and
                $_.VirtualDirectoryName -eq "Default Web Site/Powershell" -and
                $_.Configuration.SslSettings.RequireSsl -eq $true)
            }

        if ($null -ne $notSupportedExtendedProtectionDirectories) {
            foreach ($entry in $notSupportedExtendedProtectionDirectories) {
                $expectedValue = if ($entry.MitigationSupported -and $entry.MitigationEnabled) { "None" } else { $entry.ExpectedExtendedConfiguration }
                $params = $baseParams + @{
                    Details                = "$($entry.VirtualDirectoryName) - Current Value: '$($entry.ExtendedProtection)'   Expected Value: '$expectedValue'"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                    TestingName            = "EP - $($entry.VirtualDirectoryName)"
                    DisplayTestingValue    = ($entry.ExtendedProtection)
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details          = "`r`n`t`tThe current Extended Protection settings may cause issues with some clients types on $(if(@($notSupportedExtendedProtectionDirectories).Count -eq 1) { "this protocol."} else { "these protocols."})" +
                "`r`n`t`tIt is recommended to set the EP setting to the recommended value if you are having issues with that protocol." +
                "`r`n`t`tMore Information: https://aka.ms/ExchangeEPDoc"
                DisplayWriteType = "Yellow"
            }
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "All virtual directories are supported for the Extended Protection value."
        }
    }

    if ($null -ne $exchangeInformation.SettingOverrides) {

        $overridesDetected = $null -ne $exchangeInformation.SettingOverrides.SettingOverrides
        $params = $baseParams + @{
            Name    = "Setting Overrides Detected"
            Details = $overridesDetected
        }
        Add-AnalyzedResultInformation @params

        if ($overridesDetected) {
            $params = $baseParams + @{
                OutColumns = ([PSCustomObject]@{
                        DisplayObject = $exchangeInformation.SettingOverrides.SimpleSettingOverrides
                        IndentSpaces  = 12
                    })
                HtmlName   = "Setting Overrides"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($null -ne $exchangeInformation.EdgeTransportResourceThrottling) {
        try {
            # SystemMemory does not block mail flow.
            $resourceThrottling = ([xml]$exchangeInformation.EdgeTransportResourceThrottling).Diagnostics.Components.ResourceThrottling.ResourceTracker.ResourceMeter |
                Where-Object { $_.Resource -ne "SystemMemory" -and $_.CurrentResourceUse -ne "Low" }
        } catch {
            Invoke-CatchActions
        }

        if ($null -ne $resourceThrottling) {
            $resourceThrottlingList = @($resourceThrottling.Resource |
                    ForEach-Object {
                        $index = $_.IndexOf("[")
                        if ($index -eq -1) {
                            $_
                        } else {
                            $_.Substring(0, $index)
                        }
                    })
            $params = $baseParams + @{
                Name             = "Transport Back Pressure"
                Details          = "--ERROR-- The following resources are causing back pressure: $([string]::Join(", ", $resourceThrottlingList))"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Working on Exchange Server Maintenance"
    $serverMaintenance = $exchangeInformation.ServerMaintenance
    $getMailboxServer = $exchangeInformation.GetMailboxServer

    if (($serverMaintenance.InactiveComponents).Count -eq 0 -and
        ($null -eq $serverMaintenance.GetClusterNode -or
        $serverMaintenance.GetClusterNode.State -eq "Up") -and
        ($null -eq $getMailboxServer -or
            ($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -eq $false -and
        $getMailboxServer.DatabaseCopyAutoActivationPolicy.ToString() -eq "Unrestricted"))) {
        $params = $baseParams + @{
            Name             = "Exchange Server Maintenance"
            Details          = "Server is not in Maintenance Mode"
            DisplayWriteType = "Green"
        }
        Add-AnalyzedResultInformation @params
    } else {
        Add-AnalyzedResultInformation -Details "Exchange Server Maintenance" @baseParams

        if (($serverMaintenance.InactiveComponents).Count -ne 0) {
            foreach ($inactiveComponent in $serverMaintenance.InactiveComponents) {
                $params = $baseParams + @{
                    Name                   = "Component"
                    Details                = $inactiveComponent
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Red"
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details                = "For more information: https://aka.ms/HC-ServerComponentState"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
        }

        if ($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -or
            $getMailboxServer.DatabaseCopyAutoActivationPolicy -eq "Blocked") {
            $displayValue = "`r`n`t`tDatabaseCopyActivationDisabledAndMoveNow: $($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow) --- should be 'false'"
            $displayValue += "`r`n`t`tDatabaseCopyAutoActivationPolicy: $($getMailboxServer.DatabaseCopyAutoActivationPolicy) --- should be 'unrestricted'"

            $params = $baseParams + @{
                Name                   = "Database Copy Maintenance"
                Details                = $displayValue
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $serverMaintenance.GetClusterNode -and
            $serverMaintenance.GetClusterNode.State -ne "Up") {
            $params = $baseParams + @{
                Name                   = "Cluster Node"
                Details                = "'$($serverMaintenance.GetClusterNode.State)' --- should be 'Up'"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    }
}

function Invoke-AnalyzerHybridInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = Get-DisplayResultsGroupingKey -Name "Hybrid Information"  -DisplayOrder $Order
    }
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $getHybridConfiguration = $HealthServerObject.OrganizationInformation.GetHybridConfiguration

    if ($exchangeInformation.BuildInformation.VersionInformation.BuildVersion -ge "15.0.0.0" -and
        $null -ne $getHybridConfiguration) {

        $params = $baseParams + @{
            Name    = "Organization Hybrid Enabled"
            Details = "True"
        }
        Add-AnalyzedResultInformation @params

        if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.OnPremisesSmartHost))) {
            $onPremSmartHostDomain = ($getHybridConfiguration.OnPremisesSmartHost).ToString()
            $onPremSmartHostWriteType = "Grey"
        } else {
            $onPremSmartHostDomain = "No on-premises smart host domain configured for hybrid use"
            $onPremSmartHostWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "On-Premises Smart Host Domain"
            Details          = $onPremSmartHostDomain
            DisplayWriteType = $onPremSmartHostWriteType
        }
        Add-AnalyzedResultInformation @params

        if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.Domains))) {
            $domainsConfiguredForHybrid = $getHybridConfiguration.Domains
            $domainsConfiguredForHybridWriteType = "Grey"
        } else {
            $domainsConfiguredForHybridWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "Domain(s) configured for Hybrid use"
            DisplayWriteType = $domainsConfiguredForHybridWriteType
        }
        Add-AnalyzedResultInformation @params

        if ($domainsConfiguredForHybrid.Count -ge 1) {
            foreach ($domain in $domainsConfiguredForHybrid) {
                $params = $baseParams + @{
                    Details                = $domain
                    DisplayWriteType       = $domainsConfiguredForHybridWriteType
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            $params = $baseParams + @{
                Details                = "No domain configured for Hybrid use"
                DisplayWriteType       = $domainsConfiguredForHybridWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.EdgeTransportServers))) {
            Add-AnalyzedResultInformation -Name "Edge Transport Server(s)" @baseParams

            foreach ($edgeServer in $getHybridConfiguration.EdgeTransportServers) {
                $params = $baseParams + @{
                    Details                = $edgeServer
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.ReceivingTransportServers)) -or
            (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.SendingTransportServers)))) {
                $params = $baseParams + @{
                    Details                = "When configuring the EdgeTransportServers parameter, you must configure the ReceivingTransportServers and SendingTransportServers parameter values to null"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            Add-AnalyzedResultInformation -Name "Receiving Transport Server(s)" @baseParams

            if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.ReceivingTransportServers))) {
                foreach ($receivingTransportSrv in $getHybridConfiguration.ReceivingTransportServers) {
                    $params = $baseParams + @{
                        Details                = $receivingTransportSrv
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            } else {
                $params = $baseParams + @{
                    Details                = "No Receiving Transport Server configured for Hybrid use"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            Add-AnalyzedResultInformation -Name "Sending Transport Server(s)" @baseParams

            if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.SendingTransportServers))) {
                foreach ($sendingTransportSrv in $getHybridConfiguration.SendingTransportServers) {
                    $params = $baseParams + @{
                        Details                = $sendingTransportSrv
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            } else {
                $params = $baseParams + @{
                    Details                = "No Sending Transport Server configured for Hybrid use"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        }

        if ($getHybridConfiguration.ServiceInstance -eq 1) {
            $params = $baseParams + @{
                Name    = "Service Instance"
                Details = "Office 365 operated by 21Vianet"
            }
            Add-AnalyzedResultInformation @params
        } elseif ($getHybridConfiguration.ServiceInstance -ne 0) {
            $params = $baseParams + @{
                Name             = "Service Instance"
                Details          = $getHybridConfiguration.ServiceInstance
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details          = "You are using an invalid value. Please set this value to 0 (null) or re-run HCW"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.TlsCertificateName))) {
            $params = $baseParams + @{
                Name    = "TLS Certificate Name"
                Details = ($getHybridConfiguration.TlsCertificateName).ToString()
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name             = "TLS Certificate Name"
                Details          = "No valid certificate found"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        Add-AnalyzedResultInformation -Name "Feature(s) enabled for Hybrid use" @baseParams

        if (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.Features))) {
            foreach ($feature in $getHybridConfiguration.Features) {
                $params = $baseParams + @{
                    Details                = $feature
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            $params = $baseParams + @{
                Details                = "No feature(s) enabled for Hybrid use"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $exchangeInformation.ExchangeConnectors) {
            foreach ($connector in $exchangeInformation.ExchangeConnectors) {
                $cloudConnectorWriteType = "Yellow"
                if (($connector.TransportRole -ne "HubTransport") -and
                    ($connector.CloudEnabled -eq $true)) {

                    $params = $baseParams + @{
                        Details          = "`r"
                        AddHtmlDetailRow = $false
                    }
                    Add-AnalyzedResultInformation @params

                    if (($connector.CertificateDetails.CertificateMatchDetected) -and
                        ($connector.CertificateDetails.GoodTlsCertificateSyntax)) {
                        $cloudConnectorWriteType = "Green"
                    }

                    $params = $baseParams + @{
                        Name    = "Connector Name"
                        Details = $connector.Name
                    }
                    Add-AnalyzedResultInformation @params

                    $cloudConnectorEnabledWriteType = "Gray"
                    if ($connector.Enabled -eq $false) {
                        $cloudConnectorEnabledWriteType = "Yellow"
                    }

                    $params = $baseParams + @{
                        Name             = "Connector Enabled"
                        Details          = $connector.Enabled
                        DisplayWriteType = $cloudConnectorEnabledWriteType
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Name    = "Cloud Mail Enabled"
                        Details = $connector.CloudEnabled
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Name    = "Connector Type"
                        Details = $connector.ConnectorType
                    }
                    Add-AnalyzedResultInformation @params

                    if (($connector.ConnectorType -eq "Send") -and
                        ($null -ne $connector.TlsAuthLevel)) {
                        # Check if send connector is configured to relay mails to the internet via M365
                        switch ($connector) {
                            { ($_.SmartHosts -like "*.mail.protection.outlook.com") } {
                                $smartHostsPointToExo = $true
                            }
                            { ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_.AddressSpaces)) } {
                                $addressSpacesContainsWildcard = $true
                            }
                        }

                        if (($smartHostsPointToExo -eq $false) -or
                            ($addressSpacesContainsWildcard -eq $false)) {

                            $tlsAuthLevelWriteType = "Gray"
                            if ($connector.TlsAuthLevel -eq "DomainValidation") {
                                # DomainValidation: In addition to channel encryption and certificate validation,
                                # the Send connector also verifies that the FQDN of the target certificate matches
                                # the domain specified in the TlsDomain parameter. If no domain is specified in the TlsDomain parameter,
                                # the FQDN on the certificate is compared with the recipient's domain.
                                $tlsAuthLevelWriteType = "Green"
                                if ($null -eq $connector.TlsDomain) {
                                    $tlsAuthLevelWriteType = "Yellow"
                                    $tlsAuthLevelAdditionalInfo = "'TlsDomain' is empty which means that the FQDN of the certificate is compared with the recipient's domain.`r`n`t`tMore information: https://aka.ms/HC-HybridConnector"
                                }
                            }

                            $params = $baseParams + @{
                                Name             = "TlsAuthLevel"
                                Details          = $connector.TlsAuthLevel
                                DisplayWriteType = $tlsAuthLevelWriteType
                            }
                            Add-AnalyzedResultInformation @params

                            if ($null -ne $tlsAuthLevelAdditionalInfo) {
                                $params = $baseParams + @{
                                    Details                = $tlsAuthLevelAdditionalInfo
                                    DisplayWriteType       = $tlsAuthLevelWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }
                        }
                    }

                    if (($smartHostsPointToExo) -and
                        ($addressSpacesContainsWildcard)) {
                        # Seems like this send connector is configured to relay mails to the internet via M365 - skipping some checks
                        # https://docs.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-route-mail#2-set-up-your-email-server-to-relay-mail-to-the-internet-via-microsoft-365-or-office-365
                        $params = $baseParams + @{
                            Name    = "Relay Internet Mails via M365"
                            Details = $true
                        }
                        Add-AnalyzedResultInformation @params

                        switch ($connector.TlsAuthLevel) {
                            "EncryptionOnly" {
                                $tlsAuthLevelM365RelayWriteType = "Yellow"
                                break
                            }
                            "CertificateValidation" {
                                $tlsAuthLevelM365RelayWriteType = "Green"
                                break
                            }
                            "DomainValidation" {
                                if ($null -eq $connector.TlsDomain) {
                                    $tlsAuthLevelM365RelayWriteType = "Red"
                                } else {
                                    $tlsAuthLevelM365RelayWriteType = "Green"
                                }
                                break
                            }
                            default { $tlsAuthLevelM365RelayWriteType = "Red" }
                        }

                        $params = $baseParams + @{
                            Name             = "TlsAuthLevel"
                            Details          = $connector.TlsAuthLevel
                            DisplayWriteType = $tlsAuthLevelM365RelayWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($tlsAuthLevelM365RelayWriteType -ne "Green") {
                            $params = $baseParams + @{
                                Details                = "'TlsAuthLevel' should be set to 'CertificateValidation'. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $tlsAuthLevelM365RelayWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        }

                        $requireTlsWriteType = "Red"
                        if ($connector.RequireTLS) {
                            $requireTlsWriteType = "Green"
                        }

                        $params = $baseParams + @{
                            Name             = "RequireTls Enabled"
                            Details          = $connector.RequireTLS
                            DisplayWriteType = $requireTlsWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($requireTlsWriteType -eq "Red") {
                            $params = $baseParams + @{
                                Details                = "'RequireTLS' must be set to 'true' to ensure a working mail flow. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $requireTlsWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        }
                    } else {
                        $cloudConnectorTlsCertificateName = "Not set"
                        if ($null -ne $connector.CertificateDetails.TlsCertificateName) {
                            $cloudConnectorTlsCertificateName = $connector.CertificateDetails.TlsCertificateName
                        }

                        $params = $baseParams + @{
                            Name             = "TlsCertificateName"
                            Details          = $cloudConnectorTlsCertificateName
                            DisplayWriteType = $cloudConnectorWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        $params = $baseParams + @{
                            Name             = "Certificate Found On Server"
                            Details          = $connector.CertificateDetails.CertificateMatchDetected
                            DisplayWriteType = $cloudConnectorWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($connector.CertificateDetails.TlsCertificateNameStatus -eq "TlsCertificateNameEmpty") {
                            $params = $baseParams + @{
                                Details                = "There is no 'TlsCertificateName' configured for this cloud mail enabled connector.`r`n`t`tThis will cause mail flow issues in hybrid scenarios. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $cloudConnectorWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        } elseif ($connector.CertificateDetails.CertificateMatchDetected -eq $false) {
                            $params = $baseParams + @{
                                Details                = "The configured 'TlsCertificateName' was not found on the server.`r`n`t`tThis may cause mail flow issues. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $cloudConnectorWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        } else {
                            Add-AnalyzedResultInformation -Name "Certificate Thumbprint(s)" @baseParams

                            foreach ($thumbprint in $($connector.CertificateDetails.CertificateLifetimeInfo).keys) {
                                $params = $baseParams + @{
                                    Details                = $thumbprint
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }

                            Add-AnalyzedResultInformation -Name "Lifetime In Days" @baseParams

                            foreach ($thumbprint in $($connector.CertificateDetails.CertificateLifetimeInfo).keys) {
                                switch ($($connector.CertificateDetails.CertificateLifetimeInfo)[$thumbprint]) {
                                    { ($_ -ge 60) } { $certificateLifetimeWriteType = "Green"; break }
                                    { ($_ -ge 30) } { $certificateLifetimeWriteType = "Yellow"; break }
                                    default { $certificateLifetimeWriteType = "Red" }
                                }

                                $params = $baseParams + @{
                                    Details                = ($connector.CertificateDetails.CertificateLifetimeInfo)[$thumbprint]
                                    DisplayWriteType       = $certificateLifetimeWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }

                            $connectorCertificateMatchesHybridCertificate = $false
                            $connectorCertificateMatchesHybridCertificateWritingType = "Yellow"
                            if (($connector.CertificateDetails.TlsCertificateSet) -and
                                (-not([System.String]::IsNullOrEmpty($getHybridConfiguration.TlsCertificateName))) -and
                                ($connector.CertificateDetails.TlsCertificateName -eq $getHybridConfiguration.TlsCertificateName)) {
                                $connectorCertificateMatchesHybridCertificate = $true
                                $connectorCertificateMatchesHybridCertificateWritingType = "Green"
                            }

                            $params = $baseParams + @{
                                Name             = "Certificate Matches Hybrid Certificate"
                                Details          = $connectorCertificateMatchesHybridCertificate
                                DisplayWriteType = $connectorCertificateMatchesHybridCertificateWritingType
                            }
                            Add-AnalyzedResultInformation @params

                            if (($connector.CertificateDetails.TlsCertificateNameStatus -eq "TlsCertificateNameSyntaxInvalid") -or
                                (($connector.CertificateDetails.GoodTlsCertificateSyntax -eq $false) -and
                                    ($null -ne $connector.CertificateDetails.TlsCertificateName))) {
                                $params = $baseParams + @{
                                    Name             = "TlsCertificateName Syntax Invalid"
                                    Details          = "True"
                                    DisplayWriteType = $cloudConnectorWriteType
                                }
                                Add-AnalyzedResultInformation @params

                                $params = $baseParams + @{
                                    Details                = "The correct syntax is: '<I>X.500Issuer<S>X.500Subject'"
                                    DisplayWriteType       = $cloudConnectorWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }
                        }
                    }
                }
            }
        }
    }
}




# This function is used to determine the version of Exchange based off a build number or
# by providing the Exchange Version and CU and/or SU. This provides one location in the entire repository
# that is required to be updated for when a new release of Exchange is dropped.
function Get-ExchangeBuildVersionInformation {
    [CmdletBinding(DefaultParameterSetName = "AdminDisplayVersion")]
    param(
        [Parameter(ParameterSetName = "AdminDisplayVersion", Position = 1)]
        [object]$AdminDisplayVersion,

        [Parameter(ParameterSetName = "ExSetup")]
        [System.Version]$FileVersion,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $true)]
        [ValidateScript( { ValidateVersionParameter $_ } )]
        [string]$Version,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $true)]
        [ValidateScript( { ValidateCUParameter $_ } )]
        [string]$CU,

        [Parameter(ParameterSetName = "VersionCU", Mandatory = $false)]
        [ValidateScript( { ValidateSUParameter $_ } )]
        [string]$SU,

        [Parameter(ParameterSetName = "FindSUBuilds", Mandatory = $true)]
        [ValidateScript( { ValidateSUParameter $_ } )]
        [string]$FindBySUName,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )
    begin {

        function GetBuildVersion {
            param(
                [Parameter(Position = 1)]
                [string]$ExchangeVersion,
                [Parameter(Position = 2)]
                [string]$CU,
                [Parameter(Position = 3)]
                [string]$SU
            )
            $cuResult = $exchangeBuildDictionary[$ExchangeVersion][$CU]

            if ((-not [string]::IsNullOrEmpty($SU)) -and
                $cuResult.SU.ContainsKey($SU)) {
                return $cuResult.SU[$SU]
            } else {
                return $cuResult.CU
            }
        }

        # Dictionary of Exchange Version/CU/SU to build number
        $exchangeBuildDictionary = GetExchangeBuildDictionary

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $exchangeMajorVersion = [string]::Empty
        $exchangeVersion = $null
        $supportedBuildNumber = $false
        $latestSUBuild = $false
        $extendedSupportDate = [string]::Empty
        $cuReleaseDate = [string]::Empty
        $friendlyName = [string]::Empty
        $cuLevel = [string]::Empty
        $suName = [string]::Empty
        $orgValue = 0
        $schemaValue = 0
        $mesoValue = 0
        $ex19 = "Exchange2019"
        $ex16 = "Exchange2016"
        $ex13 = "Exchange2013"
    }
    process {
        # Convert both input types to a [System.Version]
        try {
            if ($PSCmdlet.ParameterSetName -eq "FindSUBuilds") {
                foreach ($exchangeKey in $exchangeBuildDictionary.Keys) {
                    foreach ($cuKey in $exchangeBuildDictionary[$exchangeKey].Keys) {
                        if ($null -ne $exchangeBuildDictionary[$exchangeKey][$cuKey].SU -and
                            $exchangeBuildDictionary[$exchangeKey][$cuKey].SU.ContainsKey($FindBySUName)) {
                            Get-ExchangeBuildVersionInformation -FileVersion $exchangeBuildDictionary[$exchangeKey][$cuKey].SU[$FindBySUName]
                        }
                    }
                }
                return
            } elseif ($PSCmdlet.ParameterSetName -eq "VersionCU") {
                [System.Version]$exchangeVersion = GetBuildVersion -ExchangeVersion $Version -CU $CU -SU $SU
            } elseif ($PSCmdlet.ParameterSetName -eq "AdminDisplayVersion") {
                $AdminDisplayVersion = $AdminDisplayVersion.ToString()
                Write-Verbose "Passed AdminDisplayVersion: $AdminDisplayVersion"
                $split1 = $AdminDisplayVersion.Substring(($AdminDisplayVersion.IndexOf(" ")) + 1, 4).Split(".")
                $buildStart = $AdminDisplayVersion.LastIndexOf(" ") + 1
                $split2 = $AdminDisplayVersion.Substring($buildStart, ($AdminDisplayVersion.LastIndexOf(")") - $buildStart)).Split(".")
                [System.Version]$exchangeVersion = "$($split1[0]).$($split1[1]).$($split2[0]).$($split2[1])"
            } else {
                [System.Version]$exchangeVersion = $FileVersion
            }
        } catch {
            Write-Verbose "Failed to convert to system.version"
            Invoke-CatchActionError $CatchActionFunction
        }

        <#
            Exchange Build Numbers: https://learn.microsoft.com/en-us/exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
            Exchange 2016 & 2019 AD Changes: https://learn.microsoft.com/en-us/exchange/plan-and-deploy/prepare-ad-and-domains?view=exchserver-2019
            Exchange 2013 AD Changes: https://learn.microsoft.com/en-us/exchange/prepare-active-directory-and-domains-exchange-2013-help
        #>
        if ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 2) {
            Write-Verbose "Exchange 2019 is detected"
            $exchangeMajorVersion = "Exchange2019"
            $extendedSupportDate = "10/14/2025"
            $friendlyName = "Exchange 2019"

            #Latest Version AD Settings
            $schemaValue = 17003
            $mesoValue = 13243
            $orgValue = 16762

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex19 "CU14") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "02/13/2024"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex19 "CU14" -SU "Apr24HU") { $latestSUBuild = $true }
                (GetBuildVersion $ex19 "CU14" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex19 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "05/03/2023"
                    $supportedBuildNumber = $true
                    $orgValue = 16761
                }
                (GetBuildVersion $ex19 "CU13" -SU "Apr24HU") { $latestSUBuild = $true }
                (GetBuildVersion $ex19 "CU13" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex19 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "04/20/2022"
                    $supportedBuildNumber = $false
                    $orgValue = 16760
                }
                { $_ -lt (GetBuildVersion $ex19 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "09/28/2021"
                    $mesoValue = 13242
                    $orgValue = 16759
                }
                (GetBuildVersion $ex19 "CU11" -SU "May22SU") { $mesoValue = 13243 }
                { $_ -lt (GetBuildVersion $ex19 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "06/29/2021"
                    $mesoValue = 13241
                    $orgValue = 16758
                }
                { $_ -lt (GetBuildVersion $ex19 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "03/16/2021"
                    $schemaValue = 17002
                    $mesoValue = 13240
                    $orgValue = 16757
                }
                { $_ -lt (GetBuildVersion $ex19 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "12/15/2020"
                    $mesoValue = 13239
                    $orgValue = 16756
                }
                { $_ -lt (GetBuildVersion $ex19 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "09/15/2020"
                    $schemaValue = 17001
                    $mesoValue = 13238
                    $orgValue = 16755
                }
                { $_ -lt (GetBuildVersion $ex19 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "06/16/2020"
                    $mesoValue = 13237
                    $orgValue = 16754
                }
                { $_ -lt (GetBuildVersion $ex19 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "03/17/2020"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "12/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "09/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "06/18/2019"
                }
                { $_ -lt (GetBuildVersion $ex19 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "02/12/2019"
                    $schemaValue = 17000
                    $mesoValue = 13236
                    $orgValue = 16752
                }
                { $_ -lt (GetBuildVersion $ex19 "CU1") } {
                    $cuLevel = "RTM"
                    $cuReleaseDate = "10/22/2018"
                    $orgValue = 16751
                }
            }
        } elseif ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 1) {
            Write-Verbose "Exchange 2016 is detected"
            $exchangeMajorVersion = "Exchange2016"
            $extendedSupportDate = "10/14/2025"
            $friendlyName = "Exchange 2016"

            #Latest Version AD Settings
            $schemaValue = 15334
            $mesoValue = 13243
            $orgValue = 16223

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex16 "CU23") } {
                    $cuLevel = "CU23"
                    $cuReleaseDate = "04/20/2022"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex16 "CU23" -SU "Apr24HU") { $latestSUBuild = $true }
                (GetBuildVersion $ex16 "CU23" -SU "Mar24SU") { $latestSUBuild = $true }
                { $_ -lt (GetBuildVersion $ex16 "CU23") } {
                    $cuLevel = "CU22"
                    $cuReleaseDate = "09/28/2021"
                    $supportedBuildNumber = $false
                    $mesoValue = 13242
                    $orgValue = 16222
                }
                (GetBuildVersion $ex16 "CU22" -SU "May22SU") { $mesoValue = 13243 }
                { $_ -lt (GetBuildVersion $ex16 "CU22") } {
                    $cuLevel = "CU21"
                    $cuReleaseDate = "06/29/2021"
                    $mesoValue = 13241
                    $orgValue = 16221
                }
                { $_ -lt (GetBuildVersion $ex16 "CU21") } {
                    $cuLevel = "CU20"
                    $cuReleaseDate = "03/16/2021"
                    $schemaValue = 15333
                    $mesoValue = 13240
                    $orgValue = 16220
                }
                { $_ -lt (GetBuildVersion $ex16 "CU20") } {
                    $cuLevel = "CU19"
                    $cuReleaseDate = "12/15/2020"
                    $mesoValue = 13239
                    $orgValue = 16219
                }
                { $_ -lt (GetBuildVersion $ex16 "CU19") } {
                    $cuLevel = "CU18"
                    $cuReleaseDate = "09/15/2020"
                    $schemaValue = 15332
                    $mesoValue = 13238
                    $orgValue = 16218
                }
                { $_ -lt (GetBuildVersion $ex16 "CU18") } {
                    $cuLevel = "CU17"
                    $cuReleaseDate = "06/16/2020"
                    $mesoValue = 13237
                    $orgValue = 16217
                }
                { $_ -lt (GetBuildVersion $ex16 "CU17") } {
                    $cuLevel = "CU16"
                    $cuReleaseDate = "03/17/2020"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU16") } {
                    $cuLevel = "CU15"
                    $cuReleaseDate = "12/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU15") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "09/17/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "06/18/2019"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "02/12/2019"
                    $mesoValue = 13236
                    $orgValue = 16215
                }
                { $_ -lt (GetBuildVersion $ex16 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "10/16/2018"
                    $orgValue = 16214
                }
                { $_ -lt (GetBuildVersion $ex16 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "06/19/2018"
                    $orgValue = 16213
                }
                { $_ -lt (GetBuildVersion $ex16 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "03/20/2018"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "12/19/2017"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "09/16/2017"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "06/24/2017"
                    $schemaValue = 15330
                }
                { $_ -lt (GetBuildVersion $ex16 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "03/21/2017"
                    $schemaValue = 15326
                }
                { $_ -lt (GetBuildVersion $ex16 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "12/13/2016"
                }
                { $_ -lt (GetBuildVersion $ex16 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "09/20/2016"
                    $orgValue = 16212
                }
                { $_ -lt (GetBuildVersion $ex16 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "06/21/2016"
                    $schemaValue = 15325
                }
                { $_ -lt (GetBuildVersion $ex16 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "03/15/2016"
                    $schemaValue = 15323
                    $orgValue = 16211
                }
            }
        } elseif ($exchangeVersion.Major -eq 15 -and $exchangeVersion.Minor -eq 0) {
            Write-Verbose "Exchange 2013 is detected"
            $exchangeMajorVersion = "Exchange2013"
            $extendedSupportDate = "04/11/2023"
            $friendlyName = "Exchange 2013"

            #Latest Version AD Settings
            $schemaValue = 15312
            $mesoValue = 13237
            $orgValue = 16133

            switch ($exchangeVersion) {
                { $_ -ge (GetBuildVersion $ex13 "CU23") } {
                    $cuLevel = "CU23"
                    $cuReleaseDate = "06/18/2019"
                    $supportedBuildNumber = $true
                }
                (GetBuildVersion $ex13 "CU23" -SU "May22SU") { $mesoValue = 13238 }
                { $_ -lt (GetBuildVersion $ex13 "CU23") } {
                    $cuLevel = "CU22"
                    $cuReleaseDate = "02/12/2019"
                    $mesoValue = 13236
                    $orgValue = 16131
                    $supportedBuildNumber = $false
                }
                { $_ -lt (GetBuildVersion $ex13 "CU22") } {
                    $cuLevel = "CU21"
                    $cuReleaseDate = "06/19/2018"
                    $orgValue = 16130
                }
                { $_ -lt (GetBuildVersion $ex13 "CU21") } {
                    $cuLevel = "CU20"
                    $cuReleaseDate = "03/20/2018"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU20") } {
                    $cuLevel = "CU19"
                    $cuReleaseDate = "12/19/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU19") } {
                    $cuLevel = "CU18"
                    $cuReleaseDate = "09/16/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU18") } {
                    $cuLevel = "CU17"
                    $cuReleaseDate = "06/24/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU17") } {
                    $cuLevel = "CU16"
                    $cuReleaseDate = "03/21/2017"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU16") } {
                    $cuLevel = "CU15"
                    $cuReleaseDate = "12/13/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU15") } {
                    $cuLevel = "CU14"
                    $cuReleaseDate = "09/20/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU14") } {
                    $cuLevel = "CU13"
                    $cuReleaseDate = "06/21/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU13") } {
                    $cuLevel = "CU12"
                    $cuReleaseDate = "03/15/2016"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU12") } {
                    $cuLevel = "CU11"
                    $cuReleaseDate = "12/15/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU11") } {
                    $cuLevel = "CU10"
                    $cuReleaseDate = "09/15/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU10") } {
                    $cuLevel = "CU9"
                    $cuReleaseDate = "06/17/2015"
                    $orgValue = 15965
                }
                { $_ -lt (GetBuildVersion $ex13 "CU9") } {
                    $cuLevel = "CU8"
                    $cuReleaseDate = "03/17/2015"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU8") } {
                    $cuLevel = "CU7"
                    $cuReleaseDate = "12/09/2014"
                }
                { $_ -lt (GetBuildVersion $ex13 "CU7") } {
                    $cuLevel = "CU6"
                    $cuReleaseDate = "08/26/2014"
                    $schemaValue = 15303
                }
                { $_ -lt (GetBuildVersion $ex13 "CU6") } {
                    $cuLevel = "CU5"
                    $cuReleaseDate = "05/27/2014"
                    $schemaValue = 15300
                    $orgValue = 15870
                }
                { $_ -lt (GetBuildVersion $ex13 "CU5") } {
                    $cuLevel = "CU4"
                    $cuReleaseDate = "02/25/2014"
                    $schemaValue = 15292
                    $orgValue = 15844
                }
                { $_ -lt (GetBuildVersion $ex13 "CU4") } {
                    $cuLevel = "CU3"
                    $cuReleaseDate = "11/25/2013"
                    $schemaValue = 15283
                    $orgValue = 15763
                }
                { $_ -lt (GetBuildVersion $ex13 "CU3") } {
                    $cuLevel = "CU2"
                    $cuReleaseDate = "07/09/2013"
                    $schemaValue = 15281
                    $orgValue = 15688
                }
                { $_ -lt (GetBuildVersion $ex13 "CU2") } {
                    $cuLevel = "CU1"
                    $cuReleaseDate = "04/02/2013"
                    $schemaValue = 15254
                    $orgValue = 15614
                }
            }
        } else {
            Write-Verbose "Unknown version of Exchange is detected."
        }

        # Now get the SU Name
        if ([string]::IsNullOrEmpty($exchangeMajorVersion) -or
            [string]::IsNullOrEmpty($cuLevel)) {
            Write-Verbose "Can't lookup when keys aren't set"
            return
        }

        $currentSUInfo = $exchangeBuildDictionary[$exchangeMajorVersion][$cuLevel].SU
        $compareValue = $exchangeVersion.ToString()
        if ($null -ne $currentSUInfo -and
            $currentSUInfo.ContainsValue($compareValue)) {
            foreach ($key in $currentSUInfo.Keys) {
                if ($compareValue -eq $currentSUInfo[$key]) {
                    $suName = $key
                }
            }
        }
    }
    end {

        if ($PSCmdlet.ParameterSetName -eq "FindSUBuilds") {
            Write-Verbose "Return nothing here, results were already returned on the pipeline"
            return
        }

        $friendlyName = "$friendlyName $cuLevel $suName".Trim()
        Write-Verbose "Determined Build Version $friendlyName"
        return [PSCustomObject]@{
            MajorVersion        = $exchangeMajorVersion
            FriendlyName        = $friendlyName
            BuildVersion        = $exchangeVersion
            CU                  = $cuLevel
            ReleaseDate         = if (-not([System.String]::IsNullOrEmpty($cuReleaseDate))) { ([System.Convert]::ToDateTime([DateTime]$cuReleaseDate, [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) } else { $null }
            ExtendedSupportDate = if (-not([System.String]::IsNullOrEmpty($extendedSupportDate))) { ([System.Convert]::ToDateTime([DateTime]$extendedSupportDate, [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) } else { $null }
            Supported           = $supportedBuildNumber
            LatestSU            = $latestSUBuild
            ADLevel             = [PSCustomObject]@{
                SchemaValue = $schemaValue
                MESOValue   = $mesoValue
                OrgValue    = $orgValue
            }
        }
    }
}

function GetExchangeBuildDictionary {

    function NewCUAndSUObject {
        param(
            [string]$CUBuildNumber,
            [Hashtable]$SUBuildNumber
        )
        return @{
            "CU" = $CUBuildNumber
            "SU" = $SUBuildNumber
        }
    }

    @{
        "Exchange2013" = @{
            "CU1"  = (NewCUAndSUObject "15.0.620.29")
            "CU2"  = (NewCUAndSUObject "15.0.712.24")
            "CU3"  = (NewCUAndSUObject "15.0.775.38")
            "CU4"  = (NewCUAndSUObject "15.0.847.32")
            "CU5"  = (NewCUAndSUObject "15.0.913.22")
            "CU6"  = (NewCUAndSUObject "15.0.995.29")
            "CU7"  = (NewCUAndSUObject "15.0.1044.25")
            "CU8"  = (NewCUAndSUObject "15.0.1076.9")
            "CU9"  = (NewCUAndSUObject "15.0.1104.5")
            "CU10" = (NewCUAndSUObject "15.0.1130.7")
            "CU11" = (NewCUAndSUObject "15.0.1156.6")
            "CU12" = (NewCUAndSUObject "15.0.1178.4")
            "CU13" = (NewCUAndSUObject "15.0.1210.3")
            "CU14" = (NewCUAndSUObject "15.0.1236.3")
            "CU15" = (NewCUAndSUObject "15.0.1263.5")
            "CU16" = (NewCUAndSUObject "15.0.1293.2")
            "CU17" = (NewCUAndSUObject "15.0.1320.4")
            "CU18" = (NewCUAndSUObject "15.0.1347.2" @{
                    "Mar18SU" = "15.0.1347.5"
                })
            "CU19" = (NewCUAndSUObject "15.0.1365.1" @{
                    "Mar18SU" = "15.0.1365.3"
                    "May18SU" = "15.0.1365.7"
                })
            "CU20" = (NewCUAndSUObject "15.0.1367.3" @{
                    "May18SU" = "15.0.1367.6"
                    "Aug18SU" = "15.0.1367.9"
                })
            "CU21" = (NewCUAndSUObject "15.0.1395.4" @{
                    "Aug18SU" = "15.0.1395.7"
                    "Oct18SU" = "15.0.1395.8"
                    "Jan19SU" = "15.0.1395.10"
                    "Mar21SU" = "15.0.1395.12"
                })
            "CU22" = (NewCUAndSUObject "15.0.1473.3" @{
                    "Feb19SU" = "15.0.1473.3"
                    "Apr19SU" = "15.0.1473.4"
                    "Jun19SU" = "15.0.1473.5"
                    "Mar21SU" = "15.0.1473.6"
                })
            "CU23" = (NewCUAndSUObject "15.0.1497.2" @{
                    "Jul19SU" = "15.0.1497.3"
                    "Nov19SU" = "15.0.1497.4"
                    "Feb20SU" = "15.0.1497.6"
                    "Oct20SU" = "15.0.1497.7"
                    "Nov20SU" = "15.0.1497.8"
                    "Dec20SU" = "15.0.1497.10"
                    "Mar21SU" = "15.0.1497.12"
                    "Apr21SU" = "15.0.1497.15"
                    "May21SU" = "15.0.1497.18"
                    "Jul21SU" = "15.0.1497.23"
                    "Oct21SU" = "15.0.1497.24"
                    "Nov21SU" = "15.0.1497.26"
                    "Jan22SU" = "15.0.1497.28"
                    "Mar22SU" = "15.0.1497.33"
                    "May22SU" = "15.0.1497.36"
                    "Aug22SU" = "15.0.1497.40"
                    "Oct22SU" = "15.0.1497.42"
                    "Nov22SU" = "15.0.1497.44"
                    "Jan23SU" = "15.0.1497.45"
                    "Feb23SU" = "15.0.1497.47"
                    "Mar23SU" = "15.0.1497.48"
                })
        }
        "Exchange2016" = @{
            "CU1"  = (NewCUAndSUObject "15.1.396.30")
            "CU2"  = (NewCUAndSUObject "15.1.466.34")
            "CU3"  = (NewCUAndSUObject "15.1.544.27")
            "CU4"  = (NewCUAndSUObject "15.1.669.32")
            "CU5"  = (NewCUAndSUObject "15.1.845.34")
            "CU6"  = (NewCUAndSUObject "15.1.1034.26")
            "CU7"  = (NewCUAndSUObject "15.1.1261.35" @{
                    "Mar18SU" = "15.1.1261.39"
                })
            "CU8"  = (NewCUAndSUObject "15.1.1415.2" @{
                    "Mar18SU" = "15.1.1415.4"
                    "May18SU" = "15.1.1415.7"
                    "Mar21SU" = "15.1.1415.8"
                })
            "CU9"  = (NewCUAndSUObject "15.1.1466.3" @{
                    "May18SU" = "15.1.1466.8"
                    "Aug18SU" = "15.1.1466.9"
                    "Mar21SU" = "15.1.1466.13"
                })
            "CU10" = (NewCUAndSUObject "15.1.1531.3" @{
                    "Aug18SU" = "15.1.1531.6"
                    "Oct18SU" = "15.1.1531.8"
                    "Jan19SU" = "15.1.1531.10"
                    "Mar21SU" = "15.1.1531.12"
                })
            "CU11" = (NewCUAndSUObject "15.1.1591.10" @{
                    "Dec18SU" = "15.1.1591.11"
                    "Jan19SU" = "15.1.1591.13"
                    "Apr19SU" = "15.1.1591.16"
                    "Jun19SU" = "15.1.1591.17"
                    "Mar21SU" = "15.1.1591.18"
                })
            "CU12" = (NewCUAndSUObject "15.1.1713.5" @{
                    "Feb19SU" = "15.1.1713.5"
                    "Apr19SU" = "15.1.1713.6"
                    "Jun19SU" = "15.1.1713.7"
                    "Jul19SU" = "15.1.1713.8"
                    "Sep19SU" = "15.1.1713.9"
                    "Mar21SU" = "15.1.1713.10"
                })
            "CU13" = (NewCUAndSUObject "15.1.1779.2" @{
                    "Jul19SU" = "15.1.1779.4"
                    "Sep19SU" = "15.1.1779.5"
                    "Nov19SU" = "15.1.1779.7"
                    "Mar21SU" = "15.1.1779.8"
                })
            "CU14" = (NewCUAndSUObject "15.1.1847.3" @{
                    "Nov19SU" = "15.1.1847.5"
                    "Feb20SU" = "15.1.1847.7"
                    "Mar20SU" = "15.1.1847.10"
                    "Mar21SU" = "15.1.1847.12"
                })
            "CU15" = (NewCUAndSUObject "15.1.1913.5" @{
                    "Feb20SU" = "15.1.1913.7"
                    "Mar20SU" = "15.1.1913.10"
                    "Mar21SU" = "15.1.1913.12"
                })
            "CU16" = (NewCUAndSUObject "15.1.1979.3" @{
                    "Sep20SU" = "15.1.1979.6"
                    "Mar21SU" = "15.1.1979.8"
                })
            "CU17" = (NewCUAndSUObject "15.1.2044.4" @{
                    "Sep20SU" = "15.1.2044.6"
                    "Oct20SU" = "15.1.2044.7"
                    "Nov20SU" = "15.1.2044.8"
                    "Dec20SU" = "15.1.2044.12"
                    "Mar21SU" = "15.1.2044.13"
                })
            "CU18" = (NewCUAndSUObject "15.1.2106.2" @{
                    "Oct20SU" = "15.1.2106.3"
                    "Nov20SU" = "15.1.2106.4"
                    "Dec20SU" = "15.1.2106.6"
                    "Feb21SU" = "15.1.2106.8"
                    "Mar21SU" = "15.1.2106.13"
                })
            "CU19" = (NewCUAndSUObject "15.1.2176.2" @{
                    "Feb21SU" = "15.1.2176.4"
                    "Mar21SU" = "15.1.2176.9"
                    "Apr21SU" = "15.1.2176.12"
                    "May21SU" = "15.1.2176.14"
                })
            "CU20" = (NewCUAndSUObject "15.1.2242.4" @{
                    "Apr21SU" = "15.1.2242.8"
                    "May21SU" = "15.1.2242.10"
                    "Jul21SU" = "15.1.2242.12"
                })
            "CU21" = (NewCUAndSUObject "15.1.2308.8" @{
                    "Jul21SU" = "15.1.2308.14"
                    "Oct21SU" = "15.1.2308.15"
                    "Nov21SU" = "15.1.2308.20"
                    "Jan22SU" = "15.1.2308.21"
                    "Mar22SU" = "15.1.2308.27"
                })
            "CU22" = (NewCUAndSUObject "15.1.2375.7" @{
                    "Oct21SU" = "15.1.2375.12"
                    "Nov21SU" = "15.1.2375.17"
                    "Jan22SU" = "15.1.2375.18"
                    "Mar22SU" = "15.1.2375.24"
                    "May22SU" = "15.1.2375.28"
                    "Aug22SU" = "15.1.2375.31"
                    "Oct22SU" = "15.1.2375.32"
                    "Nov22SU" = "15.1.2375.37"
                })
            "CU23" = (NewCUAndSUObject "15.1.2507.6" @{
                    "May22SU"   = "15.1.2507.9"
                    "Aug22SU"   = "15.1.2507.12"
                    "Oct22SU"   = "15.1.2507.13"
                    "Nov22SU"   = "15.1.2507.16"
                    "Jan23SU"   = "15.1.2507.17"
                    "Feb23SU"   = "15.1.2507.21"
                    "Mar23SU"   = "15.1.2507.23"
                    "Jun23SU"   = "15.1.2507.27"
                    "Aug23SU"   = "15.1.2507.31"
                    "Aug23SUv2" = "15.1.2507.32"
                    "Oct23SU"   = "15.1.2507.34"
                    "Nov23SU"   = "15.1.2507.35"
                    "Mar24SU"   = "15.1.2507.37"
                    "Apr24HU"   = "15.1.2507.39"
                })
        }
        "Exchange2019" = @{
            "CU1"  = (NewCUAndSUObject "15.2.330.5" @{
                    "Feb19SU" = "15.2.330.5"
                    "Apr19SU" = "15.2.330.7"
                    "Jun19SU" = "15.2.330.8"
                    "Jul19SU" = "15.2.330.9"
                    "Sep19SU" = "15.2.330.10"
                    "Mar21SU" = "15.2.330.11"
                })
            "CU2"  = (NewCUAndSUObject "15.2.397.3" @{
                    "Jul19SU" = "15.2.397.5"
                    "Sep19SU" = "15.2.397.6"
                    "Nov19SU" = "15.2.397.9"
                    "Mar21SU" = "15.2.397.11"
                })
            "CU3"  = (NewCUAndSUObject "15.2.464.5" @{
                    "Nov19SU" = "15.2.464.7"
                    "Feb20SU" = "15.2.464.11"
                    "Mar20SU" = "15.2.464.14"
                    "Mar21SU" = "15.2.464.15"
                })
            "CU4"  = (NewCUAndSUObject "15.2.529.5" @{
                    "Feb20SU" = "15.2.529.8"
                    "Mar20SU" = "15.2.529.11"
                    "Mar21SU" = "15.2.529.13"
                })
            "CU5"  = (NewCUAndSUObject "15.2.595.3" @{
                    "Sep20SU" = "15.2.595.6"
                    "Mar21SU" = "15.2.595.8"
                })
            "CU6"  = (NewCUAndSUObject "15.2.659.4" @{
                    "Sep20SU" = "15.2.659.6"
                    "Oct20SU" = "15.2.659.7"
                    "Nov20SU" = "15.2.659.8"
                    "Dec20SU" = "15.2.659.11"
                    "Mar21SU" = "15.2.659.12"
                })
            "CU7"  = (NewCUAndSUObject "15.2.721.2" @{
                    "Oct20SU" = "15.2.721.3"
                    "Nov20SU" = "15.2.721.4"
                    "Dec20SU" = "15.2.721.6"
                    "Feb21SU" = "15.2.721.8"
                    "Mar21SU" = "15.2.721.13"
                })
            "CU8"  = (NewCUAndSUObject "15.2.792.3" @{
                    "Feb21SU" = "15.2.792.5"
                    "Mar21SU" = "15.2.792.10"
                    "Apr21SU" = "15.2.792.13"
                    "May21SU" = "15.2.792.15"
                })
            "CU9"  = (NewCUAndSUObject "15.2.858.5" @{
                    "Apr21SU" = "15.2.858.10"
                    "May21SU" = "15.2.858.12"
                    "Jul21SU" = "15.2.858.15"
                })
            "CU10" = (NewCUAndSUObject "15.2.922.7" @{
                    "Jul21SU" = "15.2.922.13"
                    "Oct21SU" = "15.2.922.14"
                    "Nov21SU" = "15.2.922.19"
                    "Jan22SU" = "15.2.922.20"
                    "Mar22SU" = "15.2.922.27"
                })
            "CU11" = (NewCUAndSUObject "15.2.986.5" @{
                    "Oct21SU" = "15.2.986.9"
                    "Nov21SU" = "15.2.986.14"
                    "Jan22SU" = "15.2.986.15"
                    "Mar22SU" = "15.2.986.22"
                    "May22SU" = "15.2.986.26"
                    "Aug22SU" = "15.2.986.29"
                    "Oct22SU" = "15.2.986.30"
                    "Nov22SU" = "15.2.986.36"
                    "Jan23SU" = "15.2.986.37"
                    "Feb23SU" = "15.2.986.41"
                    "Mar23SU" = "15.2.986.42"
                })
            "CU12" = (NewCUAndSUObject "15.2.1118.7" @{
                    "May22SU"   = "15.2.1118.9"
                    "Aug22SU"   = "15.2.1118.12"
                    "Oct22SU"   = "15.2.1118.15"
                    "Nov22SU"   = "15.2.1118.20"
                    "Jan23SU"   = "15.2.1118.21"
                    "Feb23SU"   = "15.2.1118.25"
                    "Mar23SU"   = "15.2.1118.26"
                    "Jun23SU"   = "15.2.1118.30"
                    "Aug23SU"   = "15.2.1118.36"
                    "Aug23SUv2" = "15.2.1118.37"
                    "Oct23SU"   = "15.2.1118.39"
                    "Nov23SU"   = "15.2.1118.40"
                })
            "CU13" = (NewCUAndSUObject "15.2.1258.12" @{
                    "Jun23SU"   = "15.2.1258.16"
                    "Aug23SU"   = "15.2.1258.23"
                    "Aug23SUv2" = "15.2.1258.25"
                    "Oct23SU"   = "15.2.1258.27"
                    "Nov23SU"   = "15.2.1258.28"
                    "Mar24SU"   = "15.2.1258.32"
                    "Apr24HU"   = "15.2.1258.34"
                })
            "CU14" = (NewCUAndSUObject "15.2.1544.4" @{
                    "Mar24SU" = "15.2.1544.9"
                    "Apr24HU" = "15.2.1544.11"
                })
        }
    }
}

# Must be outside function to use it as a validate script
function GetValidatePossibleParameters {
    $exchangeBuildDictionary = GetExchangeBuildDictionary
    $suNames = New-Object 'System.Collections.Generic.HashSet[string]'
    $cuNames = New-Object 'System.Collections.Generic.HashSet[string]'
    $versionNames = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($exchangeKey in $exchangeBuildDictionary.Keys) {
        [void]$versionNames.Add($exchangeKey)
        foreach ($cuKey in $exchangeBuildDictionary[$exchangeKey].Keys) {
            [void]$cuNames.Add($cuKey)
            if ($null -eq $exchangeBuildDictionary[$exchangeKey][$cuKey].SU) { continue }
            foreach ($suKey in $exchangeBuildDictionary[$exchangeKey][$cuKey].SU.Keys) {
                [void]$suNames.Add($suKey)
            }
        }
    }
    return [PSCustomObject]@{
        Version = $versionNames
        CU      = $cuNames
        SU      = $suNames
    }
}

function ValidateSUParameter {
    param($name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.SU.Contains($Name)
}

function ValidateCUParameter {
    param($Name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.CU.Contains($Name)
}

function ValidateVersionParameter {
    param($Name)

    $possibleParameters = GetValidatePossibleParameters
    $possibleParameters.Version.Contains($Name)
}
function Test-ExchangeBuildGreaterOrEqualThanBuild {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CurrentExchangeBuild,
        [Parameter(Mandatory = $true)]
        [string]$Version,
        [Parameter(Mandatory = $true)]
        [string]$CU,
        [Parameter(Mandatory = $false)]
        [string]$SU
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $testResult = $false
    } process {
        if ($CurrentExchangeBuild.MajorVersion -eq $Version) {
            $params = @{
                Version = $Version
                CU      = $CU
            }

            if (-not([string]::IsNullOrEmpty($SU))) {
                $params.SU = $SU
            }
            $testBuild = Get-ExchangeBuildVersionInformation @params
            $testResult = $CurrentExchangeBuild.BuildVersion -ge $testBuild.BuildVersion
        }
    } end {
        Write-Verbose "Result $testResult"
        return $testResult
    }
}

function Test-ExchangeBuildLessThanBuild {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CurrentExchangeBuild,
        [Parameter(Mandatory = $true)]
        [string]$Version,
        [Parameter(Mandatory = $true)]
        [string]$CU,
        [Parameter(Mandatory = $false)]
        [string]$SU
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $testResult = $false
    } process {
        if ($CurrentExchangeBuild.MajorVersion -eq $Version) {
            $params = @{
                Version = $Version
                CU      = $CU
            }

            if (-not([string]::IsNullOrEmpty($SU))) {
                $params.SU = $SU
            }

            $testBuild = Get-ExchangeBuildVersionInformation @params
            $testResult = $CurrentExchangeBuild.BuildVersion -lt $testBuild.BuildVersion
        }
    } end {
        Write-Verbose "Result $testResult"
        return $testResult
    }
}

function Test-ExchangeBuildEqualBuild {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$CurrentExchangeBuild,
        [Parameter(Mandatory = $true)]
        [string]$Version,
        [Parameter(Mandatory = $true)]
        [string]$CU,
        [Parameter(Mandatory = $false)]
        [string]$SU
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $testResult = $false
    } process {
        if ($CurrentExchangeBuild.MajorVersion -eq $Version) {
            $params = @{
                Version = $Version
                CU      = $CU
            }

            if (-not([string]::IsNullOrEmpty($SU))) {
                $params.SU = $SU
            }
            $testBuild = Get-ExchangeBuildVersionInformation @params
            $testResult = $CurrentExchangeBuild.BuildVersion -eq $testBuild.BuildVersion
        }
    } end {
        Write-Verbose "Result $testResult"
        return $testResult
    }
}

function Test-ExchangeBuildGreaterOrEqualThanSecurityPatch {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [object]$CurrentExchangeBuild,
        [string]$SUName
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $testResult = $false
    } process {
        $allSecurityPatches = Get-ExchangeBuildVersionInformation -FindBySUName $SUName |
            Where-Object { $_.MajorVersion -eq $CurrentExchangeBuild.MajorVersion } |
            Sort-Object ReleaseDate -Descending

        if ($null -eq $allSecurityPatches -or
            $allSecurityPatches.Count -eq 0) {
            Write-Verbose "We didn't find a security path for this version of Exchange."
            Write-Verbose "We assume this means that this version of Exchange $($CurrentExchangeBuild.MajorVersion) isn't vulnerable for this SU $SUName"
            $testResult = $true
            return
        }

        # The first item in the list should be the latest CU for this security patch.
        # If the current exchange build is greater than the latest CU + security patch, then we are good.
        # Otherwise, we need to look at the CU that we are on to make sure we are patched.
        if ($CurrentExchangeBuild.BuildVersion -ge $allSecurityPatches[0].BuildVersion) {
            $testResult = $true
            return
        }
        Write-Verbose "Need to look at particular CU match"
        $matchCU = $allSecurityPatches | Where-Object { $_.CU -eq $CurrentExchangeBuild.CU }
        Write-Verbose "Found match CU $($null -ne $matchCU)"
        $testResult = $null -ne $matchCU -and $CurrentExchangeBuild.BuildVersion -ge $matchCU.BuildVersion
    } end {
        Write-Verbose "Result $testResult"
        return $testResult
    }
}



function Get-RemoteRegistrySubKey {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Attempting to open the Base Key $RegistryHive on Machine $MachineName"
        $regKey = $null
    }
    process {

        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegistryHive, $MachineName)
            Write-Verbose "Attempting to open the Sub Key '$SubKey'"
            $regKey = $reg.OpenSubKey($SubKey)
            Write-Verbose "Opened Sub Key"
        } catch {
            Write-Verbose "Failed to open the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        return $regKey
    }
}

function Get-RemoteRegistryValue {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [string]$GetValue,
        [string]$ValueType,
        [ScriptBlock]$CatchActionFunction
    )

    <#
    Valid ValueType return values (case-sensitive)
    (https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registryvaluekind?view=net-5.0)
    Binary = REG_BINARY
    DWord = REG_DWORD
    ExpandString = REG_EXPAND_SZ
    MultiString = REG_MULTI_SZ
    None = No data type
    QWord = REG_QWORD
    String = REG_SZ
    Unknown = An unsupported registry data type
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $registryGetValue = $null
    }
    process {

        try {

            $regSubKey = Get-RemoteRegistrySubKey -RegistryHive $RegistryHive `
                -MachineName $MachineName `
                -SubKey $SubKey

            if (-not ([System.String]::IsNullOrWhiteSpace($regSubKey))) {
                Write-Verbose "Attempting to get the value $GetValue"
                $registryGetValue = $regSubKey.GetValue($GetValue)
                Write-Verbose "Finished running GetValue()"

                if ($null -ne $registryGetValue -and
                    (-not ([System.String]::IsNullOrWhiteSpace($ValueType)))) {
                    Write-Verbose "Validating ValueType $ValueType"
                    $registryValueType = $regSubKey.GetValueKind($GetValue)
                    Write-Verbose "Finished running GetValueKind()"

                    if ($ValueType -ne $registryValueType) {
                        Write-Verbose "ValueType: $ValueType is different to the returned ValueType: $registryValueType"
                        $registryGetValue = $null
                    } else {
                        Write-Verbose "ValueType matches: $ValueType"
                    }
                }
            }
        } catch {
            Write-Verbose "Failed to get the value on the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        if ($registryGetValue.Length -le 100) {
            Write-Verbose "$($MyInvocation.MyCommand) Return Value: '$registryGetValue'"
        } else {
            Write-Verbose "$($MyInvocation.MyCommand) Return Value is too long to log"
        }
        return $registryGetValue
    }
}

function Get-NETFrameworkVersion {
    [CmdletBinding(DefaultParameterSetName = "CollectFromServer")]
    param(
        [Parameter(ParameterSetName = "CollectFromServer", Position = 1)]
        [string]$MachineName = $env:COMPUTERNAME,

        [Parameter(ParameterSetName = "NetKey")]
        [int]$NetVersionKey = -1,

        [Parameter(ParameterSetName = "NetName")]
        [ValidateScript({ ValidateNetNameParameter $_ })]
        [string]$NetVersionShortName,

        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $friendlyName = [string]::Empty
        $minValue = -1
        $netVersionDictionary = GetNetVersionDictionary

        if ($PSCmdlet.ParameterSetName -eq "NetName") {
            $NetVersionKey = $netVersionDictionary[$NetVersionShortName]
        }
    }
    process {

        if ($NetVersionKey -eq -1) {
            $params = @{
                MachineName         = $MachineName
                SubKey              = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
                GetValue            = "Release"
                CatchActionFunction = $CatchActionFunction
            }
            [int]$NetVersionKey = Get-RemoteRegistryValue @params
        }

        #Using Minimum Version as per https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed?redirectedfrom=MSDN#minimum-version
        if ($NetVersionKey -lt $netVersionDictionary["Net4d5"]) {
            $friendlyName = "Unknown"
            $minValue = -1
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d5d1"]) {
            $friendlyName = "4.5"
            $minValue = $netVersionDictionary["Net4d5"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d5d2"]) {
            $friendlyName = "4.5.1"
            $minValue = $netVersionDictionary["Net4d5d1"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d6"]) {
            $friendlyName = "4.5.2"
            $minValue = $netVersionDictionary["Net4d5d2"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d6d1"]) {
            $friendlyName = "4.6"
            $minValue = $netVersionDictionary["Net4d6"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d6d2"]) {
            $friendlyName = "4.6.1"
            $minValue = $netVersionDictionary["Net4d6d1"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d7"]) {
            $friendlyName = "4.6.2"
            $minValue = $netVersionDictionary["Net4d6d2"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d7d1"]) {
            $friendlyName = "4.7"
            $minValue = $netVersionDictionary["Net4d7"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d7d2"]) {
            $friendlyName = "4.7.1"
            $minValue = $netVersionDictionary["Net4d7d1"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d8"]) {
            $friendlyName = "4.7.2"
            $minValue = $netVersionDictionary["Net4d7d2"]
        } elseif ($NetVersionKey -lt $netVersionDictionary["Net4d8d1"]) {
            $friendlyName = "4.8"
            $minValue = $netVersionDictionary["Net4d8"]
        } elseif ($NetVersionKey -ge $netVersionDictionary["Net4d8d1"]) {
            $friendlyName = "4.8.1"
            $minValue = $netVersionDictionary["Net4d8d1"]
        }
    }
    end {
        Write-Verbose "FriendlyName: $friendlyName | RegistryValue: $netVersionKey | MinimumValue: $minValue"
        return [PSCustomObject]@{
            FriendlyName  = $friendlyName
            RegistryValue = $NetVersionKey
            MinimumValue  = $minValue
        }
    }
}

function GetNetVersionDictionary {
    return @{
        "Net4d5"       = 378389
        "Net4d5d1"     = 378675
        "Net4d5d2"     = 379893
        "Net4d5d2wFix" = 380035
        "Net4d6"       = 393295
        "Net4d6d1"     = 394254
        "Net4d6d1wFix" = 394294
        "Net4d6d2"     = 394802
        "Net4d7"       = 460798
        "Net4d7d1"     = 461308
        "Net4d7d2"     = 461808
        "Net4d8"       = 528040
        "Net4d8d1"     = 533320
    }
}

function ValidateNetNameParameter {
    param($name)
    $netVersionNames = @((GetNetVersionDictionary).Keys)
    $netVersionNames.Contains($name)
}
function Invoke-AnalyzerOsInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Operating System Information"  -DisplayOrder $Order)
    }

    $params = $baseParams + @{
        Name                  = "Version"
        Details               = $osInformation.BuildInformation.FriendlyName
        AddHtmlOverviewValues = $true
        HtmlName              = "OS Version"
    }
    Add-AnalyzedResultInformation @params

    $upTime = "{0} day(s) {1} hour(s) {2} minute(s) {3} second(s)" -f $osInformation.ServerBootUp.Days,
    $osInformation.ServerBootUp.Hours,
    $osInformation.ServerBootUp.Minutes,
    $osInformation.ServerBootUp.Seconds

    $params = $baseParams + @{
        Name                = "System Up Time"
        Details             = $upTime
        DisplayTestingValue = $osInformation.ServerBootUp
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Time Zone"
        Details               = $osInformation.TimeZone.CurrentTimeZone
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $writeValue = $false
    $warning = @("Windows can not properly detect any DST rule changes in your time zone. Set 'Adjust for daylight saving time automatically to on'")

    if ($osInformation.TimeZone.DstIssueDetected) {
        $writeType = "Red"
    } elseif ($osInformation.TimeZone.DynamicDaylightTimeDisabled -ne 0) {
        $writeType = "Yellow"
    } else {
        $warning = [string]::Empty
        $writeValue = $true
        $writeType = "Grey"
    }

    $params = $baseParams + @{
        Name             = "Dynamic Daylight Time Enabled"
        Details          = $writeValue
        DisplayWriteType = $writeType
    }
    Add-AnalyzedResultInformation @params

    if ($warning -ne [string]::Empty) {
        $params = $baseParams + @{
            Details                = $warning
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if ([string]::IsNullOrEmpty($osInformation.TimeZone.TimeZoneKeyName)) {
        $params = $baseParams + @{
            Name             = "Time Zone Key Name"
            Details          = "Empty --- Warning Need to switch your current time zone to a different value, then switch it back to have this value populated again."
            DisplayWriteType = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    # .NET Supported Levels
    $currentExchangeBuild = $exchangeInformation.BuildInformation.VersionInformation
    $ex2019 = "Exchange2019"
    $ex2016 = "Exchange2016"
    $ex2013 = "Exchange2013"
    $osVersion = $osInformation.BuildInformation.MajorVersion
    $recommendedNetVersion = $null
    $netVersionDictionary = GetNetVersionDictionary

    Write-Verbose "Checking $($exchangeInformation.BuildInformation.MajorVersion) .NET Framework Support Versions"

    if ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU4")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d5"]
    } elseif ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU13") -or
    (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU2")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d5d2wFix"]
    } elseif ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU15") -or
    (Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU2") -or
    ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU3") -and
        $osVersion -ne "Windows2016")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d6d1wFix"]
    } elseif ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU19") -or
    (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU8")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d6d2"]
    } elseif ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU21") -or
    (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU11")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d7d1"]
    } elseif ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU21") -or
    (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU13") -or
    (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU2")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d7d2"]
    } elseif ((Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU14") -and
        ($osVersion -ne "Windows2019")) {
        $recommendedNetVersion = $netVersionDictionary["Net4d8d1"]
    } else {
        $recommendedNetVersion = $netVersionDictionary["Net4d8"]
    }

    Write-Verbose "Recommended NET Version: $recommendedNetVersion"

    if ($osInformation.NETFramework.MajorVersion -eq $recommendedNetVersion) {
        $params = $baseParams + @{
            Name                  = ".NET Framework"
            Details               = $osInformation.NETFramework.FriendlyName
            DisplayWriteType      = "Green"
            AddHtmlOverviewValues = $true
        }
        Add-AnalyzedResultInformation @params
    } else {
        $displayFriendly = Get-NETFrameworkVersion -NetVersionKey $recommendedNetVersion
        $displayValue = "{0} - Warning Recommended .NET Version is {1}" -f $osInformation.NETFramework.FriendlyName, $displayFriendly.FriendlyName
        $testValue = [PSCustomObject]@{
            CurrentValue        = $osInformation.NETFramework.FriendlyName
            MaxSupportedVersion = $recommendedNetVersion
        }
        $params = $baseParams + @{
            Name                   = ".NET Framework"
            Details                = $displayValue
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $testValue
            HtmlDetailsCustomValue = $osInformation.NETFramework.FriendlyName
            AddHtmlOverviewValues  = $true
        }
        Add-AnalyzedResultInformation @params

        if ($osInformation.NETFramework.MajorVersion -gt $recommendedNetVersion) {
            # Generic information stating we are looking into supporting this version of .NET
            # But don't use it till we update the supportability matrix
            $displayValue = "Microsoft is working on .NET $($osInformation.NETFramework.FriendlyName) validation with Exchange" +
            " and the recommendation is to not use .NET $($osInformation.NETFramework.FriendlyName) until it is officially added to the supportability matrix."

            $params = $baseParams + @{
                Details                = $displayValue
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-NetFrameworkSupport"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $displayValue = [string]::Empty
    $displayWriteType = "Yellow"
    $totalPhysicalMemory = [Math]::Round($hardwareInformation.TotalMemory / 1MB)
    $instanceCount = 0
    Write-Verbose "Evaluating PageFile Information"
    Write-Verbose "Total Memory: $totalPhysicalMemory"

    foreach ($pageFile in $osInformation.PageFile) {

        $pageFileDisplayTemplate = "{0} Size: {1}MB"
        $pageFileAdditionalDisplayValue = $null

        Write-Verbose "Working on PageFile: $($pageFile.Name)"
        Write-Verbose "Max PageFile Size: $($pageFile.MaximumSize)"
        $pageFileObj = [PSCustomObject]@{
            Name                = $pageFile.Name
            TotalPhysicalMemory = $totalPhysicalMemory
            MaxPageSize         = $pageFile.MaximumSize
            MultiPageFile       = (($osInformation.PageFile).Count -gt 1)
            RecommendedPageFile = 0
        }

        if ($pageFileObj.MaxPageSize -eq 0) {
            Write-Verbose "Unconfigured PageFile detected"
            if ([System.String]::IsNullOrEmpty($pageFileObj.Name)) {
                Write-Verbose "System-wide automatically managed PageFile detected"
                $displayValue = ($pageFileDisplayTemplate -f "System is set to automatically manage the PageFile", $pageFileObj.MaxPageSize)
            } else {
                Write-Verbose "Specific system-managed PageFile detected"
                $displayValue = ($pageFileDisplayTemplate -f $pageFileObj.Name, $pageFileObj.MaxPageSize)
            }
            $displayWriteType = "Red"
        } else {
            Write-Verbose "Configured PageFile detected"
            $displayValue = ($pageFileDisplayTemplate -f $pageFileObj.Name, $pageFileObj.MaxPageSize)
        }

        if ($exchangeInformation.BuildInformation.VersionInformation.BuildVersion -ge "15.2.0.0") {
            $recommendedPageFile = [Math]::Round($totalPhysicalMemory / 4)
            $pageFileObj.RecommendedPageFile = $recommendedPageFile
            Write-Verbose "System is running Exchange 2019. Recommended PageFile Size: $recommendedPageFile"

            $recommendedPageFileWording2019 = "On Exchange 2019, the recommended PageFile size is 25% ({0}MB) of the total system memory ({1}MB)."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = ("Error: $recommendedPageFileWording2019" -f $recommendedPageFile, $totalPhysicalMemory)
            } elseif ($recommendedPageFile -ne $pageFileObj.MaxPageSize) {
                $pageFileAdditionalDisplayValue = ("Warning: $recommendedPageFileWording2019" -f $recommendedPageFile, $totalPhysicalMemory)
            } else {
                $displayWriteType = "Grey"
            }
        } elseif ($totalPhysicalMemory -ge 32768) {
            Write-Verbose "System is not running Exchange 2019 and has more than 32GB memory. Recommended PageFile Size: 32778MB"

            $recommendedPageFileWording32GBPlus = "PageFile should be capped at 32778MB for 32GB plus 10MB."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = "Error: $recommendedPageFileWording32GBPlus"
            } elseif ($pageFileObj.MaxPageSize -eq 32778) {
                $displayWriteType = "Grey"
            } else {
                $pageFileAdditionalDisplayValue = "Warning: $recommendedPageFileWording32GBPlus"
            }
        } else {
            $recommendedPageFile = $totalPhysicalMemory + 10
            Write-Verbose "System is not running Exchange 2019 and has less than 32GB of memory. Recommended PageFile Size: $recommendedPageFile"

            $recommendedPageFileWordingBelow32GB = "PageFile is not set to total system memory plus 10MB which should be {0}MB."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = ("Error: $recommendedPageFileWordingBelow32GB" -f $recommendedPageFile)
            } elseif ($recommendedPageFile -ne $pageFileObj.MaxPageSize) {
                $pageFileAdditionalDisplayValue = ("Warning: $recommendedPageFileWordingBelow32GB" -f $recommendedPageFile)
            } else {
                $displayWriteType = "Grey"
            }
        }

        $params = $baseParams + @{
            Name                = "PageFile"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            TestingName         = "PageFile Size $instanceCount"
            DisplayTestingValue = $pageFileObj
        }
        Add-AnalyzedResultInformation @params

        if ($null -ne $pageFileAdditionalDisplayValue) {
            $params = $baseParams + @{
                Details                = $pageFileAdditionalDisplayValue
                DisplayWriteType       = $displayWriteType
                TestingName            = "PageFile Additional Information"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "More information: https://aka.ms/HC-PageFile"
                DisplayWriteType       = $displayWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $instanceCount++
    }

    if ($null -ne $osInformation.PageFile -and
        $osInformation.PageFile.Count -gt 1) {
        $params = $baseParams + @{
            Details                = "`r`n`t`tError: Multiple PageFiles detected. This has been known to cause performance issues, please address this."
            DisplayWriteType       = "Red"
            TestingName            = "Multiple PageFile Detected"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($osInformation.PowerPlan.HighPerformanceSet) {
        $params = $baseParams + @{
            Name             = "Power Plan"
            Details          = $osInformation.PowerPlan.PowerPlanSetting
            DisplayWriteType = "Green"
        }
        Add-AnalyzedResultInformation @params
    } else {
        $params = $baseParams + @{
            Name             = "Power Plan"
            Details          = "$($osInformation.PowerPlan.PowerPlanSetting) --- Error"
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    $displayWriteType = "Grey"
    $displayValue = $osInformation.NetworkInformation.HttpProxy.ProxyAddress

    if (($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne "None") -and
        ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false)) {
        $displayValue = "$($osInformation.NetworkInformation.HttpProxy.ProxyAddress) --- Warning this can cause client connectivity issues."
        $displayWriteType = "Yellow"
    }

    $params = $baseParams + @{
        Name                = "Http Proxy Setting"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $osInformation.NetworkInformation.HttpProxy
    }
    Add-AnalyzedResultInformation @params

    if ($displayWriteType -eq "Yellow") {
        $params = $baseParams + @{
            Name             = "Http Proxy By Pass List"
            Details          = "$($osInformation.NetworkInformation.HttpProxy.ByPassList)"
            DisplayWriteType = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    if (($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne "None") -and
        ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) -and
        ($null -ne $exchangeInformation.GetExchangeServer.InternetWebProxy) -and
        ($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne
        "$($exchangeInformation.GetExchangeServer.InternetWebProxy.Host):$($exchangeInformation.GetExchangeServer.InternetWebProxy.Port)")) {
        $params = $baseParams + @{
            Details                = "Error: Exchange Internet Web Proxy doesn't match OS Web Proxy."
            DisplayWriteType       = "Red"
            TestingName            = "Proxy Doesn't Match"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $displayWriteType2012 = $displayWriteType2013 = "Red"
    $displayValue2012 = $displayValue2013 = $defaultValue = "Error --- Unknown"

    if ($null -ne $osInformation.VcRedistributable) {

        $installed2012 = Get-VisualCRedistributableLatest 2012 $osInformation.VcRedistributable
        $installed2013 = Get-VisualCRedistributableLatest 2013 $osInformation.VcRedistributable

        if (Test-VisualCRedistributableUpToDate -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2012 = "Green"
            $displayValue2012 = "$($installed2012.DisplayVersion) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayValue2012 = "Redistributable ($($installed2012.DisplayVersion)) is outdated"
            $displayWriteType2012 = "Yellow"
        }

        if (Test-VisualCRedistributableUpToDate -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2013 = "Green"
            $displayValue2013 = "$($installed2013.DisplayVersion) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayValue2013 = "Redistributable ($($installed2013.DisplayVersion)) is outdated"
            $displayWriteType2013 = "Yellow"
        }
    }

    $params = $baseParams + @{
        Name             = "Visual C++ 2012 x64"
        Details          = $displayValue2012
        DisplayWriteType = $displayWriteType2012
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) {
        $params = $baseParams + @{
            Name             = "Visual C++ 2013 x64"
            Details          = $displayValue2013
            DisplayWriteType = $displayWriteType2013
        }
        Add-AnalyzedResultInformation @params
    }

    if (($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false -and
            ($displayWriteType2012 -eq "Yellow" -or
            $displayWriteType2013 -eq "Yellow")) -or
        $displayWriteType2012 -eq "Yellow") {

        $params = $baseParams + @{
            Details                = "Note: For more information about the latest C++ Redistributable please visit: https://aka.ms/HC-LatestVC`r`n`t`tThis is not a requirement to upgrade, only a notification to bring to your attention."
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($defaultValue -eq $displayValue2012 -or
        ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false -and
        $displayValue2013 -eq $defaultValue)) {

        $params = $baseParams + @{
            Details                = "ERROR: Unable to find one of the Visual C++ Redistributable Packages. This can cause a wide range of issues on the server.`r`n`t`tPlease install the missing package as soon as possible. Latest C++ Redistributable please visit: https://aka.ms/HC-LatestVC"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $displayValue = "False"
    $writeType = "Grey"

    if ($osInformation.ServerPendingReboot.PendingReboot) {
        $displayValue = "True --- Warning a reboot is pending and can cause issues on the server."
        $writeType = "Yellow"
    }

    $params = $baseParams + @{
        Name                = "Server Pending Reboot"
        Details             = $displayValue
        DisplayWriteType    = $writeType
        DisplayTestingValue = $osInformation.ServerPendingReboot.PendingReboot
    }
    Add-AnalyzedResultInformation @params

    if ($osInformation.ServerPendingReboot.PendingReboot -and
        $osInformation.ServerPendingReboot.PendingRebootLocations.Count -gt 0) {

        foreach ($line in $osInformation.ServerPendingReboot.PendingRebootLocations) {
            $params = $baseParams + @{
                Details                = $line
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
                TestingName            = $line
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-RebootPending"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Reboot More Information"
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerHardwareInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Processor/Hardware Information"  -DisplayOrder $Order)
    }

    $params = $baseParams + @{
        Name                  = "Type"
        Details               = $hardwareInformation.ServerType
        AddHtmlOverviewValues = $true
        HtmlName              = "Hardware Type"
    }
    Add-AnalyzedResultInformation @params

    if ($hardwareInformation.ServerType -eq "Physical" -or
        $hardwareInformation.ServerType -eq "AmazonEC2") {
        $params = $baseParams + @{
            Name    = "Manufacturer"
            Details = $hardwareInformation.Manufacturer
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name    = "Model"
            Details = $hardwareInformation.Model
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "Processor"
        Details = $hardwareInformation.Processor.Name
    }
    Add-AnalyzedResultInformation @params

    $numberOfProcessors = $hardwareInformation.Processor.NumberOfProcessors
    $displayWriteType = "Green"
    $displayValue = $numberOfProcessors

    if ($hardwareInformation.ServerType -ne "Physical") {
        $displayWriteType = "Grey"
    } elseif ($numberOfProcessors -gt 2) {
        $displayWriteType = "Red"
        $displayValue = "$numberOfProcessors - Error: Recommended to only have 2 Processors"
    }

    $params = $baseParams + @{
        Name                = "Number of Processors"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $numberOfProcessors
    }
    Add-AnalyzedResultInformation @params

    $physicalValue = $hardwareInformation.Processor.NumberOfPhysicalCores
    $logicalValue = $hardwareInformation.Processor.NumberOfLogicalCores
    $physicalValueDisplay = $physicalValue
    $logicalValueDisplay = $logicalValue
    $displayWriteTypeLogic = $displayWriteTypePhysical = "Green"

    if (($logicalValue -gt 24 -and
            $exchangeInformation.BuildInformation.VersionInformation.BuildVersion -lt "15.2.0.0") -or
        $logicalValue -gt 48) {
        $displayWriteTypeLogic = "Red"

        if (($physicalValue -gt 24 -and
                $exchangeInformation.BuildInformation.VersionInformation.BuildVersion -lt "15.2.0.0") -or
            $physicalValue -gt 48) {
            $physicalValueDisplay = "$physicalValue - Error"
            $displayWriteTypePhysical = "Red"
        }

        $logicalValueDisplay = "$logicalValue - Error"
    }

    $params = $baseParams + @{
        Name             = "Number of Physical Cores"
        Details          = $physicalValueDisplay
        DisplayWriteType = $displayWriteTypePhysical
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Number of Logical Cores"
        Details               = $logicalValueDisplay
        DisplayWriteType      = $displayWriteTypeLogic
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $displayValue = "Disabled"
    $displayWriteType = "Green"
    $displayTestingValue = $false
    $additionalDisplayValue = [string]::Empty
    $additionalWriteType = "Red"

    if ($logicalValue -gt $physicalValue) {

        if ($hardwareInformation.ServerType -ne "HyperV") {
            $displayValue = "Enabled --- Error: Having Hyper-Threading enabled goes against best practices and can cause performance issues. Please disable as soon as possible."
            $displayTestingValue = $true
            $displayWriteType = "Red"
        } else {
            $displayValue = "Enabled --- Not Applicable"
            $displayTestingValue = $true
            $displayWriteType = "Grey"
        }

        if ($hardwareInformation.ServerType -eq "AmazonEC2") {
            $additionalDisplayValue = "Error: For high-performance computing (HPC) application, like Exchange, Amazon recommends that you have Hyper-Threading Technology disabled in their service. More information: https://aka.ms/HC-EC2HyperThreading"
        }

        if ($hardwareInformation.Processor.Name.StartsWith("AMD")) {
            $additionalDisplayValue = "This script may incorrectly report that Hyper-Threading is enabled on certain AMD processors. Check with the manufacturer to see if your model supports SMT."
            $additionalWriteType = "Yellow"
        }
    }

    $params = $baseParams + @{
        Name                = "Hyper-Threading"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $displayTestingValue
    }
    Add-AnalyzedResultInformation @params

    if (!([string]::IsNullOrEmpty($additionalDisplayValue))) {
        $params = $baseParams + @{
            Details                = $additionalDisplayValue
            DisplayWriteType       = $additionalWriteType
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    #NUMA BIOS CHECK - AKA check to see if we can properly see all of our cores on the box
    $displayWriteType = "Yellow"
    $testingValue = "Unknown"
    $displayValue = [string]::Empty

    if ($hardwareInformation.Model.Contains("ProLiant")) {
        $name = "NUMA Group Size Optimization"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If this is set to Clustered, this can cause multiple types of issues on the server"
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Clustered `r`n`t`tError: This setting should be set to Flat. By having this set to Clustered, we will see multiple different types of issues."
            $testingValue = "Clustered"
            $displayWriteType = "Red"
        } else {
            $displayValue = "Flat"
            $testingValue = "Flat"
            $displayWriteType = "Green"
        }
    } else {
        $name = "All Processor Cores Visible"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If we aren't able to see all processor cores from Exchange, we could see performance related issues."
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Failed `r`n`t`tError: Not all Processor Cores are visible to Exchange and this will cause a performance impact"
            $displayWriteType = "Red"
            $testingValue = "Failed"
        } else {
            $displayWriteType = "Green"
            $displayValue = "Passed"
            $testingValue = "Passed"
        }
    }

    $params = $baseParams + @{
        Name                = $name
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $testingValue
    }
    Add-AnalyzedResultInformation @params

    if ($displayWriteType -ne "Green") {
        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-NUMA"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "Max Processor Speed"
        Details = $hardwareInformation.Processor.MaxMegacyclesPerCore
    }
    Add-AnalyzedResultInformation @params

    if ($hardwareInformation.Processor.ProcessorIsThrottled) {
        $params = $baseParams + @{
            Name                = "Current Processor Speed"
            Details             = "$($hardwareInformation.Processor.CurrentMegacyclesPerCore) --- Error: Processor appears to be throttled."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $hardwareInformation.Processor.CurrentMegacyclesPerCore
        }
        Add-AnalyzedResultInformation @params

        $displayValue = "Error: Power Plan is NOT set to `"High Performance`". This change doesn't require a reboot and takes affect right away. Re-run script after doing so"

        if ($osInformation.PowerPlan.HighPerformanceSet) {
            $displayValue = "Error: Power Plan is set to `"High Performance`", so it is likely that we are throttling in the BIOS of the computer settings."
        }

        $params = $baseParams + @{
            Details             = $displayValue
            DisplayWriteType    = "Red"
            TestingName         = "HighPerformanceSet"
            DisplayTestingValue = $osInformation.PowerPlan.HighPerformanceSet
            AddHtmlDetailRow    = $false
        }
        Add-AnalyzedResultInformation @params
    }

    $totalPhysicalMemory = [System.Math]::Round($hardwareInformation.TotalMemory / 1024 / 1024 / 1024)
    $displayWriteType = "Yellow"
    $displayDetails = [string]::Empty

    if ($exchangeInformation.BuildInformation.VersionInformation.BuildVersion -ge "15.2.0.0") {

        if ($totalPhysicalMemory -gt 256) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 256 GB of Memory" -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 64 -and
            $exchangeInformation.GetExchangeServer.IsEdgeServer -eq $true) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 64GB of RAM installed on the machine." -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 128) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 128GB of RAM installed on the machine." -f $totalPhysicalMemory
        } else {
            $displayDetails = "{0} GB" -f $totalPhysicalMemory
            $displayWriteType = "Grey"
        }
    } elseif ($totalPhysicalMemory -gt 192 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq "Exchange2016") {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 192 GB of Memory." -f $totalPhysicalMemory
    } elseif ($totalPhysicalMemory -gt 96 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq "Exchange2013") {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 96GB of Memory." -f $totalPhysicalMemory
    } else {
        $displayDetails = "{0} GB" -f $totalPhysicalMemory
        $displayWriteType = "Grey"
    }

    $params = $baseParams + @{
        Name                  = "Physical Memory"
        Details               = $displayDetails
        DisplayWriteType      = $displayWriteType
        DisplayTestingValue   = $totalPhysicalMemory
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params
}


function Get-IISAuthenticationType {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig
    )
    begin {

        function GetAuthTypeName {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$AuthType,

                [object]$CurrentAuthLocation,

                [Parameter(Mandatory = $true)]
                [string]$MainLocation,

                [Parameter(Mandatory = $true)]
                [ref]$Completed
            )
            begin {
                $Completed.Value = $false
                $CurrentAuthLocation = $CurrentAuthLocation.$AuthType
                $returnValue = [string]::Empty
            }
            process {
                if ($null -ne $CurrentAuthLocation -and
                    $null -ne $CurrentAuthLocation.enabled) {
                    # Found setting here, set to completed
                    $Completed.Value = $true

                    if ($CurrentAuthLocation.enabled -eq "false") {
                        Write-Verbose "Disabled auth type."
                        return
                    }

                    # evaluate auth types to add to list of enabled.
                    if ($AuthType -eq "anonymousAuthentication") {
                        # provided 'anonymous (default setting)' for the locations that are expected.
                        # API, Autodiscover, ecp, ews, OWA (BE), Default Web Site, Exchange Back End
                        # use MainLocation because that is the location we are evaluating
                        if ($MainLocation -like "*/API" -or
                            $MainLocation -like "*/Autodiscover" -or
                            $MainLocation -like "*/ecp" -or
                            $MainLocation -like "*/EWS" -or
                            $MainLocation -eq "Exchange Back End/OWA" -or
                            $MainLocation -eq "Default Web Site" -or
                            $MainLocation -eq "Exchange Back End") {
                            $returnValue = "anonymous (default setting)"
                        } else {
                            $returnValue = "anonymous (NOT default setting)"
                        }
                    } elseif ($AuthType -eq "windowsAuthentication") {
                        # If clear is set, we only use the value here
                        # If clear is set, we add to the default location of provider types.

                        if ($null -ne $CurrentAuthLocation.providers.clear -or
                            $null -eq $defaultWindowsAuthProviders -or
                            $defaultWindowsAuthProviders.Count -eq 0) {

                            if ($null -ne $CurrentAuthLocation.providers.add.value) {
                                $returnValue = "Windows ($($CurrentAuthLocation.providers.add.value -join ","))"
                            } else {
                                $returnValue = "Windows (No providers)" # This could be a problem??
                            }
                        } else {
                            $localAuthProviders = @($defaultWindowsAuthProviders)

                            if ($null -ne $CurrentAuthLocation.providers.add.value) {
                                $localAuthProviders += $CurrentAuthLocation.providers.add.value
                            }

                            $returnValue = "Windows ($($localAuthProviders -join ","))"
                        }
                    } else {
                        $returnValue = $AuthType.Replace("Authentication", "").Replace("ClientCertificateMapping", "Cert")
                    }
                } else {
                    # If not set here, we need to look at the parent
                    Write-Verbose "Not set at current location. Need to look at parent."
                }
            } end {
                if (-not ([string]::IsNullOrEmpty($returnValue))) { Write-Verbose "Value Set: $returnValue" }

                return $returnValue
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $getIisAuthenticationType = @{}
        $appHostConfigLocations = $ApplicationHostConfig.configuration.Location.path | Where-Object { $_ -ne "" }
        $defaultWindowsAuthProviders = @($ApplicationHostConfig.configuration.'system.webServer'.security.authentication.windowsAuthentication.providers.add.value)
        $authenticationTypes = @("windowsAuthentication", "anonymousAuthentication", "digestAuthentication", "basicAuthentication",
            "clientCertificateMappingAuthentication", "iisClientCertificateMappingAuthentication")
        $failedKey = "FailedLocations"
        $getIisAuthenticationType.Add($failedKey, (New-Object System.Collections.Generic.List[object]))
    }
    process {
        # foreach each location, we need to look for each $authenticationTypes up the stack ordering to determine if it is enabled or not.
        # for this configuration type, clear flag doesn't appear to be used at all.
        foreach ($appKey in $appHostConfigLocations) {
            Write-Verbose "Working on appKey: $appKey"

            if (-not ($getIisAuthenticationType.ContainsKey($appKey))) {
                $getIisAuthenticationType.Add($appKey, [string]::Empty)
            }

            $currentKey = $appKey
            $authentication = @()
            $continue = $true
            $authenticationTypeCompleted = @{}
            $authenticationTypes | ForEach-Object { $authenticationTypeCompleted.Add($_, $false) }

            do {
                # to avoid doing a lot of loops, evaluate each location for all the authentication types before moving up a level.
                Write-Verbose "Working on currentKey: $currentKey"
                $location = $ApplicationHostConfig.SelectNodes("/configuration/location[@path = '$currentKey']")

                if ($null -ne $location -and
                    $null -ne $location.path) {
                    $authLocation = $location.'system.webServer'.security.authentication

                    if ($null -ne $authLocation) {
                        # look over each auth type
                        foreach ($authenticationType in $authenticationTypes) {
                            if ($authenticationTypeCompleted[$authenticationType]) {
                                # we already have this auth type evaluated don't use this setting here.
                                continue
                            }

                            Write-Verbose "Evaluating current authenticationType: $authenticationType"
                            $didComplete = $false
                            $params = @{
                                AuthType            = $authenticationType
                                CurrentAuthLocation = $authLocation
                                MainLocation        = $appKey
                                Completed           = [ref]$didComplete
                            }

                            $value = GetAuthTypeName @params
                            if ($didComplete) {
                                $authenticationTypeCompleted[$authenticationType] = $true

                                if (-not ([string]::IsNullOrEmpty($value))) {
                                    $authentication += $value
                                }
                            }
                        }
                        $continue = $null -ne ($authenticationTypeCompleted.Values | Where-Object { $_ -eq $false })

                        if ($continue) {
                            $index = $currentKey.LastIndexOf("/")

                            if ($index -eq -1) {
                                $continue = $false
                                $defaultAuthLocation = $ApplicationHostConfig.configuration.'system.webServer'.security.authentication

                                foreach ($authenticationType in $authenticationTypes) {
                                    if ($authenticationTypeCompleted[$authenticationType]) {
                                        # we already have this auth type evaluated don't use this setting here.
                                        continue
                                    }

                                    Write-Verbose "Evaluating global current authenticationType: $authenticationType"
                                    $didComplete = $false
                                    $params = @{
                                        AuthType            = $authenticationType
                                        CurrentAuthLocation = $defaultAuthLocation
                                        MainLocation        = $appKey
                                        Completed           = [ref]$didComplete
                                    }

                                    $value = GetAuthTypeName @params
                                    if ($didComplete) {
                                        $authenticationTypeCompleted[$authenticationType] = $true

                                        if (-not ([string]::IsNullOrEmpty($value))) {
                                            $authentication += $value
                                        }
                                    }
                                }
                            } else {
                                $currentKey = $currentKey.Substring(0, $index)
                            }
                        }
                    } else {
                        Write-Verbose "authLocation was NULL, but shouldn't be a problem we just use the parent."
                        $index = $currentKey.LastIndexOf("/")

                        if ($index -eq -1) {
                            $continue = $false
                            Write-Verbose "No parent location. Need to determine how to address."
                            $getIisAuthenticationType[$failedKey].Add($appKey)
                        } else {
                            $currentKey = $currentKey.Substring(0, $index)
                        }
                    }
                } elseif ($currentKey -ne $appKey) {
                    # If we are at a parent location we might not have all the locations in the config. So this could be okay.
                    Write-Verbose "Couldn't find location for '$currentKey'. Keep on looking"
                    $index = $currentKey.LastIndexOf("/")

                    if ($index -eq -1) {
                        Write-Verbose "Didn't have root parent in the config file, this is odd."
                        $getIisAuthenticationType[$failedKey].Add($appKey)
                        $continue = $false
                    } else {
                        $currentKey = $currentKey.Substring(0, $index)
                    }
                } else {
                    Write-Verbose "Couldn't find location. This shouldn't occur."
                    # Add to failed key to display issue
                    $getIisAuthenticationType[$failedKey].Add($appKey)
                }
            } while ($continue)

            $getIisAuthenticationType[$appKey] = $authentication
            Write-Verbose "Found auth types for enabled for '$appKey': $($authentication -join ",")"
        }
    }
    end {
        return $getIisAuthenticationType
    }
}

function Get-IPFilterSetting {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $locationPaths = $ApplicationHostConfig.configuration.location.path | Where-Object { $_ -ne "" }
        $ipFilterSettings = @{}
    }
    process {
        foreach ($appKey in $locationPaths) {
            Write-Verbose "Working on appKey: $appKey"

            if (-not ($ipFilterSettings.ContainsKey($appKey))) {
                $ipFilterSettings.Add($appKey, (New-Object System.Collections.Generic.List[object]))
            }

            $currentKey = $appKey
            $continue = $true

            do {
                Write-Verbose "Working on currentKey: $currentKey"
                $location = $ApplicationHostConfig.SelectNodes("/configuration/location[@path = '$currentKey']")

                if ($null -ne $location) {
                    $ipSecurity = $location.'system.webServer'.security.ipSecurity

                    if ($null -ne $ipSecurity) {
                        $clear = $null -ne $ipSecurity.clear
                        $ipFilterSettings[$appKey].Add($ipSecurity)
                    }
                } else {
                    Write-Verbose "Couldn't find location. This shouldn't occur."
                }

                if ($clear) {
                    Write-Verbose "Clear was set, don't need to know what else was set."
                    $continue = $false
                } else {
                    $index = $currentKey.LastIndexOf("/")

                    if ($index -eq -1) {
                        $continue = $false

                        # look at the global configuration applicationHost.config
                        $ipSecurity = $ApplicationHostConfig.configuration.'system.webServer'.security.ipSecurity

                        # Need to check for if it is an empty string, if it is, we don't need to worry about it.
                        if ($null -ne $ipSecurity -and
                            $ipSecurity.GetType().Name -ne "string") {
                            $add = $null -ne ($ipSecurity | Get-Member | Where-Object { $_.MemberType -eq "Property" -and $_.Name -ne "allowUnlisted" })

                            if ($add) {
                                $ipFilterSettings[$appKey].Add($ipSecurity)
                            }
                        } else {
                            Write-Verbose "No ipSecurity set globally"
                        }
                    } else {
                        $currentKey = $currentKey.Substring(0, $index)
                    }
                }
            } while ($continue)
        }
    }
    end {
        return $ipFilterSettings
    }
}

<#
.SYNOPSIS
 Pulls out URL Rewrite Rules from the web.config and applicationHost.config file to return a Hashtable of those settings.
.DESCRIPTION
 This is a function that is designed to pull out the URL Rewrite Rules that are set on a location of IIS.
 Because you can set it on an individual web.config file or the parent site(s), or the ApplicationHostConfig file for the location
 We need to check all locations to properly determine what is all set.
 The ApplicationHostConfig file must be able to be converted to Xml, but the web.config file doesn't.
 The order goes like this it appears based off testing done, if overrides are allowed which by default for URL Rewrite that is true.
    1. Current IIS Location for web.config for virtual directory
    2. ApplicationHost.config file for the same location
    3. Then move up one level (Default Web Site/mapi -> Default Web Site) and repeat 1 and 2 till no more locations.
        a. If the 'clear' flag was set at any point, we stop at that location in the process.
    4. Then there is a global setting in the ApplicationHost.config file.
#>
function Get-URLRewriteRule {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        # Key = IIS Location (Example: Default Web Site/mapi)
        # Value = web.config content
        [Parameter(Mandatory = $true)]
        [hashtable]$WebConfigContent
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $urlRewriteRules = @{}
        $appHostConfigLocations = $ApplicationHostConfig.configuration.Location.path
    }
    process {
        foreach ($key in $WebConfigContent.Keys) {
            Write-Verbose "Working on key: $key"
            $continue = $true
            $clear = $false
            $currentKey = $key
            $urlRewriteRules.Add($key, (New-Object System.Collections.Generic.List[object]))

            do {
                Write-Verbose "Working on currentKey: $currentKey"
                try {
                    # the Web.config is looked at first
                    [xml]$content = $WebConfigContent[$currentKey]
                    $rules = $content.configuration.'system.webServer'.rewrite.rules

                    if ($null -ne $rules) {
                        $clear = $null -ne $rules.clear
                        $urlRewriteRules[$key].Add($rules)
                    } else {
                        Write-Verbose "No rewrite rules in the config file"
                    }
                } catch {
                    Write-Verbose "Failed to convert to xml"
                    Invoke-CatchActions
                }

                if (-not $clear) {
                    # Now need to look at the applicationHost.config file to determine what is set at that location.
                    # need to do this because of the case sensitive query to get the xmlNode
                    Write-Verbose "clear not set on config. Looking at the applicationHost.config file"
                    $appKey = $appHostConfigLocations | Where-Object { $_ -eq $currentKey }

                    if ($appKey.Count -eq 1) {
                        $location = $ApplicationHostConfig.SelectNodes("/configuration/location[@path = '$appKey']")

                        if ($null -ne $location) {
                            $rules = $location.'system.webServer'.rewrite.rules

                            if ($null -ne $rules) {
                                $clear = $null -ne $rules.clear
                                $urlRewriteRules[$key].Add($rules)
                            } else {
                                Write-Verbose 'No rewrite rules in the applicationHost.config file'
                            }
                        } else {
                            Write-Verbose "We didn't find the location for '$appKey' in the applicationHostConfig. This shouldn't occur."
                        }
                    } else {
                        Write-Verbose "Multiple appKeys locations found for currentKey"
                    }
                }

                if ($clear) {
                    Write-Verbose "Clear was set, don't need to know what else was set."
                    $continue = $false
                } else {
                    $index = $currentKey.LastIndexOf("/")

                    if ($index -eq -1) {
                        $continue = $false
                        # look at the global configuration of the applicationHost.config file
                        $rules = $ApplicationHostConfig.configuration.'system.webServer'.rewrite.rules

                        if ($null -ne $rules) {
                            $urlRewriteRules[$key].Add($rules)
                        } else {
                            Write-Verbose "No global configuration for rewrite rules."
                        }
                    } else {
                        $currentKey = $currentKey.Substring(0, $index)
                    }
                }
            } while ($continue)

            Write-Verbose "Completed key: $key"
        }
    }
    end {
        return $urlRewriteRules
    }
}
function Invoke-AnalyzerIISInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Exchange IIS Information"  -DisplayOrder $Order)
    }

    if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $true) {
        Write-Verbose "No IIS information to review on an Edge Server"
        return
    }

    if ($null -eq $exchangeInformation.IISSettings.IISWebApplication -and
        $null -eq $exchangeInformation.IISSettings.IISWebSite -and
        $null -eq $exchangeInformation.IISSettings.IISSharedWebConfig) {
        Write-Verbose "Wasn't able find any other IIS settings, likely due to application host config file being messed up."

        if ($null -ne $exchangeInformation.IISSettings.ApplicationHostConfig) {
            Write-Verbose "Wasn't able find any other IIS settings, likely due to application host config file being messed up."
            try {
                [xml]$exchangeInformation.IISSettings.ApplicationHostConfig | Out-Null
                Write-Verbose "Application Host Config file is in a readable file, not sure how we got here."
                $displayIISIssueToReport = $true
            } catch {
                Invoke-CatchActions
                Write-Verbose "Confirmed Application Host Config file isn't in a readable xml format."
                $params = $baseParams + @{
                    Name                = "Invalid Configuration File"
                    Details             = "Application Host Config File: '$($env:WINDIR)\System32\inetSrv\config\applicationHost.config'"
                    DisplayWriteType    = "Red"
                    TestingName         = "Invalid Configuration File - Application Host Config File"
                    DisplayTestingValue = $true
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            Write-Verbose "No application host config file was collected either. not sure how we got here."
            $displayIISIssueToReport = $true
        }

        if ($displayIISIssueToReport) {
            $params = $baseParams + @{
                Name             = "Unknown IIS configuration"
                Details          = "Please report this to ExToolsFeedback@microsoft.com"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
        # Nothing to process if we don't have the information.
        return
    }

    ###################################
    # IIS Web Sites - Standard Display
    ###################################

    Write-Verbose "Working on IIS Web Sites"
    $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]
    $problemCertList = New-Object System.Collections.Generic.List[string]
    $iisWebSites = $exchangeInformation.IISSettings.IISWebSite | Sort-Object ID
    $bindingsPropertyName = "Protocol - Bindings - Certificate"

    foreach ($webSite in $iisWebSites) {
        $protocolLength = 0
        $bindingInformationLength = 0

        $webSite.Bindings.Protocol |
            ForEach-Object { if ($protocolLength -lt $_.Length) { $protocolLength = $_.Length } }
        $webSite.Bindings.bindingInformation |
            ForEach-Object { if ($bindingInformationLength -lt $_.Length) { $bindingInformationLength = $_.Length } }

        $hstsEnabled = $webSite.Hsts.NativeHstsSettings.enabled -eq $true -or $webSite.Hsts.HstsViaCustomHeader.enabled -eq $true

        $value = @($webSite.Bindings | ForEach-Object {
                $pSpace = [string]::Empty
                $biSpace = [string]::Empty
                $certHash = "NULL"

                if (-not ([string]::IsNullOrEmpty($_.certificateHash))) {
                    $certHash = $_.certificateHash
                    $cert = $exchangeInformation.ExchangeCertificates | Where-Object { $_.Thumbprint -eq $certHash }

                    if ($null -eq $cert) {
                        $problemCertList.Add("'$certHash' Doesn't exist on the server and this will cause problems.")
                    } elseif ($cert.LifetimeInDays -lt 0) {
                        $problemCertList.Add("'$certHash' Has expired and will cause problems.")
                    }
                }

                1..(($protocolLength - $_.Protocol.Length) + 1) | ForEach-Object { $pSpace += " " }
                1..(($bindingInformationLength - $_.bindingInformation.Length) + 1 ) | ForEach-Object { $biSpace += " " }
                return "$($_.Protocol)$($pSpace)- $($_.bindingInformation)$($biSpace)- $certHash"
            })

        $outputObjectDisplayValue.Add([PSCustomObject]@{
                Name                  = $webSite.Name
                State                 = $webSite.State
                "HSTS Enabled"        = $hstsEnabled
                $bindingsPropertyName = $value
            })
    }

    #Used for Web App Pools as well
    $sbStarted = { param($o, $p) if ($p -eq "State") { if ($o."$p" -eq "Started") { "Green" } else { "Red" } } }

    $params = $baseParams + @{
        OutColumns           = ([PSCustomObject]@{
                DisplayObject      = $outputObjectDisplayValue
                ColorizerFunctions = @($sbStarted)
                IndentSpaces       = 8
            })
        OutColumnsColorTests = @($sbStarted)
        HtmlName             = "IIS Sites Information"
    }
    Add-AnalyzedResultInformation @params

    if ($problemCertList.Count -gt 0) {

        foreach ($details in $problemCertList) {
            $params = $baseParams + @{
                Name             = "Certificate Binding Issue Detected"
                Details          = $details
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    ########################
    # IIS Web Sites - Issues
    ########################

    if (($iisWebSites.Hsts.NativeHstsSettings.enabled -notcontains $true) -and
        ($iisWebSites.Hsts.HstsViaCustomHeader.enabled -notcontains $true)) {
        Write-Verbose "Skipping over HSTS issues, as it isn't enabled"
    } else {
        $showAdditionalHstsInformation = $false

        foreach ($webSite in $iisWebSites) {
            $hstsConfiguration = $null
            $isExchangeBackEnd = $webSite.Name -eq "Exchange Back End"
            $hstsMaxAgeWriteType = "Green"

            if (($webSite.Hsts.NativeHstsSettings.enabled) -or
                ($webSite.Hsts.HstsViaCustomHeader.enabled)) {
                $params = $baseParams + @{
                    Name                = "HSTS Enabled"
                    Details             = "$($webSite.Name)"
                    TestingName         = "hsts-Enabled-$($webSite.Name)"
                    DisplayTestingValue = $true
                    DisplayWriteType    = if ($isExchangeBackEnd) { "Red" } else { "Green" }
                }
                Add-AnalyzedResultInformation @params

                if ($isExchangeBackEnd) {
                    $showAdditionalHstsInformation = $true
                    $params = $baseParams + @{
                        Details                = "HSTS on 'Exchange Back End' is not supported and can cause issues"
                        DisplayWriteType       = "Red"
                        TestingName            = "hsts-BackendNotSupported"
                        DisplayTestingValue    = $true
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                if (($webSite.Hsts.NativeHstsSettings.enabled) -and
                ($webSite.Hsts.HstsViaCustomHeader.enabled)) {
                    $showAdditionalHstsInformation = $true
                    Write-Verbose "HSTS conflict detected"
                    $params = $baseParams + @{
                        Details                = ("HSTS configured via customHeader and native IIS control - please remove one configuration" +
                            "`r`n`t`tHSTS native IIS control has a higher weight than the customHeader and will be used")
                        DisplayWriteType       = "Yellow"
                        TestingName            = "hsts-conflict"
                        DisplayTestingValue    = $true
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                if ($webSite.Hsts.NativeHstsSettings.enabled) {
                    Write-Verbose "HSTS configured via native IIS control"
                    $hstsConfiguration = $webSite.Hsts.NativeHstsSettings
                } else {
                    Write-Verbose "HSTS configured via customHeader"
                    $hstsConfiguration = $webSite.Hsts.HstsViaCustomHeader
                }

                $maxAgeValue = $hstsConfiguration.'max-age'
                if ($maxAgeValue -lt 31536000) {
                    $showAdditionalHstsInformation = $true
                    $hstsMaxAgeWriteType = "Yellow"
                }
                $params = $baseParams + @{
                    Details                = "max-age: $maxAgeValue"
                    DisplayWriteType       = $hstsMaxAgeWriteType
                    TestingName            = "hsts-max-age-$($webSite.Name)"
                    DisplayTestingValue    = $maxAgeValue
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params

                $params = $baseParams + @{
                    Details                = "includeSubDomains: $($hstsConfiguration.includeSubDomains)"
                    TestingName            = "hsts-includeSubDomains-$($webSite.Name)"
                    DisplayTestingValue    = $hstsConfiguration.includeSubDomains
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params

                $params = $baseParams + @{
                    Details                = "preload: $($hstsConfiguration.preload)"
                    TestingName            = "hsts-preload-$($webSite.Name)"
                    DisplayTestingValue    = $hstsConfiguration.preload
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params

                $redirectHttpToHttpsConfigured = $hstsConfiguration.redirectHttpToHttps
                $params = $baseParams + @{
                    Details                = "redirectHttpToHttps: $redirectHttpToHttpsConfigured"
                    TestingName            = "hsts-redirectHttpToHttps-$($webSite.Name)"
                    DisplayTestingValue    = $redirectHttpToHttpsConfigured
                    DisplayCustomTabNumber = 2
                }
                if ($redirectHttpToHttpsConfigured) {
                    $showAdditionalHstsInformation = $true
                    $params.Add("DisplayWriteType", "Red")
                }
                Add-AnalyzedResultInformation @params
            }
        }

        if ($showAdditionalHstsInformation) {
            $params = $baseParams + @{
                Details                = "`r`n`t`tMore Information about HSTS: https://aka.ms/HC-HSTS"
                DisplayWriteType       = "Yellow"
                TestingName            = 'hsts-MoreInfo'
                DisplayTestingValue    = $true
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    }

    ########################
    # IIS Web App Pools
    ########################

    Write-Verbose "Working on Exchange Web App GC Mode"

    $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

    foreach ($webAppKey in $exchangeInformation.ApplicationPools.Keys) {

        $appPool = $exchangeInformation.ApplicationPools[$webAppKey]
        $appRestarts = $appPool.AppSettings.add.recycling.periodicRestart
        $appRestartSet = ($appRestarts.PrivateMemory -ne "0" -or
            $appRestarts.Memory -ne "0" -or
            $appRestarts.Requests -ne "0" -or
            $null -ne $appRestarts.Schedule -or
            ($appRestarts.Time -ne "00:00:00" -and
                ($webAppKey -ne "MSExchangeOWAAppPool" -and
            $webAppKey -ne "MSExchangeECPAppPool")))

        $outputObjectDisplayValue.Add(([PSCustomObject]@{
                    AppPoolName         = $webAppKey
                    State               = $appPool.AppSettings.state
                    GCServerEnabled     = $appPool.GCServerEnabled
                    RestartConditionSet = $appRestartSet
                })
        )
    }

    $sbRestart = { param($o, $p) if ($p -eq "RestartConditionSet") { if ($o."$p") { "Red" } else { "Green" } } }
    $params = $baseParams + @{
        OutColumns           = ([PSCustomObject]@{
                DisplayObject      = $outputObjectDisplayValue
                ColorizerFunctions = @($sbStarted, $sbRestart)
                IndentSpaces       = 8
            })
        OutColumnsColorTests = @($sbStarted, $sbRestart)
        HtmlName             = "Application Pool Information"
    }
    Add-AnalyzedResultInformation @params

    $periodicStartAppPools = $outputObjectDisplayValue | Where-Object { $_.RestartConditionSet -eq $true }

    if ($null -ne $periodicStartAppPools) {

        $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

        foreach ($appPool in $periodicStartAppPools) {
            $periodicRestart = $exchangeInformation.ApplicationPools[$appPool.AppPoolName].AppSettings.add.recycling.periodicRestart
            $schedule = $periodicRestart.Schedule

            if ([string]::IsNullOrEmpty($schedule)) {
                $schedule = "null"
            }

            $outputObjectDisplayValue.Add(([PSCustomObject]@{
                        AppPoolName   = $appPool.AppPoolName
                        PrivateMemory = $periodicRestart.PrivateMemory
                        Memory        = $periodicRestart.Memory
                        Requests      = $periodicRestart.Requests
                        Schedule      = $schedule
                        Time          = $periodicRestart.Time
                    }))
        }

        $sbColorizer = {
            param($o, $p)
            switch ($p) {
                { $_ -in "PrivateMemory", "Memory", "Requests" } {
                    if ($o."$p" -eq "0") { "Green" } else { "Red" }
                }
                "Time" {
                    if ($o."$p" -eq "00:00:00") { "Green" } else { "Red" }
                }
                "Schedule" {
                    if ($o."$p" -eq "null") { "Green" } else { "Red" }
                }
            }
        }

        $params = $baseParams + @{
            OutColumns           = ([PSCustomObject]@{
                    DisplayObject      = $outputObjectDisplayValue
                    ColorizerFunctions = @($sbColorizer)
                    IndentSpaces       = 8
                })
            OutColumnsColorTests = @($sbColorizer)
            HtmlName             = "Application Pools Restarts"
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details          = "Error: The above app pools currently have the periodic restarts set. This restart will cause disruption to end users."
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    ########################################
    # Virtual Directories - Standard display
    ########################################

    $applicationHostConfig = $exchangeInformation.IISSettings.ApplicationHostConfig
    $defaultWebSitePowerShellSslEnabled = $false
    $defaultWebSitePowerShellAuthenticationEnabled = $false
    $iisWebSettings = @($exchangeInformation.IISSettings.IISWebApplication)
    $iisWebSettings += @($exchangeInformation.IISSettings.IISWebSite)
    $iisConfigurationSettings = @($exchangeInformation.IISSettings.IISWebApplication.ConfigurationFileInfo)
    $iisConfigurationSettings += $iisWebSiteConfigs = @($exchangeInformation.IISSettings.IISWebSite.ConfigurationFileInfo)
    $iisConfigurationSettings += @($exchangeInformation.IISSettings.IISSharedWebConfig)
    $extendedProtectionConfiguration = $exchangeInformation.ExtendedProtectionConfig.ExtendedProtectionConfiguration
    $displayMainSitesList = @("Default Web Site", "API", "Autodiscover", "ecp", "EWS", "mapi", "Microsoft-Server-ActiveSync", "Proxy", "OAB", "owa",
        "PowerShell", "Rpc", "Exchange Back End", "emsmdb", "nspi", "RpcWithCert")
    $iisVirtualDirectoriesDisplay = New-Object 'System.Collections.Generic.List[System.Object]'
    $iisWebConfigContent = @{}
    $iisLocations = ([xml]$applicationHostConfig).configuration.Location | Sort-Object Path

    $iisWebSettings | ForEach-Object {
        $key = if ($null -ne $_.FriendlyName) { $_.FriendlyName } else { $_.Name }

        if ($null -ne $key) {
            $iisWebConfigContent.Add($key, $_.ConfigurationFileInfo.Content)
        } else {
            Write-Verbose "Failed to set Key for iisWebConfigContent hashtable because it was null."
        }
    }

    $ruleParams = @{
        ApplicationHostConfig = [xml]$applicationHostConfig
        WebConfigContent      = $iisWebConfigContent
    }

    $urlRewriteRules = Get-URLRewriteRule @ruleParams
    $ipFilterSettings = Get-IPFilterSetting -ApplicationHostConfig ([xml]$applicationHostConfig)
    $authTypeSettings = Get-IISAuthenticationType -ApplicationHostConfig ([xml]$applicationHostConfig)
    $failedLocationsForAuth = @()
    Write-Verbose "Evaluating the IIS Locations for display"

    foreach ($location in $iisLocations) {

        if ([string]::IsNullOrEmpty($location.Path)) { continue }

        if ($displayMainSitesList -notcontains ($location.Path.Split("/")[-1])) { continue }

        Write-Verbose "Working on IIS Path: $($location.Path)"
        $sslFlag = [string]::Empty
        $displayRewriteRules = [string]::Empty
        #TODO: This is not 100% accurate because you can have a disabled rule here.
        # However, not sure how common this is going to be so going to ignore this for now.
        $ipFilterEnabled = $ipFilterSettings[$location.Path].Count -ne 0
        $epValue = "None"
        $ep = $extendedProtectionConfiguration | Where-Object { $_.VirtualDirectoryName -eq $location.Path }
        $currentRewriteRules = $urlRewriteRules[$location.Path]
        $authentication = $authTypeSettings[$location.Path]

        if ($currentRewriteRules.Count -ne 0) {
            # Need to loop through all the rules first to find the excluded rules
            # then find the rules to display
            $excludeRules = @()
            foreach ($rule in $currentRewriteRules) {
                $remove = $rule.Remove

                if ($null -ne $remove) {
                    $excludeRules += $remove.Name
                }
            }

            $displayRewriteRules = ($currentRewriteRules.rule | Where-Object { $_.enabled -ne "false" }).name |
                Where-Object { $_ -notcontains $excludeRules }
        }

        if ($null -ne $ep) {
            Write-Verbose "Using EP settings to determine sslFlags"
            $sslSettings = $ep.Configuration.SslSettings
            $sslFlag = "$($sslSettings.RequireSSL) $(if($sslSettings.Ssl128Bit) { "(128-bit)" })".Trim()

            if ($sslSettings.ClientCertificate -ne "Ignore") {
                $sslFlag = @($sslFlag, "Cert($($sslSettings.ClientCertificate))")
            }

            $epValue = $ep.ExtendedProtection
        } else {
            Write-Verbose "Not using EP settings to determine sslFlags, skipping over cert auth logic."
            $ssl = $location.'system.webServer'.security.access.SslFlags
            $sslFlag = "$($ssl -contains "ssl") $(if(($ssl -contains "ssl128")) { "(128-bit)" })".Trim()
        }

        if ($location.Path -eq "Default Web Site/PowerShell") {
            $defaultWebSitePowerShellSslEnabled = $sslFlag -contains "true"
            $defaultWebSitePowerShellAuthenticationEnabled = -not [string]::IsNullOrEmpty($authentication)
        }

        $iisVirtualDirectoriesDisplay.Add([PSCustomObject]@{
                Name               = $location.Path
                ExtendedProtection = $epValue
                SslFlags           = $sslFlag
                IPFilteringEnabled = $ipFilterEnabled
                URLRewrite         = $displayRewriteRules
                Authentication     = $authentication
            })
    }

    $params = $baseParams + @{
        OutColumns = ([PSCustomObject]@{
                DisplayObject = $iisVirtualDirectoriesDisplay
                IndentSpaces  = 8
            })
        HtmlName   = "Virtual Directory Locations"
    }
    Add-AnalyzedResultInformation @params

    if ($failedLocationsForAuth.Count -gt 0) {
        $params = $baseParams + @{
            Name             = "Inaccurate display of authentication types"
            Details          = $failedLocationsForAuth -join ","
            DisplayWriteType = "Yellow"
        }

        Add-AnalyzedResultInformation @params
    }

    ###############################
    # Virtual Directories - Issues
    ###############################

    # Invalid configuration files are ones that we can't convert to xml.
    $invalidConfigurationFile = $iisConfigurationSettings | Where-Object { $_.Valid -eq $false -and $_.Exist -eq $true }
    # If a web application config file doesn't truly exists, we end up using the parent web.config file
    # If any of the web application config file paths match a parent path, that is a problem.
    # only collect the ones that are valid, if not valid we will assume that the child web apps will point to it and can be misleading.
    $siteConfigPaths = $iisWebSiteConfigs |
        Where-Object { $_.Valid -eq $true -and $_.Exist -eq $true } |
        ForEach-Object { $_.Location }

    $iisWebApplications = $exchangeInformation.IISSettings.IISWebApplication

    if ($null -ne $siteConfigPaths) {
        $missingWebApplicationConfigFile = $iisWebApplications |
            Where-Object { $siteConfigPaths -contains "$($_.ConfigurationFileInfo.Location)" }
    }

    $correctLocations = @{
        "Default Web Site/owa"                          = "FrontEnd\HttpProxy\owa"
        "Default Web Site/ecp"                          = "FrontEnd\HttpProxy\ecp"
        "Default Web Site/EWS"                          = "FrontEnd\HttpProxy\EWS"
        "Default Web Site/API"                          = "FrontEnd\HttpProxy\Rest"
        "Default Web Site/Autodiscover"                 = "FrontEnd\HttpProxy\Autodiscover"
        "Default Web Site/Microsoft-Server-ActiveSync"  = "FrontEnd\HttpProxy\sync"
        "Default Web Site/OAB"                          = "FrontEnd\HttpProxy\OAB"
        "Default Web Site/PowerShell"                   = "FrontEnd\HttpProxy\PowerShell"
        "Default Web Site/mapi"                         = "FrontEnd\HttpProxy\mapi"
        "Default Web Site/Rpc"                          = "FrontEnd\HttpProxy\rpc"
        "Exchange Back End/PowerShell"                  = "ClientAccess\PowerShell-Proxy"
        "Exchange Back End/mapi/emsmdb"                 = "ClientAccess\mapi\emsmdb"
        "Exchange Back End/mapi/nspi"                   = "ClientAccess\mapi\nspi"
        "Exchange Back End/API"                         = "ClientAccess\rest"
        "Exchange Back End/owa"                         = "ClientAccess\owa"
        "Exchange Back End/OAB"                         = "ClientAccess\OAB"
        "Exchange Back End/ecp"                         = "ClientAccess\ecp"
        "Exchange Back End/Autodiscover"                = "ClientAccess\Autodiscover"
        "Exchange Back End/Microsoft-Server-ActiveSync" = "ClientAccess\sync"
        "Exchange Back End/EWS"                         = "ClientAccess\exchWeb\EWS"
        "Exchange Back End/EWS/bin"                     = "ClientAccess\exchWeb\EWS\bin"
        "Exchange Back End/Rpc"                         = "RpcProxy"
        "Exchange Back End/RpcWithCert"                 = "RpcProxy"
        "Exchange Back End/PushNotifications"           = "ClientAccess\PushNotifications"
    }

    # Missing config file should really only occur for SharedWebConfig files, as the web application would go back to the parent site.
    $missingSharedConfigFile = @($exchangeInformation.IISSettings.IISSharedWebConfig) | Where-Object { $_.Exist -eq $false }
    $missingConfigFiles = $iisWebSettings | Where-Object { $_.ConfigurationFileInfo.Exist -eq $false }
    $defaultVariableDetected = $iisConfigurationSettings | Where-Object { $null -ne ($_.Content | Select-String "%ExchangeInstallDir%") }
    $binSearchFoldersNotFound = $iisConfigurationSettings |
        Where-Object { $_.Location -like "*\ClientAccess\ecp\web.config" -and $_.Exist -eq $true -and $_.Valid -eq $true } |
        Where-Object {
            $binSearchFolders = (([xml]($_.Content)).configuration.appSettings.add | Where-Object {
                    $_.key -eq "BinSearchFolders"
                }).value
            $paths = $binSearchFolders.Split(";").Trim()
            $paths | ForEach-Object { Write-Verbose "BinSearchFolder: $($_)" }
            $installPath = $exchangeInformation.RegistryValues.MsiInstallPath
            foreach ($binTestPath in  @("bin", "bin\CmdletExtensionAgents", "ClientAccess\Owa\bin")) {
                $testPath = [System.IO.Path]::Combine($installPath, $binTestPath)
                Write-Verbose "Testing path: $testPath"
                if (-not ($paths -contains $testPath)) {
                    return $_
                }
            }
        }

    # Display URL Rewrite Rules.
    # To save on space, don't display rules that are on multiple vDirs by same name.
    # Use 'DisplayKey' for the display results.
    $alreadyDisplayedUrlRewriteRules = @{}
    $alreadyDisplayedUrlKey = "DisplayKey"
    $urlMatchProblem = "UrlMatchProblem"
    $alreadyDisplayedUrlRewriteRules.Add($alreadyDisplayedUrlKey, (New-Object System.Collections.Generic.List[object]))
    $alreadyDisplayedUrlRewriteRules.Add($urlMatchProblem, (New-Object System.Collections.Generic.List[string]))

    foreach ($key in $urlRewriteRules.Keys) {
        $currentSection = $urlRewriteRules[$key]

        if ($currentSection.Count -ne 0) {
            foreach ($rule in $currentSection.rule) {

                if ($null -eq $rule) {
                    Write-Verbose "Rule is NULL skipping."
                    continue
                } elseif ($rule.enabled -eq "false") {
                    # skip over disabled rules.
                    Write-Verbose "skipping over disabled rule: $($rule.Name) for vDir '$key'"
                    continue
                }

                #multiple match type possibilities, but should only be one per rule.
                $propertyType = ($rule.match | Get-Member | Where-Object { $_.MemberType -eq "Property" }).Name
                $isUrlMatchProblem = $propertyType -eq "url" -and $rule.match.$propertyType -eq "*"
                $matchProperty = "$propertyType - $($rule.match.$propertyType)"

                $displayObject = [PSCustomObject]@{
                    RewriteRuleName = $rule.name
                    Pattern         = $rule.conditions.add.pattern
                    MatchProperty   = $matchProperty
                    ActionType      = $rule.action.type
                }

                #.ContainsValue() and .ContainsKey() doesn't find the complex object it seems. Need to find it by a key and a simple name.
                if (-not ($alreadyDisplayedUrlRewriteRules.ContainsKey((($displayObject.RewriteRuleName))))) {
                    $alreadyDisplayedUrlRewriteRules.Add($displayObject.RewriteRuleName, $displayObject)
                    $alreadyDisplayedUrlRewriteRules[$alreadyDisplayedUrlKey].Add($displayObject)

                    if ($isUrlMatchProblem) {
                        $alreadyDisplayedUrlRewriteRules[$urlMatchProblem].Add($rule.Name)
                    }
                }
            }
        }
    }

    if ($alreadyDisplayedUrlRewriteRules[$alreadyDisplayedUrlKey].Count -gt 0) {
        $params = $baseParams + @{
            OutColumns       = ([PSCustomObject]@{
                    DisplayObject = $alreadyDisplayedUrlRewriteRules[$alreadyDisplayedUrlKey]
                    IndentSpaces  = 8
                })
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params

        if ($alreadyDisplayedUrlRewriteRules[$urlMatchProblem].Count -gt 0) {
            $params = $baseParams + @{
                Name             = "Misconfigured URL Rewrite Rule - URL Match Problem Rules"
                Details          = "$([string]::Join(",", $alreadyDisplayedUrlRewriteRules[$urlMatchProblem]))" +
                "`r`n`t`tURL Match is set only a wild card which will result in a HTTP 500." +
                "`r`n`t`tIf the rule is required, the URL match should be '.*' to avoid issues."
                DisplayWriteType = "Red"
            }

            Add-AnalyzedResultInformation @params
        }
    }

    foreach ($webApp in $iisWebApplications) {
        if ($correctLocations.ContainsKey($webApp.FriendlyName)) {
            if ($webApp.PhysicalPath -notlike "*$($correctLocations[$webApp.FriendlyName])") {
                $params = $baseParams + @{
                    Name             = "Incorrect Virtual Directory Path"
                    Details          = "Error: '$($webApp.FriendlyName)' location for the virtual directory configuration is incorrect." +
                    "`r`n`t`tCurrently pointing to '$($webApp.PhysicalPath)', which is incorrect for this protocol and will cause problems."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($null -ne $missingWebApplicationConfigFile) {
        $params = $baseParams + @{
            Name                = "Missing Web Application Configuration File"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        foreach ($webApp in $missingWebApplicationConfigFile) {
            $params = $baseParams + @{
                Details                = "Web Application: '$($webApp.FriendlyName)' Attempting to use config: '$($webApp.ConfigurationFileInfo.Location)'"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
                TestingName            = "Web Application: '$($webApp.FriendlyName)'"
                DisplayTestingValue    = $($webApp.ConfigurationFileInfo.Location)
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($null -ne $invalidConfigurationFile) {
        $params = $baseParams + @{
            Name                = "Invalid Configuration File"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        $alreadyDisplayConfigs = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($configFile in $invalidConfigurationFile) {
            if ($alreadyDisplayConfigs.Add($configFile.Location)) {
                $params = $baseParams + @{
                    Details                = "Invalid: $($configFile.Location)"
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                    TestingName            = "Invalid: $($configFile.Location)"
                    DisplayTestingValue    = $true
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($null -ne $missingSharedConfigFile) {
        $params = $baseParams + @{
            Name                = "Missing Shared Configuration File"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        foreach ($file in $missingSharedConfigFile) {
            $params = $baseParams + @{
                Details                = "Missing: $($file.Location)"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-MissingConfig"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $missingConfigFiles) {
        $params = $baseParams + @{
            Name                = "Couldn't Find Config File"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        foreach ($file in $missingConfigFiles) {
            $params = $baseParams + @{
                Details                = "Friendly Name: $($file.FriendlyName)"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($null -ne $defaultVariableDetected) {
        $params = $baseParams + @{
            Name                = "Default Variable Detected"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        foreach ($file in $defaultVariableDetected) {
            $params = $baseParams + @{
                Details                = "$($file.Location)"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-DefaultVariableDetected"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $binSearchFoldersNotFound) {
        $params = $baseParams + @{
            Name                = "Bin Search Folder Not Found"
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params

        foreach ($file in $binSearchFoldersNotFound) {
            $params = $baseParams + @{
                Details                = "$($file.Location)"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-BinSearchFolder"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($defaultWebSitePowerShellSslEnabled) {
        $params = $baseParams + @{
            Details             = "Default Web Site/PowerShell has ssl enabled, which is unsupported."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params
    }

    if ($defaultWebSitePowerShellAuthenticationEnabled) {
        $params = $baseParams + @{
            Details             = "Default Web Site/PowerShell has authentication set, which is unsupported."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params
    }

    ########################
    # IIS Module Information
    ########################

    Write-Verbose "Working on IIS Module information"

    # If TokenCacheModule is not loaded, we highlight that it could be added back again as Windows provided a fix to address CVE-2023-36434 (also tracked as CVE-2023-21709)
    if ($null -eq $exchangeInformation.IISSettings.IISModulesInformation.ModuleList.Name) {
        Write-Verbose "Module List is null, unable to provide accurate check for this."
    } elseif ($exchangeInformation.IISSettings.IISModulesInformation.ModuleList.Name -notcontains "TokenCacheModule") {
        Write-Verbose "TokenCacheModule wasn't detected (vulnerability mitigated) and as a result, system is not vulnerable to CVE-2023-21709 / CVE-2023-36434"

        $params = $baseParams + @{
            Name                = "TokenCacheModule loaded"
            Details             = ("$false
                `r`t`tThe module wasn't found and as a result, CVE-2023-21709 and CVE-2023-36434 are mitigated. Windows has released a Security Update that addresses the vulnerability.
                `r`t`tIt should be installed on all Exchange servers and then, the TokenCacheModule can be added back to IIS (by running .\CVE-2023-21709.ps1 -Rollback).
                `r`t`tMore Information: https://aka.ms/CVE-2023-21709ScriptDoc"
            )
            DisplayWriteType    = "Yellow"
            AddHtmlDetailRow    = $true
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerNicSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "NIC Settings Per Active Adapter"  -DisplayOrder $Order -DefaultTabNumber 2)
    }
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    foreach ($adapter in $osInformation.NetworkInformation.NetworkAdapters) {

        if ($adapter.Description -eq "Remote NDIS Compatible Device") {
            Write-Verbose "Remote NDIS Compatible Device found. Ignoring NIC."
            continue
        }

        $params = $baseParams + @{
            Name                   = "Interface Description"
            Details                = "$($adapter.Description) [$($adapter.Name)]"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        if ($osInformation.BuildInformation.MajorVersion -notlike "Windows2008*" -and
            $osInformation.BuildInformation.MajorVersion -ne "Windows2012") {
            Write-Verbose "On Windows 2012 R2 or new. Can provide more details on the NICs"

            $driverDate = $adapter.DriverDate
            $detailsValue = $driverDate

            if ($hardwareInformation.ServerType -eq "Physical" -or
                $hardwareInformation.ServerType -eq "AmazonEC2") {

                if ($null -eq $driverDate -or
                    $driverDate -eq [DateTime]::MaxValue) {
                    $detailsValue = "Unknown"
                } elseif ((New-TimeSpan -Start $date -End $driverDate).Days -lt [int]-365) {
                    $params = $baseParams + @{
                        Details          = "Warning: NIC driver is over 1 year old. Verify you are at the latest version."
                        DisplayWriteType = "Yellow"
                        AddHtmlDetailRow = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
            }

            $params = $baseParams + @{
                Name    = "Driver Date"
                Details = $detailsValue
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Driver Version"
                Details = $adapter.DriverVersion
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "MTU Size"
                Details = $adapter.MTUSize
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Max Processors"
                Details = $adapter.NetAdapterRss.MaxProcessors
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Max Processor Number"
                Details = $adapter.NetAdapterRss.MaxProcessorNumber
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Number of Receive Queues"
                Details = $adapter.NetAdapterRss.NumberOfReceiveQueues
            }
            Add-AnalyzedResultInformation @params

            $writeType = "Yellow"
            $testingValue = $null

            if ($adapter.RssEnabledValue -eq 0) {
                $detailsValue = "False --- Warning: Enabling RSS is recommended."
                $testingValue = $false
            } elseif ($adapter.RssEnabledValue -eq 1) {
                $detailsValue = "True"
                $testingValue = $true
                $writeType = "Green"
            } else {
                $detailsValue = "No RSS Feature Detected."
            }

            $params = $baseParams + @{
                Name                = "RSS Enabled"
                Details             = $detailsValue
                DisplayWriteType    = $writeType
                DisplayTestingValue = $testingValue
            }
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "On Windows 2012 or older and can't get advanced NIC settings"
        }

        $linkSpeed = $adapter.LinkSpeed
        $displayValue = "{0} --- This may not be accurate due to virtualized hardware" -f $linkSpeed

        if ($hardwareInformation.ServerType -eq "Physical" -or
            $hardwareInformation.ServerType -eq "AmazonEC2") {
            $displayValue = $linkSpeed
        }

        $params = $baseParams + @{
            Name                = "Link Speed"
            Details             = $displayValue
            DisplayTestingValue = $linkSpeed
        }
        Add-AnalyzedResultInformation @params

        $displayValue = "{0}" -f $adapter.IPv6Enabled
        $displayWriteType = "Grey"
        $testingValue = $adapter.IPv6Enabled

        if ($osInformation.RegistryValues.IPv6DisabledComponents -ne 255 -and
            $adapter.IPv6Enabled -eq $false) {
            $displayValue = "{0} --- Warning" -f $adapter.IPv6Enabled
            $displayWriteType = "Yellow"
            $testingValue = $false
        }

        $params = $baseParams + @{
            Name                = "IPv6 Enabled"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            DisplayTestingValue = $testingValue
        }
        Add-AnalyzedResultInformation @params

        Add-AnalyzedResultInformation -Name "IPv4 Address" @baseParams

        foreach ($address in $adapter.IPv4Addresses) {
            $displayValue = "{0}/{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $params = $baseParams + @{
                Name                   = "Address"
                Details                = $displayValue
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        Add-AnalyzedResultInformation -Name "IPv6 Address" @baseParams

        foreach ($address in $adapter.IPv6Addresses) {
            $displayValue = "{0}\{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $params = $baseParams + @{
                Name                   = "Address"
                Details                = $displayValue
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Name    = "DNS Server"
            Details = $adapter.DnsServer
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name    = "Registered In DNS"
            Details = $adapter.RegisteredInDns
        }
        Add-AnalyzedResultInformation @params

        #Assuming that all versions of Hyper-V doesn't allow sleepy NICs
        if (($hardwareInformation.ServerType -ne "HyperV") -and ($adapter.PnPCapabilities -ne "MultiplexorNoPnP")) {
            $displayWriteType = "Grey"
            $displayValue = $adapter.SleepyNicDisabled

            if (!$adapter.SleepyNicDisabled) {
                $displayWriteType = "Yellow"
                $displayValue = "False --- Warning: It's recommended to disable NIC power saving options`r`n`t`t`tMore Information: https://aka.ms/HC-NICPowerManagement"
            }

            $params = $baseParams + @{
                Name                = "Sleepy NIC Disabled"
                Details             = $displayValue
                DisplayWriteType    = $displayWriteType
                DisplayTestingValue = $adapter.SleepyNicDisabled
            }
            Add-AnalyzedResultInformation @params
        }

        $adapterDescription = $adapter.Description
        $cookedValue = 0
        $foundCounter = $false

        if ($null -eq $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            Write-Verbose "PacketsReceivedDiscarded is null"
            continue
        }

        foreach ($prdInstance in $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            $instancePath = $prdInstance.Path
            $startIndex = $instancePath.IndexOf("(") + 1
            $charLength = $instancePath.Substring($startIndex, ($instancePath.IndexOf(")") - $startIndex)).Length
            $instanceName = $instancePath.Substring($startIndex, $charLength)
            $possibleInstanceName = $adapterDescription.Replace("#", "_")

            if ($instanceName -eq $adapterDescription -or
                $instanceName -eq $possibleInstanceName) {
                $cookedValue = $prdInstance.CookedValue
                $foundCounter = $true
                break
            }
        }

        $displayWriteType = "Yellow"
        $displayValue = $cookedValue
        $baseDisplayValue = "{0} --- {1}: This value should be at 0."
        $knownIssue = $false

        if ($foundCounter) {

            if ($cookedValue -eq 0) {
                $displayWriteType = "Green"
            } elseif ($cookedValue -lt 1000) {
                $displayValue = $baseDisplayValue -f $cookedValue, "Warning"
            } else {
                $displayWriteType = "Red"
                $displayValue = [string]::Concat(($baseDisplayValue -f $cookedValue, "Error"), "We are also seeing this value being rather high so this can cause a performance impacted on a system.")
            }

            if ($adapterDescription -like "*vmxnet3*" -and
                $cookedValue -gt 0) {
                $knownIssue = $true
            }
        } else {
            $displayValue = "Couldn't find value for the counter."
            $cookedValue = $null
            $displayWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name                = "Packets Received Discarded"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            DisplayTestingValue = $cookedValue
        }
        Add-AnalyzedResultInformation @params

        if ($knownIssue) {
            $params = $baseParams + @{
                Details                = "Known Issue with vmxnet3: 'Large packet loss at the guest operating system level on the VMXNET3 vNIC in ESXi (2039495)' - https://aka.ms/HC-VMwareLostPackets"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 3
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($osInformation.NetworkInformation.NetworkAdapters.Count -gt 1) {
        $params = $baseParams + @{
            Details          = "Multiple active network adapters detected. Exchange 2013 or greater may not need separate adapters for MAPI and replication traffic.  For details please refer to https://aka.ms/HC-PlanHA#network-requirements"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if ($osInformation.NetworkInformation.IPv6DisabledOnNICs) {
        $displayWriteType = "Grey"
        $displayValue = "True"
        $testingValue = $true

        if ($osInformation.RegistryValues.IPv6DisabledComponents -eq -1) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not correctly disabled via DisabledComponents registry value. It is currently set to '-1'. `r`n`t`tThis setting cause a system startup delay of 5 seconds. For details please refer to: `r`n`t`thttps://aka.ms/HC-ConfigureIPv6"
        } elseif ($osInformation.RegistryValues.IPv6DisabledComponents -ne 255) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not fully disabled. DisabledComponents registry value currently set to '{0}'. For details please refer to the following articles: `r`n`t`thttps://aka.ms/HC-DisableIPv6`r`n`t`thttps://aka.ms/HC-ConfigureIPv6" -f $osInformation.RegistryValues.IPv6DisabledComponents
        }

        $params = $baseParams + @{
            Name                   = "Disable IPv6 Correctly"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params
    }

    $noDNSRegistered = ($osInformation.NetworkInformation.NetworkAdapters | Where-Object { $_.RegisteredInDns -eq $true }).Count -eq 0

    if ($noDNSRegistered) {
        $params = $baseParams + @{
            Name                   = "No NIC Registered In DNS"
            Details                = "Error: This will cause server to crash and odd mail flow issues. Exchange Depends on the primary NIC to have the setting Registered In DNS set."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params
    }
}


function Invoke-AnalyzerOrganizationInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $organizationInformation = $HealthServerObject.OrganizationInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Organization Information"  -DisplayOrder $Order)
    }

    $params = $baseParams + @{
        Name    = "MAPI/HTTP Enabled"
        Details = $organizationInformation.MapiHttpEnabled
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "Enable Download Domains"
        Details = $organizationInformation.EnableDownloadDomains
    }
    Add-AnalyzedResultInformation @params

    if ($null -ne $organizationInformation.GetOrganizationConfig -and
        $organizationInformation.EnableDownloadDomains.ToString() -eq "Unknown") {
        $params = $baseParams + @{
            Details                = "This is 'Unknown' because EMS is connected to an Exchange Version that doesn't know about Enable Download Domains in Get-OrganizationConfig"
            DisplayCustomTabNumber = 2
            DisplayWriteType       = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "AD Split Permissions"
        Details = $organizationInformation.IsSplitADPermissions
    }
    Add-AnalyzedResultInformation @params

    $displayWriteType = "Green"

    if ($organizationInformation.ADSiteCount -ge 750) {
        $displayWriteType = "Yellow"
    } elseif ( $organizationInformation.ADSiteCount -ge 1000) {
        $displayWriteType = "Red"
    }

    $params = $baseParams + @{
        Name             = "Total AD Site Count"
        Details          = $organizationInformation.ADSiteCount
        DisplayWriteType = $displayWriteType
    }
    Add-AnalyzedResultInformation @params

    if ($displayWriteType -ne "Green") {
        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-ADSiteCount"
            DisplayCustomTabNumber = 2
            DisplayWriteType       = $displayWriteType
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $organizationInformation.GetDynamicDgPublicFolderMailboxes -and
        $organizationInformation.GetDynamicDgPublicFolderMailboxes.Count -ne 0) {
        $displayWriteType = "Green"

        if ($organizationInformation.GetDynamicDgPublicFolderMailboxes.Count -gt 1) {
            $displayWriteType = "Red"
        }

        $params = $baseParams + @{
            Name             = "Dynamic Distribution Group Public Folder Mailboxes Count"
            Details          = $organizationInformation.GetDynamicDgPublicFolderMailboxes.Count
            DisplayWriteType = $displayWriteType
        }

        Add-AnalyzedResultInformation @params

        if ($displayWriteType -ne "Green") {
            $params = $baseParams + @{
                Details                = "More Information: https://aka.ms/HC-DynamicDgPublicFolderMailboxes"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }

            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "No Dynamic Distribution Group Public Folder Mailboxes found to review."
    }
}

function Invoke-AnalyzerFrequentConfigurationIssues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $tcpKeepAlive = $osInformation.RegistryValues.TCPKeepAlive
    $organizationInformation = $HealthServerObject.OrganizationInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Frequent Configuration Issues"  -DisplayOrder $Order)
    }

    if ($tcpKeepAlive -eq 0) {
        $displayValue = "Not Set `r`n`t`tError: Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration. `r`n`t`tMore details: https://aka.ms/HC-TcpIpSettingsCheck"
        $displayWriteType = "Red"
    } elseif ($tcpKeepAlive -lt 900000 -or
        $tcpKeepAlive -gt 1800000) {
        $displayValue = "$tcpKeepAlive `r`n`t`tWarning: Not configured optimally, recommended value between 15 to 30 minutes (900000 and 1800000 decimal). `r`n`t`tMore details: https://aka.ms/HC-TcpIpSettingsCheck"
        $displayWriteType = "Yellow"
    } else {
        $displayValue = $tcpKeepAlive
        $displayWriteType = "Green"
    }

    $params = $baseParams + @{
        Name                = "TCP/IP Settings"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $tcpKeepAlive
        HtmlName            = "TCPKeepAlive"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                = "RPC Min Connection Timeout"
        Details             = "$($osInformation.RegistryValues.RpcMinConnectionTimeout) `r`n`t`tMore Information: https://aka.ms/HC-RPCSetting"
        DisplayTestingValue = $osInformation.RegistryValues.RpcMinConnectionTimeout
        HtmlName            = "RPC Minimum Connection Timeout"
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.RegistryValues.DisableGranularReplication -ne 0) {
        $params = $baseParams + @{
            Name                = "DisableGranularReplication"
            Details             = "$($exchangeInformation.RegistryValues.DisableGranularReplication) - Error this can cause work load management issues."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name     = "FIPS Algorithm Policy Enabled"
        Details  = $exchangeInformation.RegistryValues.FipsAlgorithmPolicyEnabled
        HtmlName = "FipsAlgorithmPolicy-Enabled"
    }
    Add-AnalyzedResultInformation @params

    $displayValue = $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    $displayWriteType = "Green"

    if ($exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage -ne 0) {
        $displayWriteType = "Red"
        $displayValue = "{0} `r`n`t`tError: This can cause an impact to the server's search performance. This should only be used a temporary fix if no other options are available vs a long term solution." -f $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    }

    $params = $baseParams + @{
        Name                = "CTS Processor Affinity Percentage"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
        HtmlName            = "CtsProcessorAffinityPercentage"
    }
    Add-AnalyzedResultInformation @params

    $displayValue = $exchangeInformation.RegistryValues.DisableAsyncNotification
    $displayWriteType = "Grey"

    if ($displayValue -ne 0) {
        $displayWriteType = "Yellow"
        $displayValue = "$($exchangeInformation.RegistryValues.DisableAsyncNotification) Warning: This value should be set back to 0 after you no longer need it for the workaround described in http://support.microsoft.com/kb/5013118"
    }

    $params = $baseParams + @{
        Name                = "Disable Async Notification"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $displayValue -ne 0
    }
    Add-AnalyzedResultInformation @params

    $credGuardRunning = $false
    $credGuardUnknown = $osInformation.CredentialGuardCimInstance -eq "Unknown"

    if (-not ($credGuardUnknown)) {
        # CredentialGuardCimInstance is an array type and not sure if we can have multiple here, so just going to loop thru and handle it this way.
        $credGuardRunning = $null -ne ($osInformation.CredentialGuardCimInstance | Where-Object { $_ -eq 1 })
    }

    $displayValue = $credentialGuardValue = $osInformation.RegistryValues.CredentialGuard -ne 0 -or $credGuardRunning
    $displayWriteType = "Grey"

    if ($credentialGuardValue) {
        $displayValue = "{0} `r`n`t`tError: Credential Guard is not supported on an Exchange Server. This can cause a performance hit on the server." -f $credentialGuardValue
        $displayWriteType = "Red"
    }

    if ($credGuardUnknown -and (-not ($credentialGuardValue))) {
        $displayValue = "Unknown `r`n`t`tWarning: Unable to determine Credential Guard status. If enabled, this can cause a performance hit on the server."
        $displayWriteType = "Yellow"
    }

    $params = $baseParams + @{
        Name                = "Credential Guard Enabled"
        Details             = $displayValue
        DisplayTestingValue = $credentialGuardValue
        DisplayWriteType    = $displayWriteType
    }
    Add-AnalyzedResultInformation @params

    if ($null -ne $exchangeInformation.ApplicationConfigFileStatus -and
        $exchangeInformation.ApplicationConfigFileStatus.Count -ge 1) {

        # Only need to display a particular list all the time. Don't need every config that we want to possibly look at for issues.
        $alwaysDisplayConfigs = @("EdgeTransport.exe.config")
        $skipEdgeOnlyConfigs = @("noderunner.exe.config")
        $keyList = $exchangeInformation.ApplicationConfigFileStatus.Keys | Sort-Object

        foreach ($configKey in $keyList) {

            $configStatus = $exchangeInformation.ApplicationConfigFileStatus[$configKey]
            $fileName = $configStatus.FileName
            $writeType = "Green"
            [string]$writeValue = $configStatus.Present

            if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $true -and
                $skipEdgeOnlyConfigs -contains $fileName) {
                continue
            }

            if (-not $configStatus.Present) {
                $writeType = "Red"
                $writeValue += " --- Error"
            }

            $params = $baseParams + @{
                Name             = "$fileName Present"
                Details          = $writeValue
                DisplayWriteType = $writeType
            }

            if ($alwaysDisplayConfigs -contains $fileName -or
                -not $configStatus.Present) {
                Add-AnalyzedResultInformation @params
            }

            # if not a valid configuration file, provide that.
            try {
                if ($configStatus.Present) {
                    $content = [xml]($configStatus.Content)

                    # Additional checks of configuration files.
                    if ($fileName -eq "noderunner.exe.config") {
                        $memoryLimitMegabytes = $content.configuration.nodeRunnerSettings.memoryLimitMegabytes
                        $writeValue = "$memoryLimitMegabytes MB"
                        $writeType = "Green"

                        if ($null -eq $memoryLimitMegabytes) {
                            $writeType = "Yellow"
                            $writeValue = "Unconfigured. This may cause problems."
                        } elseif ($memoryLimitMegabytes -ne 0) {
                            $writeType = "Yellow"
                            $writeValue = "$memoryLimitMegabytes MB will limit the performance of search and can be more impactful than helpful if not configured correctly for your environment."
                        }

                        $params = $baseParams + @{
                            Name             = "NodeRunner.exe memory limit"
                            Details          = $writeValue
                            DisplayWriteType = $writeType
                        }

                        Add-AnalyzedResultInformation @params

                        if ($writeType -ne "Green") {
                            $params = $baseParams + @{
                                Details                = "More Information: https://aka.ms/HC-NodeRunnerMemoryCheck"
                                DisplayWriteType       = "Yellow"
                                DisplayCustomTabNumber = 2
                            }

                            Add-AnalyzedResultInformation @params
                        }
                    }
                }
            } catch {
                $params = $baseParams + @{
                    Name                = "$fileName Invalid Config Format"
                    Details             = "True --- Error: Not able to convert to xml which means it is in an incorrect format that will cause problems with the process."
                    DisplayTestingValue = $true
                    DisplayWriteType    = "Red"
                }

                Add-AnalyzedResultInformation @params
            }
        }
    }

    $displayWriteType = "Yellow"
    $displayValue = "Unknown - Unable to run Get-AcceptedDomain"
    $additionalDisplayValue = [string]::Empty

    if ($null -ne $organizationInformation.GetAcceptedDomain -and
        $organizationInformation.GetAcceptedDomain -ne "Unknown") {

        $wildCardAcceptedDomain = $organizationInformation.GetAcceptedDomain | Where-Object { $_.DomainName.ToString() -eq "*" }

        if ($null -eq $wildCardAcceptedDomain) {
            $displayValue = "Not Set"
            $displayWriteType = "Grey"
        } else {
            $displayWriteType = "Red"
            $displayValue = "Error --- Accepted Domain `"$($wildCardAcceptedDomain.Id)`" is set to a Wild Card (*) Domain Name with a domain type of $($wildCardAcceptedDomain.DomainType.ToString()). This is not recommended as this is an open relay for the entire environment.`r`n`t`tMore Information: https://aka.ms/HC-OpenRelayDomain"

            if ($wildCardAcceptedDomain.DomainType.ToString() -eq "InternalRelay" -and
                ((Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2016" -CU "CU22") -or
                (Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2019" -CU "CU11"))) {
                $additionalDisplayValue = "`r`n`t`tERROR: You have an open relay set as Internal Replay Type and on a CU that is known to cause issues with transport services crashing. Follow the above article for more information."
            } elseif ($wildCardAcceptedDomain.DomainType.ToString() -eq "InternalRelay") {
                $additionalDisplayValue = "`r`n`t`tWARNING: You have an open relay set as Internal Relay Type. You are not on a CU yet that is having issue, recommended to change this prior to upgrading. Follow the above article for more information."
            }
        }
    }

    $params = $baseParams + @{
        Name             = "Open Relay Wild Card Domain"
        Details          = $displayValue
        DisplayWriteType = $displayWriteType
    }
    Add-AnalyzedResultInformation @params

    if ($additionalDisplayValue -ne [string]::Empty) {
        $params = $baseParams + @{
            Details          = $additionalDisplayValue
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "DisablePreservation"
        Details = $exchangeInformation.RegistryValues.DisablePreservation
    }
    Add-AnalyzedResultInformation @params

    if ($osInformation.RegistryValues.SuppressExtendedProtection -ne 0) {
        $params = $baseParams + @{
            Name             = "SuppressExtendedProtection"
            Details          = "Value set to $($osInformation.RegistryValues.SuppressExtendedProtection), which disables EP resulting it to not work correctly and causes problems. --- ERROR"
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    # Detect Send Connector sending to EXO
    $exoConnector = New-Object System.Collections.Generic.List[object]
    $sendConnectors = $exchangeInformation.ExchangeConnectors | Where-Object { $_.ConnectorType -eq "Send" }

    foreach ($sendConnector in $sendConnectors) {
        $smartHostMatch = ($sendConnector.SmartHosts -like "*.mail.protection.outlook.com").Count -gt 0
        $dnsMatch = $sendConnector.SmartHosts -eq 0 -and ($sendConnector.AddressSpaces.Address -like "*.mail.onmicrosoft.com").Count -gt 0

        if ($dnsMatch -or $smartHostMatch) {
            $exoConnector.Add($sendConnector)
        }
    }

    $params = $baseParams + @{
        Name    = "EXO Connector Present"
        Details = ($exoConnector.Count -gt 0)
    }
    Add-AnalyzedResultInformation @params
    $showMoreInfo = $false

    foreach ($connector in $exoConnector) {
        # Misconfigured connector is if TLSCertificateName is not set or CloudServicesMailEnabled not set to true
        if ($connector.CloudEnabled -eq $false -or
            $connector.CertificateDetails.TlsCertificateNameStatus -eq "TlsCertificateNameEmpty") {
            $params = $baseParams + @{
                Name                   = "Send Connector - $($connector.Identity.ToString())"
                Details                = "Misconfigured to send authenticated internal mail to M365." +
                "`r`n`t`t`tCloudServicesMailEnabled: $($connector.CloudEnabled)" +
                "`r`n`t`t`tTLSCertificateName set: $($connector.CertificateDetails.TlsCertificateNameStatus -ne "TlsCertificateNameEmpty")"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Red"
            }
            Add-AnalyzedResultInformation @params
            $showMoreInfo = $true
        }

        if ($connector.TlsAuthLevel -ne "DomainValidation" -and
            $connector.TlsAuthLevel -ne "CertificateValidation") {
            $params = $baseParams + @{
                Name                   = "Send Connector - $($connector.Identity.ToString())"
                Details                = "TlsAuthLevel not set to CertificateValidation or DomainValidation"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
            $showMoreInfo = $true
        }

        if ($connector.TlsDomain -ne "mail.protection.outlook.com" -and
            $connector.TlsAuthLevel -eq "DomainValidation") {
            $params = $baseParams + @{
                Name                   = "Send Connector - $($connector.Identity.ToString())"
                Details                = "TLSDomain  not set to mail.protection.outlook.com"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
            $showMoreInfo = $true
        }
    }

    if ($showMoreInfo) {
        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-ExoConnectorIssue"
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }
}


function Invoke-AnalyzerSecurityExchangeCertificates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    foreach ($certificate in $exchangeInformation.ExchangeCertificates) {

        if ($certificate.LifetimeInDays -ge 60) {
            $displayColor = "Green"
        } elseif ($certificate.LifetimeInDays -ge 30) {
            $displayColor = "Yellow"
        } else {
            $displayColor = "Red"
        }

        $params = $baseParams + @{
            Name                   = "Certificate"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "FriendlyName"
            Details                = $certificate.FriendlyName
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Thumbprint"
            Details                = $certificate.Thumbprint
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Lifetime in days"
            Details                = $certificate.LifetimeInDays
            DisplayCustomTabNumber = 2
            DisplayWriteType       = $displayColor
        }
        Add-AnalyzedResultInformation @params

        $displayValue = $false
        $displayWriteType = "Grey"
        if ($certificate.LifetimeInDays -lt 0) {
            $displayValue = $true
            $displayWriteType = "Red"
        }

        $params = $baseParams + @{
            Name                   = "Certificate has expired"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $certStatusWriteType = [string]::Empty

        if ($null -ne $certificate.Status) {
            switch ($certificate.Status) {
                ("Unknown") { $certStatusWriteType = "Yellow" }
                ("Valid") { $certStatusWriteType = "Grey" }
                ("Revoked") { $certStatusWriteType = "Red" }
                ("DateInvalid") { $certStatusWriteType = "Red" }
                ("Untrusted") { $certStatusWriteType = "Yellow" }
                ("Invalid") { $certStatusWriteType = "Red" }
                ("RevocationCheckFailure") { $certStatusWriteType = "Yellow" }
                ("PendingRequest") { $certStatusWriteType = "Yellow" }
                default { $certStatusWriteType = "Yellow" }
            }

            $params = $baseParams + @{
                Name                   = "Certificate status"
                Details                = $certificate.Status
                DisplayCustomTabNumber = 2
                DisplayWriteType       = $certStatusWriteType
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Certificate status"
                Details                = "Unknown"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        # We show the 'Key Size' if a certificate is RSA or DSA based but not for ECC certificates where it would be displayed with a value of 0
        # More information: https://stackoverflow.com/questions/32873851/load-a-certificate-using-x509certificate2-with-ecc-public-key
        if ($certificate.PublicKeySize -lt 2048 -and
            -not($certificate.IsEccCertificate)) {
            $params = $baseParams + @{
                Name                   = "Key size"
                Details                = $certificate.PublicKeySize
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "It's recommended to use a key size of at least 2048 bit"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } elseif (-not($certificate.IsEccCertificate)) {
            $params = $baseParams + @{
                Name                   = "Key size"
                Details                = $certificate.PublicKeySize
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Name                   = "ECC Certificate"
            Details                = $certificate.IsEccCertificate
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        if ($certificate.SignatureHashAlgorithmSecure -eq 1) {
            $shaDisplayWriteType = "Yellow"
        } else {
            $shaDisplayWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name                   = "Signature Algorithm"
            Details                = $certificate.SignatureAlgorithm
            DisplayWriteType       = $shaDisplayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Signature Hash Algorithm"
            Details                = $certificate.SignatureHashAlgorithm
            DisplayWriteType       = $shaDisplayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        if ($shaDisplayWriteType -eq "Yellow") {
            $params = $baseParams + @{
                Details                = "It's recommended to use a hash algorithm from the SHA-2 family `r`n`t`tMore information: https://aka.ms/HC-SSLBP"
                DisplayWriteType       = $shaDisplayWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $certificate.Services) {
            $params = $baseParams + @{
                Name                   = "Bound to services"
                Details                = $certificate.Services
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) {
            $params = $baseParams + @{
                Name                   = "Internal Transport Certificate"
                Details                = $certificate.IsInternalTransportCertificate
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name                   = "Current Auth Certificate"
                Details                = $certificate.IsCurrentAuthConfigCertificate
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name                   = "Next Auth Certificate"
                Details                = $certificate.IsNextAuthConfigCertificate
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Name                   = "SAN Certificate"
            Details                = $certificate.IsSanCertificate
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Namespaces"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        foreach ($namespace in $certificate.Namespaces) {
            $params = $baseParams + @{
                Details                = $namespace
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        if ($certificate.IsInternalTransportCertificate) {
            $internalTransportCertificate = $certificate
        }

        if ($certificate.IsCurrentAuthConfigCertificate -eq $true) {
            $currentAuthCertificate = $certificate
        } elseif ($certificate.IsNextAuthConfigCertificate -eq $true) {
            $nextAuthCertificate = $certificate
            $nextAuthCertificateEffectiveDate = $certificate.SetAsActiveAuthCertificateOn
        }
    }

    if ($null -ne $internalTransportCertificate) {
        if ($internalTransportCertificate.LifetimeInDays -gt 0) {
            $params = $baseParams + @{
                Name                   = "Valid Internal Transport Certificate Found On Server"
                Details                = $true
                DisplayWriteType       = "Green"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Valid Internal Transport Certificate Found On Server"
                Details                = $false
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "Internal Transport Certificate has expired `r`n`t`tMore Information: https://aka.ms/HC-InternalTransportCertificate"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    } elseif ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $true) {
        $params = $baseParams + @{
            Name                   = "Valid Internal Transport Certificate Found On Server"
            Details                = $false
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "We can't check for Internal Transport Certificate on Edge Transport Servers"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    } else {
        $params = $baseParams + @{
            Name                   = "Valid Internal Transport Certificate Found On Server"
            Details                = $false
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "No Internal Transport Certificate found. This may cause several problems. `r`n`t`tMore Information: https://aka.ms/HC-InternalTransportCertificate"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $currentAuthCertificate) {
        if ($currentAuthCertificate.LifetimeInDays -gt 0) {
            $params = $baseParams + @{
                Name                   = "Valid Auth Certificate Found On Server"
                Details                = $true
                DisplayWriteType       = "Green"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Valid Auth Certificate Found On Server"
                Details                = $false
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "Auth Certificate has expired `r`n`t`tMore Information: https://aka.ms/HC-OAuthExpired"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $nextAuthCertificate) {
            $params = $baseParams + @{
                Name                   = "Next Auth Certificate Staged For Rotation"
                Details                = $true
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name                   = "Next Auth Certificate Effective Date"
                Details                = $nextAuthCertificateEffectiveDate
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params
        }
    } elseif ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $true) {
        $params = $baseParams + @{
            Name                   = "Valid Auth Certificate Found On Server"
            Details                = $false
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "We can't check for Auth Certificates on Edge Transport Servers"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    } else {
        $params = $baseParams + @{
            Name                   = "Valid Auth Certificate Found On Server"
            Details                = $false
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "No valid Auth Certificate found. This may cause several problems. `r`n`t`tMore Information: https://aka.ms/HC-FindOAuthHybrid"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }
}

<#
 This function is to create a simple return value for the Setting Override you are looking for
 It will also determine if this setting override should be applied on the server or not
 You should pass in the results from Get-ExchangeSettingOverride (the true settings on the Exchange Server)
 And the Get-SettingOverride (what is stored in AD) as a fallback
 WARNING: Get-SettingOverride should really not be used as the status is only accurate for the session we are connected to for EMS.
    Caller should determine if the override is applied to the server by the Status and FromAdSettings properties.
    If FromAdSettings is set to true, the data was determined from Get-SettingOverride and to be not accurate.
#>
function Get-FilteredSettingOverrideInformation {
    [CmdletBinding()]
    param(
        [object[]]$GetSettingOverride,
        [object[]]$ExchangeSettingOverride,

        [Parameter(Mandatory = $true)]
        [string]$FilterServer,

        [Parameter(Mandatory = $true)]
        [System.Version]$FilterServerVersion,

        [Parameter(Mandatory = $true)]
        [string]$FilterComponentName,

        [Parameter(Mandatory = $true)]
        [string]$FilterSectionName,

        [Parameter(Mandatory = $true)]
        [string]$FilterParameterName
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Trying to filter down results for ComponentName: $FilterComponentName SectionName: $FilterSectionName ParameterName: $FilterParameterName"
        $results = New-Object "System.Collections.Generic.List[object]"
        $findFromOverride = $null
        $usedAdSettings = $false
        $adjustedFilterServer = $FilterServer.Split(".")[0].ToLower()
    } process {
        # Use ExchangeSettingOverride first
        if ($null -ne $ExchangeSettingOverride -and
            $ExchangeSettingOverride.SimpleSettingOverrides.Count -gt 0) {
            $findFromOverride = $ExchangeSettingOverride.SimpleSettingOverrides
        } elseif ($null -ne $GetSettingOverride -and
            $GetSettingOverride -ne "Unknown") {
            $findFromOverride = $GetSettingOverride
            $usedAdSettings = $true
        } elseif ($GetSettingOverride -eq "Unknown") {
            $results.Add("Unknown")
            return
        } else {
            Write-Verbose "No data to filter"
            return
        }

        $filteredResults = $findFromOverride | Where-Object { $_.ComponentName -eq $FilterComponentName -and $_.SectionName -eq $FilterSectionName }

        if ($null -ne $filteredResults) {
            Write-Verbose "Found $($filteredResults.Count) override(s)"
            foreach ($entry in $filteredResults) {
                Write-Verbose "Working on entry: $($entry.Name)"
                foreach ($p in [array]($entry.Parameters)) {
                    Write-Verbose "Working on parameter: $p"
                    if ($p.Contains($FilterParameterName)) {
                        $value = $p.Substring($FilterParameterName.Length + 1) # Add plus 1 for '='
                        # everything matched, however, only add it to the list for the following reasons
                        # - Status is Accepted and not from AD and a unique value in the list
                        # - Or From AD and current logic determines it applies

                        if ($usedAdSettings) {
                            # can have it apply by build and server parameter
                            if (($null -eq $entry.MinVersion -or
                                    $FilterServerVersion -ge $entry.MinVersion) -and
                                (($null -eq $entry.MaxVersion -or
                                    $FilterServerVersion -le $entry.MaxVersion)) -and
                                (($null -eq $entry.Server -or
                                    $entry.Server.ToLower().Contains($adjustedFilterServer)))) {
                                $status = $entry.Status
                            } else {
                                $status = "DoesNotApply"
                            }
                        } else {
                            $status = $entry.Status
                        }

                        if ($status -eq "Accepted" -and
                            ($results.Count -lt 1 -or
                            -not ($results.ParameterValue.ToLower().Contains($value.ToLower())))) {
                            $results.Add([PSCustomObject]@{
                                    Name           = $entry.Name
                                    Reason         = $entry.Reason
                                    ModifiedBy     = $entry.ModifiedBy
                                    ComponentName  = $entry.ComponentName
                                    SectionName    = $entry.SectionName
                                    ParameterName  = $FilterParameterName
                                    ParameterValue = $value
                                    Status         = $entry.Status
                                    TrueStatus     = $status
                                    FromAdSettings = $usedAdSettings
                                })
                        } elseif ($status -eq "Accepted") {
                            Write-Verbose "Already have 1 Accepted value added to list no need to add another one. Skip adding $($entry.Name)"
                        } else {
                            Write-Verbose "Already have parameter value added to the. Skip adding $($entry.Name)"
                        }
                    }
                }
            }
        }
    } end {
        # If no filter data is found, return null.
        # Up to the caller for how to determine this information.
        if ($results.Count -eq 0) {
            return $null
        }
        return $results
    }
}

function Invoke-AnalyzerSecurityAMSIConfigState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    $osInformation = $HealthServerObject.OSInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # AMSI integration is only available on Windows Server 2016 or higher and only on
    # Exchange Server 2016 CU21+ or Exchange Server 2019 CU10+.
    # AMSI is also not available on Edge Transport Servers (no http component available).
    if (($osInformation.BuildInformation.BuildVersion -ge [System.Version]"10.0.0.0") -and
        ((Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2016" -CU "CU21") -or
        (Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2019" -CU "CU10")) -and
        ($exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false)) {

        $params = @{
            ExchangeSettingOverride = $HealthServerObject.ExchangeInformation.SettingOverrides
            GetSettingOverride      = $HealthServerObject.OrganizationInformation.GetSettingOverride
            FilterServer            = $HealthServerObject.ServerName
            FilterServerVersion     = $exchangeInformation.BuildInformation.VersionInformation.BuildVersion
            FilterComponentName     = "Cafe"
            FilterSectionName       = "HttpRequestFiltering"
            FilterParameterName     = "Enabled"
        }

        # Only thing that is returned is Accepted values and unique
        [array]$amsiInformation = Get-FilteredSettingOverrideInformation @params

        $amsiWriteType = "Yellow"
        $amsiConfigurationWarning = "`r`n`t`tThis may pose a security risk to your servers`r`n`t`tMore Information: https://aka.ms/HC-AMSIExchange"
        $amsiConfigurationUnknown = "Exchange AMSI integration state is unknown"
        $additionalAMSIDisplayValue = $null

        if ($null -eq $amsiInformation) {
            # No results returned, no matches therefore good.
            $amsiWriteType = "Green"
            $amsiState = "True"
        } elseif ($amsiInformation -eq "Unknown") {
            $additionalAMSIDisplayValue = "Unable to query Exchange AMSI integration state"
        } elseif ($amsiInformation.Count -eq 1) {
            $amsiState = $amsiInformation.ParameterValue
            if ($amsiInformation.ParameterValue -eq "False") {
                $additionalAMSIDisplayValue = "Setting applies to the server" + $amsiConfigurationWarning
            } elseif ($amsiInformation.ParameterValue -eq "True") {
                $amsiWriteType = "Green"
            } else {
                $additionalAMSIDisplayValue = $amsiConfigurationUnknown + " - Setting Override Name: $($amsiInformation.Name)"
                $additionalAMSIDisplayValue += $amsiConfigurationWarning
            }
        } else {
            $amsiState = "Multiple overrides detected"
            $additionalAMSIDisplayValue = $amsiConfigurationUnknown + " - Multi Setting Overrides Applied: $([string]::Join(", ", $amsiInformation.Name))"
            $additionalAMSIDisplayValue += $amsiConfigurationWarning
        }

        $params = $baseParams + @{
            Name             = "AMSI Enabled"
            Details          = $amsiState
            DisplayWriteType = $amsiWriteType
        }
        Add-AnalyzedResultInformation @params

        if ($null -ne $additionalAMSIDisplayValue) {
            $params = $baseParams + @{
                Details                = $additionalAMSIDisplayValue
                DisplayWriteType       = $amsiWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "AMSI integration is not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU"
    }
}

function Invoke-AnalyzerSecurityOverrides {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    <#
        This function is used to analyze overrides which are enabled via SettingOverride or Registry Value
    #>

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $exchangeBuild = $exchangeInformation.BuildInformation.VersionInformation.BuildVersion
    $strictModeDisabledLocationsList = New-Object System.Collections.Generic.List[string]
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    if ($exchangeBuild -ge "15.1.0.0") {
        Write-Verbose "Checking SettingOverride for Strict Mode configuration state"
        $params = @{
            ExchangeSettingOverride = $exchangeInformation.SettingOverrides
            GetSettingOverride      = $HealthServerObject.OrganizationInformation.GetSettingOverride
            FilterServer            = $HealthServerObject.ServerName
            FilterServerVersion     = $exchangeBuild
            FilterComponentName     = "Data"
            FilterSectionName       = "DeserializationBinderSettings"
            FilterParameterName     = "LearningLocations"
        }

        [array]$deserializationBinderSettings = Get-FilteredSettingOverrideInformation @params

        if ($null -ne $deserializationBinderSettings) {
            foreach ($setting in $deserializationBinderSettings) {
                Write-Verbose "Strict Mode has been disabled via SettingOverride for $($setting.ParameterValue) location"
                $strictModeDisabledLocationsList.Add($setting.ParameterValue)
            }
        }

        $params = $baseParams + @{
            Name             = "Strict Mode disabled"
            Details          = $strictModeDisabledLocationsList.Count -gt 0
            DisplayWriteType = if ($strictModeDisabledLocationsList.Count -gt 0) { "Red" } else { "Green" }
        }
        Add-AnalyzedResultInformation @params

        foreach ($location in $strictModeDisabledLocationsList) {
            $params = $baseParams + @{
                Name                   = "Location"
                Details                = $location
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Checking Registry Value for BaseTypeCheckForDeserialization configuration state"
    $disableBaseTypeCheckForDeserializationSettingsState = $exchangeInformation.RegistryValues.DisableBaseTypeCheckForDeserialization -eq 1

    $params = $baseParams + @{
        Name             = "BaseTypeCheckForDeserialization disabled"
        Details          = $disableBaseTypeCheckForDeserializationSettingsState
        DisplayWriteType = if ($disableBaseTypeCheckForDeserializationSettingsState) { "Red" } else { "Green" }
    }
    Add-AnalyzedResultInformation @params

    if (($strictModeDisabledLocationsList.Count -gt 0) -or
        ($disableBaseTypeCheckForDeserializationSettingsState)) {
        $params = $baseParams + @{
            Details                = "These overrides should only be used in very limited failure scenarios" +
            "`n`t`tRollback instructions: https://aka.ms/HC-SettingOverrides"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerSecurityMitigationService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    $getExchangeServer = $exchangeInformation.GetExchangeServer
    $mitigationEnabledAtOrg = $HealthServerObject.OrganizationInformation.GetOrganizationConfig.MitigationsEnabled
    $mitigationEnabledAtServer = $getExchangeServer.MitigationsEnabled
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }
    #Description: Check for Exchange Emergency Mitigation Service (EEMS)
    #Introduced in: Exchange 2016 CU22, Exchange 2019 CU11
    if (((Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2016" -CU "CU22") -or
        (Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -Version "Exchange2019" -CU "CU11")) -and
        $exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false) {

        if (-not([String]::IsNullOrEmpty($mitigationEnabledAtOrg))) {
            if (($mitigationEnabledAtOrg) -and
                ($mitigationEnabledAtServer)) {
                $eemsWriteType = "Green"
                $eemsOverallState = "Enabled"
            } elseif (($mitigationEnabledAtOrg -eq $false) -and
                ($mitigationEnabledAtServer)) {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled on org level"
            } elseif (($mitigationEnabledAtServer -eq $false) -and
                ($mitigationEnabledAtOrg)) {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled on server level"
            } else {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled"
            }

            $params = $baseParams + @{
                Name             = "Exchange Emergency Mitigation Service"
                Details          = $eemsOverallState
                DisplayWriteType = $eemsWriteType
            }
            Add-AnalyzedResultInformation @params

            if ($eemsWriteType -ne "Green") {
                $params = $baseParams + @{
                    Details                = "More Information: https://aka.ms/HC-EEMS"
                    DisplayWriteType       = $eemsWriteType
                    DisplayCustomTabNumber = 2
                    AddHtmlDetailRow       = $false
                }
                Add-AnalyzedResultInformation @params
            }

            $eemsWinSrvWriteType = "Yellow"
            $details = "Unknown"
            $service = $exchangeInformation.DependentServices.Monitor |
                Where-Object { $_.Name -eq "MSExchangeMitigation" }

            if ($null -ne $service) {
                if ($service.Status -eq "Running" -and $service.StartType -eq "Automatic") {
                    $details = "Running"
                    $eemsWinSrvWriteType = "Grey"
                } else {
                    $details = "Investigate"
                }
            }

            $params = $baseParams + @{
                Name                   = "Windows service"
                Details                = $details
                DisplayWriteType       = $eemsWinSrvWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            if ($exchangeInformation.ExchangeEmergencyMitigationServiceResult.StatusCode -eq 200) {
                $eemsPatternServiceWriteType = "Grey"
                $eemsPatternServiceStatus = ("200 - Reachable")
            } else {
                $eemsPatternServiceWriteType = "Yellow"
                $eemsPatternServiceStatus = "Unreachable`r`n`t`tMore information: https://aka.ms/HelpConnectivityEEMS"
            }
            $params = $baseParams + @{
                Name                   = "Pattern service"
                Details                = $eemsPatternServiceStatus
                DisplayWriteType       = $eemsPatternServiceWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            if (-not([String]::IsNullOrEmpty($getExchangeServer.MitigationsApplied))) {
                foreach ($mitigationApplied in $getExchangeServer.MitigationsApplied) {
                    $params = $baseParams + @{
                        Name                   = "Mitigation applied"
                        Details                = $mitigationApplied
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                $params = $baseParams + @{
                    Details                = "Run: 'Get-Mitigations.ps1' from: '$ExScripts' to learn more."
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            if (-not([String]::IsNullOrEmpty($getExchangeServer.MitigationsBlocked))) {
                foreach ($mitigationBlocked in $getExchangeServer.MitigationsBlocked) {
                    $params = $baseParams + @{
                        Name                   = "Mitigation blocked"
                        Details                = $mitigationBlocked
                        DisplayWriteType       = "Yellow"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            }

            if (-not([String]::IsNullOrEmpty($getExchangeServer.DataCollectionEnabled))) {
                $params = $baseParams + @{
                    Name                   = "Telemetry enabled"
                    Details                = $getExchangeServer.DataCollectionEnabled
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            Write-Verbose "Unable to validate Exchange Emergency Mitigation Service state"
            $params = $baseParams + @{
                Name             = "Exchange Emergency Mitigation Service"
                Details          = "Failed to query config"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "Exchange Emergency Mitigation Service feature not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU or on Edge Transport Server"
    }
}


# Used to determine the state of the Serialized Data Signing on the server.
function Get-SerializedDataSigningState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "HealthServerObject")]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true, ParameterSetName = "SecurityObject")]
        [object]$SecurityObject
    )
    begin {
        <#
        SerializedDataSigning was introduced with the January 2023 Exchange Server Security Update
        In the first release of the feature, it was disabled by default.
        After November 2023 Exchange Server Security Update, it was enabled by default.

        Jan23SU thru Nov23SU
        - Exchange 2016/2019 > Feature must be enabled via New-SettingOverride
        - Exchange 2013 > Feature must be enabled via EnableSerializationDataSigning registry value

        Nov23SU +
        - Exchange 2016/2019 > Feature is enabled by default, but can be disabled by New-SettingOverride.

        Note:
        If the registry value is set on E16/E19, it will be ignored.
        Same goes for the SettingOverride set on E15 - it will be ignored and the feature remains off until the registry value is set.
        #>
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "ParameterSetName: $($PSCmdlet.ParameterSetName)"

        if ($PSCmdlet.ParameterSetName -eq "HealthServerObject") {
            $exchangeInformation = $HealthServerObject.ExchangeInformation
            $getSettingOverride = $HealthServerObject.OrganizationInformation.GetSettingOverride
        } else {
            $exchangeInformation = $SecurityObject.ExchangeInformation
            $getSettingOverride = $SecurityObject.OrgInformation.GetSettingOverride
        }

        $additionalInformation = [string]::Empty
        $serializedDataSigningEnabled = $false
        $supportedRole = $exchangeInformation.GetExchangeServer.IsEdgeServer -eq $false
        $supportedVersion = (Test-ExchangeBuildGreaterOrEqualThanSecurityPatch -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -SUName "Jan23SU")
        $enabledByDefaultVersion = (Test-ExchangeBuildGreaterOrEqualThanSecurityPatch -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -SUName "Nov23SU")
        $filterServer = $exchangeInformation.GetExchangeServer.Name
        $exchangeBuild = $exchangeInformation.BuildInformation.VersionInformation.BuildVersion
        Write-Verbose "Reviewing settings against build: $exchangeBuild"
    } process {

        if ($supportedVersion -and
            $supportedRole) {
            Write-Verbose "SerializedDataSigning is available on this Exchange role / version build combination"

            if ($exchangeBuild -ge "15.1.0.0") {
                Write-Verbose "Checking SettingOverride for SerializedDataSigning configuration state"
                $params = @{
                    ExchangeSettingOverride = $exchangeInformation.SettingOverrides
                    GetSettingOverride      = $getSettingOverride
                    FilterServer            = $filterServer
                    FilterServerVersion     = $exchangeBuild
                    FilterComponentName     = "Data"
                    FilterSectionName       = "EnableSerializationDataSigning"
                    FilterParameterName     = "Enabled"
                }

                [array]$serializedDataSigningSettingOverride = Get-FilteredSettingOverrideInformation @params

                if ($null -eq $serializedDataSigningSettingOverride) {
                    Write-Verbose "No Setting Override Found"
                    $serializedDataSigningEnabled = $enabledByDefaultVersion
                } elseif ($serializedDataSigningSettingOverride.Count -eq 1) {
                    $stateValue = $serializedDataSigningSettingOverride.ParameterValue

                    if ($stateValue -eq "False") {
                        $additionalInformation = "SerializedDataSigning is explicitly disabled"
                        Write-Verbose $additionalInformation
                    } elseif ($stateValue -eq "True") {
                        Write-Verbose "SerializedDataSigning is explicitly enabled"
                        $serializedDataSigningEnabled = $true
                    } else {
                        Write-Verbose "Unknown value provided"
                        $additionalInformation = "SerializedDataSigning is unknown"
                    }
                } else {
                    Write-Verbose "Multi overrides detected"
                    $additionalInformation = "SerializedDataSigning is unknown - Multi Setting Overrides Applied: $([string]::Join(", ", $serializedDataSigningSettingOverride.Name))"
                }
            } else {
                Write-Verbose "Checking Registry Value for SerializedDataSigning configuration state"

                if ($exchangeInformation.RegistryValues.SerializedDataSigning -eq 1) {
                    $serializedDataSigningEnabled = $true
                    Write-Verbose "SerializedDataSigning enabled via Registry Value"
                } else {
                    Write-Verbose "SerializedDataSigning not configured or explicitly disabled via Registry Value"
                }
            }
        } else {
            Write-Verbose "SerializedDataSigning isn't available because we are on role: $($exchangeInformation.BuildInformation.ServerRole) build: $exchangeBuild"
        }
    } end {
        return [PSCustomObject]@{
            Enabled                 = $serializedDataSigningEnabled
            SupportedVersion        = $supportedVersion
            SupportedRole           = $supportedRole
            EnabledByDefaultVersion = $enabledByDefaultVersion
            AdditionalInformation   = $additionalInformation
        }
    }
}
function Invoke-AnalyzerSecuritySerializedDataSigningState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    $getSerializedDataSigningState = Get-SerializedDataSigningState -HealthServerObject $HealthServerObject
    # Because this is tied to public CVEs now, everything must be Red unless configured correctly
    # We must also show it even if not on the correct build of Exchange.
    $serializedDataSigningWriteType = "Red"
    $serializedDataSigningState = $false

    if ($getSerializedDataSigningState.SupportedRole -eq $false) {
        Write-Verbose "Not on a supported role, skipping over displaying this information."
        return
    }

    if ($getSerializedDataSigningState.SupportedVersion -eq $false) {
        Write-Verbose "Not on a supported version of Exchange that has serialized data signing option."
        $serializedDataSigningState = "Unsupported Version"
    } elseif ($getSerializedDataSigningState.Enabled) {
        $serializedDataSigningState = $true
        $serializedDataSigningWriteType = "Green"
    }

    $params = $baseParams + @{
        Name             = "SerializedDataSigning Enabled"
        Details          = $serializedDataSigningState
        DisplayWriteType = $serializedDataSigningWriteType
    }
    Add-AnalyzedResultInformation @params

    # Always display if not true
    if (-not ($serializedDataSigningState -eq $true)) {
        $addLine = "This may pose a security risk to your servers`r`n`t`tMore Information: https://aka.ms/HC-SerializedDataSigning"

        if (-not ([string]::IsNullOrEmpty($getSerializedDataSigningState.AdditionalInformation))) {
            $details = "$($getSerializedDataSigningState.AdditionalInformation)`r`n`t`t$addLine"
        } else {
            $details = $addLine
        }

        $params = $baseParams + @{
            Details                = $details
            DisplayWriteType       = $serializedDataSigningWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }
}
function Invoke-AnalyzerSecuritySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $osInformation = $HealthServerObject.OSInformation
    $aes256CbcInformation = $HealthServerObject.ExchangeInformation.AES256CBCInformation
    $keySecuritySettings = (Get-DisplayResultsGroupingKey -Name "Security Settings"  -DisplayOrder $Order)
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keySecuritySettings
    }

    ##############
    # TLS Settings
    ##############
    Write-Verbose "Working on TLS Settings"

    function NewDisplayObject {
        param (
            [string]$RegistryKey,
            [string]$Location,
            [object]$Value
        )
        return [PSCustomObject]@{
            RegistryKey = $RegistryKey
            Location    = $Location
            Value       = if ($null -eq $Value) { "NULL" } else { $Value }
        }
    }

    $tlsVersions = @("1.0", "1.1", "1.2", "1.3")
    $currentNetVersion = $osInformation.TLSSettings.Registry.NET["NETv4"]

    $tlsSettings = $osInformation.TLSSettings.Registry.TLS
    $misconfiguredClientServerSettings = ($tlsSettings.Values | Where-Object { $_.TLSMisconfigured -eq $true }).Count -ne 0
    $displayLinkToDocsPage = ($tlsSettings.Values | Where-Object { $_.TLSConfiguration -ne "Enabled" -and $_.TLSConfiguration -ne "Disabled" }).Count -ne 0
    $lowerTlsVersionDisabled = ($tlsSettings.Values | Where-Object { $_.TLSVersionDisabled -eq $true -and ($_.TLSVersion -ne "1.2" -and $_.TLSVersion -ne "1.3") }).Count -ne 0
    $tls13NotDisabled = ($tlsSettings.Values | Where-Object { $_.TLSConfiguration -ne "Disabled" -and $_.TLSVersion -eq "1.3" }).Count -gt 0

    $sbValue = {
        param ($o, $p)
        if ($p -eq "Value") {
            if ($o.$p -eq "NULL" -and -not $o.Location.Contains("1.3")) {
                "Red"
            }
        }
    }

    foreach ($tlsKey in $tlsVersions) {
        $currentTlsVersion = $osInformation.TLSSettings.Registry.TLS[$tlsKey]
        $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]
        $outputObjectDisplayValue.Add((NewDisplayObject "Enabled" -Location $currentTlsVersion.ServerRegistryPath -Value $currentTlsVersion.ServerEnabledValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "DisabledByDefault" -Location $currentTlsVersion.ServerRegistryPath -Value $currentTlsVersion.ServerDisabledByDefaultValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "Enabled" -Location $currentTlsVersion.ClientRegistryPath -Value $currentTlsVersion.ClientEnabledValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "DisabledByDefault" -Location $currentTlsVersion.ClientRegistryPath -Value $currentTlsVersion.ClientDisabledByDefaultValue))
        $displayWriteType = "Green"

        # Any TLS version is Misconfigured or Half Disabled is Red
        # Only TLS 1.2 being Disabled is Red
        # Currently TLS 1.3 being Enabled is Red
        # TLS 1.0 or 1.1 being Enabled is Yellow as we recommend to disable this weak protocol versions
        if (($currentTlsVersion.TLSConfiguration -eq "Misconfigured" -or
                $currentTlsVersion.TLSConfiguration -eq "Half Disabled") -or
                ($tlsKey -eq "1.2" -and $currentTlsVersion.TLSConfiguration -eq "Disabled") -or
                ($tlsKey -eq "1.3" -and $currentTlsVersion.TLSConfiguration -eq "Enabled")) {
            $displayWriteType = "Red"
        } elseif ($currentTlsVersion.TLSConfiguration -eq "Enabled" -and
            ($tlsKey -eq "1.1" -or $tlsKey -eq "1.0")) {
            $displayWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "TLS $tlsKey"
            Details          = $currentTlsVersion.TLSConfiguration
            DisplayWriteType = $displayWriteType
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            OutColumns           = ([PSCustomObject]@{
                    DisplayObject      = $outputObjectDisplayValue
                    ColorizerFunctions = @($sbValue)
                    IndentSpaces       = 8
                })
            OutColumnsColorTests = @($sbValue)
            HtmlName             = "TLS Settings $tlsKey"
            TestingName          = "TLS Settings Group $tlsKey"
        }
        Add-AnalyzedResultInformation @params
    }

    $netVersions = @("NETv4", "NETv2")
    $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

    $sbValue = {
        param ($o, $p)
        if ($p -eq "Value") {
            if ($o.$p -eq "NULL" -and $o.Location -like "*v4.0.30319") {
                "Red"
            }
        }
    }

    foreach ($netVersion in $netVersions) {
        $currentNetVersion = $osInformation.TLSSettings.Registry.NET[$netVersion]
        $outputObjectDisplayValue.Add((NewDisplayObject "SystemDefaultTlsVersions" -Location $currentNetVersion.MicrosoftRegistryLocation -Value $currentNetVersion.SystemDefaultTlsVersionsValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SchUseStrongCrypto" -Location $currentNetVersion.MicrosoftRegistryLocation -Value $currentNetVersion.SchUseStrongCryptoValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SystemDefaultTlsVersions" -Location $currentNetVersion.WowRegistryLocation -Value $currentNetVersion.WowSystemDefaultTlsVersionsValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SchUseStrongCrypto" -Location $currentNetVersion.WowRegistryLocation -Value $currentNetVersion.WowSchUseStrongCryptoValue))
    }

    $params = $baseParams + @{
        OutColumns  = ([PSCustomObject]@{
                DisplayObject      = $outputObjectDisplayValue
                ColorizerFunctions = @($sbValue)
                IndentSpaces       = 8
            })
        HtmlName    = "TLS NET Settings"
        TestingName = "NET TLS Settings Group"
    }
    Add-AnalyzedResultInformation @params

    $testValues = @("ServerEnabledValue", "ClientEnabledValue", "ServerDisabledByDefaultValue", "ClientDisabledByDefaultValue")

    foreach ($testValue in $testValues) {

        # if value not defined, we should call that out.
        $results = $tlsSettings.Values | Where-Object { $null -eq $_."$testValue" -and $_.TLSVersion -ne "1.3" }

        if ($null -ne $results) {
            $displayLinkToDocsPage = $true
            foreach ($result in $results) {
                $params = $baseParams + @{
                    Name             = "$($result.TLSVersion) $testValue"
                    Details          = "NULL --- Error: Value should be defined in registry for consistent results."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    # Check for NULL values on NETv4 registry settings
    $testValues = @("SystemDefaultTlsVersionsValue", "SchUseStrongCryptoValue", "WowSystemDefaultTlsVersionsValue", "WowSchUseStrongCryptoValue")

    foreach ($testValue in $testValues) {
        $results = $osInformation.TLSSettings.Registry.NET["NETv4"] | Where-Object { $null -eq $_."$testValue" }
        if ($null -ne $results) {
            $displayLinkToDocsPage = $true
            foreach ($result in $results) {
                $params = $baseParams + @{
                    Name             = "$($result.NetVersion) $testValue"
                    Details          = "NULL --- Error: Value should be defined in registry for consistent results."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($lowerTlsVersionDisabled -and
        ($osInformation.TLSSettings.Registry.NET["NETv4"].SystemDefaultTlsVersions -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].WowSystemDefaultTlsVersions -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].SchUseStrongCrypto -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].WowSchUseStrongCrypto -eq $false)) {
        $params = $baseParams + @{
            Details                = "Error: SystemDefaultTlsVersions or SchUseStrongCrypto is not set to the recommended value. Please visit on how to properly enable TLS 1.2 https://aka.ms/HC-TLSGuide"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($misconfiguredClientServerSettings) {
        $params = $baseParams + @{
            Details                = "Error: Mismatch in TLS version for client and server. Exchange can be both client and a server. This can cause issues within Exchange for communication."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "For More Information on how to properly set TLS follow this guide: https://aka.ms/HC-TLSGuide"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Detected TLS Mismatch Display More Info"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($tls13NotDisabled) {
        $displayLinkToDocsPage = $true
        $params = $baseParams + @{
            Details                = "Error: TLS 1.3 is not disabled and not supported currently on Exchange and is known to cause issues within the cluster."
            DisplayWriteType       = "Red"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "TLS 1.3 not disabled"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($lowerTlsVersionDisabled -eq $false) {
        $displayLinkToDocsPage = $true
        $params = $baseParams + @{
            Name = "TLS hardening recommendations"
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "Microsoft recommends customers proactively address weak TLS usage by removing TLS 1.0/1.1 dependencies in their environments and disabling TLS 1.0/1.1 at the operating system level where possible."
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($displayLinkToDocsPage) {
        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-TLSConfigDocs"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Display Link to Docs Page"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "SecurityProtocol"
        Details = $osInformation.TLSSettings.SecurityProtocol
    }
    Add-AnalyzedResultInformation @params

    if ($null -ne $osInformation.TLSSettings.TlsCipherSuite) {
        $outputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'

        foreach ($tlsCipher in $osInformation.TLSSettings.TlsCipherSuite) {
            $outputObjectDisplayValue.Add(([PSCustomObject]@{
                        TlsCipherSuiteName = $tlsCipher.Name
                        CipherSuite        = $tlsCipher.CipherSuite
                        Cipher             = $tlsCipher.Cipher
                        Certificate        = $tlsCipher.Certificate
                        Protocols          = $tlsCipher.Protocols
                    })
            )
        }

        $params = $baseParams + @{
            OutColumns  = ([PSCustomObject]@{
                    DisplayObject = $outputObjectDisplayValue
                    IndentSpaces  = 8
                })
            HtmlName    = "TLS Cipher Suite"
            TestingName = "TLS Cipher Suite Group"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "AllowInsecureRenegoClients Value"
        Details = $osInformation.RegistryValues.AllowInsecureRenegoClients
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "AllowInsecureRenegoServers Value"
        Details = $osInformation.RegistryValues.AllowInsecureRenegoServers
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "LmCompatibilityLevel Settings"
        Details = $osInformation.RegistryValues.LmCompatibilityLevel
    }
    Add-AnalyzedResultInformation @params

    $description = [string]::Empty
    switch ($osInformation.RegistryValues.LmCompatibilityLevel) {
        0 { $description = "Clients use LM and NTLM authentication, but they never use NTLMv2 session security. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        1 { $description = "Clients use LM and NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        2 { $description = "Clients use only NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controller accepts LM, NTLM, and NTLMv2 authentication." }
        3 { $description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        4 { $description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM authentication responses, but it accepts NTLM and NTLMv2." }
        5 { $description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM and NTLM authentication responses, but it accepts NTLMv2." }
    }

    $params = $baseParams + @{
        Name                   = "Description"
        Details                = $description
        DisplayCustomTabNumber = 2
        AddHtmlDetailRow       = $false
    }
    Add-AnalyzedResultInformation @params

    # AES256-CBC encryption support check
    $sp = "Supported Build"
    $vc = "Valid Configuration"
    $params = $baseParams + @{
        Name                = "AES256-CBC Protected Content Support"
        Details             = $true
        DisplayWriteType    = "Green"
        DisplayTestingValue = "$sp and $vc"
    }

    $irmConfig = $HealthServerObject.OrganizationInformation.GetIrmConfiguration

    if (($aes256CbcInformation.AES256CBCSupportedBuild) -and
        ($aes256CbcInformation.ValidAESConfiguration -eq $false) -and
        ($irmConfig.InternalLicensingEnabled -eq $true -or
        $irmConfig.ExternalLicensingEnabled -eq $true)) {
        $params.DisplayTestingValue = "$sp and not $vc"
        $params.Details = ("True" +
            "`r`n`t`tThis build supports AES256-CBC protected content, but the configuration is not complete. Exchange Server is not able to decrypt" +
            "`r`n`t`tprotected messages which could impact eDiscovery and Journaling tasks. If you use Rights Management Service (RMS) on-premises," +
            "`r`n`t`tplease follow the instructions as outlined in the documentation: https://aka.ms/ExchangeCBCKB")

        if ($irmConfig.InternalLicensingEnabled -eq $true) {
            $params.DisplayWriteType = "Red"
        } else {
            $params.DisplayWriteType = "Yellow"
        }
    } elseif ($aes256CbcInformation.AES256CBCSupportedBuild -eq $false) {
        $params.DisplayTestingValue = "Not $sp"
        $params.Details = ("False" +
            "`r`n`t`tThis could lead to scenarios where Exchange Server is no longer able to decrypt protected messages," +
            "`r`n`t`tfor example, when sending rights management protected messages using AES256-CBC encryption algorithm," +
            "`r`n`t`tor when performing eDiscovery and Journaling tasks." +
            "`r`n`t`tMore Information: https://aka.ms/Purview/CBCDetails")
        $params.DisplayWriteType = "Red"
    }
    Add-AnalyzedResultInformation @params

    $additionalDisplayValue = [string]::Empty
    $smb1Settings = $osInformation.Smb1ServerSettings

    if ($osInformation.BuildInformation.BuildVersion -ge "10.0.0.0" -or
        $osInformation.BuildInformation.MajorVersion -eq "Windows2012R2") {
        $displayValue = "False"
        $writeType = "Green"

        if (-not ($smb1Settings.SuccessfulGetInstall)) {
            $displayValue = "Failed to get install status"
            $writeType = "Yellow"
        } elseif ($smb1Settings.Installed) {
            $displayValue = "True"
            $writeType = "Red"
            $additionalDisplayValue = "SMB1 should be uninstalled"
        }

        $params = $baseParams + @{
            Name             = "SMB1 Installed"
            Details          = $displayValue
            DisplayWriteType = $writeType
        }
        Add-AnalyzedResultInformation @params
    }

    $writeType = "Green"
    $displayValue = "True"

    if (-not ($smb1Settings.SuccessfulGetBlocked)) {
        $displayValue = "Failed to get block status"
        $writeType = "Yellow"
    } elseif (-not($smb1Settings.IsBlocked)) {
        $displayValue = "False"
        $writeType = "Red"
        $additionalDisplayValue += " SMB1 should be blocked"
    }

    $params = $baseParams + @{
        Name             = "SMB1 Blocked"
        Details          = $displayValue
        DisplayWriteType = $writeType
    }
    Add-AnalyzedResultInformation @params

    if ($additionalDisplayValue -ne [string]::Empty) {
        $additionalDisplayValue += "`r`n`t`tMore Information: https://aka.ms/HC-SMB1"

        $params = $baseParams + @{
            Details                = $additionalDisplayValue.Trim()
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    Invoke-AnalyzerSecurityExchangeCertificates -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecurityAMSIConfigState -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecuritySerializedDataSigningState -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecurityOverrides -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecurityMitigationService -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings

    if ($null -ne $HealthServerObject.ExchangeInformation.FIPFSUpdateIssue) {
        $fipFsInfoObject = $HealthServerObject.ExchangeInformation.FIPFSUpdateIssue
        $highestVersion = $fipFsInfoObject.HighestVersionNumberDetected
        $fipFsIssueBaseParams = @{
            Name             = "FIP-FS Update Issue Detected"
            Details          = $true
            DisplayWriteType = "Red"
        }
        $moreInformation = "More Information: https://aka.ms/HC-FIPFSUpdateIssue"

        if ($fipFsInfoObject.ServerRoleAffected -eq $false) {
            # Server role is not affected by the FIP-FS issue so we don't need to check for the other conditions.
            Write-Verbose "The Exchange server runs a role which is not affected by the FIP-FS issue"
        } elseif (($fipFsInfoObject.FIPFSFixedBuild -eq $false) -and
            ($fipFsInfoObject.BadVersionNumberDirDetected)) {
            # Exchange doesn't run a build which is resistent against the problematic pattern
            # and a folder with the problematic version number was detected on the computer.
            $params = $baseParams + $fipFsIssueBaseParams
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = $moreInformation
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } elseif (($fipFsInfoObject.FIPFSFixedBuild) -and
            ($fipFsInfoObject.BadVersionNumberDirDetected)) {
            # Exchange runs a build that can handle the problematic pattern. However, we found
            # a high-version folder which should be removed (recommendation).
            $fipFsIssueBaseParams.DisplayWriteType = "Yellow"
            $params = $baseParams + $fipFsIssueBaseParams
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "Detected problematic FIP-FS version $highestVersion directory`r`n`t`tAlthough it should not cause any problems, we recommend performing a FIP-FS reset`r`n`t`t$moreInformation"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } elseif ($null -eq $fipFsInfoObject.HighestVersionNumberDetected) {
            # No scan engine was found on the Exchange server. This will cause multiple issues on transport.
            $fipFsIssueBaseParams.Details = "Error: Failed to find the scan engines on server, this can cause issues with transport rules as well as the malware agent."
            $params = $baseParams + $fipFsIssueBaseParams
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "Server runs a FIP-FS fixed build: $($fipFsInfoObject.FIPFSFixedBuild) - Highest version number: $highestVersion"
        }
    } else {
        $fipFsIssueBaseParams = $baseParams + @{
            Name             = "FIP-FS Update Issue Detected"
            Details          = "Warning: Unable to check if the system is vulnerable to the FIP-FS bad pattern issue. Please re-run. $moreInformation"
            DisplayWriteType = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }
}




<#
.DESCRIPTION
    Check for ADV24199947 Outside In Module vulnerability
    Must be on March 2024 SU and no overrides in place to be considered secure.
    Overrides are found in the Configuration.xml file with appending flag of |NO
    This only needs to occur on the Mailbox Servers Roles
#>
function Invoke-AnalyzerSecurityADV24199947 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $params = @{
            AnalyzedInformation = $AnalyzeResults
            DisplayGroupingKey  = $DisplayGroupingKey
            Name                = "Security Vulnerability"
            DisplayWriteType    = "Red"
            Details             = "{0}"
            DisplayTestingValue = "ADV24199947"
        }

        if ($SecurityObject.IsEdgeServer) {
            Write-Verbose "Skipping over test as this is an edge server."
            return
        }

        $isVulnerable = (-not (Test-ExchangeBuildGreaterOrEqualThanSecurityPatch -CurrentExchangeBuild $SecurityObject.BuildInformation -SUName "Mar24SU"))

        # if patch is installed, need to check for the override.
        if ($isVulnerable -eq $false) {
            Write-Verbose "Mar24SU is installed, checking to see if override is set"
            # Key for the file content information
            $key = [System.IO.Path]::Combine($SecurityObject.ExchangeInformation.RegistryValues.FipFsDatabasePath, "Configuration.xml")
            $unknownError = [string]::IsNullOrEmpty($SecurityObject.ExchangeInformation.RegistryValues.FipFsDatabasePath) -or
                ($null -eq $SecurityObject.ExchangeInformation.FileContentInformation[$key])

            if ($unknownError) {
                $params.Details += " Unable to determine if override is set due to no data to review."
                $params.DisplayWriteType = "Yellow"
                $isVulnerable = $true
            } else {
                $isVulnerable = $null -ne ($SecurityObject.ExchangeInformation.FileContentInformation[$key] | Select-String "\|NO")
            }
        }

        if ($isVulnerable) {
            $params.Details = ("$($params.Details)`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information." -f "ADV24199947")
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "Not vulnerable to ADV24199947"
        }
    }
}

function Invoke-AnalyzerSecurityCve-2020-0796 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    #Description: Check for CVE-2020-0796 SMBv3 vulnerability
    #Affected OS versions: Windows 10 build 1903 and 1909
    #Fix: KB4551762
    #Workaround: Disable SMBv3 compression

    if ($SecurityObject.MajorVersion -eq "Exchange2019") {
        Write-Verbose "Testing CVE: CVE-2020-0796"
        $buildNumber = $SecurityObject.OsInformation.BuildInformation.BuildVersion.Build

        if (($buildNumber -eq 18362 -or
                $buildNumber -eq 18363) -and
            ($SecurityObject.OsInformation.RegistryValues.CurrentVersionUbr -lt 720)) {
            Write-Verbose "Build vulnerable to CVE-2020-0796. Checking if workaround is in place."
            $writeType = "Red"
            $writeValue = "System Vulnerable"

            if ($SecurityObject.OsInformation.RegistryValues.LanManServerDisabledCompression -eq 1) {
                Write-Verbose "Workaround to disable affected SMBv3 compression is in place."
                $writeType = "Yellow"
                $writeValue = "Workaround is in place"
            } else {
                Write-Verbose "Workaround to disable affected SMBv3 compression is NOT in place."
            }

            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "CVE-2020-0796"
                Details             = "$writeValue`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796 for more information."
                DisplayWriteType    = $writeType
                DisplayTestingValue = "CVE-2020-0796"
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2020-0796. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796"
        }
    } else {
        Write-Verbose "Operating System NOT vulnerable to CVE-2020-0796."
    }
}

function Invoke-AnalyzerSecurityCve-2020-1147 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    #Description: Check for CVE-2020-1147
    #Affected OS versions: Every OS supporting .NET Core 2.1 and 3.1 and .NET Framework 2.0 SP2 or above
    #Fix: https://portal.msrc.microsoft.com/en-US/security-guidance/advisory/CVE-2020-1147
    #Workaround: N/A
    $dllFileBuildPartToCheckAgainst = 3630

    if ($SecurityObject.OsInformation.NETFramework.MajorVersion -eq ((GetNetVersionDictionary)["Net4d8"])) {
        $dllFileBuildPartToCheckAgainst = 4190
    }

    $systemDataDll = $SecurityObject.OsInformation.NETFramework.FileInformation["System.Data.dll"]
    $systemConfigurationDll = $SecurityObject.OsInformation.NETFramework.FileInformation["System.Configuration.dll"]
    Write-Verbose "System.Data.dll FileBuildPart: $($systemDataDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemDataDll.LastWriteTimeUtc)"
    Write-Verbose "System.Configuration.dll FileBuildPart: $($systemConfigurationDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemConfigurationDll.LastWriteTimeUtc)"

    if ($systemDataDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemConfigurationDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemDataDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) -and
        $systemConfigurationDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo))) {
        Write-Verbose ("System NOT vulnerable to {0}. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0}" -f "CVE-2020-1147")
    } else {
        $params = @{
            AnalyzedInformation = $AnalyzeResults
            DisplayGroupingKey  = $DisplayGroupingKey
            Name                = "Security Vulnerability"
            Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0} for more information." -f "CVE-2020-1147")
            DisplayWriteType    = "Red"
            DisplayTestingValue = "CVE-2020-1147"
            AddHtmlDetailRow    = $false
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerSecurityCve-2021-1730 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: Check for CVE-2021-1730 vulnerability
    #Fix available for: Exchange 2016 CU18+, Exchange 2019 CU7+
    #Fix: Configure Download Domains feature
    #Workaround: N/A

    if (((Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $SecurityObject.BuildInformation -Version "Exchange2016" -CU "CU18") -or
            (Test-ExchangeBuildGreaterOrEqualThanBuild -CurrentExchangeBuild $SecurityObject.BuildInformation -Version "Exchange2019" -CU "CU7")) -and
        $SecurityObject.IsEdgeServer -eq $false) {

        $downloadDomainsEnabled = $SecurityObject.OrgInformation.EnableDownloadDomains
        $owaVDirObject = $SecurityObject.ExchangeInformation.VirtualDirectories.GetOwaVirtualDirectory |
            Where-Object { $_.Name -eq "owa (Default Web Site)" }
        $displayWriteType = "Green"

        if (-not ($downloadDomainsEnabled)) {
            $downloadDomainsOrgDisplayValue = "Download Domains are not configured. You should configure them to be protected against CVE-2021-1730.`r`n`t`tConfiguration instructions: https://aka.ms/HC-DownloadDomains"
            $displayWriteType = "Red"
        } else {
            if (-not ([String]::IsNullOrEmpty($OwaVDirObject.ExternalDownloadHostName))) {
                if (($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                            ($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadExternalDisplayValue = "Set to the same as Internal Or External URL as OWA."
                    $displayWriteType = "Red"
                } else {
                    $downloadExternalDisplayValue = "Set Correctly."
                }
            } else {
                $downloadExternalDisplayValue = "Not Configured"
                $displayWriteType = "Red"
            }

            if (-not ([string]::IsNullOrEmpty($owaVDirObject.InternalDownloadHostName))) {
                if (($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                            ($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadInternalDisplayValue = "Set to the same as Internal Or External URL as OWA."
                    $displayWriteType = "Red"
                } else {
                    $downloadInternalDisplayValue = "Set Correctly."
                }
            } else {
                $displayWriteType = "Red"
                $downloadInternalDisplayValue = "Not Configured"
            }

            $downloadDomainsOrgDisplayValue = "Download Domains are configured.`r`n`t`tExternalDownloadHostName: $downloadExternalDisplayValue`r`n`t`tInternalDownloadHostName: $downloadInternalDisplayValue`r`n`t`tConfiguration instructions: https://aka.ms/HC-DownloadDomains"
        }

        #Only display security vulnerability if present
        if ($displayWriteType -eq "Red") {
            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "Security Vulnerability"
                Details             = $downloadDomainsOrgDisplayValue
                DisplayWriteType    = "Red"
                TestingName         = "CVE-2021-1730"
                DisplayTestingValue = ([PSCustomObject]@{
                        DownloadDomainsEnabled   = $downloadDomainsEnabled
                        ExternalDownloadHostName = $downloadExternalDisplayValue
                        InternalDownloadHostName = $downloadInternalDisplayValue
                    })
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "Download Domains feature not available because we are on: $($SecurityObject.MajorVersion) $($SecurityObject.CU) or on Edge Transport Server"
    }
}

function Invoke-AnalyzerSecurityCve-2021-34470 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: Check for CVE-2021-34470 rights elevation vulnerability
    #Affected Exchange versions: 2013, 2016, 2019
    #Fix:
    ##Exchange 2013 CU23 + July 2021 SU + /PrepareSchema,
    ##Exchange 2016 CU20 + July 2021 SU + /PrepareSchema or CU21,
    ##Exchange 2019 CU9 + July 2021 SU + /PrepareSchema or CU10
    #Workaround: N/A

    if (($SecurityObject.MajorVersion -eq "Exchange2013") -or
        ((Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $SecurityObject.BuildInformation -Version "Exchange2016" -CU "CU21") -or
            (Test-ExchangeBuildLessThanBuild -CurrentExchangeBuild $SecurityObject.BuildInformation -Version "Exchange2019" -CU "CU10")) -and
        $SecurityObject.IsEdgeServer -eq $false) {
        Write-Verbose "Testing CVE: CVE-2021-34470"

        $displayWriteTypeColor = $null
        if ($SecurityObject.MajorVersion -eq "Exchange2013") {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision "$($SecurityObject.ExchangeInformation.BuildInformation.ExchangeSetup.FileBuildPart).$($SecurityObject.ExchangeInformation.BuildInformation.ExchangeSetup.FilePrivatePart)" -SecurityFixedBuilds "1497.23" -CVENames "CVE-2021-34470"
        }

        if ($SecurityObject.OrgInformation.SecurityResults.CVE202134470.Unknown -or
            $SecurityObject.OrgInformation.SecurityResults.CVE202134470.IsVulnerable.ToString() -eq "Unknown") {
            Write-Verbose "Unable to query classSchema: 'ms-Exch-Storage-Group' information"
            $details = "CVE-2021-34470`r`n`t`tWarning: Unable to query classSchema: 'ms-Exch-Storage-Group' to perform testing."
            $displayWriteTypeColor = "Yellow"
        } elseif ($SecurityObject.OrgInformation.SecurityResults.CVE202134470.IsVulnerable -eq $true) {
            Write-Verbose "Attribute: 'possSuperiors' with value: 'computer' detected in classSchema: 'ms-Exch-Storage-Group'"
            $details = "CVE-2021-34470`r`n`t`tPrepareSchema required: https://aka.ms/HC-July21SU"
            $displayWriteTypeColor = "Red"
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2021-34470"
        }

        if ($null -ne $displayWriteTypeColor) {
            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "Security Vulnerability"
                Details             = $details
                DisplayWriteType    = $displayWriteTypeColor
                DisplayTestingValue = "CVE-2021-34470"
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "System NOT vulnerable to CVE-2021-34470"
    }
}

function Invoke-AnalyzerSecurityCve-2022-21978 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    # Description: Check for CVE-2022-21978 vulnerability
    # Affected Exchange versions: 2013, 2016, 2019
    # Fix:
    # Exchange 2013 CU23 + May 2022 SU + /PrepareDomain or /PrepareAllDomains,
    # Exchange 2016 CU22/CU23 + May 2022 SU + /PrepareDomain or /PrepareAllDomains,
    # Exchange 2019 CU11/CU12 + May 2022 SU + /PrepareDomain or /PrepareAllDomains
    # Workaround: N/A

    # Because this is a security vulnerability in the domain, doesn't matter what version of Exchange is installed, still need to check each domain.
    if ($SecurityObject.IsEdgeServer -eq $false) {
        Write-Verbose "Testing CVE: CVE-2022-21978"

        $cveResults = $SecurityObject.OrgInformation.SecurityResults.CVE202221978
        $domainFailedResults = New-Object 'System.Collections.Generic.List[string]'
        $domainUnknownResults = New-Object 'System.Collections.Generic.List[string]'
        $params = @{
            AnalyzedInformation = $AnalyzeResults
            DisplayGroupingKey  = $DisplayGroupingKey
            Name                = "Security Vulnerability"
            Details             = $null
            DisplayWriteType    = $null
            DisplayTestingValue = "CVE-2022-21978"
        }
        if ($null -ne $cveResults -or
            $cveResults.Count -gt 0) {
            Write-Verbose "Exchange AD permission information found - performing vulnerability testing"
            foreach ($entry in $cveResults) {

                if ($entry.DomainPassed -eq $false -and $entry.UnknownDomain -eq $false) {
                    $domainFailedResults.Add($entry.DomainName)
                } elseif ($entry.UnknownDomain) {
                    $domainUnknownResults.Add($entry.DomainName)
                }
            }
        } else {
            Write-Verbose "Unable to perform CVE-2022-21978 vulnerability testing"
            $params.details = "CVE-2022-21978`r`n`t`tUnable to perform vulnerability testing. If Exchange admins do not have domain permissions this might be expected, please re-run with domain or enterprise admin account. - See: https://aka.ms/HC-May22SU"
            $params.displayWriteType = "Yellow"
        }

        if ($domainFailedResults.Count -gt 0 -or
            $domainUnknownResults.Count -gt 0) {

            $params.Details = "CVE-2022-21978"

            if ($domainFailedResults.Count -eq 0) {
                $params.DisplayWriteType = "Yellow"
            } else {
                $params.DisplayWriteType = "Red"
            }

            if ($domainFailedResults.Count -gt 0) {
                $params.Details += "`r`n`t`tDetected the following domains that are vulnerable: $([string]::Join(",", $domainFailedResults))"
            }

            if ($domainUnknownResults.Count -gt 0) {
                $params.Details += "`r`n`t`tUnable to perform vulnerability testing of the following domains: $([string]::Join(",", $domainUnknownResults))"
                $params.Details += "`r`n`t`tIf Exchange admins do not have domain permissions this might be expected, please re-run with domain or enterprise admin account."
            }

            $params.Details += "`r`n`t`tMore Information: https://aka.ms/HC-May22SU"
        }

        if ($null -ne $params.Details) {
            Add-AnalyzedResultInformation @params
        }
    }
}

function Invoke-AnalyzerSecurityCve-2023-36434 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    <#
        Description: Check for CVE-2023-36434 vulnerability (also tracked as CVE-2023-21709)
        Affected Exchange versions: 2016, 2019
        Fix: Install October 2023 Windows Security Update
        Workaround: Remove TokenCacheModule from IIS by running the CVE-2023-21709.ps1 script
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $tokenCacheModuleVersionInformation = $SecurityObject.ExchangeInformation.IISSettings.IISTokenCacheModuleInformation
        $tokenCacheFixedVersionNumber = $null
        $tokenCacheVersionGreaterOrEqual = $false
    }
    process {
        if ($SecurityObject.IsEdgeServer -eq $false) {
            Write-Verbose "Testing CVE: CVE-2023-21709 / CVE-2023-36434"

            if ($SecurityObject.ExchangeInformation.IISSettings.IISModulesInformation.ModuleList.Name -contains "TokenCacheModule") {
                Write-Verbose "TokenCacheModule detected - system could be vulnerable to CVE-2023-21709 / CVE-2023-36434 vulnerability"

                if ($null -ne $tokenCacheModuleVersionInformation) {
                    Write-Verbose "TokenCacheModule build information found - performing build analysis now..."
                    switch ($tokenCacheModuleVersionInformation.FileBuildPart) {
                        9200 { $tokenCacheFixedVersionNumber = "8.0.9200.24514"; break } # Windows Server 2012
                        9600 { $tokenCacheFixedVersionNumber = "8.5.9600.21613"; break } # Windows Server 2012 R2
                        14393 { $tokenCacheFixedVersionNumber = "10.0.14393.6343"; break } # Windows Server 2016
                        17763 { $tokenCacheFixedVersionNumber = "10.0.17763.4968"; break } # Windows Server 2019
                        20348 { $tokenCacheFixedVersionNumber = "10.0.20348.2029"; break } # Windows Server 2022
                        default { Write-Verbose "No fixed TokenCacheModule version available for Windows OS build: $($tokenCacheModuleVersionInformation.FileBuildPart)" }
                    }

                    if ($null -ne $tokenCacheFixedVersionNumber) {
                        Write-Verbose "Build: $($tokenCacheModuleVersionInformation.FileBuildPart) found - testing against version: $tokenCacheFixedVersionNumber"
                        $tokenCacheVersionGreaterOrEqual = ([system.version]$tokenCacheModuleVersionInformation.ProductVersion -ge $tokenCacheFixedVersionNumber)
                        Write-Verbose "Version: $($tokenCacheModuleVersionInformation.ProductVersion) is greater or equal the expected version? $tokenCacheVersionGreaterOrEqual"
                    }
                } else {
                    Write-Verbose "We were unable to query TokenCacheModule build information - as the module is loaded, we're assuming that it's vulnerable"
                }

                if ($tokenCacheVersionGreaterOrEqual -eq $false) {
                    $params = @{
                        AnalyzedInformation = $AnalyzeResults
                        DisplayGroupingKey  = $DisplayGroupingKey
                        Name                = "Security Vulnerability"
                        Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information." -f "CVE-2023-36434")
                        DisplayWriteType    = "Red"
                        DisplayTestingValue = "CVE-2023-36434"
                        AddHtmlDetailRow    = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
            }
        } else {
            Write-Verbose "Edge Server Role is not affected by this vulnerability as it has no IIS installed"
        }
    }
}

function Invoke-AnalyzerSecurityCveAddressedBySerializedDataSigning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    <#
        Description: Check for vulnerabilities that are addressed by turning serialized data signing for PowerShell payload on
        Affected Exchange versions: 2016, 2019
        Fix: Enable Serialized Data Signing for PowerShell payload if disabled or install Exchange update if running an unsupported build
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        function NewCveFixedBySDSObject {
            param()

            begin {
                Write-Verbose "Calling: $($MyInvocation.MyCommand)"
                $cveList = New-Object 'System.Collections.Generic.List[object]'

                # Add all CVE that are addressed by turning Serialized Data Signing for PowerShell payload on
                # Add true or false as an indicator as some fixes needs to be done via code fix + SDS on
                $cveFixedBySDS = @(
                    "CVE-2023-36050, $true",
                    "CVE-2023-36039, $true",
                    "CVE-2023-36035, $true",
                    "CVE-2023-36439, $true")
            } process {
                foreach ($cve in $cveFixedBySDS) {
                    $entry = $($cve.Split(",")[0]).Trim()
                    $fixIndicator = $($cve.Split(",")[1]).Trim()
                    $cveList.Add([PSCustomObject]@{
                            CVE             = $entry
                            CodeFixRequired = $fixIndicator
                        })
                }
            } end {
                return $cveList
            }
        }

        function FindCveEntryInAnalyzeResults {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Value is used')]
            param (
                [Parameter(Mandatory = $true)]
                [ref]$AnalyzeResults,

                [Parameter(Mandatory = $true)]
                [string]$CVE,

                [Parameter(Mandatory = $false)]
                [switch]$RemoveWhenFound
            )

            begin {
                Write-Verbose "Calling: $($MyInvocation.MyCommand)"

                $key = $null
                $cveFound = $false
            } process {
                ($AnalyzeResults.Value.DisplayResults.Values | Where-Object {
                    # Find the 'Security Vulnerability' section
                    ($_.Name -eq "Security Vulnerability")
                }) | ForEach-Object {
                    if ($_.CustomValue -match $CVE) {
                        # Loop through each entry and check if the value is equal the CVE that we're looking for
                        Write-Verbose ("$CVE was found in the CVE list!")
                        $key = $_
                    }
                }

                $cveFound = ($null -ne $key)

                if ($RemoveWhenFound -and
                    $cveFound) {
                    # Remove the entry if found and if RemovedWhenFound parameter was used
                    Write-Verbose ("Removing $CVE from the list")
                    $AnalyzeResults.Value.DisplayResults.Values.Remove($key)
                }
            } end {
                Write-Verbose ("Was $CVE found in the list? $cveFound")
                return $cveFound
            }
        }

        $params = @{
            AnalyzedInformation = $AnalyzeResults
            DisplayGroupingKey  = $DisplayGroupingKey
            Name                = "Security Vulnerability"
            DisplayWriteType    = "Red"
        }

        $detailsString = "{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information."

        $getSerializedDataSigningState = Get-SerializedDataSigningState -SecurityObject $SecurityObject
        $cveFixedBySerializedDataSigning = NewCveFixedBySDSObject
    }
    process {
        if ($getSerializedDataSigningState.SupportedRole -ne $false) {
            if ($cveFixedBySerializedDataSigning.Count -ge 1) {
                Write-Verbose ("Testing CVEs: {0}" -f [string]::Join(", ", $cveFixedBySerializedDataSigning.CVE))

                if (($getSerializedDataSigningState.SupportedVersion) -and
                    ($getSerializedDataSigningState.Enabled)) {
                    Write-Verbose ("Serialized Data Signing is supported and enabled - removing any CVE that is mitigated by this feature")

                    foreach ($entry in $cveFixedBySerializedDataSigning) {
                        $buildIsVulnerable = $null
                        # If we find it on the AnalyzedResults list, it means that the build is outdated and as a result vulnerable
                        $buildIsVulnerable = FindCveEntryInAnalyzeResults -AnalyzeResults $AnalyzeResults -CVE $($entry.CVE)
                        if ($entry.CodeFixRequired -and
                            $buildIsVulnerable) {
                            # SDS is configured but there is a code change required that comes as part of a newer Exchange build.
                            # We consider this version as vulnerable since it's running an outdated build.
                            Write-Verbose ("To be fully protected against this vulnerability, a fixed Exchange build is required")
                        } elseif (($entry.CodeFixRequired -eq $false) -and
                            ($buildIsVulnerable)) {
                            # SDS is configured as expected and there is no code change required.
                            # We consider this combination as secure since the Exchange build was vulnerable but SDS mitigates.
                            Write-Verbose ("CVE was on this list but was removed since SDS mitigates the vulnerability")
                            FindCveEntryInAnalyzeResults -AnalyzeResults $AnalyzeResults -CVE $($entry.CVE) -RemoveWhenFound
                        } else {
                            # We end up here if build is not vulnerable
                            Write-Verbose ("CVE wasn't on the list - system seems not to be vulnerable")
                        }
                    }
                } elseif (($getSerializedDataSigningState.SupportedVersion -eq $false) -or
                    ($getSerializedDataSigningState.Enabled -eq $false)) {

                    foreach ($entry in $cveFixedBySerializedDataSigning) {
                        Write-Verbose ("System is vulnerable to: $($entry.CVE)")

                        if ((FindCveEntryInAnalyzeResults -AnalyzeResults $AnalyzeResults -CVE $($entry.CVE)) -eq $false) {
                            Write-Verbose ("CVE wasn't found in the results list and will be added now as it requires SDS to be mitigated")
                            $params.Details = $detailsString -f $($entry.CVE)
                            $params.DisplayTestingValue = $($entry.CVE)
                            Add-AnalyzedResultInformation @params
                        } else {
                            # We end up here in case the CVE is already on the list
                            Write-Verbose ("CVE is already on the results list")
                        }
                    }
                }
            } else {
                Write-Verbose "There are no vulnerabilities that have been addressed by enabling serialized data signing"
            }
        } else {
            Write-Verbose "Exchange server role is not affected by these vulnerabilities"
        }
    }
}

function Invoke-AnalyzerSecurityCve-MarchSuSpecial {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: March 2021 Exchange vulnerabilities Security Update (SU) check for outdated version (CUs)
    #Affected Exchange versions: Exchange 2013, Exchange 2016, Exchange 2016 (we only provide this special SU for these versions)
    #Fix: Update to a supported CU and apply KB5000871
    $march2021SUInstalled = $null -ne $SecurityObject.ExchangeInformation.BuildInformation.KBsInstalled -and
    $SecurityObject.ExchangeInformation.BuildInformation.KBsInstalled -like "*KB5000871*"
    $ex2019 = "Exchange2019"
    $ex2016 = "Exchange2016"
    $ex2013 = "Exchange2013"
    $currentExchangeBuild = $SecurityObject.BuildInformation

    if (($march2021SUInstalled) -and
        ($SecurityObject.ExchangeInformation.BuildInformation.VersionInformation.SupportedBuild -eq $false)) {

        if ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU21")) {
            $KBCveComb = @{KB4340731 = "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2013 -CU "CU22")) {
            $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU8")) {
            $KBCveComb = @{KB4073392 = "CVE-2018-0924", "CVE-2018-0940", "CVE-2018-0941"; KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU9")) {
            $KBCveComb = @{KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159"; KB4340731 = "CVE-2018-8374", "CVE-2018-8302" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU10")) {
            $KBCveComb = @{KB4340731 = "CVE-2018-8374", "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU11")) {
            $KBCveComb = @{KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588"; KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU12")) {
            $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU13")) {
            $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU14")) {
            $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU15")) {
            $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU16")) {
            $KBCveComb = @{KB4577352 = "CVE-2020-16875" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2016 -CU "CU17")) {
            $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU1")) {
            $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018"; KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU2")) {
            $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU3")) {
            $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU4")) {
            $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU5")) {
            $KBCveComb = @{KB4577352 = "CVE-2020-16875" }
        } elseif ((Test-ExchangeBuildEqualBuild -CurrentExchangeBuild $currentExchangeBuild -Version $ex2019 -CU "CU6")) {
            $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" }
        } else {
            Write-Verbose "No need to call 'Show-March2021SUOutdatedCUWarning'"
        }

        if ($null -ne $KBCveComb) {
            foreach ($kbName in $KBCveComb.Keys) {
                foreach ($cveName in $KBCveComb[$kbName]) {
                    $params = @{
                        AnalyzedInformation = $AnalyzeResults
                        DisplayGroupingKey  = $DisplayGroupingKey
                        Name                = "March 2021 Exchange Security Update for unsupported CU detected"
                        Details             = "`r`n`t`tPlease make sure $kbName is installed to be fully protected against: $cveName"
                        DisplayWriteType    = "Yellow"
                        DisplayTestingValue = $cveName
                        AddHtmlDetailRow    = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
            }
        }
    }
}

function Invoke-AnalyzerSecurityExtendedProtectionConfigState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $extendedProtection = $SecurityObject.ExchangeInformation.ExtendedProtectionConfig
    # Adding CVE-2024-21410 for the updated CVE for release with CU14
    $cveList = "CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134, CVE-2024-21410"

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Supported server roles are: Mailbox and ClientAccess
    if ($SecurityObject.IsEdgeServer -eq $false) {

        if ($null -ne $extendedProtection) {
            Write-Verbose "Exchange extended protection information found - performing vulnerability testing"

            # Description: Check for CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134, CVE-2024-21410 vulnerability
            # Affected Exchange versions: 2013, 2016, 2019
            # Fix: Install Aug 2022 SU & enable extended protection
            # Extended protection is available with IIS 7.5 or higher
            Write-Verbose "Testing CVE: $cveList"
            if (($extendedProtection.ExtendedProtectionConfiguration.ProperlySecuredConfiguration.Contains($false)) -or
                ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                Write-Verbose "At least one vDir is not configured properly and so, the system may be at risk"
                if (($extendedProtection.ExtendedProtectionConfiguration.SupportedExtendedProtection.Contains($false)) -and
                    ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                    # This combination means that EP is configured for at least one vDir, but the Exchange build doesn't support it.
                    # Such a combination can break several things like mailbox access, EMS... .
                    # Recommended action: Disable EP, upgrade to a supported build (Aug 2022 SU+) and enable afterwards.
                    $epDetails = "Extended Protection is configured, but not supported on this Exchange Server build"
                } elseif ((-not($extendedProtection.ExtendedProtectionConfiguration.SupportedExtendedProtection.Contains($false))) -and
                    ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                    # This combination means that EP is not configured and the Exchange build doesn't support it.
                    # Recommended action: Upgrade to a supported build (Aug 2022 SU+) and enable EP afterwards.
                    $epDetails = "Your Exchange server is at risk. Install the latest SU and enable Extended Protection"
                } elseif ($extendedProtection.ExtendedProtectionConfigured) {
                    # This means that EP is supported but not configured for at least one vDir.
                    # Recommended action: Enable EP for each vDir on the system by using the script provided by us.
                    $epDetails = "Extended Protection isn't configured as expected"
                } else {
                    # No Extended Protection is configured, provide a slightly different wording to avoid confusion of possible misconfigured EP.
                    $epDetails = "Extended Protection is not configured"
                }

                $epCveParams = $baseParams + @{
                    Name                = "Security Vulnerability"
                    Details             = $cveList
                    DisplayWriteType    = "Red"
                    TestingName         = "Extended Protection Vulnerable"
                    CustomName          = $cveList
                    DisplayTestingValue = $true
                }
                $epBasicParams = $baseParams + @{
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                    Details                = "$epDetails"
                    TestingName            = "Extended Protection Vulnerable Details"
                    DisplayTestingValue    = $epDetails
                }
                Add-AnalyzedResultInformation @epCveParams
                Add-AnalyzedResultInformation @epBasicParams

                $epFrontEndOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'
                $epBackEndOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'
                $mitigationOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'

                foreach ($entry in $extendedProtection.ExtendedProtectionConfiguration) {
                    $vDirArray = $entry.VirtualDirectoryName.Split("/", 2)
                    $ssl = $entry.Configuration.SslSettings

                    $listToAdd = $epFrontEndOutputObjectDisplayValue
                    if ($vDirArray[0] -eq "Exchange Back End") {
                        $listToAdd = $epBackEndOutputObjectDisplayValue
                    }

                    $listToAdd.Add(([PSCustomObject]@{
                                $vDirArray[0]     = $vDirArray[1]
                                Value             = $entry.ExtendedProtection
                                SupportedValue    = if ($entry.MitigationSupported -and $entry.MitigationEnabled) { "None" } else { $entry.ExpectedExtendedConfiguration }
                                ConfigSupported   = $entry.SupportedExtendedProtection
                                ConfigSecure      = $entry.ProperlySecuredConfiguration
                                RequireSSL        = "$($ssl.RequireSSL) $(if($ssl.Ssl128Bit) { "(128-bit)" })".Trim()
                                ClientCertificate = $ssl.ClientCertificate
                                IPFilterEnabled   = $entry.MitigationEnabled
                            })
                    )

                    if ($entry.MitigationEnabled) {
                        $mitigationOutputObjectDisplayValue.Add([PSCustomObject]@{
                                VirtualDirectory = $entry.VirtualDirectoryName
                                Details          = $entry.Configuration.MitigationSettings.Restrictions
                            })
                    }
                }

                $epConfig = {
                    param ($o, $p)
                    if ($p -eq "ConfigSupported") {
                        if ($o.$p -ne $true) {
                            "Red"
                        }
                    } elseif ($p -eq "IPFilterEnabled") {
                        if ($o.$p -eq $true) {
                            "Green"
                        }
                    } elseif ($p -eq "ConfigSecure") {
                        if ($o.$p -ne $true) {
                            "Red"
                        } else {
                            "Green"
                        }
                    }
                }

                $epFrontEndParams = $baseParams + @{
                    Name                = "Security Vulnerability"
                    OutColumns          = ([PSCustomObject]@{
                            DisplayObject      = $epFrontEndOutputObjectDisplayValue
                            ColorizerFunctions = @($epConfig)
                            IndentSpaces       = 8
                        })
                    DisplayTestingValue = $cveList
                }

                $epBackEndParams = $baseParams + @{
                    Name                = "Security Vulnerability"
                    OutColumns          = ([PSCustomObject]@{
                            DisplayObject      = $epBackEndOutputObjectDisplayValue
                            ColorizerFunctions = @($epConfig)
                            IndentSpaces       = 8
                        })
                    DisplayTestingValue = $cveList
                }

                Add-AnalyzedResultInformation @epFrontEndParams
                Add-AnalyzedResultInformation @epBackEndParams
                if ($mitigationOutputObjectDisplayValue.Count -ge 1) {
                    foreach ($mitigation in $mitigationOutputObjectDisplayValue) {
                        $epMitigationVDir = $baseParams + @{
                            Details          = "$($mitigation.Details.Count) IPs in filter list on vDir: '$($mitigation.VirtualDirectory)'"
                            DisplayWriteType = "Yellow"
                        }
                        Add-AnalyzedResultInformation @epMitigationVDir
                        $mitigationOutputObjectDisplayValue.Details.GetEnumerator() | ForEach-Object {
                            Write-Verbose "IP Address: $($_.key) is allowed to connect? $($_.value)"
                        }
                    }
                }

                $moreInformationParams = $baseParams + @{
                    DisplayWriteType = "Red"
                    Details          = "For more information about Extended Protection and how to configure, please read this article:`n`thttps://aka.ms/HC-ExchangeEPDoc"
                }
                Add-AnalyzedResultInformation @moreInformationParams
            } elseif ($SecurityObject.OsInformation.RegistryValues.SuppressExtendedProtection -ne 0) {
                # If this key is set, we need to flag it as the server being vulnerable.
                $params = $baseParams + @{
                    Name                = "Security Vulnerability"
                    Details             = $cveList
                    DisplayWriteType    = "Red"
                    TestingName         = "Extended Protection Vulnerable"
                    CustomName          = $cveList
                    DisplayTestingValue = $true
                }
                Add-AnalyzedResultInformation @params
            } else {
                Write-Verbose "System NOT vulnerable to $cveList"
            }
        } else {
            Write-Verbose "No Extended Protection configuration found - check will be skipped"
        }
    }
}


function Invoke-AnalyzerSecurityIISModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $SecurityObject.ExchangeInformation
    $moduleInformation = $exchangeInformation.IISSettings.IISModulesInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Description: Check for modules which are loaded by IIS and not signed by Microsoft or not signed at all
    if ($SecurityObject.IsEdgeServer -eq $false) {
        if ($null -ne $moduleInformation) {
            $iisModulesOutputList = New-Object 'System.Collections.Generic.List[object]'
            $modulesWriteType = "Grey"

            foreach ($m in $moduleInformation.ModuleList) {
                if ($m.Signed -eq $false) {
                    $modulesWriteType = "Red"

                    $iisModulesOutputList.Add([PSCustomObject]@{
                            Module = $m.Name
                            Path   = $m.Path
                            Signer = "N/A"
                            Status = "Not signed"
                        })
                } elseif (($m.SignatureDetails.IsMicrosoftSigned -eq $false) -or
                    ($m.SignatureDetails.SignatureStatus -ne 0) -and
                    ($m.SignatureDetails.SignatureStatus -ne -1)) {
                    if ($modulesWriteType -ne "Red") {
                        $modulesWriteType = "Yellow"
                    }

                    $iisModulesOutputList.Add([PSCustomObject]@{
                            Module = $m.Name
                            Path   = $m.Path
                            Signer = $m.SignatureDetails.Signer
                            Status = $m.SignatureDetails.SignatureStatus
                        })
                }
            }
            $params = $baseParams + @{
                Name             = "IIS module anomalies detected"
                Details          = ($iisModulesOutputList.Count -ge 1)
                DisplayWriteType = $modulesWriteType
            }
            Add-AnalyzedResultInformation @params

            if ($iisModulesOutputList.Count -ge 1) {
                if ($moduleInformation.AllModulesSigned -eq $false) {
                    $params = $baseParams + @{
                        Details                = "Modules that are loaded by IIS but NOT SIGNED - possibly a security risk"
                        DisplayCustomTabNumber = 2
                        DisplayWriteType       = "Red"
                    }
                    Add-AnalyzedResultInformation @params
                }

                if (($moduleInformation.AllSignedModulesSignedByMSFT -eq $false) -or
                    ($moduleInformation.AllSignaturesValid -eq $false)) {
                    $params = $baseParams + @{
                        Details                = "Modules that are loaded but NOT SIGNED BY Microsoft OR that have a problem with their signature"
                        DisplayCustomTabNumber = 2
                        DisplayWriteType       = "Yellow"
                    }
                    Add-AnalyzedResultInformation @params
                }

                $iisModulesConfig = {
                    param ($o, $p)
                    if ($p -eq "Signer") {
                        if ($o.$p -eq "N/A") {
                            "Red"
                        } else {
                            "Yellow"
                        }
                    } elseif ($p -eq "Status") {
                        if ($o.$p -eq "Not signed") {
                            "Red"
                        } elseif ($o.$p -ne 0) {
                            "Yellow"
                        }
                    }
                }

                $iisModulesParams = $baseParams + @{
                    Name       = "IIS Modules"
                    OutColumns = ([PSCustomObject]@{
                            DisplayObject      = $iisModulesOutputList
                            ColorizerFunctions = @($iisModulesConfig)
                            IndentSpaces       = 8
                        })
                }
                Add-AnalyzedResultInformation @iisModulesParams
            }
        } else {
            Write-Verbose "No modules were returned by previous call"
        }
    } else {
        Write-Verbose "IIS is not available on Edge Transport Server - check will be skipped"
    }
}
function Invoke-AnalyzerSecurityCveCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    function TestVulnerabilitiesByBuildNumbersForDisplay {
        param(
            [Parameter(Mandatory = $true)][string]$ExchangeBuildRevision,
            [Parameter(Mandatory = $true)][array]$SecurityFixedBuilds,
            [Parameter(Mandatory = $true)][array]$CVENames
        )
        [int]$fileBuildPart = ($split = $ExchangeBuildRevision.Split("."))[0]
        [int]$filePrivatePart = $split[1]
        $Script:breakpointHit = $false

        foreach ($securityFixedBuild in $SecurityFixedBuilds) {
            [int]$securityFixedBuildPart = ($split = $securityFixedBuild.Split("."))[0]
            [int]$securityFixedPrivatePart = $split[1]

            if ($fileBuildPart -eq $securityFixedBuildPart) {
                $Script:breakpointHit = $true
            }

            if (($fileBuildPart -lt $securityFixedBuildPart) -or
                    ($fileBuildPart -eq $securityFixedBuildPart -and
                $filePrivatePart -lt $securityFixedPrivatePart)) {
                foreach ($cveName in $CVENames) {
                    $params = @{
                        AnalyzedInformation = $AnalyzeResults
                        DisplayGroupingKey  = $DisplayGroupingKey
                        Name                = "Security Vulnerability"
                        Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information." -f $cveName)
                        DisplayWriteType    = "Red"
                        DisplayTestingValue = $cveName
                        AddHtmlDetailRow    = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
                break
            }

            if ($Script:breakpointHit) {
                break
            }
        }
    }

    function NewCveEntry {
        param(
            [string[]]$CVENames,
            [string[]]$ExchangeVersion
        )
        foreach ($cve in $CVENames) {
            [PSCustomObject]@{
                CVE     = $cve
                Version = $ExchangeVersion
            }
        }
    }

    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation

    [string]$buildRevision = ("{0}.{1}" -f $exchangeInformation.BuildInformation.ExchangeSetup.FileBuildPart, `
            $exchangeInformation.BuildInformation.ExchangeSetup.FilePrivatePart)
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    Write-Verbose "Exchange Build Revision: $buildRevision"
    Write-Verbose "Exchange CU: $exchangeCU"
    # This dictionary is a list of how to crawl through the list and add all the vulnerabilities to display
    # only place CVEs here that are fix by code fix only. If special checks are required, we need to check for that manually.
    $ex131619 = @("Exchange2013", "Exchange2016", "Exchange2019")
    $ex2013 = "Exchange2013"
    $ex2016 = "Exchange2016"
    $ex2019 = "Exchange2019"
    $suNameDictionary = @{
        "Mar18SU" = ((NewCveEntry @("CVE-2018-0924", "CVE-2018-0940") @($ex2013, $ex2016)) + (NewCveEntry "CVE-2018-0941" $ex2016))
        "May18SU" = ((NewCveEntry @("CVE-2018-8151", "CVE-2018-8154", "CVE-2018-8159") @($ex2013, $ex2016)) + (NewCveEntry @("CVE-2018-8152", "CVE-2018-8153") $ex2016))
        "Aug18SU" = (@((NewCveEntry "CVE-2018-8302" @($ex2013, $ex2016))) + (NewCveEntry "CVE-2018-8374" $ex2016))
        "Oct18SU" = (NewCveEntry @("CVE-2018-8265", "CVE-2018-8448") @($ex2013, $ex2016))
        "Dec18SU" = (@(NewCveEntry "CVE-2018-8604" $ex2016))
        "Jan19SU" = (NewCveEntry @("CVE-2019-0586", "CVE-2019-0588") @($ex2013, $ex2016))
        "Feb19SU" = (NewCveEntry @("CVE-2019-0686", "CVE-2019-0724") $ex131619)
        "Apr19SU" = (NewCveEntry @("CVE-2019-0817", "CVE-2019-0858") $ex131619)
        "Jun19SU" = (@(NewCveEntry @("ADV190018") $ex131619))
        "Jul19SU" = ((NewCveEntry @("CVE-2019-1084", "CVE-2019-1137") $ex131619) + (NewCveEntry "CVE-2019-1136" @($ex2013, $ex2016)))
        "Sep19SU" = (NewCveEntry @("CVE-2019-1233", "CVE-2019-1266") @($ex2016, $ex2019))
        "Nov19SU" = (@(NewCveEntry "CVE-2019-1373" $ex131619))
        "Feb20SU" = (NewCveEntry @("CVE-2020-0688", "CVE-2020-0692") $ex131619)
        "Mar20SU" = (@(NewCveEntry "CVE-2020-0903" @($ex2016, $ex2019)))
        "Sep20SU" = (@(NewCveEntry "CVE-2020-16875" @($ex2016, $ex2019)))
        "Oct20SU" = (@(NewCveEntry "CVE-2020-16969" $ex131619))
        "Nov20SU" = (NewCveEntry @("CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085") $ex131619)
        "Dec20SU" = ((NewCveEntry @("CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17142", "CVE-2020-17143") $ex131619) + (NewCveEntry "CVE-2020-17141" @($ex2016, $ex2019)))
        "Feb21SU" = (@(NewCveEntry "CVE-2021-24085" @($ex2016, $ex2019)))
        "Mar21SU" = (NewCveEntry @("CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065", "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854") $ex131619)
        "Apr21SU" = (NewCveEntry @("CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483") $ex131619)
        "May21SU" = (NewCveEntry @("CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209") $ex131619)
        "Jul21SU" = (NewCveEntry @("CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768") $ex131619)
        "Oct21SU" = (@((NewCveEntry "CVE-2021-26427" $ex131619)) + (NewCveEntry @("CVE-2021-41350", "CVE-2021-41348", "CVE-2021-34453") @($ex2016, $ex2019)))
        "Nov21SU" = ((NewCveEntry @("CVE-2021-42305", "CVE-2021-41349") $ex131619) + (NewCveEntry "CVE-2021-42321" @($ex2016, $ex2019)))
        "Jan22SU" = (NewCveEntry @("CVE-2022-21855", "CVE-2022-21846", "CVE-2022-21969") $ex131619)
        "Mar22SU" = (@((NewCveEntry "CVE-2022-23277" $ex131619)) + (NewCveEntry "CVE-2022-24463" @($ex2016, $ex2019)))
        "Aug22SU" = (@(NewCveEntry "CVE-2022-34692" @($ex2016, $ex2019)))
        "Nov22SU" = ((NewCveEntry @("CVE-2022-41040", "CVE-2022-41082", "CVE-2022-41079", "CVE-2022-41078", "CVE-2022-41080") $ex131619) + (NewCveEntry "CVE-2022-41123" @($ex2016, $ex2019)))
        "Jan23SU" = (@((NewCveEntry "CVE-2023-21762" $ex131619)) + (NewCveEntry @("CVE-2023-21745", "CVE-2023-21761", "CVE-2023-21763", "CVE-2023-21764") @($ex2016, $ex2019)))
        "Feb23SU" = (@(NewCveEntry @("CVE-2023-21529", "CVE-2023-21706", "CVE-2023-21707") $ex131619) + (NewCveEntry "CVE-2023-21710" @($ex2016, $ex2019)))
        "Mar23SU" = (@(NewCveEntry ("CVE-2023-21707") $ex131619))
        "Jun23SU" = (NewCveEntry @("CVE-2023-28310", "CVE-2023-32031") @($ex2016, $ex2019))
        "Aug23SU" = (NewCveEntry @("CVE-2023-38181", "CVE-2023-38182", "CVE-2023-38185", "CVE-2023-35368", "CVE-2023-35388", "CVE-2023-36777", "CVE-2023-36757", "CVE-2023-36756", "CVE-2023-36745", "CVE-2023-36744") @($ex2016, $ex2019))
        "Oct23SU" = (NewCveEntry @("CVE-2023-36778") @($ex2016, $ex2019))
        "Nov23SU" = (NewCveEntry @("CVE-2023-36050", "CVE-2023-36039", "CVE-2023-36035", "CVE-2023-36439") @($ex2016, $ex2019))
        "Mar24SU" = (NewCveEntry @("CVE-2024-26198") @($ex2016, $ex2019))
    }

    # Need to organize the list so oldest CVEs come out first.
    $monthOrder = @{
        "Jan" = 1
        "Feb" = 2
        "Mar" = 3
        "Apr" = 4
        "May" = 5
        "Jun" = 6
        "Jul" = 7
        "Aug" = 8
        "Sep" = 9
        "Oct" = 10
        "Nov" = 11
        "Dec" = 12
    }
    $unsortedKeys = @($suNameDictionary.Keys)
    $sortedKeys = New-Object System.Collections.Generic.List[string]

    foreach ($value in $unsortedKeys) {
        $month = $value.Substring(0, 3)
        $year = [int]$value.Substring(3, 2)
        $insertAt = 0
        while ($insertAt -lt $sortedKeys.Count) {

            $compareMonth = $sortedKeys[$insertAt].Substring(0, 3)
            $compareYear = [int]$sortedKeys[$insertAt].Substring(3, 2)
            # break to add at current spot in list
            if ($compareYear -gt $year) { break }
            elseif ( $compareYear -eq $year -and
                $monthOrder[$month] -lt $monthOrder[$compareMonth]) { break }

            $insertAt++
        }

        $sortedKeys.Insert($insertAt, $value)
    }

    foreach ($key in $sortedKeys) {
        if (-not (Test-ExchangeBuildGreaterOrEqualThanSecurityPatch -CurrentExchangeBuild $exchangeInformation.BuildInformation.VersionInformation -SUName $key)) {
            Write-Verbose "Tested that we aren't on SU $key or greater"
            $cveNames = ($suNameDictionary[$key] | Where-Object { $_.Version.Contains($exchangeInformation.BuildInformation.MajorVersion) }).CVE
            foreach ($cveName in $cveNames) {
                $params = @{
                    AnalyzedInformation = $AnalyzeResults
                    DisplayGroupingKey  = $DisplayGroupingKey
                    Name                = "Security Vulnerability"
                    Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information." -f $cveName)
                    DisplayWriteType    = "Red"
                    DisplayTestingValue = $cveName
                    AddHtmlDetailRow    = $false
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    $securityObject = [PSCustomObject]@{
        BuildInformation    = $exchangeInformation.BuildInformation.VersionInformation
        MajorVersion        = $exchangeInformation.BuildInformation.MajorVersion
        IsEdgeServer        = $exchangeInformation.GetExchangeServer.IsEdgeServer
        ExchangeInformation = $exchangeInformation
        OsInformation       = $osInformation
        OrgInformation      = $HealthServerObject.OrganizationInformation
    }

    Invoke-AnalyzerSecurityIISModules -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2020-0796 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2020-1147 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2021-1730 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2021-34470 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2022-21978 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2023-36434 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCveAddressedBySerializedDataSigning -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-MarchSuSpecial -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityADV24199947 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    # Make sure that these stay as the last one to keep the output more readable
    Invoke-AnalyzerSecurityExtendedProtectionConfigState -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
}
function Invoke-AnalyzerSecurityVulnerability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $keySecurityVulnerability = Get-DisplayResultsGroupingKey -Name "Security Vulnerability"  -DisplayOrder $Order
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keySecurityVulnerability
    }

    Invoke-AnalyzerSecurityCveCheck -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecurityVulnerability

    $allSecurityVulnerabilities = $AnalyzeResults.Value.DisplayResults[$keySecurityVulnerability]
    $securityVulnerabilities = $allSecurityVulnerabilities | Where-Object { $_.Name -ne "IIS module anomalies detected" }
    $iisModule = $allSecurityVulnerabilities | Where-Object { $_.Name -eq "IIS module anomalies detected" }
    $buildVersion = $HealthServerObject.ExchangeInformation.BuildInformation.VersionInformation
    $noLongerSecureExchange = ($buildVersion.ExtendedSupportDate -le ([DateTime]::Now)) -and $buildVersion.LatestSU -eq $false

    if ((($null -eq $securityVulnerabilities -and
                $HealthServerObject.ExchangeInformation.GetExchangeServer.IsEdgeServer) -or
        ($null -eq $securityVulnerabilities -and
        ($null -ne $iisModule -or $iisModule.DisplayValue -eq $false))) -and
        (-not $noLongerSecureExchange)) {
        $params = $baseParams + @{
            Details          = "All known security issues in this version of the script passed."
            DisplayWriteType = "Green"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                      = "Vulnerability Detected"
            Details                   = "None"
            AddDisplayResultsLineInfo = $false
            AddHtmlOverviewValues     = $true
        }
        Add-AnalyzedResultInformation @params
    } elseif ($null -ne $securityVulnerabilities -or
        ($null -ne $iisModule -and $iisModule.DisplayValue -eq $true)) {

        $details = $securityVulnerabilities.DisplayValue |
            ForEach-Object {
                return $_ + "<br>"
            }

        # If details are null, but iisModule is showing a vulnerability,
        # then just provide see IIS Module section
        if ($null -eq $details) { $details = "See IIS module anomalies detected section above" }

        $params = $baseParams + @{
            Name                      = "Security Vulnerabilities"
            Details                   = $details
            DisplayWriteType          = "Red"
            AddDisplayResultsLineInfo = $false
        }
        Add-AnalyzedResultInformation @params

        # Only add this if on current supported and releasing SU for Exchange.
        if (-not ($noLongerSecureExchange)) {

            $params = $baseParams + @{
                Name                      = "Vulnerability Detected"
                Details                   = $true
                AddDisplayResultsLineInfo = $false
                DisplayWriteType          = "Red"
                AddHtmlOverviewValues     = $true
                AddHtmlDetailRow          = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($noLongerSecureExchange) {
        $friendlyName = ($buildVersion.MajorVersion.ToString()).Insert(($buildVersion.MajorVersion.IndexOf("2")), " ").Trim()
        $details = "`r`n`t$friendlyName is out of support and will not receive any further security updates." +
        "`r`n`tWe do not perform any vulnerability testing against this version of Exchange any more." +
        "`r`n`t$friendlyName is likely vulnerable to any vulnerabilities disclosed after $($buildVersion.ExtendedSupportDate.ToString("yyyy/MM/dd"))" +
        "`r`n`tYou should migrate to the latest Exchange Server for On Prem or Exchange Online as soon as possible and decommission $friendlyName from your environment."
        $params = $baseParams + @{
            Details          = $details
            DisplayWriteType = "Red"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                      = "Vulnerability Detected"
            Details                   = "Expired Version of Exchange"
            DisplayWriteType          = "Red"
            AddDisplayResultsLineInfo = $false
            AddHtmlOverviewValues     = $true
        }
        Add-AnalyzedResultInformation @params
    }
}
function Invoke-AnalyzerEngine {
    [CmdletBinding()]
    param(
        [object]$HealthServerObject
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $analyzedResults = [PSCustomObject]@{
        HealthCheckerExchangeServer = $HealthServerObject
        HtmlServerValues            = @{}
        DisplayResults              = @{}
    }

    #Display Grouping Keys
    $order = 1
    $baseParams = @{
        AnalyzedInformation = $analyzedResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "BeginningInfo" -DisplayGroupName $false -DisplayOrder 0 -DefaultTabNumber 0)
    }

    if (!$Script:DisplayedScriptVersionAlready) {
        $params = $baseParams + @{
            Name             = "Exchange Health Checker Version"
            Details          = $BuildVersion
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    $VirtualizationWarning = @"
Virtual Machine detected.  Certain settings about the host hardware cannot be detected from the virtual machine.  Verify on the VM Host that:

    - There is no more than a 1:1 Physical Core to Virtual CPU ratio (no oversubscribing)
    - If Hyper-Threading is enabled do NOT count Hyper-Threaded cores as physical cores
    - Do not oversubscribe memory or use dynamic memory allocation

Although Exchange technically supports up to a 2:1 physical core to vCPU ratio, a 1:1 ratio is strongly recommended for performance reasons.  Certain third party Hyper-Visors such as VMWare have their own guidance.

VMWare recommends a 1:1 ratio.  Their guidance can be found at https://aka.ms/HC-VMwareBP2019.
Related specifically to VMWare, if you notice you are experiencing packet loss on your VMXNET3 adapter, you may want to review the following article from VMWare:  https://aka.ms/HC-VMwareLostPackets.

For further details, please review the virtualization recommendations on Microsoft Docs here: https://aka.ms/HC-Virtualization.

"@

    if ($HealthServerObject.HardwareInformation.ServerType -eq "VMWare" -or
        $HealthServerObject.HardwareInformation.ServerType -eq "HyperV") {
        $params = $baseParams + @{
            Details          = $VirtualizationWarning
            DisplayWriteType = "Yellow"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    # Can't do a Hash Table pass param due to [ref]
    Invoke-AnalyzerExchangeInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerOrganizationInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerHybridInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerOsInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerHardwareInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerNicSettings -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerFrequentConfigurationIssues -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerSecuritySettings -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerSecurityVulnerability -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerIISInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Write-Debug("End of Analyzer Engine")
    return $analyzedResults
}

function Get-ErrorsThatOccurred {

    function WriteErrorInformation {
        [CmdletBinding()]
        param(
            [object]$CurrentError
        )
        Write-VerboseErrorInformation $CurrentError
        Write-Verbose "-----------------------------------`r`n`r`n"
    }

    if ($Error.Count -gt 0) {
        Write-Host ""
        Write-Host ""
        function Write-Errors {
            Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

            Get-UnhandledErrors | ForEach-Object {
                Write-Verbose "Error Index: $($_.Index)"
                WriteErrorInformation $_.ErrorInformation
            }

            Write-Verbose "`r`n`r`nErrors that were handled"
            Get-HandledErrors | ForEach-Object {
                Write-Verbose "Error Index: $($_.Index)"
                WriteErrorInformation $_.ErrorInformation
            }
        }

        if ((Test-UnhandledErrorsOccurred)) {
            Write-Red("There appears to have been some errors in the script. To assist with debugging of the script, please send the HealthChecker-Debug_*.txt, HealthChecker-Errors.json, and .xml file to ExToolsFeedback@microsoft.com.")
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
            #Need to convert Error to Json because running into odd issues with trying to export $Error out in my lab. Got StackOverflowException for one of the errors i always see there.
            try {
                $Error |
                    ConvertTo-Json |
                    Out-File ("$Script:OutputFilePath\HealthChecker-Errors.json")
            } catch {
                Write-Red("Failed to export the HealthChecker-Errors.json")
                Invoke-CatchActions
            }
        } elseif ($Script:VerboseEnabled -or
            $SaveDebugLog) {
            Write-Verbose "All errors that occurred were in try catch blocks and was handled correctly."
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

function Get-ExportedHealthCheckerFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $importedItems = New-Object System.Collections.Generic.List[object]
        $customFileObject = New-Object System.Collections.Generic.List[object]
    }
    process {
        $allItems = @(Get-ChildItem $Directory |
                Where-Object { $_.Name -like "HealthChecker-*-*.xml" -and $_.Name -notlike "HealthChecker-ExchangeDCCoreRatio-*.xml" })

        if ($null -eq $allItems) {
            Write-Verbose "No items were found like HealthChecker-*-*.xml"
            return
        }

        $allItems |
            ForEach-Object {
                [string]$name = $_.Name
                $startIndex = $name.IndexOf("-")
                $serverName = $name.Substring(($startIndex + 1), ($name.LastIndexOf("-") - $startIndex - 1))
                $customFileObject.Add([PSCustomObject]@{
                        ServerName = $serverName
                        FileName   = $name
                        FileObject = $_
                    })
            }

        # Group the items by server name and then get the latest one and import that file.
        $groupResults = $customFileObject | Group-Object ServerName

        $groupResults |
            ForEach-Object {
                $sortedGroup = $_.Group | Sort-Object FileName -Descending
                $index = 0
                $continueLoop = $true

                do {
                    $fileName = $sortedGroup[$index].FileObject.VersionInfo.FileName
                    $data = Import-Clixml -Path $fileName

                    if ($null -ne $data -and
                        $null -ne $data.HealthCheckerExchangeServer) {
                        Write-Verbose "For Server $($_.Group[0].ServerName) using file: $fileName"
                        $importedItems.Add($data)
                        $continueLoop = $false
                    } else {
                        $index++
                        if ($index -ge $_.Count) {
                            $continueLoop = $false
                            Write-Red "Failed to find proper Health Checker data to import for server $($_.Group[0].ServerName)"
                        }
                    }
                } while ($continueLoop)
            }
    }
    end {
        if ($importedItems.Count -eq 0) {
            return $null
        }
        return $importedItems
    }
}



# Confirm that either Remote Shell or EMS is loaded from an Edge Server, Exchange Server, or a Tools box.
# It does this by also initializing the session and running Get-EventLogLevel. (Server Management RBAC right)
# All script that require Confirm-ExchangeShell should be at least using Server Management RBAC right for the user running the script.
function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell"
        $currentErrors = $Error.Count
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        $remoteShell = (-not(Test-Path $setupKey))
        $toolsServer = (Test-Path $setupKey) -and
            (-not(Test-Path $edgeTransportKey)) -and
            ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        function IsExchangeManagementSession {
            [OutputType("System.Boolean")]
            param(
                [ScriptBlock]$CatchActionFunction
            )

            $getEventLogLevelCallSuccessful = $false
            $isExchangeManagementShell = $false

            try {
                $currentErrors = $Error.Count
                $attempts = 0
                do {
                    $eventLogLevel = Get-EventLogLevel -ErrorAction Stop | Select-Object -First 1
                    $attempts++
                    if ($attempts -ge 5) {
                        throw "Failed to run Get-EventLogLevel too many times."
                    }
                } while ($null -eq $eventLogLevel)
                $getEventLogLevelCallSuccessful = $true
                foreach ($e in $eventLogLevel) {
                    Write-Verbose "Type is: $($e.GetType().Name) BaseType is: $($e.GetType().BaseType)"
                    if (($e.GetType().Name -eq "EventCategoryObject") -or
                        (($e.GetType().Name -eq "PSObject") -and
                            ($null -ne $e.SerializationData))) {
                        $isExchangeManagementShell = $true
                    }
                }
                Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
            } catch {
                Write-Verbose "Failed to run Get-EventLogLevel"
                Invoke-CatchActionError $CatchActionFunction
            }

            return [PSCustomObject]@{
                CallWasSuccessful = $getEventLogLevelCallSuccessful
                IsManagementShell = $isExchangeManagementShell
            }
        }
    }
    process {
        $isEMS = IsExchangeManagementSession $CatchActionFunction
        if ($isEMS.CallWasSuccessful) {
            Write-Verbose "Exchange PowerShell Module already loaded."
        } else {
            if (-not ($LoadExchangeShell)) { return }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    $currentErrors = $Error.Count
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exShell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }
                    Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

                    Write-Verbose "Imported Module. Trying Get-EventLogLevel Again"
                    $isEMS = IsExchangeManagementSession $CatchActionFunction
                    if (($isEMS.CallWasSuccessful) -and
                        ($isEMS.IsManagementShell)) {
                        Write-Verbose "Successfully loaded Exchange Management Shell"
                    } else {
                        Write-Warning "Something went wrong while loading the Exchange Management Shell"
                    }
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $returnObject = [PSCustomObject]@{
            ShellLoaded = $isEMS.CallWasSuccessful
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $isEMS.CallWasSuccessful -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $isEMS.CallWasSuccessful -and $toolsServer
            RemoteShell = $isEMS.CallWasSuccessful -and $remoteShell
            EMS         = $isEMS.IsManagementShell
        }

        return $returnObject
    }
}
function Invoke-ConfirmExchangeShell {

    $Script:ExchangeShellComputer = Confirm-ExchangeShell -CatchActionFunction ${Function:Invoke-CatchActions}

    if (-not ($Script:ExchangeShellComputer.ShellLoaded)) {
        Write-Warning "Failed to load Exchange Shell... stopping script"
        $Script:Logger.PreventLogCleanup = $true
        exit
    }

    if ($Script:ExchangeShellComputer.EdgeServer -and
        ($Script:ServerNameList.Count -gt 1 -or
        (-not ($Script:ServerNameList.ToLower().Contains($env:COMPUTERNAME.ToLower()))))) {
        Write-Warning "Can't run Exchange Health Checker from an Edge Server against anything but the local Edge Server."
        $Script:Logger.PreventLogCleanup = $true
        exit
    }

    if ($Script:ExchangeShellComputer.ToolsOnly -and
        $Script:ServerNameList.ToLower().Contains($env:COMPUTERNAME.ToLower()) -and
        -not ($LoadBalancingReport)) {
        Write-Warning "Can't run Exchange Health Checker Against a Tools Server. Use the -Server Parameter and provide the server you want to run the script against."
        $Script:Logger.PreventLogCleanup = $true
        exit
    }

    Write-Verbose("Script Executing on Server $env:COMPUTERNAME")
    Write-Verbose("ToolsOnly: $($Script:ExchangeShellComputer.ToolsOnly) | RemoteShell $($Script:ExchangeShellComputer.RemoteShell)")
}

function Invoke-SetOutputInstanceLocation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $false)]
        [string]$Server,

        [Parameter(Mandatory = $false)]
        [bool]$IncludeServerName = $false
    )
    $endName = "-{0}.txt" -f $Script:dateTimeStringFormat

    if ($IncludeServerName) {
        $endName = "-{0}{1}" -f $Server, $endName
    }

    $Script:OutputFullPath = "{0}\{1}{2}" -f $Script:OutputFilePath, $FileName, $endName
    $Script:OutXmlFullPath = $Script:OutputFullPath.Replace(".txt", ".xml")
}

function Write-ResultsToScreen {
    param(
        [Hashtable]$ResultsToWrite
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $indexOrderGroupingToKey = @{}

    foreach ($keyGrouping in $ResultsToWrite.Keys) {
        $indexOrderGroupingToKey[$keyGrouping.DisplayOrder] = $keyGrouping
    }

    $sortedIndexOrderGroupingToKey = $indexOrderGroupingToKey.Keys | Sort-Object

    foreach ($key in $sortedIndexOrderGroupingToKey) {
        Write-Verbose "Working on Key: $key"
        $keyGrouping = $indexOrderGroupingToKey[$key]
        Write-Verbose "Working on Key Group: $($keyGrouping.Name)"
        Write-Verbose "Total lines to write: $($ResultsToWrite[$keyGrouping].Count)"

        try {
            if ($keyGrouping.DisplayGroupName) {
                Write-Grey($keyGrouping.Name)
                $dashes = [string]::empty
                1..($keyGrouping.Name.Length) | ForEach-Object { $dashes = $dashes + "-" }
                Write-Grey($dashes)
            }

            foreach ($line in $ResultsToWrite[$keyGrouping]) {
                try {
                    $tab = [string]::Empty

                    if ($line.TabNumber -ne 0) {
                        1..($line.TabNumber) | ForEach-Object { $tab = $tab + "`t" }
                    }

                    if ([string]::IsNullOrEmpty($line.Name)) {
                        $displayLine = $line.DisplayValue
                    } else {
                        $displayLine = [string]::Concat($line.Name, ": ", $line.DisplayValue)
                    }

                    $writeValue = "{0}{1}" -f $tab, $displayLine
                    switch ($line.WriteType) {
                        "Grey" { Write-Grey($writeValue) }
                        "Yellow" { Write-Yellow($writeValue) }
                        "Green" { Write-Green($writeValue) }
                        "Red" { Write-Red($writeValue) }
                        "OutColumns" { Write-OutColumns($line.OutColumns) }
                    }
                } catch {
                    # We do not want to call Invoke-CatchActions here because we want the issues reported.
                    Write-Verbose "Failed inside the section loop writing. Writing out a blank line and continuing. Inner Exception: $_"
                    Write-Grey ""
                }
            }

            Write-Grey ""
        } catch {
            # We do not want to call Invoke-CatchActions here because we want the issues reported.
            Write-Verbose "Failed in $($MyInvocation.MyCommand) outside section writing loop. Inner Exception: $_"
            Write-Grey ""
        }
    }
}


<#
.SYNOPSIS
    Outputs a table of objects with certain values colorized.
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
function Out-Columns {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $InputObject,

        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]
        $Properties,

        [Parameter(Mandatory = $false, Position = 1)]
        [ScriptBlock[]]
        $ColorizerFunctions = @(),

        [Parameter(Mandatory = $false)]
        [int]
        $IndentSpaces = 0,

        [Parameter(Mandatory = $false)]
        [int]
        $LinesBetweenObjects = 0,

        [Parameter(Mandatory = $false)]
        [ref]
        $StringOutput
    )

    begin {
        function WrapLine {
            param([string]$line, [int]$width)
            if ($line.Length -le $width -and $line.IndexOf("`n") -lt 0) {
                return $line
            }

            $lines = New-Object System.Collections.ArrayList

            $noLF = $line.Replace("`r", "")
            $lineSplit = $noLF.Split("`n")
            foreach ($l in $lineSplit) {
                if ($l.Length -le $width) {
                    [void]$lines.Add($l)
                } else {
                    $split = $l.Split(" ")
                    $sb = New-Object System.Text.StringBuilder
                    for ($i = 0; $i -lt $split.Length; $i++) {
                        if ($sb.Length -eq 0 -and $sb.Length + $split[$i].Length -lt $width) {
                            [void]$sb.Append($split[$i])
                        } elseif ($sb.Length -gt 0 -and $sb.Length + $split[$i].Length + 1 -lt $width) {
                            [void]$sb.Append(" " + $split[$i])
                        } elseif ($sb.Length -gt 0) {
                            [void]$lines.Add($sb.ToString())
                            [void]$sb.Clear()
                            $i--
                        } else {
                            if ($split[$i].Length -le $width) {
                                [void]$lines.Add($split[$i])
                            } else {
                                [void]$lines.Add($split[$i].Substring(0, $width))
                                $split[$i] = $split[$i].Substring($width)
                                $i--
                            }
                        }
                    }

                    if ($sb.Length -gt 0) {
                        [void]$lines.Add($sb.ToString())
                    }
                }
            }

            return $lines
        }

        function GetLineObjects {
            param($obj, $props, $colWidths)
            $linesNeededForThisObject = 1
            $multiLineProps = @{}
            for ($i = 0; $i -lt $props.Length; $i++) {
                $p = $props[$i]
                $val = $obj."$p"

                if ($val -isnot [array]) {
                    $val = WrapLine -line $val -width $colWidths[$i]
                } elseif ($val -is [array]) {
                    $val = $val | Where-Object { $null -ne $_ }
                    $val = $val | ForEach-Object { WrapLine -line $_ -width $colWidths[$i] }
                }

                if ($val -is [array]) {
                    $multiLineProps[$p] = $val
                    if ($val.Length -gt $linesNeededForThisObject) {
                        $linesNeededForThisObject = $val.Length
                    }
                }
            }

            if ($linesNeededForThisObject -eq 1) {
                $obj
            } else {
                for ($i = 0; $i -lt $linesNeededForThisObject; $i++) {
                    $lineProps = @{}
                    foreach ($p in $props) {
                        if ($null -ne $multiLineProps[$p] -and $multiLineProps[$p].Length -gt $i) {
                            $lineProps[$p] = $multiLineProps[$p][$i]
                        } elseif ($i -eq 0) {
                            $lineProps[$p] = $obj."$p"
                        } else {
                            $lineProps[$p] = $null
                        }
                    }

                    [PSCustomObject]$lineProps
                }
            }
        }

        function GetColumnColors {
            param($obj, $props, $functions)

            $consoleHost = (Get-Host).Name -eq "ConsoleHost"
            $colColors = New-Object string[] $props.Count
            for ($i = 0; $i -lt $props.Count; $i++) {
                if ($consoleHost) {
                    $fgColor = (Get-Host).ui.RawUi.ForegroundColor
                } else {
                    $fgColor = "White"
                }
                foreach ($func in $functions) {
                    $result = $func.Invoke($obj, $props[$i])
                    if (-not [string]::IsNullOrEmpty($result)) {
                        $fgColor = $result
                        break # The first colorizer that takes action wins
                    }
                }

                $colColors[$i] = $fgColor
            }

            $colColors
        }

        function GetColumnWidths {
            param($objects, $props)

            $colWidths = New-Object int[] $props.Count

            # Start with the widths of the property names
            for ($i = 0; $i -lt $props.Count; $i++) {
                $colWidths[$i] = $props[$i].Length
            }

            # Now check the widths of the widest values
            foreach ($thing in $objects) {
                for ($i = 0; $i -lt $props.Count; $i++) {
                    $val = $thing."$($props[$i])"
                    if ($null -ne $val) {
                        $width = 0
                        if ($val -isnot [array]) {
                            $val = $val.ToString().Split("`n")
                        }

                        $width = ($val | ForEach-Object {
                                if ($null -ne $_) { $_.ToString() } else { "" }
                            } | Sort-Object Length -Descending | Select-Object -First 1).Length

                        if ($width -gt $colWidths[$i]) {
                            $colWidths[$i] = $width
                        }
                    }
                }
            }

            # If we're within the window width, we're done
            $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
            $windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
            if ($windowWidth -lt 1 -or $totalColumnWidth -lt $windowWidth) {
                return $colWidths
            }

            # Take size away from one or more columns to make them fit
            while ($totalColumnWidth -ge $windowWidth) {
                $startingTotalWidth = $totalColumnWidth
                $widest = $colWidths | Sort-Object -Descending | Select-Object -First 1
                $newWidest = [Math]::Floor($widest * 0.95)
                for ($i = 0; $i -lt $colWidths.Length; $i++) {
                    if ($colWidths[$i] -eq $widest) {
                        $colWidths[$i] = $newWidest
                        break
                    }
                }

                $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
                if ($totalColumnWidth -ge $startingTotalWidth) {
                    # Somehow we didn't reduce the size at all, so give up
                    break
                }
            }

            return $colWidths
        }

        $objects = New-Object System.Collections.ArrayList
        $padding = 2
        $stb = New-Object System.Text.StringBuilder
    }

    process {
        foreach ($thing in $InputObject) {
            [void]$objects.Add($thing)
        }
    }

    end {
        if ($objects.Count -gt 0) {
            $props = $null

            if ($null -ne $Properties) {
                $props = $Properties
            } else {
                $props = $objects[0].PSObject.Properties.Name
            }

            $colWidths = GetColumnWidths $objects $props

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i]) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i])
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length)) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length))
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            foreach ($o in $objects) {
                $colColors = GetColumnColors -obj $o -props $props -functions $ColorizerFunctions
                $lineObjects = @(GetLineObjects -obj $o -props $props -colWidths $colWidths)
                foreach ($lineObj in $lineObjects) {
                    Write-Host (" " * $IndentSpaces) -NoNewline
                    [void]$stb.Append(" " * $IndentSpaces)
                    for ($i = 0; $i -lt $props.Count; $i++) {
                        $val = $lineObj."$($props[$i])"
                        if ($val.Count -eq 0) { $val = "" }
                        Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $val) -NoNewline -ForegroundColor $colColors[$i]
                        [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $val)
                    }

                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }

                for ($i = 0; $i -lt $LinesBetweenObjects; $i++) {
                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            if ($null -ne $StringOutput) {
                $StringOutput.Value = $stb.ToString()
            }
        }
    }
}
function Write-Red($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Red
    Write-HostLog $message
}

function Write-Yellow($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Yellow
    Write-HostLog $message
}

function Write-Green($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Green
    Write-HostLog $message
}

function Write-Grey($message) {
    Write-DebugLog $message
    Write-Host $message
    Write-HostLog $message
}

function Write-DebugLog($message) {
    if (![string]::IsNullOrEmpty($message)) {
        $Script:Logger = $Script:Logger | Write-LoggerInstance $message
    }
}

function Write-HostLog ($message) {
    if ($Script:OutputFullPath) {
        $message | Out-File ($Script:OutputFullPath) -Append
    }
}

function Write-OutColumns($OutColumns) {
    if ($null -ne $OutColumns) {
        try {
            $stringOutput = $null
            $params = @{
                Properties         = $OutColumns.SelectProperties
                ColorizerFunctions = $OutColumns.ColorizerFunctions
                IndentSpaces       = $OutColumns.IndentSpaces
                StringOutput       = ([ref]$stringOutput)
            }
            $OutColumns.DisplayObject | Out-Columns @params
            $stringOutput | Out-File ($Script:OutputFullPath) -Append
            Write-DebugLog $stringOutput
        } catch {
            # We do not want to call Invoke-CatchActions here because we want the issues reported.
            Write-Verbose "Failed to export Out-Columns. Inner Exception: $_"
            $s = $OutColumns.DisplayObject | Out-String
            Write-DebugLog $s
        }
    }
}

function Write-Break {
    Write-Host ""
}

function Get-HtmlServerReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AnalyzedHtmlServerValues,

        [string]$HtmlOutFilePath
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function GetOutColumnHtmlTable {
        param(
            [object]$OutColumn
        )
        # this keeps the order of the columns
        $headerValues = $OutColumn[0].PSObject.Properties.Name
        $htmlTableValue = "<table>"

        foreach ($header in $headerValues) {
            $htmlTableValue += "<th>$header</th>"
        }

        foreach ($dataRow in $OutColumn) {
            $htmlTableValue += "$([System.Environment]::NewLine)<tr>"

            foreach ($header in $headerValues) {
                $htmlTableValue += "<td class=`"$($dataRow.$header.DisplayColor)`">$($dataRow.$header.Value)</td>"
            }
            $htmlTableValue += "$([System.Environment]::NewLine)</tr>"
        }
        $htmlTableValue += "</table>"
        return $htmlTableValue
    }

    $htmlHeader = "<html>
        <style>
        BODY{font-family: Arial; font-size: 8pt;}
        H1{font-size: 16px;}
        H2{font-size: 14px;}
        H3{font-size: 12px;}
        TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
        TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
        TD{border: 1px solid black; padding: 5px; }
        td.Green{background: #7FFF00;}
        td.Yellow{background: #FFE600;}
        td.Red{background: #FF0000; color: #ffffff;}
        td.Info{background: #85D4FF;}
        </style>
        <body>
        <h1 align=""center"">Exchange Health Checker v$($BuildVersion)</h1><br>
        <h2>Servers Overview</h2>"

    [array]$htmlOverviewTable += "<p>
        <table>
        <tr>$([System.Environment]::NewLine)"

    foreach ($tableHeaderName in $AnalyzedHtmlServerValues[0]["OverviewValues"].Name) {
        $htmlOverviewTable += "<th>{0}</th>$([System.Environment]::NewLine)" -f $tableHeaderName
    }

    $htmlOverviewTable += "</tr>$([System.Environment]::NewLine)"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        $htmlTableRow = @()
        [array]$htmlTableRow += "<tr>$([System.Environment]::NewLine)"
        foreach ($htmlTableDataRow in $serverHtmlServerValues["OverviewValues"]) {
            $htmlTableRow += "<td class=`"{0}`">{1}</td>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                $htmlTableDataRow.DetailValue
        }

        $htmlTableRow += "</tr>$([System.Environment]::NewLine)"
        $htmlOverviewTable += $htmlTableRow
    }

    $htmlOverviewTable += "</table>$([System.Environment]::NewLine)</p>$([System.Environment]::NewLine)"

    [array]$htmlServerDetails += "<p>$([System.Environment]::NewLine)<h2>Server Details</h2>$([System.Environment]::NewLine)<table>"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        foreach ($htmlTableDataRow in $serverHtmlServerValues["ServerDetails"]) {
            if ($htmlTableDataRow.Name -eq "Server Name") {
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<th>{0}</th>$([System.Environment]::NewLine)<th>{1}</th>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            } elseif ($null -ne $htmlTableDataRow.TableValue) {
                $htmlTable = GetOutColumnHtmlTable $htmlTableDataRow.TableValue
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<td class=`"{0}`">{1}</td><td class=`"{0}`">{2}</td>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                    $htmlTableDataRow.Name, `
                    $htmlTable
            } else {
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<td class=`"{0}`">{1}</td><td class=`"{0}`">{2}</td>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                    $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            }
        }
    }
    $htmlServerDetails += "$([System.Environment]::NewLine)</table>$([System.Environment]::NewLine)</p>$([System.Environment]::NewLine)"

    $htmlReport = $htmlHeader + $htmlOverviewTable + $htmlServerDetails + "</body>$([System.Environment]::NewLine)</html>"

    $htmlReport | Out-File $HtmlOutFilePath -Encoding UTF8

    Write-Host "HTML Report Location: $HtmlOutFilePath"
}



# Use this after the counters have been localized.
function Get-CounterSamples {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$MachineName,

        [Parameter(Mandatory = $true)]
        [string[]]$Counter,

        [string]$CustomErrorAction = "Stop"
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    try {
        return (Get-Counter -ComputerName $MachineName -Counter $Counter -ErrorAction $CustomErrorAction).CounterSamples
    } catch {
        Write-Verbose "Failed ot get counter samples"
        Invoke-CatchActions
    }
}

# Use this to localize the counters provided
function Get-LocalizedCounterSamples {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$MachineName,

        [Parameter(Mandatory = $true)]
        [string[]]$Counter,

        [string]$CustomErrorAction = "Stop"
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $localizedCounters = @()

    foreach ($computer in $MachineName) {

        foreach ($currentCounter in $Counter) {
            $counterObject = Get-CounterFullNameToCounterObject -FullCounterName $currentCounter
            $localizedCounterName = Get-LocalizedPerformanceCounterName -ComputerName $computer -PerformanceCounterName $counterObject.CounterName
            $localizedObjectName = Get-LocalizedPerformanceCounterName -ComputerName $computer -PerformanceCounterName $counterObject.ObjectName
            $localizedFullCounterName = ($counterObject.FullName.Replace($counterObject.CounterName, $localizedCounterName)).Replace($counterObject.ObjectName, $localizedObjectName)

            if (-not ($localizedCounters.Contains($localizedFullCounterName))) {
                $localizedCounters += $localizedFullCounterName
            }
        }
    }

    return (Get-CounterSamples -MachineName $MachineName -Counter $localizedCounters -CustomErrorAction $CustomErrorAction)
}

function Get-LocalizedPerformanceCounterName {
    [CmdletBinding()]
    [OutputType('System.String')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$PerformanceCounterName
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        MachineName         = $ComputerName
        CatchActionFunction = ${Function:Invoke-CatchActions}
    }

    if ($null -eq $Script:EnglishOnlyOSCache) {
        $Script:EnglishOnlyOSCache = @{}
    }

    if ($null -eq $Script:Counter009Cache) {
        $Script:Counter009Cache = @{}
    }

    if ($null -eq $Script:CounterCurrentLanguageCache) {
        $Script:CounterCurrentLanguageCache = @{}
    }

    if (-not ($Script:EnglishOnlyOSCache.ContainsKey($ComputerName))) {
        $perfLib = Get-RemoteRegistrySubKey @baseParams -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib"

        if ($null -eq $perfLib) {
            Write-Verbose "No Perflib on computer. Assume EnglishOnlyOS for Get-Counter attempt"
            $Script:EnglishOnlyOSCache.Add($ComputerName, $true)
        } else {
            try {
                $englishOnlyOS = ($perfLib.GetSubKeyNames() |
                        Where-Object { $_ -like "0*" }).Count -eq 1
                Write-Verbose "Determined computer '$ComputerName' is englishOnlyOS: $englishOnlyOS"
                $Script:EnglishOnlyOSCache.Add($ComputerName, $englishOnlyOS)
            } catch {
                Write-Verbose "Failed to run GetSubKeyNames() on the opened key. Assume EnglishOnlyOS for Get-Counter attempt"
                $Script:EnglishOnlyOSCache.Add($ComputerName, $true)
                Invoke-CatchActions
            }
        }
    }

    if ($Script:EnglishOnlyOSCache[$ComputerName]) {
        Write-Verbose "English Only Machine, return same value"
        return $PerformanceCounterName
    }

    if (-not ($Script:Counter009Cache.ContainsKey($ComputerName))) {
        $params = $baseParams + @{
            SubKey    = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009"
            GetValue  = "Counter"
            ValueType = "MultiString"
        }
        $enUSCounterKeys = Get-RemoteRegistryValue @params

        if ($null -eq $enUSCounterKeys) {
            Write-Verbose "No 'en-US' (009) 'Counter' registry value found."
            Write-Verbose "Set Computer to English OS to just return PerformanceCounterName"
            $Script:EnglishOnlyOSCache[$ComputerName] = $true
            return $PerformanceCounterName
        } else {
            $Script:Counter009Cache.Add($ComputerName, $enUSCounterKeys)
        }
    }

    if (-not ($Script:CounterCurrentLanguageCache.ContainsKey($ComputerName))) {
        $params = $baseParams + @{
            SubKey    = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage"
            GetValue  = "Counter"
            ValueType = "MultiString"
        }
        $currentCounterKeys = Get-RemoteRegistryValue @params

        if ($null -eq $currentCounterKeys) {
            Write-Verbose "No 'localized' (CurrentLanguage) 'Counter' registry value found"
            Write-Verbose "Set Computer to English OS to just return PerformanceCounterName"
            $Script:EnglishOnlyOSCache[$ComputerName] = $true
            return $PerformanceCounterName
        } else {
            $Script:CounterCurrentLanguageCache.Add($ComputerName, $currentCounterKeys)
        }
    }

    $counterName = $PerformanceCounterName.ToLower()
    Write-Verbose "Trying to query ID index for Performance Counter: $counterName"
    $enUSCounterKeys = $Script:Counter009Cache[$ComputerName]
    $currentCounterKeys = $Script:CounterCurrentLanguageCache[$ComputerName]
    $counterIdIndex = ($enUSCounterKeys.ToLower().IndexOf("$counterName") - 1)

    if ($counterIdIndex -ge 0) {
        Write-Verbose "Counter ID Index: $counterIdIndex"
        Write-Verbose "Verify Value: $($enUSCounterKeys[$counterIdIndex + 1])"
        $counterId = $enUSCounterKeys[$counterIdIndex]
        Write-Verbose "Counter ID: $counterId"
        $localizedCounterNameIndex = ($currentCounterKeys.IndexOf("$counterId") + 1)

        if ($localizedCounterNameIndex -gt 0) {
            $localCounterName = $currentCounterKeys[$localizedCounterNameIndex]
            Write-Verbose "Found Localized Counter Index: $localizedCounterNameIndex"
            Write-Verbose "Localized Counter Name: $localCounterName"
            return $localCounterName
        } else {
            Write-Verbose "Failed to find Localized Counter Index"
            return $PerformanceCounterName
        }
    } else {
        Write-Verbose "Failed to find the counter ID."
        return $PerformanceCounterName
    }
}

function Get-CounterFullNameToCounterObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullCounterName
    )

    # Supported Scenarios
    # \\adt-e2k13aio1\LogicalDisk(HardDiskVolume1)\avg. disk sec/read
    # \\adt-e2k13aio1\\LogicalDisk(HardDiskVolume1)\avg. disk sec/read
    # \LogicalDisk(HardDiskVolume1)\avg. disk sec/read
    if (-not ($FullCounterName.StartsWith("\"))) {
        throw "Full Counter Name Should start with '\'"
    } elseif ($FullCounterName.StartsWith("\\")) {
        $endOfServerIndex = $FullCounterName.IndexOf("\", 2)
        $serverName = $FullCounterName.Substring(2, $endOfServerIndex - 2)
    } else {
        $endOfServerIndex = 0
    }
    $startOfCounterIndex = $FullCounterName.LastIndexOf("\") + 1
    $endOfCounterObjectIndex = $FullCounterName.IndexOf("(")

    if ($endOfCounterObjectIndex -eq -1) {
        $endOfCounterObjectIndex = $startOfCounterIndex - 1
    } else {
        $instanceName = $FullCounterName.Substring($endOfCounterObjectIndex + 1, ($FullCounterName.IndexOf(")") - $endOfCounterObjectIndex - 1))
    }

    $doubleSlash = 0
    if (($FullCounterName.IndexOf("\\", 2) -ne -1)) {
        $doubleSlash = 1
    }

    return [PSCustomObject]@{
        FullName     = $FullCounterName
        ServerName   = $serverName
        ObjectName   = ($FullCounterName.Substring($endOfServerIndex + 1 + $doubleSlash, $endOfCounterObjectIndex - $endOfServerIndex - 1 - $doubleSlash))
        InstanceName = $instanceName
        CounterName  = $FullCounterName.Substring($startOfCounterIndex)
    }
}
function Get-LoadBalancingReport {
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $CASServers = @()
    $MBXServers = @()
    $getExchangeServer = Get-ExchangeServer | Select-Object Name, Site, IsClientAccessServer, IsMailboxServer, AdminDisplayVersion, FQDN

    if ($SiteName -ne [string]::Empty) {
        Write-Grey("Site filtering ON.  Only Exchange 2013+ CAS servers in {0} will be used in the report." -f $SiteName)
        $CASServers = $getExchangeServer | Where-Object {
            ($_.IsClientAccessServer -eq $true) -and
            ($_.AdminDisplayVersion -Match "^Version 15") -and
            ([System.Convert]::ToString($_.Site).Split("/")[-1] -eq $SiteName) } | Select-Object Name, Site | Sort-Object Name
        Write-Grey("Site filtering ON.  Only Exchange 2013+ MBX servers in {0} will be used in the report." -f $SiteName)
        $MBXServers = $getExchangeServer | Where-Object {
                ($_.IsMailboxServer -eq $true) -and
                ($_.AdminDisplayVersion -Match "^Version 15") -and
                ([System.Convert]::ToString($_.Site).Split("/")[-1] -eq $SiteName) } | Select-Object Name, Site | Sort-Object Name
    } else {
        if ( ($null -eq $ServerList) ) {
            Write-Grey("Filtering OFF.  All Exchange 2013+ servers will be used in the report.")
            $CASServers = $getExchangeServer | Where-Object { ($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15") } | Select-Object Name, Site | Sort-Object Name
            $MBXServers = $getExchangeServer | Where-Object { ($_.IsMailboxServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15") } | Select-Object Name, Site | Sort-Object Name
        } else {
            Write-Grey("Custom server list is being used. Only servers specified after the -ServerList parameter will be used in the report.")
            $CASServers = $getExchangeServer | Where-Object { ($_.IsClientAccessServer -eq $true) -and ( ($_.Name -in $ServerList) -or ($_.FQDN -in $ServerList) ) } | Select-Object Name, Site | Sort-Object Name
            $MBXServers = $getExchangeServer | Where-Object { ($_.IsMailboxServer -eq $true) -and ( ($_.Name -in $ServerList) -or ($_.FQDN -in $ServerList) ) } | Select-Object Name, Site | Sort-Object Name
        }
    }

    if ($CASServers.Count -eq 0) {
        Write-Red("Error: No CAS servers found using the specified search criteria.")
        exit
    }

    if ($MBXServers.Count -eq 0) {
        Write-Red("Error: No MBX servers found using the specified search criteria.")
        exit
    }

    foreach ($server in $ServerList) {
        if ($server -notin $CASServers.Name -and $server -notin $MBXServers.Name) {
            Write-Warning "$server was not found as an Exchange server."
        }
    }

    function DisplayKeyMatching {
        param(
            [string]$CounterValue,
            [string]$DisplayValue
        )
        return [PSCustomObject]@{
            Counter = $CounterValue
            Display = $DisplayValue
        }
    }

    #Request stats from perfmon for all CAS
    $displayKeys = @{
        1  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_Total" "Load Distribution"
        2  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT" "root"
        3  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_API" "API"
        4  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_Autodiscover" "AutoDiscover"
        5  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_ecp" "ECP"
        6  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_EWS" "EWS"
        7  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_mapi" "MapiHttp"
        8  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_Microsoft-Server-ActiveSync" "EAS"
        9  = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_OAB" "OAB"
        10 = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_owa" "OWA"
        11 = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_owa_Calendar" "OWA-Calendar"
        12 = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_PowerShell" "PowerShell"
        13 = DisplayKeyMatching "_LM_W3SVC_DefaultSite_ROOT_Rpc" "RpcHttp"
    }

    #Request stats from perfmon for all MBX
    $displayKeysBackend = @{
        1  = DisplayKeyMatching "_LM_W3SVC_BackendSite_Total" "Load Distribution-BackEnd"
        2  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_API" "API-BackEnd"
        3  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_Autodiscover" "AutoDiscover-BackEnd"
        4  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_ecp" "ECP-BackEnd"
        5  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_EWS" "EWS-BackEnd"
        6  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_mapi_emsmdb" "MapiHttp_emsmdb-BackEnd"
        7  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_mapi_nspi" "MapiHttp_nspi-BackEnd"
        8  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_Microsoft-Server-ActiveSync" "EAS-BackEnd"
        9  = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_owa" "OWA-BackEnd"
        10 = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_PowerShell" "PowerShell-BackEnd"
        11 = DisplayKeyMatching "_LM_W3SVC_BackendSite_ROOT_Rpc" "RpcHttp-BackEnd"
    }

    $perServerStats = [ordered]@{}
    $perServerBackendStats = [ordered]@{}
    $totalStats = [ordered]@{}
    $totalBackendStats = [ordered]@{}

    #TODO: Improve performance here #1770
    #This is very slow loop against Each Server to collect this information.
    #Should be able to improve the speed by running 1 or 2 script blocks against the servers.
    foreach ( $CASServer in $CASServers.Name) {
        $currentErrors = $Error.Count
        $DefaultIdSite = Invoke-Command -ComputerName $CASServer -ScriptBlock { (Get-Website "Default Web Site").Id }

        $params = @{
            MachineName       = $CASServer
            Counter           = "\ASP.NET Apps v4.0.30319(_lm_w3svc_$($DefaultIdSite)_*)\Requests Executing"
            CustomErrorAction = "SilentlyContinue"
        }

        $FECounters = Get-LocalizedCounterSamples @params
        Invoke-CatchActionErrorLoop $currentErrors ${Function:Invoke-CatchActions}

        if ($null -eq $FECounters -or
            $FECounters.Count -eq 0) {
            Write-Verbose "Didn't find any counters on the server that matched."
            continue
        }

        foreach ( $sample in $FECounters) {
            $sample.Path = $sample.Path.Replace("_$($DefaultIdSite)_", "_DefaultSite_")
            $sample.InstanceName = $sample.InstanceName.Replace("_$($DefaultIdSite)_", "_DefaultSite_")
        }

        $counterSamples += $FECounters
    }

    foreach ($counterSample in $counterSamples) {
        $counterObject = Get-CounterFullNameToCounterObject -FullCounterName $counterSample.Path

        if (-not ($perServerStats.Contains($counterObject.ServerName))) {
            $perServerStats.Add($counterObject.ServerName, @{})
        }
        if (-not ($perServerStats[$counterObject.ServerName].Contains($counterObject.InstanceName))) {
            $perServerStats[$counterObject.ServerName].Add($counterObject.InstanceName, $counterSample.CookedValue)
        } else {
            Write-Verbose "This shouldn't occur...."
            $perServerStats[$counterObject.ServerName][$counterObject.InstanceName] += $counterSample.CookedValue
        }
        if (-not ($totalStats.Contains($counterObject.InstanceName))) {
            $totalStats.Add($counterObject.InstanceName, 0)
        }
        $totalStats[$counterObject.InstanceName] += $counterSample.CookedValue
    }

    $totalStats.Add("_lm_w3svc_DefaultSite_total", ($totalStats.Values | Measure-Object -Sum).Sum)

    for ($i = 0; $i -lt $perServerStats.count; $i++) {
        $perServerStats[$i].Add("_lm_w3svc_DefaultSite_total", ($perServerStats[$i].Values | Measure-Object -Sum).Sum)
    }

    $keyOrders = $displayKeys.Keys | Sort-Object

    foreach ( $MBXServer in $MBXServers.Name) {
        $currentErrors = $Error.Count
        $BackendIdSite = Invoke-Command -ComputerName $MBXServer -ScriptBlock { (Get-Website "Exchange Back End").Id }

        $params = @{
            MachineName       = $MBXServer
            Counter           = "\ASP.NET Apps v4.0.30319(_lm_w3svc_$($BackendIdSite)_*)\Requests Executing"
            CustomErrorAction = "SilentlyContinue"
        }

        $BECounters = Get-LocalizedCounterSamples @params
        Invoke-CatchActionErrorLoop $currentErrors ${Function:Invoke-CatchActions}

        if ($null -eq $BECounters -or
            $BECounters.Count -eq 0) {
            Write-Verbose "Didn't find any counters on the server that matched."
            continue
        }

        foreach ( $sample in $BECounters) {
            $sample.Path = $sample.Path.Replace("_$($BackendIdSite)_", "_BackendSite_")
            $sample.InstanceName = $sample.InstanceName.Replace("_$($BackendIdSite)_", "_BackendSite_")
        }

        $counterBackendSamples += $BECounters
    }

    foreach ($counterSample in $counterBackendSamples) {
        $counterObject = Get-CounterFullNameToCounterObject -FullCounterName $counterSample.Path

        if (-not ($perServerBackendStats.Contains($counterObject.ServerName))) {
            $perServerBackendStats.Add($counterObject.ServerName, @{})
        }
        if (-not ($perServerBackendStats[$counterObject.ServerName].Contains($counterObject.InstanceName))) {
            $perServerBackendStats[$counterObject.ServerName].Add($counterObject.InstanceName, $counterSample.CookedValue)
        } else {
            Write-Verbose "This shouldn't occur...."
            $perServerBackendStats[$counterObject.ServerName][$counterObject.InstanceName] += $counterSample.CookedValue
        }
        if (-not ($totalBackendStats.Contains($counterObject.InstanceName))) {
            $totalBackendStats.Add($counterObject.InstanceName, 0)
        }
        $totalBackendStats[$counterObject.InstanceName] += $counterSample.CookedValue
    }

    $totalBackendStats.Add("_lm_w3svc_BackendSite_total", ($totalBackendStats.Values | Measure-Object -Sum).Sum)

    for ($i = 0; $i -lt $perServerBackendStats.count; $i++) {
        $perServerBackendStats[$i].Add("_lm_w3svc_BackendSite_total", ($perServerBackendStats[$i].Values | Measure-Object -Sum).Sum)
    }

    $keyOrdersBackend = $displayKeysBackend.Keys | Sort-Object

    $htmlHeader = "<html>
    <style>
    BODY{font-family: Arial; font-size: 8pt;}
    H1{font-size: 16px;}
    H2{font-size: 14px;}
    H3{font-size: 12px;}
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
    TD{border: 1px solid black; padding: 5px; }
    td.Green{background: #7FFF00;}
    td.Yellow{background: #FFE600;}
    td.Red{background: #FF0000; color: #ffffff;}
    td.Info{background: #85D4FF;}
    </style>
    <body>
    <h1 align=""center"">Exchange Health Checker v$($BuildVersion)</h1>
    <h1 align=""center"">Domain : $(($(Get-ADDomain).DNSRoot).toUpper())</h1>
    <h2 align=""center"">Load balancer run finished : $((Get-Date).ToString("yyyy-MM-dd HH:mm"))</h2><br>"

    [array]$htmlLoadDetails += "<table>
    <tr><th>Server</th>
    <th>Site</th>
    "
    #Load the key Headers
    $keyOrders | ForEach-Object {
        if ( $totalStats[$displayKeys[$_].counter] -gt 0) {
            $htmlLoadDetails += "$([System.Environment]::NewLine)<th><center>$($displayKeys[$_].Display) Requests</center></th>
            <th><center>$($displayKeys[$_].Display) %</center></th>"
        }
    }
    $htmlLoadDetails += "$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)"

    foreach ($server in $CASServers) {
        $serverKey = $server.Name
        Write-Verbose "Working Server for HTML report $serverKey"
        $htmlLoadDetails += "<tr>
        <td>$($serverKey)</td>
        <td><center>$($server.Site)</center></td>"

        foreach ($key in $keyOrders) {
            if ( $totalStats[$displayKeys[$key].counter] -gt 0) {
                $currentDisplayKey = $displayKeys[$key]
                $totalRequests = $totalStats[$currentDisplayKey.Counter]

                if ($perServerStats.Contains($serverKey)) {
                    $serverValue = $perServerStats[$serverKey][$currentDisplayKey.Counter]
                    if ($null -eq $serverValue) { $serverValue = 0 }
                } else {
                    $serverValue = 0
                }
                if ($perServerStats.Contains($serverKey)) {
                    $serverValue = $perServerStats[$serverKey][$currentDisplayKey.Counter]
                    if ($null -eq $serverValue) { $serverValue = 0 }
                } else {
                    $serverValue = 0
                }
                if (($totalRequests -eq 0) -or
                ($null -eq $totalRequests)) {
                    $percentageLoad = 0
                } else {
                    $percentageLoad = [math]::Round((($serverValue / $totalRequests) * 100))
                    Write-Verbose "$($currentDisplayKey.Display) Server Value $serverValue Percentage usage $percentageLoad"

                    $htmlLoadDetails += "$([System.Environment]::NewLine)<td><center>$($serverValue)</center></td>
                    <td><center>$percentageLoad</center></td>"
                }
            }
        }
        $htmlLoadDetails += "$([System.Environment]::NewLine)</tr>"
    }

    # Totals
    $htmlLoadDetails += "$([System.Environment]::NewLine)<tr>
        <td><center>Totals</center></td>
        <td></td>"
    $keyOrders | ForEach-Object {
        if ( $totalStats[$displayKeys[$_].counter] -gt 0) {
            $htmlLoadDetails += "$([System.Environment]::NewLine)<td><center>$($totalStats[(($displayKeys[$_]).Counter)])</center></td>
            <td></td>"
        }
    }

    $htmlLoadDetails += "$([System.Environment]::NewLine)</table>"

    $htmlHeaderBackend = "<h2 align=""center"">BackEnd - Mailbox Role</h2><br>"

    [array]$htmlLoadDetailsBackend = "<table>
        <tr><th>Server</th>
        <th>Site</th>
        "
    #Load the key Headers
    $keyOrdersBackend | ForEach-Object {
        if ( $totalBackendStats[$displayKeysBackend[$_].counter] -gt 0) {
            $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)<th><center>$($displayKeysBackend[$_].Display) Requests</center></th>
            <th><center>$($displayKeysBackend[$_].Display) %</center></th>"
        }
    }
    $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)"

    foreach ($server in $MBXServers) {
        $serverKey = $server.Name
        Write-Verbose "Working Server for HTML report $serverKey"
        $htmlLoadDetailsBackend += "<tr>
            <td>$($serverKey)</td>
            <td><center>$($server.Site)</center></td>"

        foreach ($key in $keyOrdersBackend) {
            if ( $totalBackendStats[$displayKeysBackend[$key].counter] -gt 0) {
                $currentDisplayKey = $displayKeysBackend[$key]
                $totalRequests = $totalBackendStats[$currentDisplayKey.Counter]

                if ($perServerBackendStats.Contains($serverKey)) {
                    $serverValue = $perServerBackendStats[$serverKey][$currentDisplayKey.Counter]
                    if ($null -eq $serverValue) { $serverValue = 0 }
                } else {
                    $serverValue = 0
                }
                if ($perServerBackendStats.Contains($serverKey)) {
                    $serverValue = $perServerBackendStats[$serverKey][$currentDisplayKey.Counter]
                    if ($null -eq $serverValue) { $serverValue = 0 }
                } else {
                    $serverValue = 0
                }
                if (($totalRequests -eq 0) -or
                ($null -eq $totalRequests)) {
                    $percentageLoad = 0
                } else {
                    $percentageLoad = [math]::Round((($serverValue / $totalRequests) * 100))
                    Write-Verbose "$($currentDisplayKey.Display) Server Value $serverValue Percentage usage $percentageLoad"
                    $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)<td><center>$($serverValue)</center></td>
                    <td><center>$percentageLoad</center></td>"
                }
            }
        }
        $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)</tr>"
    }

    # Totals
    $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)<tr>
            <td><center>Totals</center></td>
            <td></td>"
    $keyOrdersBackend | ForEach-Object {
        if ( $totalBackendStats[$displayKeysBackend[$_].counter] -gt 0) {
            $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)<td><center>$($totalBackendStats[(($displayKeysBackend[$_]).Counter)])</center></td>
            <td></td>"
        }
    }
    $htmlLoadDetailsBackend += "$([System.Environment]::NewLine)</table>"

    $htmlReport = $htmlHeader + $htmlLoadDetails
    $htmlReport = $htmlReport + $htmlHeaderBackend + $htmlLoadDetailsBackend
    $htmlReport = $htmlReport + "</body></html>"

    $htmlFile = "$Script:OutputFilePath\HtmlLoadBalancerReport-$((Get-Date).ToString("yyyyMMddhhmmss")).html"
    $htmlReport | Out-File $htmlFile

    Write-Grey ""
    Write-Green "Client Access - FrontEnd information"
    foreach ($key in $keyOrders) {
        $currentDisplayKey = $displayKeys[$key]
        $totalRequests = $totalStats[$currentDisplayKey.Counter]

        if ($totalRequests -le 0) { continue }

        Write-Grey ""
        Write-Grey "Current $($currentDisplayKey.Display) Per Server"
        Write-Grey "Total Requests: $totalRequests"

        foreach ($serverKey in $perServerStats.Keys) {
            if ($perServerStats.Contains($serverKey)) {
                $serverValue = $perServerStats[$serverKey][$currentDisplayKey.Counter]
                Write-Grey "$serverKey : $serverValue Connections = $([math]::Round((([int]$serverValue / $totalRequests) * 100)))% Distribution"
            }
        }
    }

    Write-Grey ""
    Write-Green "Mailbox - BackEnd information"
    foreach ($key in $keyOrdersBackend) {
        $currentDisplayKey = $displayKeysBackend[$key]
        $totalRequests = $totalBackendStats[$currentDisplayKey.Counter]

        if ($totalRequests -le 0) { continue }

        Write-Grey ""
        Write-Grey "Current $($currentDisplayKey.Display) Per Server on Backend"
        Write-Grey "Total Requests: $totalRequests on Backend"

        foreach ($serverKey in $perServerBackendStats.Keys) {
            if ($perServerBackendStats.Contains($serverKey)) {
                $serverValue = $perServerBackendStats[$serverKey][$currentDisplayKey.Counter]
                Write-Grey "$serverKey : $serverValue Connections = $([math]::Round((([int]$serverValue / $totalRequests) * 100)))% Distribution on Backend"
            }
        }
    }
    Write-Grey ""
    Write-Grey "HTML File Report Written to $htmlFile"
}

function Get-ComputerCoresObject {
    param(
        [Parameter(Mandatory = $true)][string]$MachineName
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand) Passed: $MachineName"
        $errorOccurred = $false
        $numberOfCores = [int]::empty
        $exception = [string]::empty
        $exceptionType = [string]::empty
    } process {

        try {
            # rethrow the previous error to get handled here
            $wmi_obj_processor = Get-WmiObjectHandler -ComputerName $MachineName -Class "Win32_Processor" -CatchActionFunction { throw $_ }

            foreach ($processor in $wmi_obj_processor) {
                $numberOfCores += $processor.NumberOfCores
            }

            Write-Grey "Server $MachineName Cores: $numberOfCores"
        } catch {
            Invoke-CatchActions

            if ($_.Exception.GetType().FullName -eq "System.UnauthorizedAccessException") {
                Write-Yellow "Unable to get processor information from server $MachineName. You do not have the correct permissions to get this data from that server. Exception: $($_.ToString())"
            } else {
                Write-Yellow "Unable to get processor information from server $MachineName. Reason: $($_.ToString())"
            }
            $exception = $_.ToString()
            $exceptionType = $_.Exception.GetType().FullName
            $errorOccurred = $true
        }
    } end {
        return [PSCustomObject]@{
            Error         = $errorOccurred
            ComputerName  = $MachineName
            NumberOfCores = $numberOfCores
            Exception     = $exception
            ExceptionType = $exceptionType
        }
    }
}

function Get-ExchangeDCCoreRatio {

    Invoke-SetOutputInstanceLocation -FileName "HealthChecker-ExchangeDCCoreRatio"
    Invoke-ConfirmExchangeShell
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Grey("Exchange Server Health Checker Report - AD GC Core to Exchange Server Core Ratio - v{0}" -f $BuildVersion)
    $coreRatioObj = New-Object PSCustomObject

    try {
        Write-Verbose "Attempting to load Active Directory Module"
        Import-Module ActiveDirectory
        Write-Verbose "Successfully loaded"
    } catch {
        Write-Red("Failed to load Active Directory Module. Stopping the script")
        exit
    }

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    [array]$DomainControllers = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Forest.FindAllGlobalCatalogs($ADSite)

    [System.Collections.Generic.List[System.Object]]$DCList = New-Object System.Collections.Generic.List[System.Object]
    $DCCoresTotal = 0
    Write-Break
    Write-Grey("Collecting data for the Active Directory Environment in Site: {0}" -f $ADSite)
    $iFailedDCs = 0

    foreach ($DC in $DomainControllers) {
        $DCCoreObj = Get-ComputerCoresObject -MachineName $DC.Name
        $DCList.Add($DCCoreObj)

        if (-not ($DCCoreObj.Error)) {
            $DCCoresTotal += $DCCoreObj.NumberOfCores
        } else {
            $iFailedDCs++
        }
    }

    $coreRatioObj | Add-Member -MemberType NoteProperty -Name DCList -Value $DCList

    if ($iFailedDCs -eq $DomainControllers.count) {
        #Core count is going to be 0, no point to continue the script
        Write-Red("Failed to collect data from your DC servers in site {0}." -f $ADSite)
        Write-Yellow("Because we can't determine the ratio, we are going to stop the script. Verify with the above errors as to why we failed to collect the data and address the issue, then run the script again.")
        exit
    }

    [array]$ExchangeServers = Get-ExchangeServer | Where-Object { $_.Site -match $ADSite }
    $EXCoresTotal = 0
    [System.Collections.Generic.List[System.Object]]$EXList = New-Object System.Collections.Generic.List[System.Object]
    Write-Break
    Write-Grey("Collecting data for the Exchange Environment in Site: {0}" -f $ADSite)
    foreach ($svr in $ExchangeServers) {
        $EXCoreObj = Get-ComputerCoresObject -MachineName $svr.Name
        $EXList.Add($EXCoreObj)

        if (-not ($EXCoreObj.Error)) {
            $EXCoresTotal += $EXCoreObj.NumberOfCores
        }
    }
    $coreRatioObj | Add-Member -MemberType NoteProperty -Name ExList -Value $EXList

    Write-Break
    $CoreRatio = $EXCoresTotal / $DCCoresTotal
    Write-Grey("Total DC/GC Cores: {0}" -f $DCCoresTotal)
    Write-Grey("Total Exchange Cores: {0}" -f $EXCoresTotal)
    Write-Grey("You have {0} Exchange Cores for every Domain Controller Global Catalog Server Core" -f $CoreRatio)

    if ($CoreRatio -gt 8) {
        Write-Break
        Write-Red("Your Exchange to Active Directory Global Catalog server's core ratio does not meet the recommended guidelines of 8:1")
        Write-Red("Recommended guidelines for Exchange 2013/2016 for every 8 Exchange cores you want at least 1 Active Directory Global Catalog Core.")
        Write-Yellow("Documentation:")
        Write-Yellow("`thttps://aka.ms/HC-PerfSize")
        Write-Yellow("`thttps://aka.ms/HC-ADCoreCount")
    } else {
        Write-Break
        Write-Green("Your Exchange Environment meets the recommended core ratio of 8:1 guidelines.")
    }

    $XMLDirectoryPath = $Script:OutputFullPath.Replace(".txt", ".xml")
    $coreRatioObj | Export-Clixml $XMLDirectoryPath
    Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
    Write-Grey("Output XML Object file written to {0}" -f $XMLDirectoryPath)
}

function Get-MailboxDatabaseAndMailboxStatistics {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $AllDBs = Get-MailboxDatabaseCopyStatus -server $Server -ErrorAction SilentlyContinue
    $MountedDBs = $AllDBs | Where-Object { $_.ActiveCopy -eq $true }

    if ($MountedDBs.Count -gt 0) {
        Write-Grey("`tActive Database:")
        foreach ($db in $MountedDBs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating User Mailbox Total for Active Database: $_"; $TotalActiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Active User Mailboxes on server: " + $TotalActiveUserMailboxCount)
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating Public Mailbox Total for Active Database: $_"; $TotalActivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Active Public Folder Mailboxes on server: " + $TotalActivePublicFolderMailboxCount)
        Write-Grey("`tTotal Active Mailboxes on server " + $Server + ": " + ($TotalActiveUserMailboxCount + $TotalActivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Active Mailbox Databases found on server " + $Server + ".")
    }

    $HealthyDbs = $AllDBs | Where-Object { $_.Status -match 'Healthy' }

    if ($HealthyDbs.count -gt 0) {
        Write-Grey("`r`n`tPassive Databases:")
        foreach ($db in $HealthyDbs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating User Mailbox Total for Passive Healthy Databases: $_"; $TotalPassiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Passive user Mailboxes on Server: " + $TotalPassiveUserMailboxCount)
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating Passive Mailbox Total for Passive Healthy Databases: $_"; $TotalPassivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Passive Public Mailboxes on server: " + $TotalPassivePublicFolderMailboxCount)
        Write-Grey("`tTotal Passive Mailboxes on server: " + ($TotalPassiveUserMailboxCount + $TotalPassivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Passive Mailboxes found on server " + $Server + ".")
    }
}






function Get-ExtendedProtectionConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $false)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        [Parameter(Mandatory = $false)]
        [System.Version]$ExSetupVersion,

        [Parameter(Mandatory = $false)]
        [bool]$IsMailboxServer = $true,

        [Parameter(Mandatory = $false)]
        [bool]$IsClientAccessServer = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeEWS = $false,

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeEWSFe,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Exchange Back End/EWS")]
        [string[]]$SiteVDirLocations,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        function NewVirtualDirMatchingEntry {
            param(
                [Parameter(Mandatory = $true)]
                [string]$VirtualDirectory,
                [Parameter(Mandatory = $true)]
                [ValidateSet("Default Web Site", "Exchange Back End")]
                [string[]]$WebSite,
                [Parameter(Mandatory = $true)]
                [ValidateSet("None", "Allow", "Require")]
                [string[]]$ExtendedProtection,
                # Need to define this twice once for Default Web Site and Exchange Back End for the default values
                [Parameter(Mandatory = $false)]
                [string[]]$SslFlags = @("Ssl,Ssl128", "Ssl,Ssl128")
            )

            if ($WebSite.Count -ne $ExtendedProtection.Count) {
                throw "Argument count mismatch on $VirtualDirectory"
            }

            for ($i = 0; $i -lt $WebSite.Count; $i++) {
                # special conditions for Exchange 2013
                # powershell is on front and back so skip over those
                if ($IsExchange2013 -and $virtualDirectory -ne "Powershell") {
                    # No API virtual directory
                    if ($virtualDirectory -eq "API") { return }
                    if ($IsClientAccessServer -eq $false -and $WebSite[$i] -eq "Default Web Site") { continue }
                    if ($IsMailboxServer -eq $false -and $WebSite[$i] -eq "Exchange Back End") { continue }
                }
                # Set EWS VDir to None for known issues
                if ($ExcludeEWS -and $virtualDirectory -eq "EWS") { $ExtendedProtection[$i] = "None" }

                # EWS FE
                if ($ExcludeEWSFe -and $VirtualDirectory -eq "EWS" -and $WebSite[$i] -eq "Default Web Site") { $ExtendedProtection[$i] = "None" }

                if ($null -ne $SiteVDirLocations -and
                    $SiteVDirLocations.Count -gt 0) {
                    foreach ($SiteVDirLocation in $SiteVDirLocations) {
                        if ($SiteVDirLocation -eq "$($WebSite[$i])/$virtualDirectory") {
                            Write-Verbose "Set Extended Protection to None because of restriction override '$($WebSite[$i])\$virtualDirectory'"
                            $ExtendedProtection[$i] = "None"
                            break
                        }
                    }
                }

                [PSCustomObject]@{
                    VirtualDirectory   = $virtualDirectory
                    WebSite            = $WebSite[$i]
                    ExtendedProtection = $ExtendedProtection[$i]
                    SslFlags           = $SslFlags[$i]
                }
            }
        }

        # Intended for inside of Invoke-Command.
        function GetApplicationHostConfig {
            $appHostConfig = New-Object -TypeName Xml
            try {
                $appHostConfigPath = "$($env:WINDIR)\System32\inetSrv\config\applicationHost.config"
                $appHostConfig.Load($appHostConfigPath)
            } catch {
                Write-Verbose "Failed to loaded application host config file. $_"
                $appHostConfig = $null
            }
            return $appHostConfig
        }

        function GetExtendedProtectionConfiguration {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Xml.XmlNode]$Xml,
                [Parameter(Mandatory = $true)]
                [string]$Path
            )
            process {
                try {
                    $nodePath = [string]::Empty
                    $extendedProtection = "None"
                    $ipRestrictionsHashTable = @{}
                    $pathIndex = [array]::IndexOf(($Xml.configuration.location.path).ToLower(), $Path.ToLower())
                    $rootIndex = [array]::IndexOf(($Xml.configuration.location.path).ToLower(), ($Path.Split("/")[0]).ToLower())
                    $parentIndex = [array]::IndexOf(($Xml.configuration.location.path).ToLower(), ($Path.Substring(0, $Path.LastIndexOf("/")).ToLower()))

                    if ($pathIndex -ne -1) {
                        $configNode = $Xml.configuration.location[$pathIndex]
                        $nodePath = $configNode.Path
                        $ep = $configNode.'system.webServer'.security.authentication.windowsAuthentication.extendedProtection.tokenChecking
                        $ipRestrictions = $configNode.'system.webServer'.security.ipSecurity

                        if (-not ([string]::IsNullOrEmpty($ep))) {
                            Write-Verbose "Found tokenChecking: $ep"
                            $extendedProtection = $ep
                        } else {
                            if ($parentIndex -ne -1) {
                                $parentConfigNode = $Xml.configuration.location[$parentIndex]
                                $ep = $parentConfigNode.'system.webServer'.security.authentication.windowsAuthentication.extendedProtection.tokenChecking

                                if (-not ([string]::IsNullOrEmpty($ep))) {
                                    Write-Verbose "Found tokenChecking: $ep"
                                    $extendedProtection = $ep
                                } else {
                                    Write-Verbose "Failed to find tokenChecking. Using default value of None."
                                }
                            } else {
                                Write-Verbose "Failed to find tokenChecking. Using default value of None."
                            }
                        }

                        [string]$sslSettings = $configNode.'system.webServer'.security.access.sslFlags

                        if ([string]::IsNullOrEmpty($sslSettings)) {
                            Write-Verbose "Failed to find SSL settings for the path. Falling back to the root."

                            if ($rootIndex -ne -1) {
                                Write-Verbose "Found root path."
                                $rootConfigNode = $Xml.configuration.location[$rootIndex]
                                [string]$sslSettings = $rootConfigNode.'system.webServer'.security.access.sslFlags
                            }
                        }

                        if (-not([string]::IsNullOrEmpty($ipRestrictions))) {
                            Write-Verbose "IP-filtered restrictions detected"
                            foreach ($restriction in $ipRestrictions.add) {
                                $ipRestrictionsHashTable.Add($restriction.ipAddress, $restriction.allowed)
                            }
                        }

                        Write-Verbose "SSLSettings: $sslSettings"

                        if ($null -ne $sslSettings) {
                            [array]$sslFlags = ($sslSettings.Split(",").ToLower()).Trim()
                        } else {
                            $sslFlags = $null
                        }

                        # SSL flags: https://docs.microsoft.com/iis/configuration/system.webserver/security/access#attributes
                        $requireSsl = $false
                        $ssl128Bit = $false
                        $clientCertificate = "Unknown"

                        if ($null -eq $sslFlags) {
                            Write-Verbose "Failed to find SSLFlags"
                        } elseif ($sslFlags.Contains("none")) {
                            $clientCertificate = "Ignore"
                        } else {
                            if ($sslFlags.Contains("ssl")) { $requireSsl = $true }
                            if ($sslFlags.Contains("ssl128")) { $ssl128Bit = $true }
                            if ($sslFlags.Contains("sslNegotiateCert".ToLower())) {
                                $clientCertificate = "Accept"
                            } elseif ($sslFlags.Contains("sslRequireCert".ToLower())) {
                                $clientCertificate = "Require"
                            } else {
                                $clientCertificate = "Ignore"
                            }
                        }
                    }
                } catch {
                    Write-Verbose "Ran into some error trying to parse the application host config for $Path."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } end {
                return [PSCustomObject]@{
                    ExtendedProtection = $extendedProtection
                    ValidPath          = ($pathIndex -ne -1)
                    NodePath           = $nodePath
                    SslSettings        = [PSCustomObject]@{
                        RequireSsl        = $requireSsl
                        Ssl128Bit         = $ssl128Bit
                        ClientCertificate = $clientCertificate
                        Value             = $sslSettings
                    }
                    MitigationSettings = [PScustomObject]@{
                        AllowUnlisted = $ipRestrictions.allowUnlisted
                        Restrictions  = $ipRestrictionsHashTable
                    }
                }
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $computerResult = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlock { return $env:COMPUTERNAME }
        $serverConnected = $null -ne $computerResult

        if ($null -eq $computerResult) {
            Write-Verbose "Failed to connect to server $ComputerName"
            return
        }

        if ($null -eq $ExSetupVersion) {
            [System.Version]$ExSetupVersion = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlock {
                (Get-Command ExSetup.exe |
                    ForEach-Object { $_.FileVersionInfo } |
                    Select-Object -First 1).FileVersion
            }

            if ($null -eq $ExSetupVersion) {
                throw "Failed to determine Exchange build number"
            }
        } else {
            # Hopefully the caller knows what they are doing, best be from the correct server!!
            Write-Verbose "Caller passed the ExSetupVersion information"
        }

        if ($null -eq $ApplicationHostConfig) {
            Write-Verbose "Trying to load the application host config from $ComputerName"
            $params = @{
                ComputerName        = $ComputerName
                ScriptBlock         = ${Function:GetApplicationHostConfig}
                CatchActionFunction = $CatchActionFunction
            }

            $ApplicationHostConfig = Invoke-ScriptBlockHandler @params

            if ($null -eq $ApplicationHostConfig) {
                throw "Failed to load application host config from $ComputerName"
            }
        } else {
            # Hopefully the caller knows what they are doing, best be from the correct server!!
            Write-Verbose "Caller passed the application host config."
        }

        $default = "Default Web Site"
        $backend = "Exchange Back End"
        $Script:IsExchange2013 = $ExSetupVersion.Major -eq 15 -and $ExSetupVersion.Minor -eq 0
        try {
            $VirtualDirectoryMatchEntries = @(
                (NewVirtualDirMatchingEntry "API" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "Autodiscover" -WebSite $default, $backend -ExtendedProtection "None", "None")
                (NewVirtualDirMatchingEntry "ECP" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "EWS" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                (NewVirtualDirMatchingEntry "Microsoft-Server-ActiveSync" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                (NewVirtualDirMatchingEntry "Microsoft-Server-ActiveSync/Proxy" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                # This was changed due to Outlook for Mac not being able to do download the OAB.
                (NewVirtualDirMatchingEntry "OAB" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                (NewVirtualDirMatchingEntry "Powershell" -WebSite $default, $backend -ExtendedProtection "None", "Require" -SslFlags "SslNegotiateCert", "Ssl,Ssl128,SslNegotiateCert")
                (NewVirtualDirMatchingEntry "OWA" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "RPC" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "MAPI" -WebSite $default -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "PushNotifications" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "RPCWithCert" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "MAPI/emsmdb" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "MAPI/nspi" -WebSite $backend -ExtendedProtection "Require")
            )
        } catch {
            # Don't handle with Catch Error as this is a bug in the script.
            throw "Failed to create NewVirtualDirMatchingEntry. Inner Exception $_"
        }

        # Is Supported build of Exchange to have the configuration set.
        # Edge Server is not accounted for. It is the caller's job to not try to collect this info on Edge.
        $supportedVersion = $false
        $extendedProtectionList = New-Object 'System.Collections.Generic.List[object]'

        if ($ExSetupVersion.Major -eq 15) {
            if ($ExSetupVersion.Minor -eq 2) {
                $supportedVersion = $ExSetupVersion.Build -gt 1118 -or
                ($ExSetupVersion.Build -eq 1118 -and $ExSetupVersion.Revision -ge 11) -or
                ($ExSetupVersion.Build -eq 986 -and $ExSetupVersion.Revision -ge 28)
            } elseif ($ExSetupVersion.Minor -eq 1) {
                $supportedVersion = $ExSetupVersion.Build -gt 2507 -or
                ($ExSetupVersion.Build -eq 2507 -and $ExSetupVersion.Revision -ge 11) -or
                ($ExSetupVersion.Build -eq 2375 -and $ExSetupVersion.Revision -ge 30)
            } elseif ($ExSetupVersion.Minor -eq 0) {
                $supportedVersion = $ExSetupVersion.Build -gt 1497 -or
                ($ExSetupVersion.Build -eq 1497 -and $ExSetupVersion.Revision -ge 38)
            }
            Write-Verbose "Build $ExSetupVersion is supported: $supportedVersion"
        } else {
            Write-Verbose "Not on Exchange Version 15"
        }

        # Add all vDirs for which the IP filtering mitigation is supported
        $mitigationSupportedVDirs = $MyInvocation.MyCommand.Parameters["SiteVDirLocations"].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } |
            ForEach-Object { return $_.ValidValues }
        Write-Verbose "Supported mitigated virtual directories: $([string]::Join(",", $mitigationSupportedVDirs))"
    }
    process {
        try {
            foreach ($matchEntry in $VirtualDirectoryMatchEntries) {
                try {
                    Write-Verbose "Verify extended protection setting for $($matchEntry.VirtualDirectory) on web site $($matchEntry.WebSite)"

                    $extendedConfiguration = GetExtendedProtectionConfiguration -Xml $applicationHostConfig -Path "$($matchEntry.WebSite)/$($matchEntry.VirtualDirectory)"

                    # Extended Protection is a windows security feature which blocks MiTM attacks.
                    # Supported server roles are: Mailbox and ClientAccess
                    # Possible configuration settings are:
                    # <None>: This value specifies that IIS will not perform channel-binding token checking.
                    # <Allow>: This value specifies that channel-binding token checking is enabled, but not required.
                    # <Require>: This value specifies that channel-binding token checking is required.
                    # https://docs.microsoft.com/iis/configuration/system.webserver/security/authentication/windowsauthentication/extendedprotection/

                    if ($extendedConfiguration.ValidPath) {
                        Write-Verbose "Configuration was successfully returned: $($extendedConfiguration.ExtendedProtection)"
                    } else {
                        Write-Verbose "Extended protection setting was not queried because it wasn't found on the system."
                    }

                    $sslFlagsToSet = $extendedConfiguration.SslSettings.Value
                    $currentSetFlags = $sslFlagsToSet.Split(",").Trim()
                    foreach ($sslFlag in $matchEntry.SslFlags.Split(",").Trim()) {
                        if (-not($currentSetFlags.Contains($sslFlag))) {
                            Write-Verbose "Failed to find SSL Flag $sslFlag"
                            # We do not want to include None in the flags as that takes priority over the other options.
                            if ($sslFlagsToSet -eq "None") {
                                $sslFlagsToSet = "$sslFlag"
                            } else {
                                $sslFlagsToSet += ",$sslFlag"
                            }
                            Write-Verbose "Updated SSL Flags Value: $sslFlagsToSet"
                        } else {
                            Write-Verbose "SSL Flag $sslFlag set."
                        }
                    }

                    $expectedExtendedConfiguration = if ($supportedVersion) { $matchEntry.ExtendedProtection } else { "None" }
                    $virtualDirectoryName = "$($matchEntry.WebSite)/$($matchEntry.VirtualDirectory)"

                    # Supported Configuration is when the current value of Extended Protection is less than our expected extended protection value.
                    # While this isn't secure as we would like, it is still a supported state that should work.
                    $supportedExtendedConfiguration = $expectedExtendedConfiguration -eq $extendedConfiguration.ExtendedProtection

                    if ($supportedExtendedConfiguration) {
                        Write-Verbose "The EP value set to the expected value."
                    } else {
                        Write-Verbose "We are expecting a value of '$expectedExtendedConfiguration' but the current value is '$($extendedConfiguration.ExtendedProtection)'"

                        if ($expectedExtendedConfiguration -eq "Require" -or
                            ($expectedExtendedConfiguration -eq "Allow" -and
                            $extendedConfiguration.ExtendedProtection -eq "None")) {
                            $supportedExtendedConfiguration = $true
                            Write-Verbose "This is still supported because it is lower than what we recommended."
                        } else {
                            Write-Verbose "This is not supported because you are higher than the recommended value and will likely cause problems."
                        }
                    }

                    # Properly Secured Configuration is when the current Extended Protection value is equal to or greater than the Expected Extended Protection Configuration.
                    # If the Expected value is Allow, you can have the value set to Allow or Required and it will not be a security risk. However, if set to None, that is a security concern.
                    # For a mitigation scenario, like EWS BE, Required is the Expected value. Therefore, on those directories, we need to verify that IP filtering is set if not set to Require.
                    $properlySecuredConfiguration = $expectedExtendedConfiguration -eq $extendedConfiguration.ExtendedProtection

                    if ($properlySecuredConfiguration) {
                        Write-Verbose "We are 'properly' secure because we have EP set to the expected EP configuration value: $($expectedExtendedConfiguration)"
                    } elseif ($expectedExtendedConfiguration -eq "Require") {
                        Write-Verbose "Checking to see if we have mitigations enabled for the supported vDirs"
                        # Only care about virtual directories that we allow mitigation for
                        $properlySecuredConfiguration = $mitigationSupportedVDirs -contains $virtualDirectoryName -and
                        $extendedConfiguration.MitigationSettings.AllowUnlisted -eq "false"
                    } elseif ($expectedExtendedConfiguration -eq "Allow") {
                        Write-Verbose "Checking to see if Extended Protection is set to 'Require' to still be considered secure"
                        $properlySecuredConfiguration = $extendedConfiguration.ExtendedProtection -eq "Require"
                    } else {
                        Write-Verbose "Recommended EP setting is 'None' means you can have it higher, but you might run into other issues. But you are 'secure'."
                        $properlySecuredConfiguration = $true
                    }

                    Write-Verbose "Properly Secure Configuration value: $properlySecuredConfiguration"

                    $extendedProtectionList.Add([PSCustomObject]@{
                            VirtualDirectoryName          = $virtualDirectoryName
                            Configuration                 = $extendedConfiguration
                            # The current Extended Protection configuration set on the server
                            ExtendedProtection            = $extendedConfiguration.ExtendedProtection
                            # The Recommended Extended Protection is to verify that we have set the current Extended Protection
                            #   setting value to the Expected Extended Protection Value
                            RecommendedExtendedProtection = $expectedExtendedConfiguration -eq $extendedConfiguration.ExtendedProtection
                            # The supported/expected Extended Protection Configuration value that we should be set to (based off the build of Exchange)
                            ExpectedExtendedConfiguration = $expectedExtendedConfiguration
                            # Properly Secured is determined if we have a value equal to or greater than the ExpectedExtendedConfiguration value
                            # However, if we have a value greater than the expected, this could mean that we might run into a known set of issues.
                            ProperlySecuredConfiguration  = $properlySecuredConfiguration
                            # The Supported Extended Protection is a value that is equal to or lower than the Expected Extended Protection configuration.
                            # While this is not the best security setting, it is lower and shouldn't cause a connectivity issue and should still be supported.
                            SupportedExtendedProtection   = $supportedExtendedConfiguration
                            MitigationEnabled             = ($extendedConfiguration.MitigationSettings.AllowUnlisted -eq "false")
                            MitigationSupported           = $mitigationSupportedVDirs -contains $virtualDirectoryName
                            ExpectedSslFlags              = $matchEntry.SslFlags
                            SslFlagsSetCorrectly          = $sslFlagsToSet.Split(",").Trim().Count -eq $currentSetFlags.Count
                            SslFlagsToSet                 = $sslFlagsToSet
                        })
                } catch {
                    Write-Verbose "Failed to get extended protection match entry."
                    Invoke-CatchActionError $CatchActionFunction
                }
            }
        } catch {
            Write-Verbose "Failed to get get extended protection."
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            ComputerName                          = $ComputerName
            ServerConnected                       = $serverConnected
            SupportedVersionForExtendedProtection = $supportedVersion
            ApplicationHostConfig                 = $ApplicationHostConfig
            ExtendedProtectionConfiguration       = $extendedProtectionList
            ExtendedProtectionConfigured          = $null -ne ($extendedProtectionList.ExtendedProtection | Where-Object { $_ -ne "None" })
        }
    }
}

function Get-ExchangeDiagnosticInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [string]$Process,

        [Parameter(Mandatory = $true)]
        [string]$Component,

        [string]$Argument,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )
    process {
        try {
            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            $params = @{
                Process     = $Process
                Component   = $Component
                Server      = $Server
                ErrorAction = "Stop"
            }

            if (-not ([string]::IsNullOrEmpty($Argument))) {
                $params.Add("Argument", $Argument)
            }

            return (Get-ExchangeDiagnosticInfo @params)
        } catch {
            Write-Verbose "Failed to execute $($MyInvocation.MyCommand). Inner Exception: $_"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
}


function Get-ExchangeSettingOverride {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $updatedTime = [DateTime]::MinValue
        $settingOverrides = $null
        $simpleSettingOverrides = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        try {
            $params = @{
                Process             = "Microsoft.Exchange.Directory.TopologyService"
                Component           = "VariantConfiguration"
                Argument            = "Overrides"
                Server              = $Server
                CatchActionFunction = $CatchActionFunction
            }
            $diagnosticInfo = Get-ExchangeDiagnosticInformation @params

            if ($null -ne $diagnosticInfo) {
                Write-Verbose "Successfully got the Exchange Diagnostic Information"
                $xml = [xml]$diagnosticInfo.Result
                $overrides = $xml.Diagnostics.Components.VariantConfiguration.Overrides
                $updatedTime = $overrides.Updated
                $settingOverrides = $overrides.SettingOverride

                foreach ($override in $settingOverrides) {
                    Write-Verbose "Working on $($override.Name)"
                    $simpleSettingOverrides.Add([PSCustomObject]@{
                            Name          = $override.Name
                            ModifiedBy    = $override.ModifiedBy
                            Reason        = $override.Reason
                            ComponentName = $override.ComponentName
                            SectionName   = $override.SectionName
                            Status        = $override.Status
                            Parameters    = $override.Parameters.Parameter
                        })
                }
            } else {
                Write-Verbose "Failed to get Exchange Diagnostic Information"
            }
        } catch {
            Write-Verbose "Failed to get the Exchange setting override. Inner Exception: $_"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            Server                 = $Server
            LastUpdated            = $updatedTime
            SettingOverrides       = $settingOverrides
            SimpleSettingOverrides = $simpleSettingOverrides
        }
    }
}

function Get-ExSetupFileVersionInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exSetupDetails = [string]::Empty
    function Get-ExSetupDetailsScriptBlock {
        try {
            Get-Command ExSetup -ErrorAction Stop | ForEach-Object { $_.FileVersionInfo }
        } catch {
            try {
                Write-Verbose "Failed to find ExSetup by environment path locations. Attempting manual lookup."
                $installDirectory = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction Stop).MsiInstallPath

                if ($null -ne $installDirectory) {
                    Get-Command ([System.IO.Path]::Combine($installDirectory, "bin\ExSetup.exe")) -ErrorAction Stop | ForEach-Object { $_.FileVersionInfo }
                }
            } catch {
                Write-Verbose "Failed to find ExSetup, need to fallback."
            }
        }
    }

    $exSetupDetails = Invoke-ScriptBlockHandler -ComputerName $Server -ScriptBlock ${Function:Get-ExSetupDetailsScriptBlock} -ScriptBlockDescription "Getting ExSetup remotely" -CatchActionFunction $CatchActionFunction
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exSetupDetails
}

function Get-FileContentInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$FileLocation
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $allFiles = New-Object System.Collections.Generic.List[string]
    }
    process {
        foreach ($file in $FileLocation) {
            $allFiles.Add($file)
        }
    }
    end {
        $params = @{
            ComputerName           = $ComputerName
            ScriptBlockDescription = "Getting File Content Information"
            ArgumentList           = @(, $allFiles)
            ScriptBlock            = {
                param($FileLocations)
                $results = @{}
                foreach ($fileLocation in $FileLocations) {
                    $present = (Test-Path $fileLocation)

                    if ($present) {
                        $content = Get-Content $fileLocation -Raw -Encoding UTF8
                    } else {
                        $content = $null
                    }

                    $obj = [PSCustomObject]@{
                        Present  = $present
                        FileName = ([IO.Path]::GetFileName($fileLocation))
                        FilePath = $fileLocation
                        Content  = $content
                    }

                    $results.Add($fileLocation, $obj)
                }
                return $results
            }
        }
        return (Invoke-ScriptBlockHandler @params)
    }
}


function Get-AppPool {
    [CmdletBinding()]
    param ()

    begin {
        function Get-IndentLevel ($line) {
            if ($line.StartsWith(" ")) {
                ($line | Select-String "^ +").Matches[0].Length
            } else {
                0
            }
        }

        function Convert-FromAppPoolText {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string[]]
                $Text,

                [Parameter(Mandatory = $false)]
                [int]
                $Line = 0,

                [Parameter(Mandatory = $false)]
                [int]
                $MinimumIndentLevel = 2
            )

            if ($Line -ge $Text.Count) {
                return $null
            }

            $startingIndentLevel = Get-IndentLevel $Text[$Line]
            if ($startingIndentLevel -lt $MinimumIndentLevel) {
                return $null
            }

            $hash = @{}

            while ($Line -lt $Text.Count) {
                $indentLevel = Get-IndentLevel $Text[$Line]
                if ($indentLevel -gt $startingIndentLevel) {
                    # Skip until we get to the next thing at this level
                } elseif ($indentLevel -eq $startingIndentLevel) {
                    # We have a property at this level. Add it to the object.
                    if ($Text[$Line] -match "\[(\S+)\]") {
                        $name = $Matches[1]
                        $value = Convert-FromAppPoolText -Text $Text -Line ($Line + 1) -MinimumIndentLevel $startingIndentLevel
                        $hash[$name] = $value
                    } elseif ($Text[$Line] -match "\s+(\S+):`"(.*)`"") {
                        $name = $Matches[1]
                        $value = $Matches[2].Trim("`"")
                        $hash[$name] = $value
                    }
                } else {
                    # IndentLevel is less than what we started with, so return
                    [PSCustomObject]$hash
                    return
                }

                ++$Line
            }

            [PSCustomObject]$hash
        }

        $appPoolCmd = "$env:windir\System32\inetSrv\appCmd.exe"
    }

    process {
        $appPoolNames = & $appPoolCmd list appPool |
            Select-String "AppPool `"(\S+)`" " |
            ForEach-Object { $_.Matches.Groups[1].Value }

        foreach ($appPoolName in $appPoolNames) {
            $appPoolText = & $appPoolCmd list appPool $appPoolName /text:*
            Convert-FromAppPoolText -Text $appPoolText -Line 1
        }
    }
}
function Get-ExchangeAppPoolsInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $appPool = Invoke-ScriptBlockHandler -ComputerName $Server -ScriptBlock ${Function:Get-AppPool} `
        -ScriptBlockDescription "Getting App Pool information" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $exchangeAppPoolsInfo = @{}

    $appPool |
        Where-Object { $_.add.name -like "MSExchange*" } |
        ForEach-Object {
            Write-Verbose "Working on App Pool: $($_.add.name)"
            $configContent = Invoke-ScriptBlockHandler -ComputerName $Server -ScriptBlock {
                param(
                    $FilePath
                )
                if (Test-Path $FilePath) {
                    return (Get-Content $FilePath -Raw -Encoding UTF8).Trim()
                }
                return [string]::Empty
            } `
                -ScriptBlockDescription "Getting Content file for $($_.add.name)" `
                -ArgumentList $_.add.CLRConfigFile `
                -CatchActionFunction ${Function:Invoke-CatchActions}

            $gcUnknown = $true
            $gcServerEnabled = $false

            if (-not ([string]::IsNullOrEmpty($configContent))) {
                $gcSetting = ([xml]$configContent).Configuration.Runtime.gcServer.Enabled
                $gcUnknown = $gcSetting -ne "true" -and $gcSetting -ne "false"
                $gcServerEnabled = $gcSetting -eq "true"
            }
            $exchangeAppPoolsInfo.Add($_.add.Name, [PSCustomObject]@{
                    ConfigContent   = $configContent
                    AppSettings     = $_
                    GCUnknown       = $gcUnknown
                    GCServerEnabled = $gcServerEnabled
                })
        }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exchangeAppPoolsInfo
}


function Get-IISWebApplication {
    $webApplications = Get-WebApplication
    $returnList = New-Object 'System.Collections.Generic.List[object]'

    foreach ($webApplication in $webApplications) {
        try {
            $linkedConfigurationLine = $null
            $webConfigContent = $null
            $linkedConfigurationFilePath = $null
            $validWebConfig = $false # able to convert the file to xml type
            # set back to default, just incase there is an exception below
            $webConfigExists = $false
            $configurationFilePath = [string]::Empty
            $siteName = $webApplication.ItemXPath | Select-String -Pattern "site\[\@name='(.+)'\s|\]"
            $friendlyName = "$($siteName.Matches.Groups[1].Value)$($webApplication.Path)"
            Write-Verbose "Working on Web Application: $friendlyName"
            # Logic should be consistent for all ways we call Get-WebConfigFile
            try {
                $configurationFilePath = (Get-WebConfigFile "IIS:\Sites\$friendlyName").FullName
            } catch {
                $finder = "\\?\"
                if (($_.Exception.ErrorCode -eq -2147024846 -or
                        $_.Exception.ErrorCode -eq -2147024883) -and
                    $_.Exception.Message.Contains($finder)) {
                    $message = $_.Exception.Message
                    $index = $message.IndexOf($finder) + $finder.Length
                    $configurationFilePath = $message.Substring($index, ($message.IndexOf([System.Environment]::NewLine) - $index)).Trim()
                    Write-Verbose "Found possible file path from exception: $configurationFilePath"
                } else {
                    Write-Verbose "Unable to find possible file path based off exception: $($_.Exception)"
                }
            }
            $webConfigExists = Test-Path $configurationFilePath

            if ($webConfigExists) {
                $webConfigContent = (Get-Content $configurationFilePath -Raw -Encoding UTF8).Trim()

                try {
                    $linkedConfigurationLine = ([xml]$webConfigContent).configuration.assemblyBinding.linkedConfiguration.href
                    $validWebConfig = $true
                    if ($null -ne $linkedConfigurationLine) {
                        $linkedConfigurationFilePath = $linkedConfigurationLine.Substring("file://".Length)
                    }
                } catch {
                    Write-Verbose "Failed to convert '$configurationFilePath' to xml. Exception: $($_.Exception)"
                }
            }
        } catch {
            # Inside of Invoke-Command, can't use Invoke-CatchActions
            Write-Verbose "Failed to process additional context for: $($webApplication.ItemXPath). Exception: $($_.Exception)"
        }

        $returnList.Add([PSCustomObject]@{
                FriendlyName               = $friendlyName
                Path                       = $webApplication.Path
                ConfigurationFileInfo      = ([PSCustomObject]@{
                        Valid                       = $validWebConfig
                        Location                    = $configurationFilePath
                        Content                     = $webConfigContent
                        Exist                       = $webConfigExists
                        LinkedConfigurationLine     = $linkedConfigurationLine
                        LinkedConfigurationFilePath = $linkedConfigurationFilePath
                    })
                ApplicationPool            = $webApplication.applicationPool
                EnabledProtocols           = $webApplication.enabledProtocols
                ServiceAutoStartEnabled    = $webApplication.serviceAutoStartEnabled
                ServiceAutoStartProvider   = $webApplication.serviceAutoStartProvider
                PreloadEnabled             = $webApplication.preloadEnabled
                PreviouslyEnabledProtocols = $webApplication.previouslyEnabledProtocols
                ServiceAutoStartMode       = $webApplication.serviceAutoStartMode
                VirtualDirectoryDefaults   = $webApplication.virtualDirectoryDefaults
                Collection                 = $webApplication.Collection
                Location                   = $webApplication.Location
                ItemXPath                  = $webApplication.ItemXPath
                PhysicalPath               = $webApplication.PhysicalPath.Replace("%windir%", $env:windir).Replace("%SystemDrive%", $env:SystemDrive)
            })
    }

    return $returnList
}

function Get-IISWebSite {
    param(
        [array]$WebSitesToProcess
    )

    $returnList = New-Object 'System.Collections.Generic.List[object]'
    $webSites = New-Object 'System.Collections.Generic.List[object]'

    if ($null -eq $WebSitesToProcess) {
        $webSites.AddRange((Get-Website))
    } else {
        foreach ($iisWebSite in $WebSitesToProcess) {
            $webSites.Add((Get-Website -Name $($iisWebSite)))
        }
    }

    $bindings = Get-WebBinding

    foreach ($site in $webSites) {
        Write-Verbose "Working on Site: $($site.Name)"
        $siteBindings = $bindings |
            Where-Object { $_.ItemXPath -like "*@name='$($site.name)' and @id='$($site.id)'*" }
        # Logic should be consistent for all ways we call Get-WebConfigFile
        try {
            $configurationFilePath = (Get-WebConfigFile "IIS:\Sites\$($site.Name)").FullName
        } catch {
            $finder = "\\?\"
            if (($_.Exception.ErrorCode -eq -2147024846 -or
                    $_.Exception.ErrorCode -eq -2147024883) -and
                $_.Exception.Message.Contains($finder)) {
                $message = $_.Exception.Message
                $index = $message.IndexOf($finder) + $finder.Length
                $configurationFilePath = $message.Substring($index, ($message.IndexOf([System.Environment]::NewLine) - $index)).Trim()
                Write-Verbose "Found possible file path from exception: $configurationFilePath"
            } else {
                Write-Verbose "Unable to find possible file path based off exception: $($_.Exception)"
            }
        }

        $webConfigExists = Test-Path $configurationFilePath
        $webConfigContent = $null
        $webConfigContentXml = $null
        $validWebConfig = $false
        $customHeaderHstsObj = [PSCustomObject]@{
            enabled             = $false
            "max-age"           = 0
            includeSubDomains   = $false
            preload             = $false
            redirectHttpToHttps = $false
        }
        $customHeaderHsts = $null

        if ($webConfigExists) {
            $webConfigContent = (Get-Content $configurationFilePath -Raw -Encoding UTF8).Trim()

            try {
                $webConfigContentXml = [xml]$webConfigContent
                $validWebConfig = $true
            } catch {
                # Inside of Invoke-Command, can't use Invoke-CatchActions
                Write-Verbose "Failed to convert IIS web config '$configurationFilePath' to xml. Exception: $($_.Exception)"
            }
        }

        if ($validWebConfig) {
            <#
                HSTS configuration can be done in different ways:
                Via native HSTS control that comes with IIS 10.0 Version 1709.
                See: https://learn.microsoft.com/iis/get-started/whats-new-in-iis-10-version-1709/iis-10-version-1709-hsts
                The native control stores the HSTS configuration attributes in the <hsts> element which can be found under each <site> element.
                These settings are returned when running the Get-WebSite cmdlet and there is no need to prepare the data as they are ready for use.

                Via customHeader configuration (when running IIS older than the version mentioned before where the native HSTS config is not available
                or when admins prefer to do it via customHeader as there is no requirement to do it via native HSTS control instead of using customHeader).
                HSTS via customHeader configuration are stored under httpProtocol.customHeaders element. As we get the content in the previous
                call, we can simply use these data to extract the customHeader with name Strict-Transport-Security (if exists) and can then prepare
                the data for further processing.
                The following code searches for a customHeader with name Strict-Transport-Security. If we find the header, we then extract the directive
                and return them as PSCustomObject. We're looking for the following directive: max-age, includeSubDomains, preload, redirectHttpToHttps
            #>
            $customHeaderHsts = ($webConfigContentXml.configuration.'system.webServer'.httpProtocol.customHeaders.add | Where-Object {
                ($_.name -eq "Strict-Transport-Security")
                }).value
            if ($null -ne $customHeaderHsts) {
                Write-Verbose "Hsts via custom header configuration detected"
                $customHeaderHstsObj.enabled = $true
                # Make sure to ignore the case as per RFC 6797 the directives are case-insensitive
                # We ignore any other directives as these MUST be ignored by the User Agent (UA) as per RFC 6797
                # UAs MUST ignore any STS header field containing directives, or other header field value data,
                # that does not conform to the syntax defined in this specification.
                $maxAgeIndex = $customHeaderHsts.IndexOf("max-age=", [System.StringComparison]::OrdinalIgnoreCase)
                $includeSubDomainsIndex = $customHeaderHsts.IndexOf("includeSubDomains", [System.StringComparison]::OrdinalIgnoreCase)
                $preloadIndex = $customHeaderHsts.IndexOf("preload", [System.StringComparison]::OrdinalIgnoreCase)
                $redirectHttpToHttpsIndex = $customHeaderHsts.IndexOf("redirectHttpToHttps", [System.StringComparison]::OrdinalIgnoreCase)
                if ($maxAgeIndex -ne -1) {
                    Write-Verbose "max-age directive found"
                    $maxAgeValueIndex = $customHeaderHsts.IndexOf(";", $maxAgeIndex)
                    # add 8 to find the start index after 'max-age='
                    $maxAgeIndex = $maxAgeIndex + 8

                    if ($maxAgeValueIndex -ne -1) {
                        # subtract maxAgeIndex to get the length that we need to find the substring
                        $maxAgeValueIndex = $maxAgeValueIndex - $maxAgeIndex
                        $customHeaderHstsObj.'max-age' = $customHeaderHsts.Substring($maxAgeIndex, $maxAgeValueIndex)
                    } else {
                        $customHeaderHstsObj.'max-age' = $customHeaderHsts.Substring($maxAgeIndex)
                    }
                } else {
                    Write-Verbose "max-age directive not found"
                }

                if ($includeSubDomainsIndex -ne -1) {
                    Write-Verbose "includeSubDomains directive found"
                    $customHeaderHstsObj.includeSubDomains = $true
                }

                if ($preloadIndex -ne -1) {
                    Write-Verbose "preload directive found"
                    $customHeaderHstsObj.preload = $true
                }

                if ($redirectHttpToHttpsIndex -ne -1) {
                    Write-Verbose "redirectHttpToHttps directive found"
                    $customHeaderHstsObj.redirectHttpToHttps = $true
                }
            } else {
                Write-Verbose "No Hsts via custom header configuration detected"
            }
        }

        $physicalPath = [string]::Empty

        if (-not ([string]::IsNullOrEmpty($site.physicalPath))) {
            $physicalPath = $site.physicalPath.Replace("%windir%", $env:windir).Replace("%SystemDrive%", $env:SystemDrive)
        }

        $returnList.Add([PSCustomObject]@{
                Name                       = $site.Name
                Id                         = $site.Id
                State                      = $site.State
                Bindings                   = $siteBindings
                Limits                     = $site.Limits
                LogFile                    = $site.logFile
                TraceFailedRequestsLogging = $site.traceFailedRequestsLogging
                Hsts                       = [PSCustomObject]@{
                    NativeHstsSettings  = $site.hsts
                    HstsViaCustomHeader = $customHeaderHstsObj
                }
                ApplicationDefaults        = $site.applicationDefaults
                VirtualDirectoryDefaults   = $site.virtualDirectoryDefaults
                Collection                 = $site.collection
                ApplicationPool            = $site.applicationPool
                EnabledProtocols           = $site.enabledProtocols
                PhysicalPath               = $physicalPath
                ConfigurationFileInfo      = [PSCustomObject]@{
                    Location = $configurationFilePath
                    Content  = $webConfigContent
                    Exist    = $webConfigExists
                    Valid    = $validWebConfig
                }
            }
        )
    }
    return $returnList
}




function Get-ExchangeContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
    $exchangeContainerPath = ("CN=Microsoft Exchange,CN=Services," + $rootDSE.configurationNamingContext)
    $exchangeContainer = [ADSI]("LDAP://" + $exchangeContainerPath)
    Write-Verbose "Exchange Container Path: $($exchangeContainer.path)"
    return $exchangeContainer
}

function Get-OrganizationContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $exchangeContainer = Get-ExchangeContainer
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($exchangeContainer, "(objectClass=msExchOrganizationContainer)", @("distinguishedName"))
    return $searcher.FindOne().GetDirectoryEntry()
}

function Get-ExchangeProtocolContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    $ComputerName = $ComputerName.Split(".")[0]

    $organizationContainer = Get-OrganizationContainer
    $protocolContainerPath = ("CN=Protocols,CN=" + $ComputerName + ",CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups," + $organizationContainer.distinguishedName)
    $protocolContainer = [ADSI]("LDAP://" + $protocolContainerPath)
    Write-Verbose "Protocol Container Path: $($protocolContainer.Path)"
    return $protocolContainer
}

function Get-ExchangeWebSitesFromAd {
    [CmdletBinding()]
    [OutputType([System.Object])]
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    begin {
        function GetExchangeWebSiteFromCn {
            param (
                [string]$Site
            )

            if ($null -ne $Site) {
                $index = $Site.IndexOf("(") + 1
                if ($index -ne 0) {
                    return ($Site.Substring($index, ($Site.LastIndexOf(")") - $index)))
                }
            }
        }

        $processedExchangeWebSites = New-Object 'System.Collections.Generic.List[array]'
    }
    process {
        $protocolContainer = Get-ExchangeProtocolContainer -ComputerName $ComputerName
        if ($null -ne $protocolContainer) {
            $httpProtocol = $protocolContainer.Children | Where-Object {
                ($_.name -eq "HTTP")
            }

            foreach ($cn in $httpProtocol.Children.cn) {
                $processedExchangeWebSites.Add((GetExchangeWebSiteFromCn $cn))
            }
        }
    }
    end {
        return ($processedExchangeWebSites | Select-Object -Unique)
    }
}


function Get-ApplicationHostConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [ScriptBlock]$CatchActionFunction
    )

    $params = @{
        ComputerName           = $ComputerName
        ScriptBlockDescription = "Getting applicationHost.config"
        ScriptBlock            = { (Get-Content "$($env:WINDIR)\System32\inetSrv\config\applicationHost.config" -Raw -Encoding UTF8).Trim() }
        CatchActionFunction    = $CatchActionFunction
    }

    return Invoke-ScriptBlockHandler @params
}


function Get-IISModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        [Parameter(Mandatory = $false)]
        [bool]$SkipLegacyOSModulesCheck = $false,

        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $modulesToCheckList = New-Object 'System.Collections.Generic.List[object]'

        function GetModulePath {
            [CmdletBinding()]
            [OutputType([System.String])]
            param(
                [string]$Path
            )

            if (-not([String]::IsNullOrEmpty($Path))) {
                $returnPath = $Path

                if ($Path -match "\%.+\%") {
                    Write-Verbose "Environment variable found in path: $Path"
                    # Assuming that we have the env var always at the beginning of the string and no other vars within the string
                    # Example: %windir%\system32\SomeExample.dll
                    $preparedPath = ($Path.Split("%", [System.StringSplitOptions]::RemoveEmptyEntries))
                    if ($preparedPath.Count -eq 2) {
                        if ($preparedPath[0] -notmatch "\\.+\\") {
                            $varPath = [System.Environment]::GetEnvironmentVariable($preparedPath[0])
                            $returnPath = [String]::Join("", $varPath, $($preparedPath[1]))
                        }
                    }
                }
            } else {
                $returnPath = $null
            }

            return $returnPath
        }
        function GetIISModulesSignatureStatus {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [string]$ComputerName,

                [Parameter(Mandatory = $true)]
                [object[]]$Modules,

                [Parameter(Mandatory = $false)]
                [bool]$SkipLegacyOSModules = $false,

                [Parameter(Mandatory = $false)]
                [ScriptBlock]$CatchActionFunction
            )
            begin {
                # Add all modules here which should be skipped on legacy OS (pre-Windows Server 2016)
                $modulesToSkip = @(
                    "$env:windir\system32\inetSrv\cachUri.dll",
                    "$env:windir\system32\inetSrv\cachFile.dll",
                    "$env:windir\system32\inetSrv\cachtokn.dll",
                    "$env:windir\system32\inetSrv\cachHttp.dll",
                    "$env:windir\system32\inetSrv\compStat.dll",
                    "$env:windir\system32\inetSrv\defDoc.dll",
                    "$env:windir\system32\inetSrv\dirList.dll",
                    "$env:windir\system32\inetSrv\protsUp.dll",
                    "$env:windir\system32\inetSrv\redirect.dll",
                    "$env:windir\system32\inetSrv\static.dll",
                    "$env:windir\system32\inetSrv\authAnon.dll",
                    "$env:windir\system32\inetSrv\cusTerr.dll",
                    "$env:windir\system32\inetSrv\logHttp.dll",
                    "$env:windir\system32\inetSrv\iisEtw.dll",
                    "$env:windir\system32\inetSrv\iisFreb.dll",
                    "$env:windir\system32\inetSrv\iisReQs.dll",
                    "$env:windir\system32\inetSrv\isApi.dll",
                    "$env:windir\system32\inetSrv\compDyn.dll",
                    "$env:windir\system32\inetSrv\authCert.dll",
                    "$env:windir\system32\inetSrv\authBas.dll",
                    "$env:windir\system32\inetSrv\authsspi.dll",
                    "$env:windir\system32\inetSrv\authMd5.dll",
                    "$env:windir\system32\inetSrv\modRqFlt.dll",
                    "$env:windir\system32\inetSrv\filter.dll",
                    "$env:windir\system32\rpcProxy\rpcProxy.dll",
                    "$env:windir\system32\inetSrv\validCfg.dll",
                    "$env:windir\system32\wsmSvc.dll",
                    "$env:windir\system32\inetSrv\ipReStr.dll",
                    "$env:windir\system32\inetSrv\dipReStr.dll",
                    "$env:windir\system32\inetSrv\iis_ssi.dll",
                    "$env:windir\system32\inetSrv\cgi.dll",
                    "$env:windir\system32\inetSrv\iisFcGi.dll",
                    "$env:windir\system32\inetSrv\iisWSock.dll",
                    "$env:windir\system32\inetSrv\warmup.dll")

                $iisModulesList = New-Object 'System.Collections.Generic.List[object]'
                $signerSubject = "O=Microsoft Corporation, L=Redmond, S=Washington"
            }
            process {
                try {
                    $numberOfModulesFound = $Modules.Count
                    if ($numberOfModulesFound -ge 1) {
                        Write-Verbose "$numberOfModulesFound module(s) loaded by IIS"
                        Write-Verbose "SkipLegacyOSModules enabled? $SkipLegacyOSModules"
                        Write-Verbose "Checking file signing information now..."

                        $signatureParams = @{
                            ComputerName        = $ComputerName
                            ScriptBlock         = { Get-AuthenticodeSignature -FilePath $args[0] }
                            ArgumentList        = , $Modules.image # , is used to force the array to be passed as a single object
                            CatchActionFunction = $CatchActionFunction
                        }
                        $allSignatures = Invoke-ScriptBlockHandler @signatureParams

                        foreach ($m in $Modules) {
                            Write-Verbose "Now processing module: $($m.name)"
                            $signature = $null
                            $isModuleSigned = $false
                            $signatureDetails = [PSCustomObject]@{
                                Signer            = $null
                                SignatureStatus   = -1
                                IsMicrosoftSigned = $null
                            }

                            try {
                                $signature = $allSignatures | Where-Object { $_.Path -eq $m.image } | Select-Object -First 1
                                if (($SkipLegacyOSModules) -and
                                    ($m.image -in $modulesToSkip)) {
                                    Write-Verbose "Module was found in module skip list and will be skipped"
                                    # set to $null as this will indicate that the module was on the skip list
                                    $isModuleSigned = $null
                                } elseif ($null -ne $signature) {
                                    Write-Verbose "Performing signature status validation. Status: $($signature.Status)"
                                    # Signature Status Enum Values:
                                    # <0> Valid, <1> UnknownError, <2> NotSigned, <3> HashMismatch,
                                    # <4> NotTrusted, <5> NotSupportedFileFormat, <6> Incompatible,
                                    # https://docs.microsoft.com/dotnet/api/system.management.automation.signaturestatus
                                    if (($null -ne $signature.Status) -and
                                        ($signature.Status -ne 1) -and
                                        ($signature.Status -ne 2) -and
                                        ($signature.Status -ne 5) -and
                                        ($signature.Status -ne 6)) {

                                        $signatureDetails.SignatureStatus = $signature.Status
                                        $isModuleSigned = $true

                                        if ($null -ne $signature.SignerCertificate.Subject) {
                                            Write-Verbose "Signer information found. Subject: $($signature.SignerCertificate.Subject)"
                                            $signatureDetails.Signer = $signature.SignerCertificate.Subject.ToString()
                                            $signatureDetails.IsMicrosoftSigned = $signature.SignerCertificate.Subject -cmatch $signerSubject
                                        }
                                    }
                                } else {
                                    Write-Verbose "No signature information found for module $($m.name)"
                                    $isModuleSigned = $false
                                }

                                $iisModulesList.Add([PSCustomObject]@{
                                        Name             = $m.name
                                        Path             = $m.image
                                        Signed           = $isModuleSigned
                                        SignatureDetails = $signatureDetails
                                    })
                            } catch {
                                Write-Verbose "Unable to validate file signing information"
                                Invoke-CatchActionError $CatchActionFunction
                            }
                        }
                    } else {
                        Write-Verbose "No modules are loaded by IIS"
                    }
                } catch {
                    Write-Verbose "Failed to process global module information. $_"
                    Invoke-CatchActionError $CatchActionFunction
                }
            }
            end {
                return $iisModulesList
            }
        }
    }
    process {
        $ApplicationHostConfig.configuration.'system.webServer'.globalModules.add | ForEach-Object {
            $moduleFilePath = GetModulePath -Path $_.image
            # Replace the image path with the full path without environment variables
            $_.image = $moduleFilePath
            $modulesToCheckList.Add($_)
        }

        $getIISModulesSignatureStatusParams = @{
            ComputerName        = $ComputerName
            Modules             = $modulesToCheckList
            SkipLegacyOSModules = $SkipLegacyOSModulesCheck # now handled within the function as we need to return all modules which are loaded by IIS
            CatchActionFunction = $CatchActionFunction
        }
        $modules = GetIISModulesSignatureStatus @getIISModulesSignatureStatusParams

        # Validate if all modules that are loaded are digitally signed
        $allModulesAreSigned = (-not($modules.Signed.Contains($false)))
        Write-Verbose "Are all modules loaded by IIS digitally signed? $allModulesAreSigned"

        # Validate that all modules are signed by Microsoft Corp.
        $allModulesSignedByMSFT = (-not($modules.SignatureDetails.IsMicrosoftSigned.Contains($false)))
        Write-Verbose "Are all modules signed by Microsoft Corporation? $allModulesSignedByMSFT"

        # Validate if all signatures are valid (regardless of whether signed by Microsoft Corp. or not)
        $allSignaturesValid = $null -eq ($modules | Where-Object {
                ($_.Signed) -and
                ($_.SignatureDetails.SignatureStatus -ne 0)
            })
    }
    end {
        return [PSCustomObject]@{
            AllSignedModulesSignedByMSFT = $allModulesSignedByMSFT
            AllSignaturesValid           = $allSignaturesValid
            AllModulesSigned             = $allModulesAreSigned
            ModuleList                   = $modules
        }
    }
}

function Get-ExchangeServerIISSettings {
    param(
        [string]$ComputerName,
        [bool]$IsLegacyOS = $false,
        [ScriptBlock]$CatchActionFunction
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $params = @{
            ComputerName        = $ComputerName
            CatchActionFunction = $CatchActionFunction
        }

        try {
            $exchangeWebSites = Get-ExchangeWebSitesFromAd -ComputerName $ComputerName
            if ($exchangeWebSites.Count -gt 2) {
                Write-Verbose "Multiple OWA/ECP virtual directories detected"
            }
            Write-Verbose "Exchange websites detected: $([string]::Join(", " ,$exchangeWebSites))"
        } catch {
            Write-Verbose "Failed to get the Exchange Web Sites from Ad."
            $exchangeWebSites = $null
            Invoke-CatchActions
        }

        # We need to wrap the array into another array as the -WebSitesToProcess parameter expects an array object
        $webSite = Invoke-ScriptBlockHandler @params -ScriptBlock ${Function:Get-IISWebSite} -ArgumentList (, $exchangeWebSites) -ScriptBlockDescription "Get-IISWebSite"
        $webApplication = Invoke-ScriptBlockHandler @params -ScriptBlock ${Function:Get-IISWebApplication} -ScriptBlockDescription "Get-IISWebApplication"

        # Get the TokenCacheModule build information as we need it to perform version testing
        Write-Verbose "Trying to query TokenCacheModule version information"
        $tokenCacheModuleParams = @{
            ComputerName           = $Server
            ScriptBlockDescription = "Get TokenCacheModule version information"
            ScriptBlock            = { [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:windir\System32\inetsrv\cachtokn.dll") }
            CatchActionFunction    = ${Function:Invoke-CatchActions}
        }
        $tokenCacheModuleVersionInformation = Invoke-ScriptBlockHandler @tokenCacheModuleParams

        # Get the shared web configuration files
        $sharedWebConfigPaths = @($webApplication.ConfigurationFileInfo.LinkedConfigurationFilePath | Select-Object -Unique)
        $sharedWebConfig = $null

        if ($sharedWebConfigPaths.Count -gt 0) {
            $sharedWebConfig = Invoke-ScriptBlockHandler @params -ScriptBlock {
                param ($ConfigFiles)
                $ConfigFiles | ForEach-Object {
                    Write-Verbose "Working on shared config file: $_"
                    $validWebConfig = $false
                    $exist = Test-Path $_
                    $content = $null
                    try {
                        if ($exist) {
                            $content = (Get-Content $_ -Raw -Encoding UTF8).Trim()
                            [xml]$content | Out-Null # test to make sure it is valid
                            $validWebConfig = $true
                        }
                    } catch {
                        # Inside of Invoke-Command, can't use Invoke-CatchActions
                        Write-Verbose "Failed to convert shared web config '$_' to xml. Exception: $($_.Exception)"
                    }

                    [PSCustomObject]@{
                        Location = $_
                        Exist    = $exist
                        Content  = $content
                        Valid    = $validWebConfig
                    }
                }
            } -ArgumentList (, $sharedWebConfigPaths) -ScriptBlockDescription "Getting Shared Web Config Files"
        }

        Write-Verbose "Trying to query the 'applicationHost.config' file"
        $applicationHostConfig = Get-ApplicationHostConfig $ComputerName $CatchActionFunction

        if ($null -ne $applicationHostConfig) {
            Write-Verbose "Trying to query the modules which are loaded by IIS"
            try {
                [xml]$xmlApplicationHostConfig = [xml]$applicationHostConfig
            } catch {
                Write-Verbose "Failed to convert the Application Host Config to XML"
                Invoke-CatchActions
                # Don't attempt to run Get-IISModules
                return
            }
            $iisModulesParams = @{
                ComputerName             = $ComputerName
                ApplicationHostConfig    = $xmlApplicationHostConfig
                SkipLegacyOSModulesCheck = $IsLegacyOS
                CatchActionFunction      = $CatchActionFunction
            }
            $iisModulesInformation = Get-IISModules @iisModulesParams
        } else {
            Write-Verbose "No 'applicationHost.config' file returned by previous call"
        }
    } end {
        return [PSCustomObject]@{
            ApplicationHostConfig          = $applicationHostConfig
            IISModulesInformation          = $iisModulesInformation
            IISTokenCacheModuleInformation = $tokenCacheModuleVersionInformation
            IISConfigurationSettings       = $iisConfigurationSettings
            IISWebSite                     = $webSite
            IISWebApplication              = $webApplication
            IISSharedWebConfig             = $sharedWebConfig
        }
    }
}

function Get-ExchangeAES256CBCDetails {
    param(
        [Parameter(Mandatory = $false)]
        [String]$Server = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [System.Object]$VersionInformation
    )

    <#
        AES256-CBC encryption support check
        https://techcommunity.microsoft.com/t5/security-compliance-and-identity/encryption-algorithm-changes-in-microsoft-purview-information/ba-p/3831909
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        function GetRegistryAclCheckScriptBlock {
            $sbMsipcRegistryAclAsExpected = $false
            $regPathToCheck = "HKLM:\SOFTWARE\Microsoft\MSIPC\Server"
            # Translates to: "NetworkService", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow"
            # See: https://learn.microsoft.com/dotnet/api/system.security.accesscontrol.registryaccessrule.-ctor?view=net-7.0#system-security-accesscontrol-registryaccessrule-ctor(system-security-principal-identityreference-system-security-accesscontrol-registryrights-system-security-accesscontrol-inheritanceflags-system-security-accesscontrol-propagationflags-system-security-accesscontrol-accesscontroltype)
            $networkServiceAcl = New-Object System.Security.AccessControl.RegistryAccessRule(
                (New-Object System.Security.Principal.SecurityIdentifier("S-1-5-20")), 983103, 3, 0, 0
            )
            $pathExists = Test-Path $regPathToCheck

            if ($pathExists -eq $false) {
                Write-Verbose "Unable to query Acl of registry key $regPathToCheck assuming that the key doesn't exist"
            } else {
                $acl = Get-Acl -Path $regPathToCheck
                # ToDo: As we have multiple places in HC where we query acls, we should consider creating a function
                # that can be used to do the acl call, similar to what we do in Get-ExchangeRegistryValues.ps1.
                Write-Verbose "Registry key exists and Acl was successfully queried - validating Acl now"
                try {
                    $aclMatch = $acl.Access.Where({
                    ($_.RegistryRights -eq $networkServiceAcl.RegistryRights) -and
                    ($_.AccessControlType -eq $networkServiceAcl.AccessControlType) -and
                    ($_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -eq $networkServiceAcl.IdentityReference) -and
                    ($_.InheritanceFlags -eq $networkServiceAcl.InheritanceFlags) -and
                    ($_.PropagationFlags -eq $networkServiceAcl.PropagationFlags)
                        })

                    if (@($aclMatch).Count -ge 1) {
                        Write-Verbose "Acl for NetworkService is as expected"
                        $sbMsipcRegistryAclAsExpected = $true
                    } else {
                        Write-Verbose "Acl for NetworkService was not found or is not as expected"
                    }
                } catch {
                    Write-Verbose "Unable to verify Acl on registry key $regPathToCheck"
                    # Unable to use Invoke-CatchActions because of remote script block
                }
            }

            return [PSCustomObject]@{
                PathExits                       = $pathExists
                RegistryKeyConfiguredAsExpected = $sbMsipcRegistryAclAsExpected
            }
        }

        $aes256CBCSupported = $false
        $msipcRegistryAclAsExpected = $false
    } process {
        # First, check if the build running on the server supports AES256-CBC
        if (Test-ExchangeBuildGreaterOrEqualThanSecurityPatch -CurrentExchangeBuild $VersionInformation -SU "Aug23SU") {

            Write-Verbose "AES256-CBC encryption for information protection is supported by this Exchange Server build"
            $aes256CBCSupported = $true

            $params = @{
                ComputerName        = $Server
                ScriptBlock         = ${Function:GetRegistryAclCheckScriptBlock}
                CatchActionFunction = ${Function:Invoke-CatchActions}
            }
            $results = Invoke-ScriptBlockHandler @params
            Write-Verbose "Found Registry Path: $($results.PathExits)"
            Write-Verbose "Configured Correctly: $($results.RegistryKeyConfiguredAsExpected)"
            $msipcRegistryAclAsExpected = $results.RegistryKeyConfiguredAsExpected
        } else {
            Write-Verbose "AES256-CBC encryption for information protection is not supported by this Exchange Server build"
        }
    } end {
        return [PSCustomObject]@{
            AES256CBCSupportedBuild         = $aes256CBCSupported
            RegistryKeyConfiguredAsExpected = $msipcRegistryAclAsExpected
            ValidAESConfiguration           = (($aes256CBCSupported) -and ($msipcRegistryAclAsExpected))
        }
    }
}

function Get-ExchangeConnectors {
    [CmdletBinding()]
    [OutputType("System.Object[]")]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,
        [Parameter(Mandatory = $false)]
        [object]
        $CertificateObject
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: $ComputerName"
        function ExchangeConnectorObjectFactory {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [object]
                $ConnectorObject
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            $exchangeFactoryConnectorReturnObject = [PSCustomObject]@{
                Identity           = $ConnectorObject.Identity
                Name               = $ConnectorObject.Name
                Enabled            = $ConnectorObject.Enabled
                CloudEnabled       = $false
                ConnectorType      = $null
                TransportRole      = $null
                SmartHosts         = $null
                AddressSpaces      = $null
                RequireTLS         = $false
                TlsAuthLevel       = $null
                TlsDomain          = $null
                CertificateDetails = [PSCustomObject]@{
                    CertificateMatchDetected = $false
                    GoodTlsCertificateSyntax = $false
                    TlsCertificateName       = $null
                    TlsCertificateNameStatus = $null
                    TlsCertificateSet        = $false
                    CertificateLifetimeInfo  = $null
                }
            }

            Write-Verbose ("Creating object for Exchange connector: '{0}'" -f $ConnectorObject.Identity)
            if ($null -ne $ConnectorObject.Server) {
                Write-Verbose "Exchange ReceiveConnector detected"
                $exchangeFactoryConnectorReturnObject.ConnectorType =  "Receive"
                $exchangeFactoryConnectorReturnObject.TransportRole = $ConnectorObject.TransportRole
                if (-not([System.String]::IsNullOrEmpty($ConnectorObject.TlsDomainCapabilities))) {
                    $exchangeFactoryConnectorReturnObject.CloudEnabled = $true
                }
            } else {
                Write-Verbose "Exchange SendConnector detected"
                $exchangeFactoryConnectorReturnObject.ConnectorType = "Send"
                $exchangeFactoryConnectorReturnObject.CloudEnabled = $ConnectorObject.CloudServicesMailEnabled
                $exchangeFactoryConnectorReturnObject.TlsDomain = $ConnectorObject.TlsDomain
                if ($null -ne $ConnectorObject.TlsAuthLevel) {
                    $exchangeFactoryConnectorReturnObject.TlsAuthLevel = $ConnectorObject.TlsAuthLevel
                }

                if ($null -ne $ConnectorObject.SmartHosts) {
                    $exchangeFactoryConnectorReturnObject.SmartHosts = $ConnectorObject.SmartHosts
                }

                if ($null -ne $ConnectorObject.AddressSpaces) {
                    $exchangeFactoryConnectorReturnObject.AddressSpaces = $ConnectorObject.AddressSpaces
                }
            }

            if ($null -ne $ConnectorObject.TlsCertificateName) {
                Write-Verbose "TlsCertificateName is configured on this connector"
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateSet = $true
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateName = ($ConnectorObject.TlsCertificateName).ToString()
            } else {
                Write-Verbose "TlsCertificateName is not configured on this connector"
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNameEmpty"
            }

            $exchangeFactoryConnectorReturnObject.RequireTLS = $ConnectorObject.RequireTLS

            return $exchangeFactoryConnectorReturnObject
        }

        function NormalizeTlsCertificateName {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $TlsCertificateName
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                Write-Verbose ("TlsCertificateName that was passed: '{0}'" -f $TlsCertificateName)
                # RegEx to match the recommended value which is "<I>X.500Issuer<S>X.500Subject"
                if ($TlsCertificateName -match "(<i>).*(<s>).*") {
                    $expectedTlsCertificateNameDetected = $true
                    $issuerIndex = $TlsCertificateName.IndexOf("<I>", [System.StringComparison]::OrdinalIgnoreCase)
                    $subjectIndex = $TlsCertificateName.IndexOf("<S>", [System.StringComparison]::OrdinalIgnoreCase)

                    Write-Verbose "TlsCertificateName that matches the expected syntax was passed"
                } else {
                    # Failsafe to detect cases where <I> and <S> are missing in TlsCertificateName
                    $issuerIndex = $TlsCertificateName.IndexOf("CN=", [System.StringComparison]::OrdinalIgnoreCase)
                    $subjectIndex = $TlsCertificateName.LastIndexOf("CN=", [System.StringComparison]::OrdinalIgnoreCase)

                    Write-Verbose "TlsCertificateName with bad syntax was passed"
                }

                # We stop processing if Issuer OR Subject index is -1 (no match found)
                if (($issuerIndex -ne -1) -and
                    ($subjectIndex -ne -1)) {
                    if ($expectedTlsCertificateNameDetected) {
                        $issuer = $TlsCertificateName.Substring(($issuerIndex + 3), ($subjectIndex - 3))
                        $subject = $TlsCertificateName.Substring($subjectIndex + 3)
                    } else {
                        $issuer  = $TlsCertificateName.Substring($issuerIndex, $subjectIndex)
                        $subject = $TlsCertificateName.Substring($subjectIndex)
                    }
                }

                if (($null -ne $issuer) -and
                    ($null -ne $subject)) {
                    return [PSCustomObject]@{
                        Issuer     = $issuer
                        Subject    = $subject
                        GoodSyntax = $expectedTlsCertificateNameDetected
                    }
                }
            } catch {
                Write-Verbose "We hit an exception while parsing the TlsCertificateName string"
                Invoke-CatchActions
            }
        }

        function FindMatchingExchangeCertificate {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [object]
                $CertificateObject,
                [Parameter(Mandatory = $true)]
                [object]
                $ConnectorCustomObject
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                Write-Verbose ("{0} connector object(s) was/were passed to process" -f $ConnectorCustomObject.Count)
                foreach ($connectorObject in $ConnectorCustomObject) {

                    if ($null -ne $ConnectorObject.CertificateDetails.TlsCertificateName) {
                        $connectorTlsCertificateNormalizedObject = NormalizeTlsCertificateName `
                            -TlsCertificateName $ConnectorObject.CertificateDetails.TlsCertificateName

                        if ($null -eq $connectorTlsCertificateNormalizedObject) {
                            Write-Verbose "Unable to normalize TlsCertificateName - could be caused by an invalid TlsCertificateName configuration"
                            $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNameSyntaxInvalid"
                        } else {
                            if ($connectorTlsCertificateNormalizedObject.GoodSyntax) {
                                $connectorObject.CertificateDetails.GoodTlsCertificateSyntax = $connectorTlsCertificateNormalizedObject.GoodSyntax
                            }

                            $certificateMatches = 0
                            $certificateLifetimeInformation = @{}
                            foreach ($certificate in $CertificateObject) {
                                if (($certificate.Issuer -eq $connectorTlsCertificateNormalizedObject.Issuer) -and
                                    ($certificate.Subject -eq $connectorTlsCertificateNormalizedObject.Subject)) {
                                    Write-Verbose ("Certificate: '{0}' matches Connectors: '{1}' TlsCertificateName: '{2}'" -f $certificate.Thumbprint, $connectorObject.Identity, $connectorObject.CertificateDetails.TlsCertificateName)
                                    $connectorObject.CertificateDetails.CertificateMatchDetected = $true
                                    $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateMatch"
                                    $certificateLifetimeInformation.Add($certificate.Thumbprint, $certificate.LifetimeInDays)

                                    $certificateMatches++
                                }
                            }

                            if ($certificateMatches -eq 0) {
                                Write-Verbose "No matching certificate was found on the server"
                                $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNotFound"
                            } else {
                                Write-Verbose ("We found: '{0}' matching certificates on the server" -f $certificateMatches)
                                $connectorObject.CertificateDetails.CertificateLifetimeInfo = $certificateLifetimeInformation
                            }
                        }
                    }
                }
            } catch {
                Write-Verbose "Hit an exception while trying to locate the configured certificate on the system"
                Invoke-CatchActions
            }

            return $ConnectorCustomObject
        }
    }
    process {
        Write-Verbose ("Trying to query Exchange connectors for server: '{0}'" -f $ComputerName)
        try {
            $allReceiveConnectors = Get-ReceiveConnector -Server $ComputerName -ErrorAction Stop
            $allSendConnectors = Get-SendConnector -ErrorAction Stop
            $connectorCustomObject = @()

            foreach ($receiveConnector in $allReceiveConnectors) {
                $connectorCustomObject += ExchangeConnectorObjectFactory -ConnectorObject $receiveConnector
            }

            foreach ($sendConnector in $allSendConnectors) {
                $connectorCustomObject += ExchangeConnectorObjectFactory -ConnectorObject $sendConnector
            }

            if (($null -ne $connectorCustomObject) -and
                ($null -ne $CertificateObject)) {
                $connectorReturnObject = FindMatchingExchangeCertificate `
                    -CertificateObject $CertificateObject `
                    -ConnectorCustomObject $connectorCustomObject
            } else {
                Write-Verbose "No connector object which can be processed was returned"
                $connectorReturnObject = $connectorCustomObject
            }
        } catch {
            Write-Verbose "Hit an exception while processing the Exchange Send-/Receive Connectors"
            Invoke-CatchActions
        }
    }
    end {
        return $connectorReturnObject
    }
}

function Get-ExchangeDependentServices {
    [CmdletBinding()]
    param(
        [string]$MachineName
    )
    begin {

        function NewServiceObject {
            param(
                [object]$Service
            )
            $name = $Service.Name
            $status = "Unknown"
            $startType = "Unknown"
            try {
                $status = $Service.Status.ToString()
            } catch {
                Write-Verbose "Failed to set Status of service '$name'"
                Invoke-CatchActions
            }
            try {
                $startType = $Service.StartType.ToString()
            } catch {
                Write-Verbose "Failed to set Start Type of service '$name'"
                Invoke-CatchActions
            }
            return [PSCustomObject]@{
                Name      = $name
                Status    = $status
                StartType = $startType
            }
        }

        function NewMonitorServiceObject {
            param(
                [Parameter(Mandatory = $true, Position = 1)]
                [string]$ServiceName,
                [Parameter(Mandatory = $false)]
                [ValidateSet("Automatic", "Manual")]
                [string]$StartType = "Automatic",
                [Parameter(Mandatory = $false)]
                [ValidateSet("Common", "Critical")]
                [string]$Type = "Critical"
            )
            return [PSCustomObject]@{
                ServiceName = $ServiceName
                StartType   = $StartType
                Type        = $Type
            }
        }
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $servicesList = @(
            (NewMonitorServiceObject "WinMgmt"),
            (NewMonitorServiceObject "W3Svc"),
            (NewMonitorServiceObject "IISAdmin"),
            (NewMonitorServiceObject "Pla" -StartType "Manual"),
            (NewMonitorServiceObject "MpsSvc"),
            (NewMonitorServiceObject "RpcEptMapper"),
            (NewMonitorServiceObject "EventLog"),
            (NewMonitorServiceObject "MSExchangeADTopology"),
            (NewMonitorServiceObject "MSExchangeDelivery"),
            (NewMonitorServiceObject "MSExchangeFastSearch"),
            (NewMonitorServiceObject "MSExchangeFrontEndTransport"),
            (NewMonitorServiceObject "MSExchangeIS"),
            (NewMonitorServiceObject "MSExchangeRepl"),
            (NewMonitorServiceObject "MSExchangeRPC"),
            (NewMonitorServiceObject "MSExchangeServiceHost"),
            (NewMonitorServiceObject "MSExchangeSubmission"),
            (NewMonitorServiceObject "MSExchangeTransport"),
            (NewMonitorServiceObject "HostControllerService"),
            (NewMonitorServiceObject "MSExchangeAntispamUpdate" -Type "Common"),
            (NewMonitorServiceObject "MSComplianceAudit" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeCompliance" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeDagMgmt" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeDiagnostics" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeEdgeSync" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeHM" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeHMRecovery" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeMailboxAssistants" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeMailboxReplication" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeMitigation" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeThrottling" -Type "Common"),
            (NewMonitorServiceObject "MSExchangeTransportLogSearch" -Type "Common"),
            (NewMonitorServiceObject "BITS" -Type "Common" -StartType "Manual") # BITS have seen both Manual and Automatic
        )
        $notRunningCriticalServices = New-Object 'System.Collections.Generic.List[object]'
        $notRunningCommonServices = New-Object 'System.Collections.Generic.List[object]'
        $misconfiguredServices = New-Object 'System.Collections.Generic.List[object]'
        $getServicesList = New-Object 'System.Collections.Generic.List[object]'
        $monitorServicesList = New-Object 'System.Collections.Generic.List[object]'
    } process {
        try {
            $getServices = Get-Service -ComputerName $MachineName -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to get the services on the server"
            Invoke-CatchActions
            return
        }

        foreach ($service in $getServices) {

            $monitor = $servicesList | Where-Object { $_.ServiceName -eq $service.Name }

            if ($null -ne $monitor) {
                # Any critical services not running, add to list
                # Any critical or common services not set to Automatic that should be or set to disabled, add to list
                # Any common services not running, besides the ones that are set to manual, add to list
                Write-Verbose "Working on $($monitor.ServiceName)"
                $monitorServicesList.Add((NewServiceObject $service))

                if (-not ($service.Status.ToString() -eq "Running" -or
                ($monitor.Type -eq "Common" -and
                        $monitor.StartType -eq "Manual"))) {
                    if ($monitor.Type -eq "Critical") {
                        $notRunningCriticalServices.Add((NewServiceObject $service))
                    } else {
                        $notRunningCommonServices.Add((NewServiceObject $service))
                    }
                }
                try {
                    $startType = $service.StartType.ToString()
                    Write-Verbose "StartType set to $startType"

                    if ($startType -ne "Automatic") {
                        if ($monitor.StartType -eq "Manual" -and
                            $startType -eq "Manual") {
                            Write-Verbose "Good configuration"
                        } else {
                            $serviceObject = NewServiceObject $service
                            $serviceObject | Add-Member -MemberType NoteProperty -Name "CorrectStartType" -Value $monitor.StartType
                            $misconfiguredServices.Add($serviceObject)
                        }
                    }
                } catch {
                    Write-Verbose "Failed to convert StartType"
                    Invoke-CatchActions
                }
            }
            $getServicesList.Add((NewServiceObject $service))
        }
    } end {
        return [PSCustomObject]@{
            Services      = $getServicesList
            Monitor       = $monitorServicesList
            Misconfigured = $misconfiguredServices
            Critical      = $notRunningCriticalServices
            Common        = $notRunningCommonServices
        }
    }
}

function Get-ExchangeRegistryValues {
    [CmdletBinding()]
    param(
        [string]$MachineName,
        [ScriptBlock]$CatchActionFunction
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $baseParams = @{
        MachineName         = $MachineName
        CatchActionFunction = $CatchActionFunction
    }

    $ctsParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters"
        GetValue = "CtsProcessorAffinityPercentage"
    }

    $fipsParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy"
        GetValue = "Enabled"
    }

    $blockReplParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Replay\Parameters"
        GetValue = "DisableGranularReplication"
    }

    $disableAsyncParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15"
        GetValue = "DisableAsyncNotification"
    }

    $serializedDataSigningParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics"
        GetValue = "EnableSerializationDataSigning"
    }

    $installDirectoryParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
        GetValue = "MsiInstallPath"
    }

    $baseTypeCheckForDeserializationParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Diagnostics"
        GetValue = "DisableBaseTypeCheckForDeserialization"
    }

    $disablePreservationParams = $baseParams + @{
        SubKey    = "SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
        GetValue  = "DisablePreservation"
        ValueType = "String"
    }

    $fipFsDatabasePathParams = $baseParams + @{
        SubKey    = "SOFTWARE\Microsoft\ExchangeServer\v15\FIP-FS"
        GetValue  = "DatabasePath"
        ValueType = "String"
    }

    return [PSCustomObject]@{
        DisableBaseTypeCheckForDeserialization = [int](Get-RemoteRegistryValue @baseTypeCheckForDeserializationParams)
        CtsProcessorAffinityPercentage         = [int](Get-RemoteRegistryValue @ctsParams)
        FipsAlgorithmPolicyEnabled             = [int](Get-RemoteRegistryValue @fipsParams)
        DisableGranularReplication             = [int](Get-RemoteRegistryValue @blockReplParams)
        DisableAsyncNotification               = [int](Get-RemoteRegistryValue @disableAsyncParams)
        SerializedDataSigning                  = [int](Get-RemoteRegistryValue @serializedDataSigningParams)
        MsiInstallPath                         = [string](Get-RemoteRegistryValue @installDirectoryParams)
        DisablePreservation                    = [string](Get-RemoteRegistryValue @disablePreservationParams)
        FipFsDatabasePath                      = [string](Get-RemoteRegistryValue @fipFsDatabasePathParams)
    }
}



function Get-InternalTransportCertificateFromServer {
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param (
        [string]$ComputerName = $env:COMPUTERNAME,
        [Parameter(Mandatory = $false)]
        [ScriptBlock]$CatchActionFunction
    )

    <#
        Reads the certificate set as internal transport certificate (aka default SMTP certificate) from AD.
        The certificate is specified on a per-server base.

        Returns the X509Certificate2 object if we were able to query it from AD, otherwise it returns $null.
    #>

    try {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $organizationContainer = Get-OrganizationContainer
        $exchangeServerPath = ("CN=" + $($ComputerName.Split(".")[0]) + ",CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups," + $organizationContainer.distinguishedName)
        $exchangeServer = [ADSI]("LDAP://" + $exchangeServerPath)
        Write-Verbose "Exchange Server path: $($exchangeServerPath)"
        if ($null -ne $exchangeServer.msExchServerInternalTLSCert) {
            $certObject = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($exchangeServer.msExchServerInternalTLSCert)
            Write-Verbose ("Internal transport certificate on server: $($ComputerName) is: $($certObject.Thumbprint)")
        }
    } catch {
        Write-Verbose ("Unable to query the internal transport certificate - Exception: $($Error[0].Exception.Message)")
        Invoke-CatchActionError $CatchActionFunction
    }

    return $certObject
}

function Import-ExchangeCertificateFromRawData {
    [CmdletBinding()]
    param(
        [System.Object[]]$ExchangeCertificates
    )

    <#
        This helper function must be used if Serialization Data Signing is enabled, but the Auth Certificate
        which is configured has expired or isn't available on the system where the script runs.
        The 'Get-ExchangeCertificate' cmdlet fails to deserialize and so, only RawData (byte[]) will be returned.
        To workaround, we initialize the X509Certificate2 class and import the data by using the Import() method.
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $exchangeCertificatesList = New-Object 'System.Collections.Generic.List[object]'
    } process {
        if ($ExchangeCertificates.Count -ne 0) {
            Write-Verbose ("Going to process '$($ExchangeCertificates.Count )' Exchange certificates")

            foreach ($c in $ExchangeCertificates) {
                # Initialize X509Certificate2 class
                $certObject = New-Object 'System.Security.Cryptography.X509Certificates.X509Certificate2'
                # Use the Import() method to import byte[] RawData
                $certObject.Import($c.RawData)

                if ($null -ne $certObject.Thumbprint) {
                    Write-Verbose ("Certificate with thumbprint: $($certObject.Thumbprint) imported successfully")
                    $exchangeCertificatesList.Add($certObject)
                }
            }
        }
    } end {
        return $exchangeCertificatesList
    }
}

function Get-ExchangeServerCertificates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        function NewCertificateExclusionEntry {
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $IssuerOrSubjectPattern,
                [Parameter(Mandatory = $true)]
                [bool]
                $IsSelfSigned
            )

            return [PSCustomObject]@{
                IorSPattern  = $IssuerOrSubjectPattern
                IsSelfSigned = $IsSelfSigned
            }
        }

        function ShouldCertificateBeSkipped {
            [OutputType("System.Boolean")]
            param (
                [Parameter(Mandatory = $true)]
                [PSCustomObject]
                $Exclusions,
                [Parameter(Mandatory = $true)]
                [System.Security.Cryptography.X509Certificates.X509Certificate2]
                $Certificate
            )

            $certificateMatch = $Exclusions | Where-Object {
                ((($Certificate.Subject -match $_.IorSPattern) -or
                ($Certificate.Issuer -match $_.IorSPattern)) -and
                ($Certificate.IsSelfSigned -eq $_.IsSelfSigned))
            } | Select-Object -First 1

            if ($null -ne $certificateMatch) {
                return $certificateMatch.IsSelfSigned -eq $Certificate.IsSelfSigned
            }
            return $false
        }

        $certObject = New-Object 'System.Collections.Generic.List[object]'
    } process {
        try {
            Write-Verbose "Build certificate exclusion list"
            <#
                Add the certificates that should be excluded from the Exchange certificate check (we don't return an object for them)
                Exclude "MS-Organization-P2P-Access [YYYY]" certificate with one day lifetime on Azure hosted machines.
                See: What are the MS-Organization-P2P-Access certificates present on our Windows 10/11 devices?
                https://docs.microsoft.com/azure/active-directory/devices/faq
                Exclude "DC=Windows Azure CRP Certificate Generator" (TenantEncryptionCertificate)
                The certificates are built by the Azure fabric controller and passed to the Azure VM Agent.
                If you stop and start the VM every day, the fabric controller might create a new certificate.
                These certificates can be deleted. The Azure VM Agent re-creates certificates if needed.
                https://docs.microsoft.com/azure/virtual-machines/extensions/features-windows
            #>
            $certificatesToExclude = @(
                NewCertificateExclusionEntry "CN=MS-Organization-P2P-Access \[[12][0-9]{3}\]$" $false
                NewCertificateExclusionEntry "DC=Windows Azure CRP Certificate Generator" $true
            )
            Write-Verbose "Trying to receive certificates from Exchange server: $($Server)"
            $exchangeServerCertificates = Get-ExchangeCertificate -Server $Server -ErrorAction Stop

            Write-Verbose "Trying to query internal transport certificate from AD for this server"
            $internalTransportCertificate = Get-InternalTransportCertificateFromServer -ComputerName $Server -CatchActionFunction ${Function:Invoke-CatchActions}

            if ($null -ne $exchangeServerCertificates) {
                try {
                    $authConfig = Get-AuthConfig -ErrorAction Stop
                    $authConfigDetected = $true
                } catch {
                    $authConfigDetected = $false
                    Invoke-CatchActions
                }

                if ($null -ne $exchangeServerCertificates[0].Thumbprint) {
                    Write-Verbose "Deserialization of the Exchange certificate object was successful - nothing to do"
                } else {
                    Write-Verbose "Deserialization of the Exchange certificate failed - trying to import the certificate from raw data"
                    $exchangeServerCertificates = Import-ExchangeCertificateFromRawData -ExchangeCertificates $exchangeServerCertificates
                }

                foreach ($cert in $exchangeServerCertificates) {
                    $isInternalTransportCertificate = $false

                    try {
                        $certificateLifetime = ([System.Convert]::ToDateTime($cert.NotAfter, [System.Globalization.DateTimeFormatInfo]::InvariantInfo) - (Get-Date)).Days
                        $sanCertificateInfo = $false

                        $excludeCertificate = ShouldCertificateBeSkipped -Exclusions $certificatesToExclude -Certificate $cert

                        if ($excludeCertificate) {
                            Write-Verbose "Excluding certificate $($cert.Subject). Moving to next certificate"
                            continue
                        }

                        $currentErrors = $Error.Count
                        if ($null -ne $cert.DnsNameList -and
                            ($cert.DnsNameList).Count -gt 1) {
                            $sanCertificateInfo = $true
                            $certDnsNameList = $cert.DnsNameList
                        } elseif ($null -eq $cert.DnsNameList) {
                            $certDnsNameList = "None"
                        } else {
                            $certDnsNameList = $cert.DnsNameList
                        }
                        if ($currentErrors -lt $Error.Count) {
                            $i = 0
                            while ($i -lt ($Error.Count - $currentErrors)) {
                                Invoke-CatchActions $Error[$i]
                                $i++
                            }
                        }

                        if (($null -ne $internalTransportCertificate) -and
                            ($cert.Thumbprint -eq $internalTransportCertificate.Thumbprint)) {
                            $isInternalTransportCertificate = $true
                        }

                        if ($authConfigDetected) {
                            $isAuthConfigInfo = $false
                            $isNextAuthCertificate = $false

                            if ($cert.Thumbprint -eq $authConfig.CurrentCertificateThumbprint) {
                                $isAuthConfigInfo = $true
                            } elseif ($cert.Thumbprint -eq $authConfig.NextCertificateThumbprint) {
                                $isNextAuthCertificate = $true
                            }
                        } else {
                            $isAuthConfigInfo = "InvalidAuthConfig"
                            $isNextAuthCertificate = "InvalidAuthConfig"
                        }

                        if ([String]::IsNullOrEmpty($cert.FriendlyName)) {
                            $certFriendlyName = ($certDnsNameList[0]).ToString()
                        } else {
                            $certFriendlyName = $cert.FriendlyName
                        }

                        if ([String]::IsNullOrEmpty($cert.Status)) {
                            $certStatus = "Unknown"
                        } else {
                            $certStatus = ($cert.Status).ToString()
                        }

                        if ([String]::IsNullOrEmpty($cert.SignatureAlgorithm.FriendlyName)) {
                            $certSignatureAlgorithm = "Unknown"
                            $certSignatureHashAlgorithm = "Unknown"
                            $certSignatureHashAlgorithmSecure = 0
                        } else {
                            $certSignatureAlgorithm = $cert.SignatureAlgorithm.FriendlyName
                            <#
                                OID Table
                                https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpnap/a48b02b2-2a10-4eb0-bed4-1807a6d2f5ad
                                SignatureHashAlgorithmSecure = Unknown 0
                                SignatureHashAlgorithmSecure = Insecure/Weak 1
                                SignatureHashAlgorithmSecure = Secure 2
                            #>
                            switch ($cert.SignatureAlgorithm.Value) {
                                "1.2.840.113549.1.1.5" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.113549.1.1.4" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.10040.4.3" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.29" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.15" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.3" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.113549.1.1.2" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.113549.1.1.3" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.2" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.4" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.7.2.3.1" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.13" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.3.14.3.2.27" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "2.16.840.1.101.2.1.1.19" { $certSignatureHashAlgorithm = "mosaicSignature"; $certSignatureHashAlgorithmSecure = 0 }
                                "1.3.14.3.2.26" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.113549.2.5" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                                "2.16.840.1.101.3.4.2.1" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                                "2.16.840.1.101.3.4.2.2" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                                "2.16.840.1.101.3.4.2.3" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.113549.1.1.11" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.113549.1.1.12" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.113549.1.1.13" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.113549.1.1.10" { $certSignatureHashAlgorithm = "rsassa-pss"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.10045.4.1" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                                "1.2.840.10045.4.3.2" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.10045.4.3.3" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.10045.4.3.4" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                                "1.2.840.10045.4.3" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                                default { $certSignatureHashAlgorithm = "Unknown"; $certSignatureHashAlgorithmSecure = 0 }
                            }
                        }

                        $certObject.Add([PSCustomObject]@{
                                Issuer                         = $cert.Issuer
                                Subject                        = $cert.Subject
                                FriendlyName                   = $certFriendlyName
                                Thumbprint                     = $cert.Thumbprint
                                PublicKeySize                  = $cert.PublicKey.Key.KeySize
                                IsEccCertificate               = $cert.PublicKey.Oid.Value -eq "1.2.840.10045.2.1" # WellKnownOid for ECC
                                SignatureAlgorithm             = $certSignatureAlgorithm
                                SignatureHashAlgorithm         = $certSignatureHashAlgorithm
                                SignatureHashAlgorithmSecure   = $certSignatureHashAlgorithmSecure
                                IsSanCertificate               = $sanCertificateInfo
                                Namespaces                     = $certDnsNameList
                                Services                       = $cert.Services
                                IsInternalTransportCertificate = $isInternalTransportCertificate
                                IsCurrentAuthConfigCertificate = $isAuthConfigInfo
                                IsNextAuthConfigCertificate    = $isNextAuthCertificate
                                SetAsActiveAuthCertificateOn   = if ($isNextAuthCertificate) { $authConfig.NextCertificateEffectiveDate } else { $null }
                                LifetimeInDays                 = $certificateLifetime
                                Status                         = $certStatus
                                CertificateObject              = $cert
                            })
                    } catch {
                        Write-Verbose "Unable to process certificate: $($cert.Thumbprint)"
                        Invoke-CatchActions
                    }
                }
            }
        } catch {
            Write-Verbose "Failed to run 'Get-ExchangeCertificate' - Exception: $($Error[0].Exception)."
            Invoke-CatchActions
        }
    } end {
        if ($certObject.Count -ge 1) {
            Write-Verbose "Processed: $($certObject.Count) certificates"
        } else {
            Write-Verbose "Failed to find any Exchange certificates"
        }
        return $certObject
    }
}

function Get-ExchangeServerMaintenanceState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $false)]
        [array]$ComponentsToSkip
    )
    begin {
        Write-Verbose "Calling Function: $($MyInvocation.MyCommand)"
        $getClusterNode = $null
        $getServerComponentState = $null
        $inactiveComponents = @()
    } process {

        $getServerComponentState = Get-ServerComponentState -Identity $Server -ErrorAction SilentlyContinue

        try {
            $getClusterNode = Get-ClusterNode -Name $Server -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to run Get-ClusterNode"
            Invoke-CatchActions
        }

        Write-Verbose "Running ServerComponentStates checks"

        foreach ($component in $getServerComponentState) {
            if (($null -ne $ComponentsToSkip -and
                    $ComponentsToSkip.Count -ne 0) -and
                $ComponentsToSkip -notcontains $component.Component) {
                if ($component.State.ToString() -ne "Active") {
                    $latestLocalState = $null
                    $latestRemoteState = $null

                    if ($null -ne $component.LocalStates -and
                        $component.LocalStates.Count -gt 0) {
                        $latestLocalState = ($component.LocalStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                    }

                    if ($null -ne $component.RemoteStates -and
                        $component.RemoteStates.Count -gt 0) {
                        $latestRemoteState = ($component.RemoteStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                    }

                    Write-Verbose "Component: '$($component.Component)' LocalState: '$($latestLocalState.State)' RemoteState: '$($latestRemoteState.State)'"

                    if ($latestLocalState.State -eq $latestRemoteState.State) {
                        $inactiveComponents += "'{0}' is in Maintenance Mode" -f $component.Component
                    } else {
                        if (($null -ne $latestLocalState) -and
                        ($latestLocalState.State -ne "Active")) {
                            $inactiveComponents += "'{0}' is in Local Maintenance Mode only" -f $component.Component
                        }

                        if (($null -ne $latestRemoteState) -and
                        ($latestRemoteState.State -ne "Active")) {
                            $inactiveComponents += "'{0}' is in Remote Maintenance Mode only" -f $component.Component
                        }
                    }
                } else {
                    Write-Verbose "Component '$($component.Component)' is Active"
                }
            } else {
                Write-Verbose "Component: $($component.Component) will be skipped"
            }
        }
    } end {

        return [PSCustomObject]@{
            InactiveComponents      = [array]$inactiveComponents
            GetServerComponentState = $getServerComponentState
            GetClusterNode          = $getClusterNode
        }
    }
}

function Get-ExchangeUpdates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Exchange2013", "Exchange2016", "Exchange2019")]
        [string]$ExchangeMajorVersion
    )
    Write-Verbose("Calling: $($MyInvocation.MyCommand) Passed: $ExchangeMajorVersion")
    $RegLocation = [string]::Empty

    if ("Exchange2013" -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2013"
    } elseif ("Exchange2016" -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2016"
    } else {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2019"
    }

    $RegKey = Get-RemoteRegistrySubKey -MachineName $Server `
        -SubKey $RegLocation `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $RegKey) {
        $IU = $RegKey.GetSubKeyNames()
        if ($null -ne $IU) {
            Write-Verbose "Detected fixes installed on the server"
            $fixes = @()
            foreach ($key in $IU) {
                $IUKey = $RegKey.OpenSubKey($key)
                $IUName = $IUKey.GetValue("PackageName")
                Write-Verbose "Found: $IUName"
                $fixes += $IUName
            }
            return $fixes
        } else {
            Write-Verbose "No IUs found in the registry"
        }
    } else {
        Write-Verbose "No RegKey returned"
    }

    Write-Verbose "Exiting: Get-ExchangeUpdates"
    return $null
}


function Get-ExchangeVirtualDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $failedString = "Failed to get {0} virtual directory."
        $getActiveSyncVirtualDirectory = $null
        $getAutoDiscoverVirtualDirectory = $null
        $getEcpVirtualDirectory = $null
        $getMapiVirtualDirectory = $null
        $getOabVirtualDirectory = $null
        $getOutlookAnywhere = $null
        $getOwaVirtualDirectory = $null
        $getPowerShellVirtualDirectory = $null
        $getWebServicesVirtualDirectory = $null
        $paramsNoShow = @{
            Server           = $Server
            ErrorAction      = "Stop"
            ADPropertiesOnly = $true
        }
        $params = $paramsNoShow + @{
            ShowMailboxVirtualDirectories = $true
        }
    }
    process {
        try {
            $getActiveSyncVirtualDirectory = Get-ActiveSyncVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "EAS")
            Invoke-CatchActions
        }

        try {
            $getAutoDiscoverVirtualDirectory = Get-AutodiscoverVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "Autodiscover")
            Invoke-CatchActions
        }

        try {
            $getEcpVirtualDirectory = Get-EcpVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "ECP")
            Invoke-CatchActions
        }

        try {
            # Doesn't have ShowMailboxVirtualDirectories
            $getMapiVirtualDirectory = Get-MapiVirtualDirectory @paramsNoShow
        } catch {
            Write-Verbose ($failedString -f "Mapi")
            Invoke-CatchActions
        }

        try {
            $getOabVirtualDirectory = Get-OabVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "OAB")
            Invoke-CatchActions
        }

        try {
            $getOutlookAnywhere = Get-OutlookAnywhere @params
        } catch {
            Write-Verbose ($failedString -f "Outlook Anywhere")
            Invoke-CatchActions
        }

        try {
            $getOwaVirtualDirectory = Get-OwaVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "OWA")
            Invoke-CatchActions
        }

        try {
            $getPowerShellVirtualDirectory = Get-PowerShellVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "PowerShell")
            Invoke-CatchActions
        }

        try {
            $getWebServicesVirtualDirectory = Get-WebServicesVirtualDirectory @params
        } catch {
            Write-Verbose ($failedString -f "EWS")
            Invoke-CatchActions
        }
    }
    end {
        return [PSCustomObject]@{
            GetActiveSyncVirtualDirectory   = $getActiveSyncVirtualDirectory
            GetAutoDiscoverVirtualDirectory = $getAutoDiscoverVirtualDirectory
            GetEcpVirtualDirectory          = $getEcpVirtualDirectory
            GetMapiVirtualDirectory         = $getMapiVirtualDirectory
            GetOabVirtualDirectory          = $getOabVirtualDirectory
            GetOutlookAnywhere              = $getOutlookAnywhere
            GetOwaVirtualDirectory          = $getOwaVirtualDirectory
            GetPowerShellVirtualDirectory   = $getPowerShellVirtualDirectory
            GetWebServicesVirtualDirectory  = $getWebServicesVirtualDirectory
        }
    }
}


function Get-FIPFSScanEngineVersionState {
    [CmdletBinding()]
    [OutputType("System.Object")]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,
        [Parameter(Mandatory = $true)]
        [System.Version]
        $ExSetupVersion,
        [Parameter(Mandatory = $true)]
        [bool]
        $AffectedServerRole
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        function GetFolderFromExchangeInstallPath {
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $ExchangeSubDir
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                $exSetupPath = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction Stop).MsiInstallPath
            } catch {
                # since this is a script block, can't call Invoke-CatchActions
                $exSetupPath = $env:ExchangeInstallPath
            }

            $finalPath = Join-Path $exSetupPath $ExchangeSubDir

            if ($ExchangeSubDir -notmatch '\.[a-zA-Z0-9]+$') {

                if (Test-Path $finalPath) {
                    $getDir = Get-ChildItem -Path $finalPath -Attributes Directory
                }

                return ([PSCustomObject]@{
                        Name             = $getDir.Name
                        LastWriteTimeUtc = $getDir.LastWriteTimeUtc
                        Failed           = $null -eq $getDir
                    })
            }
            return $null
        }

        function GetHighestScanEngineVersionNumber {
            param (
                [string]
                $ComputerName
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"

            try {
                $scanEngineVersions = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
                    -ScriptBlock ${Function:GetFolderFromExchangeInstallPath} `
                    -ArgumentList ("FIP-FS\Data\Engines\amd64\Microsoft\Bin") `
                    -CatchActionFunction ${Function:Invoke-CatchActions}

                if ($null -ne $scanEngineVersions) {
                    if ($scanEngineVersions.Failed) {
                        Write-Verbose "Failed to find the scan engine directory"
                    } else {
                        return [Int64]($scanEngineVersions.Name | Measure-Object -Maximum).Maximum
                    }
                } else {
                    Write-Verbose "No FIP-FS scan engine version(s) detected - GetFolderFromExchangeInstallPath returned null"
                }
            } catch {
                Write-Verbose "Error occurred while processing FIP-FS scan engine version(s)"
                Invoke-CatchActions
            }
            return $null
        }

        function IsFIPFSFixedBuild {
            param (
                [System.Version]
                $BuildNumber
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"

            $fixedFIPFSBuild = $false

            # Fixed on Exchange side with March 2022 Security update
            if ($BuildNumber.Major -eq 15) {
                if ($BuildNumber.Minor -eq 2) {
                    $fixedFIPFSBuild = ($BuildNumber.Build -gt 986) -or
                        (($BuildNumber.Build -eq 986) -and ($BuildNumber.Revision -ge 22)) -or
                        (($BuildNumber.Build -eq 922) -and ($BuildNumber.Revision -ge 27))
                } elseif ($BuildNumber.Minor -eq 1) {
                    $fixedFIPFSBuild = ($BuildNumber.Build -gt 2375) -or
                        (($BuildNumber.Build -eq 2375) -and ($BuildNumber.Revision -ge 24)) -or
                        (($BuildNumber.Build -eq 2308) -and ($BuildNumber.Revision -ge 27))
                } else {
                    Write-Verbose "Looks like we're on Exchange 2013 which is not affected by this FIP-FS issue"
                    $fixedFIPFSBuild = $true
                }
            } else {
                Write-Verbose "We are not on Exchange version 15"
                $fixedFIPFSBuild = $true
            }

            return $fixedFIPFSBuild
        }
    } process {
        $isAffectedByFIPFSUpdateIssue = $false
        try {

            if ($AffectedServerRole) {
                $highestScanEngineVersionNumber = GetHighestScanEngineVersionNumber -ComputerName $ComputerName
                $fipFsIssueFixedBuild = IsFIPFSFixedBuild -BuildNumber $ExSetupVersion

                if ($null -eq $highestScanEngineVersionNumber) {
                    Write-Verbose "No scan engine version found on the computer - this can cause issues still with some transport rules"
                } elseif ($highestScanEngineVersionNumber -ge 2201010000) {
                    if ($fipFsIssueFixedBuild) {
                        Write-Verbose "Scan engine: $highestScanEngineVersionNumber detected but Exchange runs a fixed build that doesn't crash"
                    } else {
                        Write-Verbose "Scan engine: $highestScanEngineVersionNumber will cause transport queue or pattern update issues"
                    }
                    $isAffectedByFIPFSUpdateIssue = $true
                } else {
                    Write-Verbose "Scan engine: $highestScanEngineVersionNumber is safe to use"
                }
            }
        } catch {
            Write-Verbose "Failed to check for the FIP-FS update issue"
            Invoke-CatchActions
            return $null
        }
    } end {
        return [PSCustomObject]@{
            FIPFSFixedBuild              = $fipFsIssueFixedBuild
            ServerRoleAffected           = $AffectedServerRole
            HighestVersionNumberDetected = $highestScanEngineVersionNumber
            BadVersionNumberDirDetected  = $isAffectedByFIPFSUpdateIssue
        }
    }
}

function Get-ServerRole {
    param(
        [Parameter(Mandatory = $true)][object]$ExchangeServerObj
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $roles = $ExchangeServerObj.ServerRole.ToString()
    Write-Verbose "Roll: $roles"
    #Need to change this to like because of Exchange 2010 with AIO with the hub role.
    if ($roles -like "Mailbox, ClientAccess*") {
        return "MultiRole"
    } elseif ($roles -eq "Mailbox") {
        return "Mailbox"
    } elseif ($roles -eq "Edge") {
        return "Edge"
    } elseif ($roles -like "*ClientAccess*") {
        return "ClientAccess"
    } else {
        return "None"
    }
}
function Get-ExchangeInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $params = @{
            ComputerName           = $Server
            ScriptBlock            = { [environment]::OSVersion.Version -ge "10.0.0.0" }
            ScriptBlockDescription = "Windows 2016 or Greater Check"
            CatchActionFunction    = ${Function:Invoke-CatchActions}
        }
        $windows2016OrGreater = Invoke-ScriptBlockHandler @params
        $getExchangeServer = (Get-ExchangeServer -Identity $Server -Status)
        $exchangeCertificates = Get-ExchangeServerCertificates -Server $Server
        $exSetupDetails = Get-ExSetupFileVersionInfo -Server $Server -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -eq $exSetupDetails) {
            # couldn't find ExSetup.exe this should be rare so we are just going to handle this by displaying the AdminDisplayVersion from Get-ExchangeServer
            $versionInformation = (Get-ExchangeBuildVersionInformation -AdminDisplayVersion $getExchangeServer.AdminDisplayVersion)
            $exSetupDetails = [PSCustomObject]@{
                FileVersion      = $versionInformation.BuildVersion.ToString()
                FileBuildPart    = $versionInformation.BuildVersion.Build
                FilePrivatePart  = $versionInformation.BuildVersion.Revision
                FileMajorPart    = $versionInformation.BuildVersion.Major
                FileMinorPart    = $versionInformation.BuildVersion.Minor
                FailedGetExSetup = $true
            }
        } else {
            $versionInformation = (Get-ExchangeBuildVersionInformation -FileVersion ($exSetupDetails.FileVersion))
        }

        $buildInformation = [PSCustomObject]@{
            ServerRole         = (Get-ServerRole -ExchangeServerObj $getExchangeServer)
            MajorVersion       = $versionInformation.MajorVersion
            CU                 = $versionInformation.CU
            ExchangeSetup      = $exSetupDetails
            VersionInformation = $versionInformation
            KBsInstalled       = [array](Get-ExchangeUpdates -Server $Server -ExchangeMajorVersion $versionInformation.MajorVersion)
        }

        $dependentServices = (Get-ExchangeDependentServices -MachineName $Server)

        try {
            $getMailboxServer = (Get-MailboxServer -Identity $Server -ErrorAction Stop)
        } catch {
            Write-Verbose "Failed to run Get-MailboxServer"
            Invoke-CatchActions
        }

        $getExchangeVirtualDirectories = Get-ExchangeVirtualDirectories -Server $Server

        $registryValues = Get-ExchangeRegistryValues -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $serverExchangeBinDirectory = [System.Io.Path]::Combine($registryValues.MsiInstallPath, "Bin\")
        Write-Verbose "Found Exchange Bin: $serverExchangeBinDirectory"

        if ($getExchangeServer.IsEdgeServer -eq $false) {
            $applicationPools = Get-ExchangeAppPoolsInformation -Server $Server

            Write-Verbose "Query Exchange Connector settings via 'Get-ExchangeConnectors'"
            $exchangeConnectors = Get-ExchangeConnectors -ComputerName $Server -CertificateObject $exchangeCertificates

            $exchangeServerIISParams = @{
                ComputerName        = $Server
                IsLegacyOS          = ($windows2016OrGreater -eq $false)
                CatchActionFunction = ${Function:Invoke-CatchActions}
            }

            Write-Verbose "Trying to query Exchange Server IIS settings"
            $iisSettings = Get-ExchangeServerIISSettings @exchangeServerIISParams

            Write-Verbose "Query extended protection configuration for multiple CVEs testing"
            $getExtendedProtectionConfigurationParams = @{
                ComputerName        = $Server
                ExSetupVersion      = $buildInformation.ExchangeSetup.FileVersion
                CatchActionFunction = ${Function:Invoke-CatchActions}
            }

            try {
                if ($null -ne $iisSettings.ApplicationHostConfig) {
                    $getExtendedProtectionConfigurationParams.ApplicationHostConfig = [xml]$iisSettings.ApplicationHostConfig
                }
                Write-Verbose "Was able to convert the ApplicationHost.Config to XML"

                $extendedProtectionConfig = Get-ExtendedProtectionConfiguration @getExtendedProtectionConfigurationParams
            } catch {
                Write-Verbose "Failed to get the ExtendedProtectionConfig"
                Invoke-CatchActions
            }
        }

        $configParams = @{
            ComputerName = $Server
            FileLocation = @("$([System.IO.Path]::Combine($serverExchangeBinDirectory, "EdgeTransport.exe.config"))",
                "$([System.IO.Path]::Combine($serverExchangeBinDirectory, "Search\Ceres\Runtime\1.0\noderunner.exe.config"))")
        }

        if ($getExchangeServer.IsEdgeServer -eq $false -and
            (-not ([string]::IsNullOrEmpty($registryValues.FipFsDatabasePath)))) {
            $configParams.FileLocation += "$([System.IO.Path]::Combine($registryValues.FipFsDatabasePath, "Configuration.xml"))"
        }

        $getFileContentInformation = Get-FileContentInformation @configParams
        $applicationConfigFileStatus = @{}
        $fileContentInformation = @{}

        foreach ($key in $getFileContentInformation.Keys) {
            if ($key -like "*.exe.config") {
                $applicationConfigFileStatus.Add($key, $getFileContentInformation[$key])
            } else {
                $fileContentInformation.Add($key, $getFileContentInformation[$key])
            }
        }

        $serverMaintenance = Get-ExchangeServerMaintenanceState -Server $Server -ComponentsToSkip "ForwardSyncDaemon", "ProvisioningRps"
        $settingOverrides = Get-ExchangeSettingOverride -Server $Server -CatchActionFunction ${Function:Invoke-CatchActions}

        if (($getExchangeServer.IsMailboxServer) -or
        ($getExchangeServer.IsEdgeServer)) {
            try {
                $exchangeServicesNotRunning = @()
                $testServiceHealthResults = Test-ServiceHealth -Server $Server -ErrorAction Stop
                foreach ($notRunningService in $testServiceHealthResults.ServicesNotRunning) {
                    if ($exchangeServicesNotRunning -notcontains $notRunningService) {
                        $exchangeServicesNotRunning += $notRunningService
                    }
                }
            } catch {
                Write-Verbose "Failed to run Test-ServiceHealth"
                Invoke-CatchActions
            }
        }

        Write-Verbose "Checking if FIP-FS is affected by the pattern issue"
        $fipFsParams = @{
            ComputerName       = $Server
            ExSetupVersion     = $buildInformation.ExchangeSetup.FileVersion
            AffectedServerRole = $($getExchangeServer.IsMailboxServer -eq $true)
        }

        $FIPFSUpdateIssue = Get-FIPFSScanEngineVersionState @fipFsParams

        $eemsEndpointParams = @{
            ComputerName           = $Server
            ScriptBlockDescription = "Test EEMS pattern service connectivity"
            CatchActionFunction    = ${Function:Invoke-CatchActions}
            ArgumentList           = $getExchangeServer.InternetWebProxy
            ScriptBlock            = {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                if ($null -ne $args[0]) {
                    Write-Verbose "Proxy Server detected. Going to use: $($args[0])"
                    [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($args[0])
                    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
                    [System.Net.WebRequest]::DefaultWebProxy.BypassProxyOnLocal = $true
                }
                Invoke-WebRequest -Method Get -Uri "https://officeclient.microsoft.com/GetExchangeMitigations" -UseBasicParsing
            }
        }
        $eemsEndpointResults = Invoke-ScriptBlockHandler @eemsEndpointParams

        Write-Verbose "Checking AES256-CBC information protection readiness and configuration"
        $aes256CbcParams = @{
            Server             = $Server
            VersionInformation = $versionInformation
        }
        $aes256CbcDetails = Get-ExchangeAES256CBCDetails @aes256CbcParams

        Write-Verbose "Getting Exchange Diagnostic Information"
        $params = @{
            Server    = $Server
            Process   = "EdgeTransport"
            Component = "ResourceThrottling"
        }
        $edgeTransportResourceThrottling = Get-ExchangeDiagnosticInformation @params

        if ($getExchangeServer.IsEdgeServer -eq $false) {
            $params = @{
                ComputerName           = $Server
                ScriptBlockDescription = "Getting Exchange Server Members"
                CatchActionFunction    = ${Function:Invoke-CatchActions}
                ScriptBlock            = {
                    [PSCustomObject]@{
                        LocalGroupMember  =  (Get-LocalGroupMember -SID "S-1-5-32-544" -ErrorAction Stop)
                        ADGroupMembership = (Get-ADPrincipalGroupMembership (Get-ADComputer $env:COMPUTERNAME).DistinguishedName)
                    }
                }
            }
            $computerMembership = Invoke-ScriptBlockHandler @params
        }
    } end {

        Write-Verbose "Exiting: Get-ExchangeInformation"
        return [PSCustomObject]@{
            BuildInformation                         = $buildInformation
            GetExchangeServer                        = $getExchangeServer
            VirtualDirectories                       = $getExchangeVirtualDirectories
            GetMailboxServer                         = $getMailboxServer
            ExtendedProtectionConfig                 = $extendedProtectionConfig
            ExchangeConnectors                       = $exchangeConnectors
            ExchangeServicesNotRunning               = [array]$exchangeServicesNotRunning
            ApplicationPools                         = $applicationPools
            RegistryValues                           = $registryValues
            ServerMaintenance                        = $serverMaintenance
            ExchangeCertificates                     = [array]$exchangeCertificates
            ExchangeEmergencyMitigationServiceResult = $eemsEndpointResults
            EdgeTransportResourceThrottling          = $edgeTransportResourceThrottling # If we want to checkout other diagnosticInfo, we should create a new object here.
            ApplicationConfigFileStatus              = $applicationConfigFileStatus
            DependentServices                        = $dependentServices
            IISSettings                              = $iisSettings
            SettingOverrides                         = $settingOverrides
            FIPFSUpdateIssue                         = $FIPFSUpdateIssue
            AES256CBCInformation                     = $aes256CbcDetails
            FileContentInformation                   = $fileContentInformation
            ComputerMembership                       = $computerMembership
        }
    }
}





function Get-WmiObjectHandler {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '', Justification = 'This is what this function is for')]
    [CmdletBinding()]
    param(
        [string]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [string]
        $Class,

        [string]
        $Filter,

        [string]
        $Namespace,

        [ScriptBlock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | Class: '$Class' | Filter: '$Filter' | Namespace: '$Namespace'"

        $execute = @{
            ComputerName = $ComputerName
            Class        = $Class
            ErrorAction  = "Stop"
        }

        if (-not ([string]::IsNullOrEmpty($Filter))) {
            $execute.Add("Filter", $Filter)
        }

        if (-not ([string]::IsNullOrEmpty($Namespace))) {
            $execute.Add("Namespace", $Namespace)
        }
    }
    process {
        try {
            $wmi = Get-WmiObject @execute
            Write-Verbose "Return a value: $($null -ne $wmi)"
            return $wmi
        } catch {
            Write-Verbose "Failed to run Get-WmiObject on class '$class'"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
}
function Get-WmiObjectCriticalHandler {
    [CmdletBinding()]
    param(
        [string]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [string]
        $Class,

        [string]
        $Filter,

        [string]
        $Namespace,

        [ScriptBlock]
        $CatchActionFunction
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $params = @{
        ComputerName        = $ComputerName
        Class               = $Class
        Filter              = $Filter
        Namespace           = $Namespace
        CatchActionFunction = $CatchActionFunction
    }

    $wmi = Get-WmiObjectHandler @params

    if ($null -eq $wmi) {
        # Check for common issues that have been seen. If common issue, Write-Warning the re-throw the error up.

        if ($Error[0].Exception.ErrorCode -eq 0x800703FA) {
            Write-Verbose "Registry key marked for deletion."
            $message = "A registry key is marked for deletion that was attempted to read from for the cmdlet 'Get-WmiObject -Class $Class'.`r`n"
            $message += "`tThis error goes away after some time and/or a reboot of the computer. At that time you should be able to run Health Checker again."
            Write-Warning $message
        }

        # Grab the English version of hte message and/or the error code. Could get a different error code if service is not disabled.
        if ($Error[0].Exception.Message -like "The service cannot be started, either because it is disabled or because it has no enabled devices associated with it. *" -or
            $Error[0].Exception.ErrorCode -eq 0x80070422) {
            Write-Verbose "winMgmt service is disabled or not working."
            Write-Warning "The 'winMgmt' service appears to not be working correctly. Please make sure it is set to Automatic and in a running state. This script will fail unless this is working correctly."
        }

        Write-Error $($Error[0]) -ErrorAction Stop
    }

    return $wmi
}
function Get-ProcessorInformation {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $wmiObject = $null
        $processorName = [string]::Empty
        $maxClockSpeed = 0
        $numberOfLogicalCores = 0
        $numberOfPhysicalCores = 0
        $numberOfProcessors = 0
        $currentClockSpeed = 0
        $processorIsThrottled = $false
        $differentProcessorCoreCountDetected = $false
        $differentProcessorsDetected = $false
        $presentedProcessorCoreCount = 0
        $previousProcessor = $null
    }
    process {
        $wmiObject = @(Get-WmiObjectCriticalHandler -ComputerName $MachineName -Class "Win32_Processor" -CatchActionFunction $CatchActionFunction)
        $processorName = $wmiObject[0].Name
        $maxClockSpeed = $wmiObject[0].MaxClockSpeed
        Write-Verbose "Evaluating processor results"

        foreach ($processor in $wmiObject) {
            $numberOfPhysicalCores += $processor.NumberOfCores
            $numberOfLogicalCores += $processor.NumberOfLogicalProcessors
            $numberOfProcessors++

            if ($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed) {
                Write-Verbose "Processor is being throttled"
                $processorIsThrottled = $true
                $currentClockSpeed = $processor.CurrentClockSpeed
            }

            if ($null -ne $previousProcessor) {

                if ($processor.Name -ne $previousProcessor.Name -or
                    $processor.MaxClockSpeed -ne $previousProcessor.MaxClockSpeed) {
                    Write-Verbose "Different Processors are detected!!! This is an issue."
                    $differentProcessorsDetected = $true
                }

                if ($processor.NumberOfLogicalProcessors -ne $previousProcessor.NumberOfLogicalProcessors) {
                    Write-Verbose "Different Processor core count per processor socket detected. This is an issue."
                    $differentProcessorCoreCountDetected = $true
                }
            }
            $previousProcessor = $processor
        }

        $presentedProcessorCoreCount = Invoke-ScriptBlockHandler -ComputerName $MachineName `
            -ScriptBlock { [System.Environment]::ProcessorCount } `
            -ScriptBlockDescription "Trying to get the System.Environment ProcessorCount" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $presentedProcessorCoreCount) {
            Write-Verbose "Wasn't able to get Presented Processor Core Count on the Server. Setting to -1."
            $presentedProcessorCoreCount = -1
        }
    }
    end {
        Write-Verbose "PresentedProcessorCoreCount: $presentedProcessorCoreCount"
        Write-Verbose "NumberOfPhysicalCores: $numberOfPhysicalCores | NumberOfLogicalCores: $numberOfLogicalCores | NumberOfProcessors: $numberOfProcessors"
        Write-Verbose "ProcessorIsThrottled: $processorIsThrottled | CurrentClockSpeed: $currentClockSpeed"
        Write-Verbose "DifferentProcessorsDetected: $differentProcessorsDetected | DifferentProcessorCoreCountDetected: $differentProcessorCoreCountDetected"
        return [PSCustomObject]@{
            Name                                = $processorName
            MaxMegacyclesPerCore                = $maxClockSpeed
            NumberOfPhysicalCores               = $numberOfPhysicalCores
            NumberOfLogicalCores                = $numberOfLogicalCores
            NumberOfProcessors                  = $numberOfProcessors
            CurrentMegacyclesPerCore            = $currentClockSpeed
            ProcessorIsThrottled                = $processorIsThrottled
            DifferentProcessorsDetected         = $differentProcessorsDetected
            DifferentProcessorCoreCountDetected = $differentProcessorCoreCountDetected
            EnvironmentProcessorCount           = $presentedProcessorCoreCount
            ProcessorClassObject                = $wmiObject
        }
    }
}

function Get-ServerType {
    [CmdletBinding()]
    [OutputType("System.String")]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ServerType
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ServerType: $ServerType"
        $returnServerType = [string]::Empty
    }
    process {
        if ($ServerType -like "VMWare*") { $returnServerType = "VMware" }
        elseif ($ServerType -like "*Amazon EC2*") { $returnServerType = "AmazonEC2" }
        elseif ($ServerType -like "*Microsoft Corporation*") { $returnServerType = "HyperV" }
        elseif ($ServerType.Length -gt 0) { $returnServerType = "Physical" }
        else { $returnServerType = "Unknown" }
    }
    end {
        Write-Verbose "Returning: $returnServerType"
        return $returnServerType
    }
}
function Get-HardwareInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $system = Get-WmiObjectCriticalHandler -ComputerName $Server -Class "Win32_ComputerSystem" -CatchActionFunction ${Function:Invoke-CatchActions}
        $physicalMemory = Get-WmiObjectHandler -ComputerName $Server -Class "Win32_PhysicalMemory" -CatchActionFunction ${Function:Invoke-CatchActions}
        $processorInformation = Get-ProcessorInformation -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $totalMemory = 0

        if ($null -eq $physicalMemory) {
            Write-Verbose "Using memory from Win32_ComputerSystem class instead. This may cause memory calculation issues."
            $totalMemory = $system.TotalPhysicalMemory
        } else {
            foreach ($memory in $physicalMemory) {
                $totalMemory += $memory.Capacity
            }
        }
    } end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return [PSCustomObject]@{
            Manufacturer      = $system.Manufacturer
            ServerType        = (Get-ServerType -ServerType $system.Manufacturer)
            AutoPageFile      = $system.AutomaticManagedPagefile
            Model             = $system.Model
            System            = $system
            Processor         = $processorInformation
            TotalMemory       = $totalMemory
            MemoryInformation = [array]$physicalMemory
        }
    }
}



function Get-ServerRebootPending {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {

        function Get-PendingFileReboot {
            try {
                if ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" -Name PendingFileRenameOperations -ErrorAction Stop)) {
                    return $true
                }
                return $false
            } catch {
                throw
            }
        }

        function Get-UpdateExeVolatile {
            try {
                $updateExeVolatileProps = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\" -ErrorAction Stop
                if ($null -ne $updateExeVolatileProps -and $null -ne $updateExeVolatileProps.Flags) {
                    return $true
                }
                return $false
            } catch {
                throw
            }
        }

        function Get-PendingCCMReboot {
            try {
                return (Invoke-CimMethod -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending' -ErrorAction Stop)
            } catch {
                throw
            }
        }

        function Get-PathTestingReboot {
            param(
                [string]$TestingPath
            )

            return (Test-Path $TestingPath)
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $pendingRebootLocations = New-Object 'System.Collections.Generic.List[string]'
    }
    process {
        $pendingFileRenameOperationValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingFileReboot} `
            -ScriptBlockDescription "Get-PendingFileReboot" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $pendingFileRenameOperationValue) {
            $pendingFileRenameOperationValue = $false
        }

        $componentBasedServicingPendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Component Based Servicing Reboot Pending" `
            -ArgumentList "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" `
            -CatchActionFunction $CatchActionFunction

        $ccmReboot = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingCCMReboot} `
            -ScriptBlockDescription "Get-PendingSCCMReboot" `
            -CatchActionFunction $CatchActionFunction

        $autoUpdatePendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Auto Update Pending Reboot" `
            -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" `
            -CatchActionFunction $CatchActionFunction

        $updateExeVolatileValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-UpdateExeVolatile} `
            -ScriptBlockDescription "UpdateExeVolatile Reboot Pending" `
            -CatchActionFunction $CatchActionFunction

        $ccmRebootPending = $ccmReboot -and ($ccmReboot.RebootPending -or $ccmReboot.IsHardRebootPending)
        $pendingReboot = $ccmRebootPending -or $pendingFileRenameOperationValue -or $componentBasedServicingPendingRebootValue -or $autoUpdatePendingRebootValue -or $updateExeVolatileValue

        if ($ccmRebootPending) {
            Write-Verbose "RebootPending in CCM_ClientUtilities"
            $pendingRebootLocations.Add("CCM_ClientUtilities Showing Reboot Pending")
        }

        if ($pendingFileRenameOperationValue) {
            Write-Verbose "RebootPending at HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
            $pendingRebootLocations.Add("HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations")
        }

        if ($componentBasedServicingPendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            $pendingRebootLocations.Add("HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        }

        if ($autoUpdatePendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            $pendingRebootLocations.Add("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        }

        if ($updateExeVolatileValue) {
            Write-Verbose "RebootPending at HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\Flags"
            $pendingRebootLocations.Add("HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\Flags")
        }
    }
    end {
        return [PSCustomObject]@{
            PendingFileRenameOperations          = $pendingFileRenameOperationValue
            ComponentBasedServicingPendingReboot = $componentBasedServicingPendingRebootValue
            AutoUpdatePendingReboot              = $autoUpdatePendingRebootValue
            UpdateExeVolatileValue               = $updateExeVolatileValue
            CcmRebootPending                     = $ccmRebootPending
            PendingReboot                        = $pendingReboot
            PendingRebootLocations               = $pendingRebootLocations
        }
    }
}


function Get-AllTlsSettingsFromRegistry {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {

        function Get-TLSMemberValue {
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $GetKeyType,

                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter( Mandatory = $false)]
                [bool]
                $NullIsEnabled
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | GetKeyType: $GetKeyType | NullIsEnabled: $NullIsEnabled"
            switch ($GetKeyType) {
                "Enabled" {
                    return ($null -eq $KeyValue -and $NullIsEnabled) -or ($KeyValue -ne 0 -and $null -ne $KeyValue)
                }
                "DisabledByDefault" {
                    return $null -ne $KeyValue -and $KeyValue -ne 0
                }
            }
        }

        function Get-NETDefaultTLSValue {
            param(
                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter(Mandatory = $true)]
                [string]
                $NetVersion,

                [Parameter(Mandatory = $true)]
                [string]
                $KeyName
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | NetVersion: '$NetVersion' | KeyName: '$KeyName'"
            return $null -ne $KeyValue -and $KeyValue -eq 1
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - MachineName: '$MachineName'"
        $registryBase = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS {0}\{1}"
        $enabledKey = "Enabled"
        $disabledKey = "DisabledByDefault"
        $netRegistryBase = "SOFTWARE\{0}\.NETFramework\{1}"
        $allTlsObjects = [PSCustomObject]@{
            "TLS" = @{}
            "NET" = @{}
        }
    }
    process {
        foreach ($tlsVersion in @("1.0", "1.1", "1.2", "1.3")) {
            $registryServer = $registryBase -f $tlsVersion, "Server"
            $registryClient = $registryBase -f $tlsVersion, "Client"
            $baseParams = @{
                MachineName         = $MachineName
                CatchActionFunction = $CatchActionFunction
            }

            # Get the Enabled and DisabledByDefault values
            $serverEnabledValue = Get-RemoteRegistryValue @baseParams -SubKey $registryServer -GetValue $enabledKey
            $serverDisabledByDefaultValue = Get-RemoteRegistryValue @baseParams -SubKey $registryServer -GetValue $disabledKey
            $clientEnabledValue = Get-RemoteRegistryValue @baseParams -SubKey $registryClient -GetValue $enabledKey
            $clientDisabledByDefaultValue = Get-RemoteRegistryValue @baseParams -SubKey $registryClient -GetValue $disabledKey
            $serverEnabled = (Get-TLSMemberValue -GetKeyType $enabledKey -KeyValue $serverEnabledValue -NullIsEnabled ($tlsVersion -ne "1.3"))
            $serverDisabledByDefault = (Get-TLSMemberValue -GetKeyType $disabledKey -KeyValue $serverDisabledByDefaultValue)
            $clientEnabled = (Get-TLSMemberValue -GetKeyType $enabledKey -KeyValue $clientEnabledValue -NullIsEnabled ($tlsVersion -ne "1.3"))
            $clientDisabledByDefault = (Get-TLSMemberValue -GetKeyType $disabledKey -KeyValue $clientDisabledByDefaultValue)
            $disabled = $serverEnabled -eq $false -and ($serverDisabledByDefault -or $null -eq $serverDisabledByDefaultValue) -and
            $clientEnabled -eq $false -and ($clientDisabledByDefault -or $null -eq $clientDisabledByDefaultValue)
            $misconfigured = $serverEnabled -ne $clientEnabled -or $serverDisabledByDefault -ne $clientDisabledByDefault
            # only need to test server settings here, because $misconfigured will be set and will be the official status.
            # want to check for if Server is Disabled and Disabled By Default is not set or the reverse. This would be only part disabled
            # and not what we recommend on the blog post.
            $halfDisabled = ($serverEnabled -eq $false -and $serverDisabledByDefault -eq $false -and $null -ne $serverDisabledByDefaultValue) -or
                ($serverEnabled -and $serverDisabledByDefault)
            $configuration = "Enabled"

            if ($disabled) {
                Write-Verbose "TLS is Disabled"
                $configuration = "Disabled"
            }

            if ($halfDisabled) {
                Write-Verbose "TLS is only half disabled"
                $configuration = "Half Disabled"
            }

            if ($misconfigured) {
                Write-Verbose "TLS is misconfigured"
                $configuration = "Misconfigured"
            }

            $currentTLSObject = [PSCustomObject]@{
                TLSVersion                 = $tlsVersion
                "Server$enabledKey"        = $serverEnabled
                "Server$enabledKey`Value"  = $serverEnabledValue
                "Server$disabledKey"       = $serverDisabledByDefault
                "Server$disabledKey`Value" = $serverDisabledByDefaultValue
                "ServerRegistryPath"       = $registryServer
                "Client$enabledKey"        = $clientEnabled
                "Client$enabledKey`Value"  = $clientEnabledValue
                "Client$disabledKey"       = $clientDisabledByDefault
                "Client$disabledKey`Value" = $clientDisabledByDefaultValue
                "ClientRegistryPath"       = $registryClient
                "TLSVersionDisabled"       = $disabled
                "TLSMisconfigured"         = $misconfigured
                "TLSHalfDisabled"          = $halfDisabled
                "TLSConfiguration"         = $configuration
            }
            $allTlsObjects.TLS.Add($TlsVersion, $currentTLSObject)
        }

        foreach ($netVersion in @("v2.0.50727", "v4.0.30319")) {

            $msRegistryKey = $netRegistryBase -f "Microsoft", $netVersion
            $wowMsRegistryKey = $netRegistryBase -f "Wow6432Node\Microsoft", $netVersion

            $systemDefaultTlsVersionsValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $msRegistryKey `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $schUseStrongCryptoValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $msRegistryKey `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction
            $wowSystemDefaultTlsVersionsValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $wowMsRegistryKey `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $wowSchUseStrongCryptoValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $wowMsRegistryKey `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction

            $systemDefaultTlsVersions = (Get-NETDefaultTLSValue -KeyValue $SystemDefaultTlsVersionsValue -NetVersion $netVersion -KeyName "SystemDefaultTlsVersions")
            $wowSystemDefaultTlsVersions = (Get-NETDefaultTLSValue -KeyValue $wowSystemDefaultTlsVersionsValue -NetVersion $netVersion -KeyName "WowSystemDefaultTlsVersions")

            $currentNetTlsDefaultVersionObject = [PSCustomObject]@{
                NetVersion                       = $netVersion
                SystemDefaultTlsVersions         = $systemDefaultTlsVersions
                SystemDefaultTlsVersionsValue    = $systemDefaultTlsVersionsValue
                SchUseStrongCrypto               = (Get-NETDefaultTLSValue -KeyValue $schUseStrongCryptoValue -NetVersion $netVersion -KeyName "SchUseStrongCrypto")
                SchUseStrongCryptoValue          = $schUseStrongCryptoValue
                MicrosoftRegistryLocation        = $msRegistryKey
                WowSystemDefaultTlsVersions      = $wowSystemDefaultTlsVersions
                WowSystemDefaultTlsVersionsValue = $wowSystemDefaultTlsVersionsValue
                WowSchUseStrongCrypto            = (Get-NETDefaultTLSValue -KeyValue $wowSchUseStrongCryptoValue -NetVersion $netVersion -KeyName "WowSchUseStrongCrypto")
                WowSchUseStrongCryptoValue       = $wowSchUseStrongCryptoValue
                WowRegistryLocation              = $wowMsRegistryKey
                SDtvConfiguredCorrectly          = $systemDefaultTlsVersions -eq $wowSystemDefaultTlsVersions
                SDtvEnabled                      = $systemDefaultTlsVersions -and $wowSystemDefaultTlsVersions
            }

            $hashKeyName = "NET{0}" -f ($netVersion.Split(".")[0])
            $allTlsObjects.NET.Add($hashKeyName, $currentNetTlsDefaultVersionObject)
        }
        return $allTlsObjects
    }
}


function Get-TlsCipherSuiteInformation {
    [OutputType("System.Object")]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )

    begin {

        function GetProtocolNames {
            param(
                [int[]]$Protocol
            )
            $protocolNames = New-Object System.Collections.Generic.List[string]

            foreach ($p in $Protocol) {
                $name = [string]::Empty

                if ($p -eq 2) { $name = "SSL_2_0" }
                elseif ($p -eq 768) { $name = "SSL_3_0" }
                elseif ($p -eq 769) { $name = "TLS_1_0" }
                elseif ($p -eq 770) { $name = "TLS_1_1" }
                elseif ($p -eq 771) { $name = "TLS_1_2" }
                elseif ($p -eq 772) { $name = "TLS_1_3" }
                elseif ($p -eq 32528) { $name = "TLS_1_3_DRAFT_16" }
                elseif ($p -eq 32530) { $name = "TLS_1_3_DRAFT_18" }
                elseif ($p -eq 65279) { $name = "DTLS_1_0" }
                elseif ($p -eq 65277) { $name = "DTLS_1_1" }
                else {
                    Write-Verbose "Unable to determine protocol $p"
                    $name = $p
                }

                $protocolNames.Add($name)
            }
            return [string]::Join(" & ", $protocolNames)
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $tlsCipherReturnObject = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        # 'Get-TlsCipherSuite' takes account of the cipher suites which are configured by the help of GPO.
        # No need to query the ciphers defined via GPO if this call is successful.
        Write-Verbose "Trying to query TlsCipherSuites via 'Get-TlsCipherSuite'"
        $getTlsCipherSuiteParams = @{
            ComputerName        = $MachineName
            ScriptBlock         = { Get-TlsCipherSuite }
            CatchActionFunction = $CatchActionFunction
        }
        $tlsCipherSuites = Invoke-ScriptBlockHandler @getTlsCipherSuiteParams

        if ($null -eq $tlsCipherSuites) {
            # If we can't get the ciphers via cmdlet, we need to query them via registry call and need to check
            # if ciphers suites are defined via GPO as well. If there are some, these take precedence over what
            # is in the default location.
            Write-Verbose "Failed to query TlsCipherSuites via 'Get-TlsCipherSuite' fallback to registry"

            $policyTlsRegistryParams = @{
                MachineName         = $MachineName
                SubKey              = "SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002"
                GetValue            = "Functions"
                ValueType           = "String"
                CatchActionFunction = $CatchActionFunction
            }

            Write-Verbose "Trying to query cipher suites configured via GPO from registry"
            $policyDefinedCiphers = Get-RemoteRegistryValue @policyTlsRegistryParams

            if ($null -ne $policyDefinedCiphers) {
                Write-Verbose "Ciphers specified via GPO found - these take precedence over what is in the default location"
                $tlsCipherSuites = $policyDefinedCiphers.Split(",")
            } else {
                Write-Verbose "No cipher suites configured via GPO found - going to query the local TLS cipher suites"
                $tlsRegistryParams = @{
                    MachineName         = $MachineName
                    SubKey              = "SYSTEM\CurrentControlSet\Control\Cryptography\Configuration\Local\SSL\00010002"
                    GetValue            = "Functions"
                    ValueType           = "MultiString"
                    CatchActionFunction = $CatchActionFunction
                }

                $tlsCipherSuites = Get-RemoteRegistryValue @tlsRegistryParams
            }
        }

        if ($null -ne $tlsCipherSuites) {
            foreach ($cipher in $tlsCipherSuites) {
                $tlsCipherReturnObject.Add([PSCustomObject]@{
                        Name        = if ($null -eq $cipher.Name) { $cipher } else { $cipher.Name }
                        CipherSuite = if ($null -eq $cipher.CipherSuite) { "N/A" } else { $cipher.CipherSuite }
                        Cipher      = if ($null -eq $cipher.Cipher) { "N/A" } else { $cipher.Cipher }
                        Certificate = if ($null -eq $cipher.Certificate) { "N/A" } else { $cipher.Certificate }
                        Protocols   = if ($null -eq $cipher.Protocols) { "N/A" } else { (GetProtocolNames $cipher.Protocols) }
                    })
            }
        }
    }
    end {
        return $tlsCipherReturnObject
    }
}

# Gets all related TLS Settings, from registry or other factors
function Get-AllTlsSettings {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    }
    process {
        return [PSCustomObject]@{
            Registry         = (Get-AllTlsSettingsFromRegistry -MachineName $MachineName -CatchActionFunction $CatchActionFunction)
            SecurityProtocol = (Invoke-ScriptBlockHandler -ComputerName $MachineName -ScriptBlock { ([System.Net.ServicePointManager]::SecurityProtocol).ToString() } -CatchActionFunction $CatchActionFunction)
            TlsCipherSuite   = (Get-TlsCipherSuiteInformation -MachineName $MachineName -CatchActionFunction $CatchActionFunction)
        }
    }
}
function Get-AllNicInformation {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$ComputerFQDN,
        [ScriptBlock]$CatchActionFunction
    )
    begin {

        # Extract for Pester Testing - Start
        function Get-NicPnpCapabilitiesSetting {
            [CmdletBinding()]
            param(
                [ValidateNotNullOrEmpty()]
                [string]$NicAdapterComponentId
            )
            begin {
                $nicAdapterBasicPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"
                [int]$i = 0
                Write-Verbose "Probing started to detect NIC adapter registry path"
            }
            process {
                $registrySubKey = Get-RemoteRegistrySubKey -MachineName $ComputerName -SubKey $nicAdapterBasicPath
                if ($null -ne $registrySubKey) {
                    $optionalKeys = $registrySubKey.GetSubKeyNames() | Where-Object { $_ -like "0*" }
                    do {
                        $nicAdapterPnPCapabilitiesProbingKey = "$nicAdapterBasicPath\$($optionalKeys[$i])"
                        $netCfgRemoteRegistryParams = @{
                            MachineName         = $ComputerName
                            SubKey              = $nicAdapterPnPCapabilitiesProbingKey
                            GetValue            = "NetCfgInstanceId"
                            CatchActionFunction = $CatchActionFunction
                        }
                        $netCfgInstanceId = Get-RemoteRegistryValue @netCfgRemoteRegistryParams

                        if ($netCfgInstanceId -eq $NicAdapterComponentId) {
                            Write-Verbose "Matching ComponentId found - now checking for PnPCapabilitiesValue"
                            $pnpRemoteRegistryParams = @{
                                MachineName         = $ComputerName
                                SubKey              = $nicAdapterPnPCapabilitiesProbingKey
                                GetValue            = "PnPCapabilities"
                                CatchActionFunction = $CatchActionFunction
                            }
                            $nicAdapterPnPCapabilitiesValue = Get-RemoteRegistryValue @pnpRemoteRegistryParams
                            break
                        } else {
                            Write-Verbose "No matching ComponentId found"
                            $i++
                        }
                    } while ($i -lt $optionalKeys.Count)
                }
            }
            end {
                return [PSCustomObject]@{
                    PnPCapabilities   = $nicAdapterPnPCapabilitiesValue
                    SleepyNicDisabled = ($nicAdapterPnPCapabilitiesValue -eq 24 -or $nicAdapterPnPCapabilitiesValue -eq 280)
                }
            }
        }

        # Extract for Pester Testing - End

        function Get-NetworkConfiguration {
            [CmdletBinding()]
            param(
                [string]$ComputerName
            )
            begin {
                $currentErrors = $Error.Count
                $params = @{
                    ErrorAction = "Stop"
                }
            }
            process {
                try {
                    if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {
                        $cimSession = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
                        $params.Add("CimSession", $cimSession)
                    }
                    $networkIpConfiguration = Get-NetIPConfiguration @params | Where-Object { $_.NetAdapter.MediaConnectionState -eq "Connected" }
                    Invoke-CatchActionErrorLoop -CurrentErrors $currentErrors -CatchActionFunction $CatchActionFunction
                    return $networkIpConfiguration
                } catch {
                    Write-Verbose "Failed to run Get-NetIPConfiguration. Error $($_.Exception)"
                    #just rethrow as caller will handle the catch
                    throw
                }
            }
        }

        function Get-NicInformation {
            [CmdletBinding()]
            param(
                [array]$NetworkConfiguration,
                [bool]$WmiObject
            )
            begin {

                function Get-IpvAddresses {
                    return [PSCustomObject]@{
                        Address        = ([string]::Empty)
                        Subnet         = ([string]::Empty)
                        DefaultGateway = ([string]::Empty)
                    }
                }

                if ($null -eq $NetworkConfiguration) {
                    Write-Verbose "NetworkConfiguration are null in New-NicInformation. Returning a null object."
                    return $null
                }

                $nicObjects = New-Object 'System.Collections.Generic.List[object]'
            }
            process {
                if ($WmiObject) {
                    $networkAdapterConfigurationsParams = @{
                        ComputerName        = $ComputerName
                        Class               = "Win32_NetworkAdapterConfiguration"
                        Filter              = "IPEnabled = True"
                        CatchActionFunction = $CatchActionFunction
                    }
                    $networkAdapterConfigurations = Get-WmiObjectHandler @networkAdapterConfigurationsParams
                }

                foreach ($networkConfig in $NetworkConfiguration) {
                    $dnsClient = $null
                    $rssEnabledValue = 2
                    $netAdapterRss = $null
                    $mtuSize = 0
                    $driverDate = [DateTime]::MaxValue
                    $driverVersion = [string]::Empty
                    $description = [string]::Empty
                    $ipv4Address = @()
                    $ipv6Address = @()
                    $ipv6Enabled = $false
                    $isRegisteredInDns = $false
                    $dnsServerToBeUsed = $null

                    if (-not ($WmiObject)) {
                        Write-Verbose "Working on NIC: $($networkConfig.InterfaceDescription)"
                        $adapter = $networkConfig.NetAdapter

                        if ($adapter.DriverFileName -ne "NdIsImPlatform.sys") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.DeviceID
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        try {
                            $dnsClient = $adapter | Get-DnsClient -ErrorAction Stop
                            $isRegisteredInDns = $dnsClient.RegisterThisConnectionsAddress
                            Write-Verbose "Got DNS Client information"
                        } catch {
                            Write-Verbose "Failed to get the DNS client information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        try {
                            $netAdapterRss = $adapter | Get-NetAdapterRss -ErrorAction Stop
                            Write-Verbose "Got Net Adapter RSS Information"

                            if ($null -ne $netAdapterRss) {
                                [int]$rssEnabledValue = $netAdapterRss.Enabled
                            }
                        } catch {
                            Write-Verbose "Failed to get RSS Information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        foreach ($ipAddress in $networkConfig.AllIPAddresses.IPAddress) {
                            if ($ipAddress.Contains(":")) {
                                $ipv6Enabled = $true
                            }
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv4Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv4Address -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv4Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv4Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv4DefaultGateway -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv4DefaultGateway[$i].NextHop
                            }
                            $ipv4Address += $newIpvAddress
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv6Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv6Address -and
                                $i -lt $networkConfig.IPv6Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv6Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv6Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv6DefaultGateway -and
                                $i -lt $networkConfig.IPv6DefaultGateway.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv6DefaultGateway[$i].NextHop
                            }
                            $ipv6Address += $newIpvAddress
                        }

                        $mtuSize = $adapter.MTUSize
                        $driverDate = $adapter.DriverDate
                        $driverVersion = $adapter.DriverVersionString
                        $description = $adapter.InterfaceDescription
                        $dnsServerToBeUsed = $networkConfig.DNSServer.ServerAddresses
                    } else {
                        Write-Verbose "Working on NIC: $($networkConfig.Description)"
                        $adapter = $networkConfig
                        $description = $adapter.Description

                        if ($adapter.ServiceName -ne "NdIsImPlatformMp") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.Guid
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        #set the correct $adapterConfiguration to link to the correct $networkConfig that we are on
                        $adapterConfiguration = $networkAdapterConfigurations |
                            Where-Object { $_.SettingID -eq $networkConfig.GUID -or
                                $_.SettingID -eq $networkConfig.InterfaceGuid }

                        if ($null -eq $adapterConfiguration) {
                            Write-Verbose "Failed to find correct adapterConfiguration for this networkConfig."
                            Write-Verbose "GUID: $($networkConfig.GUID) | InterfaceGuid: $($networkConfig.InterfaceGuid)"
                        } else {
                            $ipv6Enabled = ($adapterConfiguration.IPAddress | Where-Object { $_.Contains(":") }).Count -ge 1

                            if ($null -ne $adapterConfiguration.DefaultIPGateway) {
                                $ipv4Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(".") }
                                $ipv6Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(":") }
                            } else {
                                $ipv4Gateway = "No default IPv4 gateway set"
                                $ipv6Gateway = "No default IPv6 gateway set"
                            }

                            for ($i = 0; $i -lt $adapterConfiguration.IPAddress.Count; $i++) {

                                if ($adapterConfiguration.IPAddress[$i].Contains(":")) {
                                    $newIpv6Address = Get-IpvAddresses
                                    if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                        $newIpv6Address.Address = $adapterConfiguration.IPAddress[$i]
                                        $newIpv6Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                    }

                                    $newIpv6Address.DefaultGateway = $ipv6Gateway
                                    $ipv6Address += $newIpv6Address
                                } else {
                                    $newIpv4Address = Get-IpvAddresses
                                    if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                        $newIpv4Address.Address = $adapterConfiguration.IPAddress[$i]
                                        $newIpv4Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                    }

                                    $newIpv4Address.DefaultGateway = $ipv4Gateway
                                    $ipv4Address += $newIpv4Address
                                }
                            }

                            $isRegisteredInDns = $adapterConfiguration.FullDNSRegistrationEnabled
                            $dnsServerToBeUsed = $adapterConfiguration.DNSServerSearchOrder
                        }
                    }

                    $nicObjects.Add([PSCustomObject]@{
                            WmiObject         = $WmiObject
                            Name              = $adapter.Name
                            LinkSpeed         = ((($adapter.Speed) / 1000000).ToString() + " Mbps")
                            DriverDate        = $driverDate
                            NetAdapterRss     = $netAdapterRss
                            RssEnabledValue   = $rssEnabledValue
                            IPv6Enabled       = $ipv6Enabled
                            Description       = $description
                            DriverVersion     = $driverVersion
                            MTUSize           = $mtuSize
                            PnPCapabilities   = $nicPnpCapabilitiesSetting.PnpCapabilities
                            SleepyNicDisabled = $nicPnpCapabilitiesSetting.SleepyNicDisabled
                            IPv4Addresses     = $ipv4Address
                            IPv6Addresses     = $ipv6Address
                            RegisteredInDns   = $isRegisteredInDns
                            DnsServer         = $dnsServerToBeUsed
                            DnsClient         = $dnsClient
                        })
                }
            }
            end {
                Write-Verbose "Found $($nicObjects.Count) active adapters on the computer."
                Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
                return $nicObjects
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | ComputerFQDN: '$ComputerFQDN'"
    }
    process {
        try {
            try {
                $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerName
            } catch {
                Invoke-CatchActionError $CatchActionFunction

                try {
                    if (-not ([string]::IsNullOrEmpty($ComputerFQDN))) {
                        $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerFQDN
                    } else {
                        $bypassCatchActions = $true
                        Write-Verbose "No FQDN was passed, going to rethrow error."
                        throw
                    }
                } catch {
                    #Just throw again
                    throw
                }
            }

            if ([String]::IsNullOrEmpty($networkConfiguration)) {
                # Throw if nothing was returned by previous calls.
                # Can be caused when executed on Server 2008 R2 where CIM namespace ROOT/StandardCiMv2 is invalid.
                Write-Verbose "No value was returned by 'Get-NetworkConfiguration'. Fallback to WMI."
                throw
            }

            return (Get-NicInformation -NetworkConfiguration $networkConfiguration)
        } catch {
            if (-not $bypassCatchActions) {
                Invoke-CatchActionError $CatchActionFunction
            }

            $wmiNetworkCardsParams = @{
                ComputerName        = $ComputerName
                Class               = "Win32_NetworkAdapter"
                Filter              = "NetConnectionStatus ='2'"
                CatchActionFunction = $CatchActionFunction
            }
            $wmiNetworkCards = Get-WmiObjectHandler @wmiNetworkCardsParams

            return (Get-NicInformation -NetworkConfiguration $wmiNetworkCards -WmiObject $true)
        }
    }
}


function Get-DotNetDllFileVersions {
    [CmdletBinding()]
    [OutputType("System.Collections.Hashtable")]
    param(
        [string]$ComputerName,
        [array]$FileNames,
        [ScriptBlock]$CatchActionFunction
    )

    begin {
        function Invoke-ScriptBlockGetItem {
            param(
                [string]$FilePath
            )
            $getItem = Get-Item $FilePath

            $returnObject = ([PSCustomObject]@{
                    GetItem          = $getItem
                    LastWriteTimeUtc = $getItem.LastWriteTimeUtc
                    VersionInfo      = ([PSCustomObject]@{
                            FileMajorPart   = $getItem.VersionInfo.FileMajorPart
                            FileMinorPart   = $getItem.VersionInfo.FileMinorPart
                            FileBuildPart   = $getItem.VersionInfo.FileBuildPart
                            FilePrivatePart = $getItem.VersionInfo.FilePrivatePart
                        })
                })

            return $returnObject
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $dotNetInstallPath = [string]::Empty
        $files = @{}
    }
    process {
        $dotNetInstallPath = Get-RemoteRegistryValue -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
            -GetValue "InstallPath" `
            -CatchActionFunction $CatchActionFunction

        if ([string]::IsNullOrEmpty($dotNetInstallPath)) {
            Write-Verbose "Failed to determine .NET install path"
            return
        }

        foreach ($fileName in $FileNames) {
            Write-Verbose "Querying for .NET DLL File $fileName"
            $getItem = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
                -ScriptBlock ${Function:Invoke-ScriptBlockGetItem} `
                -ArgumentList ("{0}\{1}" -f $dotNetInstallPath, $filename) `
                -CatchActionFunction $CatchActionFunction
            $files.Add($fileName, $getItem)
        }
    }
    end {
        return $files
    }
}

function Get-NETFrameworkInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $params = @{
            ComputerName        = $Server
            FileNames           = @("System.Data.dll", "System.Configuration.dll")
            CatchActionFunction = ${Function:Invoke-CatchActions}
        }
        $fileInformation = Get-DotNetDllFileVersions @params
        $netFramework = Get-NETFrameworkVersion -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
    } end {
        return [PSCustomObject]@{
            MajorVersion    = $netFramework.MinimumValue
            RegistryValue   = $netFramework.RegistryValue
            FriendlyName    = $netFramework.FriendlyName
            FileInformation = $fileInformation
        }
    }
}


function Get-HttpProxySetting {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function GetWinHttpSettings {
        param(
            [Parameter(Mandatory = $true)][string]$RegistryLocation
        )
        $connections = Get-ItemProperty -Path $RegistryLocation
        $byteLength = 4
        $proxyStartLocation = 16
        $proxyLength = 0
        $proxyAddress = [string]::Empty
        $byPassList = [string]::Empty

        if (($null -ne $connections) -and
            ($Connections | Get-Member).Name -contains "WinHttpSettings") {
            try {
                $bytes = $Connections.WinHttpSettings
                $proxyLength = [System.BitConverter]::ToInt32($bytes, $proxyStartLocation - $byteLength)

                if ($proxyLength -gt 0) {
                    $proxyAddress = [System.Text.Encoding]::UTF8.GetString($bytes, $proxyStartLocation, $proxyLength)
                    $byPassListLength = [System.BitConverter]::ToInt32($bytes, $proxyStartLocation + $proxyLength)

                    if ($byPassListLength -gt 0) {
                        $byPassList = [System.Text.Encoding]::UTF8.GetString($bytes, $byteLength + $proxyStartLocation + $proxyLength, $byPassListLength)
                    }
                }
            } catch {
                Write-Verbose "Failed to properly get HTTP Proxy information. Inner Exception: $_"
            }
        }

        return [PSCustomObject]@{
            ProxyAddress = $(if ($proxyAddress -eq [string]::Empty) { "None" } else { $proxyAddress })
            ByPassList   = $byPassList
        }
    }

    $httpProxy32 = Invoke-ScriptBlockHandler -ComputerName $Server `
        -ScriptBlock ${Function:GetWinHttpSettings} `
        -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" `
        -ScriptBlockDescription "Getting 32 Http Proxy Value" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $httpProxy64 = Invoke-ScriptBlockHandler -ComputerName $Server `
        -ScriptBlock ${Function:GetWinHttpSettings} `
        -ArgumentList "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" `
        -ScriptBlockDescription "Getting 64 Http Proxy Value" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $httpProxy = [PSCustomObject]@{
        ProxyAddress         = $(if ($httpProxy32.ProxyAddress -ne "None") { $httpProxy32.ProxyAddress } else { $httpProxy64.ProxyAddress })
        ByPassList           = $(if ($httpProxy32.ByPassList -ne [string]::Empty) { $httpProxy32.ByPassList } else { $httpProxy64.ByPassList })
        HttpProxyDifference  = $httpProxy32.ProxyAddress -ne $httpProxy64.ProxyAddress
        HttpByPassDifference = $httpProxy32.ByPassList -ne $httpProxy64.ByPassList
        HttpProxy32          = $httpProxy32
        HttpProxy64          = $httpProxy64
    }

    Write-Verbose "Http Proxy 32: $($httpProxy32.ProxyAddress)"
    Write-Verbose "Http By Pass List 32: $($httpProxy32.ByPassList)"
    Write-Verbose "Http Proxy 64: $($httpProxy64.ProxyAddress)"
    Write-Verbose "Http By Pass List 64: $($httpProxy64.ByPassList)"
    Write-Verbose "Proxy Address: $($httpProxy.ProxyAddress)"
    Write-Verbose "By Pass List: $($httpProxy.ByPassList)"
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $httpProxy
}

function Get-NetworkingInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $ipv6DisabledOnNICs = $false
    } process {
        $httpProxy = Get-HttpProxySetting -Server $Server
        $packetsReceivedDiscarded = (Get-LocalizedCounterSamples -MachineName $Server -Counter "\Network Interface(*)\Packets Received Discarded")
        $networkAdapters = @(Get-AllNicInformation -ComputerName $Server -CatchActionFunction ${Function:Invoke-CatchActions} -ComputerFQDN $ServerFQDN)

        foreach ($adapter in $networkAdapters) {
            if (-not ($adapter.IPv6Enabled)) {
                $ipv6DisabledOnNICs = $true
                break
            }
        }
    } end {
        return [PSCustomObject]@{
            HttpProxy                = $httpProxy
            PacketsReceivedDiscarded = $packetsReceivedDiscarded
            NetworkAdapters          = [array]$networkAdapters
            IPv6DisabledOnNICs       = $ipv6DisabledOnNICs
        }
    }
}


function Get-ServerOperatingSystemVersion {
    [CmdletBinding()]
    [OutputType("System.Object")]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,

        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $osReturnValue = [string]::Empty
        $baseParams = @{
            MachineName         = $ComputerName
            CatchActionFunction = $CatchActionFunction
        }

        # Get ProductName via registry call as this is more accurate when running on Server Core
        $productNameParams = $baseParams + @{
            SubKey   = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
            GetValue = "ProductName"
        }

        # Find out if we're running on Server Core to output on the 'Operating System Information' page
        $installationTypeParams = $baseParams + @{
            SubKey   = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
            GetValue = "InstallationType"
        }
    }
    process {
        Write-Verbose "Getting the version build information for computer: $ComputerName"
        $osCaption = Get-RemoteRegistryValue @productNameParams
        $installationType = Get-RemoteRegistryValue @installationTypeParams
        Write-Verbose "OsCaption: '$osCaption' InstallationType: '$installationType'"

        switch -Wildcard ($osCaption) {
            "*Server 2008 R2*" { $osReturnValue = "Windows2008R2"; break }
            "*Server 2008*" { $osReturnValue = "Windows2008" }
            "*Server 2012 R2*" { $osReturnValue = "Windows2012R2"; break }
            "*Server 2012*" { $osReturnValue = "Windows2012" }
            "*Server 2016*" { $osReturnValue = "Windows2016" }
            "*Server 2019*" { $osReturnValue = "Windows2019" }
            "*Server 2022*" { $osReturnValue = "Windows2022" }
            default { $osReturnValue = "Unknown" }
        }
    }
    end {
        Write-Verbose "OsReturnValue: '$osReturnValue'"
        return [PSCustomObject]@{
            MajorVersion     = $osReturnValue
            InstallationType = $installationType
            FriendlyName     = if ($installationType -eq "Server Core") { "$osCaption ($installationType)" } else { $osCaption }
        }
    }
}

function Get-OperatingSystemBuildInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $win32_OperatingSystem = Get-WmiObjectCriticalHandler -ComputerName $Server -Class Win32_OperatingSystem -CatchActionFunction ${Function:Invoke-CatchActions}
        $serverOsVersionInformation = Get-ServerOperatingSystemVersion -ComputerName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
    } end {
        return [PSCustomObject]@{
            BuildVersion     = [System.Version]$win32_OperatingSystem.Version
            MajorVersion     = $serverOsVersionInformation.MajorVersion
            InstallationType = $serverOsVersionInformation.InstallationType
            FriendlyName     = $serverOsVersionInformation.FriendlyName
            OperatingSystem  = $win32_OperatingSystem
        }
    }
}

function Get-OperatingSystemRegistryValues {
    [CmdletBinding()]
    param(
        [string]$MachineName,
        [ScriptBlock]$CatchActionFunction
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $baseParams = @{
        MachineName         = $MachineName
        CatchActionFunction = $CatchActionFunction
    }

    $lanManParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters"
        GetValue = "DisableCompression"
    }

    $ubrParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        GetValue = "UBR"
    }

    $ipv6ComponentsParams = $baseParams + @{
        SubKey    = "SYSTEM\CurrentControlSet\Services\TcpIp6\Parameters"
        GetValue  = "DisabledComponents"
        ValueType = "DWord"
    }

    $tcpKeepAliveParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Services\TcpIp\Parameters"
        GetValue = "KeepAliveTime"
    }

    $rpcMinParams = $baseParams + @{
        SubKey   = "Software\Policies\Microsoft\Windows NT\RPC\"
        GetValue = "MinimumConnectionTimeout"
    }

    $renegoClientsParams = $baseParams + @{
        SubKey    = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"
        GetValue  = "AllowInsecureRenegoClients"
        ValueType = "DWord"
    }
    $renegoClientValue = Get-RemoteRegistryValue @renegoClientsParams

    if ($null -eq $renegoClientValue) { $renegoClientValue = "NULL" }

    $renegoServersParams = $baseParams + @{
        SubKey    = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"
        GetValue  = "AllowInsecureRenegoServers"
        ValueType = "DWord"
    }
    $renegoServerValue = Get-RemoteRegistryValue @renegoServersParams

    if ($null -eq $renegoServerValue) { $renegoServerValue = "NULL" }

    $credGuardParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Control\LSA"
        GetValue = "LsaCfgFlags"
    }

    $suppressEpParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Control\LSA"
        GetValue = "SuppressExtendedProtection"
    }

    $lmCompParams = $baseParams + @{
        SubKey    = "SYSTEM\CurrentControlSet\Control\Lsa"
        GetValue  = "LmCompatibilityLevel"
        ValueType = "DWord"
    }
    $lmValue = Get-RemoteRegistryValue @lmCompParams

    if ($null -eq $lmValue) { $lmValue = 3 }

    return [PSCustomObject]@{
        SuppressExtendedProtection      = [int](Get-RemoteRegistryValue @suppressEpParams)
        LmCompatibilityLevel            = $lmValue
        CurrentVersionUbr               = [int](Get-RemoteRegistryValue @ubrParams)
        LanManServerDisabledCompression = [int](Get-RemoteRegistryValue @lanManParams)
        IPv6DisabledComponents          = [int](Get-RemoteRegistryValue @ipv6ComponentsParams)
        TCPKeepAlive                    = [int](Get-RemoteRegistryValue @tcpKeepAliveParams)
        RpcMinConnectionTimeout         = [int](Get-RemoteRegistryValue @rpcMinParams)
        AllowInsecureRenegoServers      = $renegoServerValue
        AllowInsecureRenegoClients      = $renegoClientValue
        CredentialGuard                 = [int](Get-RemoteRegistryValue @credGuardParams)
    }
}

function Get-PageFileInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $pageFiles = @(Get-WmiObjectHandler -ComputerName $Server -Class "Win32_PageFileSetting" -CatchActionFunction ${Function:Invoke-CatchActions})
    $pageFileList = New-Object 'System.Collections.Generic.List[object]'

    if ($null -eq $pageFiles -or
        $pageFiles.Count -eq 0) {
        Write-Verbose "Found No Page File Settings"
        $pageFileList.Add([PSCustomObject]@{
                Name        = [string]::Empty
                InitialSize = 0
                MaximumSize = 0
            })
    } else {
        Write-Verbose "Found $($pageFiles.Count) different page files"
    }

    foreach ($pageFile in $pageFiles) {
        $pageFileList.Add([PSCustomObject]@{
                Name        = $pageFile.Name
                InitialSize = $pageFile.InitialSize
                MaximumSize = $pageFile.MaximumSize
            })
    }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $pageFileList
}


function Get-PowerPlanSetting {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $highPerformanceSet = $false
        $powerPlanSetting = [string]::Empty
        $win32_PowerPlan = Get-WmiObjectHandler -ComputerName $Server -Class Win32_PowerPlan -Namespace 'root\ciMv2\power' -Filter "isActive='true'" -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -ne $win32_PowerPlan) {

            # Guid 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c is 'High Performance' power plan that comes with the OS
            # Guid db310065-829b-4671-9647-2261c00e86ef is 'High Performance (ConfigMgr)' power plan when configured via Configuration Manager / SCCM
            if (($win32_PowerPlan.InstanceID -eq "Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}") -or
                ($win32_PowerPlan.InstanceID -eq "Microsoft:PowerPlan\{db310065-829b-4671-9647-2261c00e86ef}")) {
                Write-Verbose "High Performance Power Plan is set to true"
                $highPerformanceSet = $true
            } else { Write-Verbose "High Performance Power Plan is NOT set to true" }
            $powerPlanSetting = $win32_PowerPlan.ElementName
        } else {
            Write-Verbose "Power Plan Information could not be read"
            $powerPlanSetting = "N/A"
        }
    } end {
        return [PSCustomObject]@{
            HighPerformanceSet = $highPerformanceSet
            PowerPlanSetting   = $powerPlanSetting
            PowerPlan          = $win32_PowerPlan
        }
    }
}

function Get-Smb1ServerSettings {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $smbServerConfiguration = $null
        $windowsFeature = $null
    }
    process {
        $smbServerConfiguration = Invoke-ScriptBlockHandler -ComputerName $ServerName `
            -ScriptBlock { Get-SmbServerConfiguration -ErrorAction Stop } `
            -CatchActionFunction $CatchActionFunction `
            -ScriptBlockDescription "Get-SmbServerConfiguration"

        try {
            $windowsFeature = Get-WindowsFeature "FS-SMB1" -ComputerName $ServerName -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to Get-WindowsFeature for FS-SMB1"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            SmbServerConfiguration = $smbServerConfiguration
            WindowsFeature         = $windowsFeature
            SuccessfulGetInstall   = $null -ne $windowsFeature
            SuccessfulGetBlocked   = $null -ne $smbServerConfiguration
            Installed              = $windowsFeature.Installed -eq $true
            IsBlocked              = $smbServerConfiguration.EnableSMB1Protocol -eq $false
        }
    }
}

function Get-TimeZoneInformation {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [ScriptBlock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $actionsToTake = @()
        $dstIssueDetected = $false
        $registryParams = @{
            MachineName         = $MachineName
            SubKey              = "SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
            CatchActionFunction = $CatchActionFunction
        }
    }
    process {
        $dynamicDaylightTimeDisabled = Get-RemoteRegistryValue @registryParams -GetValue "DynamicDaylightTimeDisabled"
        $timeZoneKeyName = Get-RemoteRegistryValue @registryParams -GetValue "TimeZoneKeyName"
        $standardStart = Get-RemoteRegistryValue @registryParams -GetValue "StandardStart"
        $daylightStart = Get-RemoteRegistryValue @registryParams -GetValue "DaylightStart"

        if ([string]::IsNullOrEmpty($timeZoneKeyName)) {
            Write-Verbose "TimeZoneKeyName is null or empty. Action should be taken to address this."
            $actionsToTake += "TimeZoneKeyName is blank. Need to switch your current time zone to a different value, then switch it back to have this value populated again."
        }

        $standardStartNonZeroValue = ($null -ne ($standardStart | Where-Object { $_ -ne 0 }))
        $daylightStartNonZeroValue = ($null -ne ($daylightStart | Where-Object { $_ -ne 0 }))

        if ($dynamicDaylightTimeDisabled -ne 0 -and
            ($standardStartNonZeroValue -or
            $daylightStartNonZeroValue)) {
            Write-Verbose "Determined that there is a chance the settings set could cause a DST issue."
            $dstIssueDetected = $true
            $actionsToTake += "High Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone. `
            It is possible that you could be running into this issue. Set 'Adjust for daylight saving time automatically to on'"
        } elseif ($dynamicDaylightTimeDisabled -ne 0) {
            Write-Verbose "Daylight savings auto adjustment is disabled."
            $actionsToTake += "Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone."
        }

        $params = @{
            ComputerName           = $Server
            ScriptBlock            = { ([System.TimeZone]::CurrentTimeZone).StandardName }
            ScriptBlockDescription = "Getting Current Time Zone"
            CatchActionFunction    = $CatchActionFunction
        }

        $currentTimeZone = Invoke-ScriptBlockHandler @params
    }
    end {
        return [PSCustomObject]@{
            DynamicDaylightTimeDisabled = $dynamicDaylightTimeDisabled
            TimeZoneKeyName             = $timeZoneKeyName
            StandardStart               = $standardStart
            DaylightStart               = $daylightStart
            DstIssueDetected            = $dstIssueDetected
            ActionsToTake               = [array]$actionsToTake
            CurrentTimeZone             = $currentTimeZone
        }
    }
}

function Get-OperatingSystemInformation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server
    )

    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $buildInformation = Get-OperatingSystemBuildInformation -Server $Server
        $currentDateTime = Get-Date
        $lastBootUpTime = [Management.ManagementDateTimeConverter]::ToDateTime($buildInformation.OperatingSystem.LastBootUpTime)
        $serverBootUp = [PSCustomObject]@{
            Days    = ($currentDateTime - $lastBootUpTime).Days
            Hours   = ($currentDateTime - $lastBootUpTime).Hours
            Minutes = ($currentDateTime - $lastBootUpTime).Minutes
            Seconds = ($currentDateTime - $lastBootUpTime).Seconds
        }

        $powerPlan = Get-PowerPlanSetting -Server $Server
        $pageFile = Get-PageFileInformation -Server $Server
        $networkInformation = Get-NetworkingInformation -Server $Server

        try {
            $hotFixes = (Get-HotFix -ComputerName $Server -ErrorAction Stop) #old school check still valid and faster and a failsafe
        } catch {
            Write-Verbose "Failed to run Get-HotFix"
            Invoke-CatchActions
        }

        $credentialGuardCimInstance = $false
        try {
            $params = @{
                ClassName    = "Win32_DeviceGuard"
                Namespace    = "root\Microsoft\Windows\DeviceGuard"
                ErrorAction  = "Stop"
                ComputerName = $Server
            }
            $credentialGuardCimInstance = (Get-CimInstance @params).SecurityServicesRunning
        } catch {
            Write-Verbose "Failed to run Get-CimInstance for Win32_DeviceGuard"
            Invoke-CatchActions
            $credentialGuardCimInstance = "Unknown"
        }

        $serverPendingReboot = (Get-ServerRebootPending -ServerName $Server -CatchActionFunction ${Function:Invoke-CatchActions})
        $timeZoneInformation = Get-TimeZoneInformation -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $tlsSettings = Get-AllTlsSettings -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $vcRedistributable = Get-VisualCRedistributableInstalledVersion -ComputerName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $smb1ServerSettings = Get-Smb1ServerSettings -ServerName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $registryValues = Get-OperatingSystemRegistryValues -MachineName $Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $netFrameworkInformation = Get-NETFrameworkInformation -Server $Server
    } end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return [PSCustomObject]@{
            BuildInformation           = $buildInformation
            NetworkInformation         = $networkInformation
            PowerPlan                  = $powerPlan
            PageFile                   = $pageFile
            ServerPendingReboot        = $serverPendingReboot
            TimeZone                   = $timeZoneInformation
            TLSSettings                = $tlsSettings
            ServerBootUp               = $serverBootUp
            VcRedistributable          = [array]$vcRedistributable
            RegistryValues             = $registryValues
            Smb1ServerSettings         = $smb1ServerSettings
            HotFixes                   = $hotFixes
            NETFramework               = $netFrameworkInformation
            CredentialGuardCimInstance = $credentialGuardCimInstance
        }
    }
}
function Get-HealthCheckerExchangeServer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ServerName
    )

    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $hardwareInformation = Get-HardwareInformation -Server $ServerName
        $osInformation = Get-OperatingSystemInformation -Server $ServerName
        $exchangeInformation = Get-ExchangeInformation -Server $ServerName
    } end {
        Write-Verbose "Finished building health Exchange Server Object for server: $ServerName"
        return [PSCustomObject]@{
            ServerName              = $ServerName
            HardwareInformation     = $hardwareInformation
            OSInformation           = $osInformation
            ExchangeInformation     = $exchangeInformation
            HealthCheckerVersion    = $BuildVersion
            OrganizationInformation = $null
            GenerationTime          = [DateTime]::Now
        }
    }
}



function Get-ExchangeAdSchemaClass {
    param(
        [Parameter(Mandatory = $true)][string]$SchemaClassName
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand) to query $SchemaClassName schema class"

    $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")

    if ([string]::IsNullOrEmpty($rootDSE.schemaNamingContext)) {
        return $null
    }

    $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher
    $directorySearcher.SearchScope = "Subtree"
    $directorySearcher.SearchRoot = [ADSI]("LDAP://" + $rootDSE.schemaNamingContext.ToString())
    $directorySearcher.Filter = "(Name={0})" -f $SchemaClassName

    $findAll = $directorySearcher.FindAll()

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $findAll
}

function Get-ExchangeAdSchemaInformation {

    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $returnObject = New-Object PSCustomObject
        $schemaClasses = @("ms-Exch-Storage-Group", "ms-Exch-Schema-Version-Pt")

        foreach ($name in $schemaClasses) {
            $propertyName = $name.Replace("-", "")
            $value = $null
            try {
                $value = Get-ExchangeAdSchemaClass -SchemaClassName $name
            } catch {
                Write-Verbose "Failed to get $name"
                Invoke-CatchActions
            }
            $returnObject | Add-Member -MemberType NoteProperty -Name $propertyName -Value $value
        }
    } end {
        return $returnObject
    }
}


function Get-ActiveDirectoryAcl {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.ActiveDirectorySecurity])]
    param (
        [Parameter()]
        [string]
        $DistinguishedName
    )

    $adEntry = [ADSI]("LDAP://$($DistinguishedName)")
    $sdFinder = New-Object System.DirectoryServices.DirectorySearcher($adEntry, "(objectClass=*)", [string[]]("distinguishedName", "ntSecurityDescriptor"), [System.DirectoryServices.SearchScope]::Base)
    $sdResult = $sdFinder.FindOne()
    $ntSdProp = $sdResult.Properties["ntSecurityDescriptor"][0]
    $adSec = New-Object System.DirectoryServices.ActiveDirectorySecurity
    $adSec.SetSecurityDescriptorBinaryForm($ntSdProp)
    return $adSec
}

# Collect the ACLs that we want from all domains where the MESO container exists within the forest.
function Get-ExchangeDomainsAclPermissions {
    [CmdletBinding()]
    param ()

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $exchangeDomains = New-Object 'System.Collections.Generic.List[object]'
    } process {

        $forest = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Forest

        foreach ($domain in $forest.Domains) {

            $domainName = $domain.Name
            Write-Verbose "Working on $domainName"

            $domainObject = [PSCustomObject]@{
                DomainName  = $domainName
                DomainDN    = $null
                Permissions = New-Object 'System.Collections.Generic.List[object]'
                MesoObject  = [PSCustomObject]@{
                    DN            = $null
                    ObjectVersion = 0
                    ACL           = $null
                    WhenChanged   = $null
                }
            }

            try {
                $domainDN = $domain.GetDirectoryEntry().distinguishedName
                $domainObject.DomainDN = $domainDN.ToString()
            } catch {
                Write-Verbose "Domain: $domainName - seems to be offline and will be skipped"
                $domainObject.DomainDN = "Unknown" # Set the domain to unknown vs not knowing it is there.
                Invoke-CatchActions
                continue
            }

            try {
                $mesoEntry = [ADSI]("LDAP://CN=Microsoft Exchange System Objects," + $domainDN)
                $sdFinder = New-Object System.DirectoryServices.DirectorySearcher($mesoEntry)
                $mesoResult = $sdFinder.FindOne()
                Write-Verbose "Found the MESO Container in domain"
            } catch {
                Write-Verbose "Failed to find MESO container in $domainDN"
                Write-Verbose "Skipping over domain"
                Invoke-CatchActions
                continue
            }
            [int]$mesoObjectVersion = $mesoResult.Properties["ObjectVersion"][0]
            $mesoWhenChangedInfo = $mesoResult.Properties["WhenChanged"]
            $mesoDN = $mesoResult.Properties["DistinguishedName"]
            Write-Verbose "Object Version: $mesoObjectVersion"
            Write-Verbose "When Changed: $mesoWhenChangedInfo"
            Write-Verbose "MESO DN: $mesoDN"
            $mesoAcl = $null

            try {
                $mesoAcl = Get-ActiveDirectoryAcl $mesoDN
                Write-Verbose "Got the MESO ACL"
            } catch {
                Write-Verbose "Failed to get the MESO ACL"
                Invoke-CatchActions
            }
            $domainObject.MesoObject.DN = $mesoDN
            $domainObject.MesoObject.ObjectVersion = $mesoObjectVersion
            $domainObject.MesoObject.ACL = $mesoAcl
            $domainObject.MesoObject.WhenChanged = $mesoWhenChangedInfo

            $permissionsCheckList = @($domainDN.ToString(), "CN=AdminSDHolder,CN=System,$domainDN")

            foreach ($permissionDN in $permissionsCheckList) {
                $acl = $null
                try {
                    $acl = Get-ActiveDirectoryAcl $permissionDN
                    Write-Verbose "Got the ACL for: $permissionDN"
                } catch {
                    Write-Verbose "Failed to get the ACL for: $permissionDN"
                    Invoke-CatchActions
                }
                $domainObject.Permissions.Add([PSCustomObject]@{
                        DN  = $permissionDN
                        Acl = $acl
                    })
            }

            $exchangeDomains.Add($domainObject)
        }
    } end {
        return $exchangeDomains
    }
}



function Get-ExchangeOtherWellKnownObjects {
    [CmdletBinding()]
    param ()

    $otherWellKnownObjectIds = @{
        "C262A929D691B74A9E068728F8F842EA" = "Organization Management"
        "DB72C41D49580A4DB304FE6981E56297" = "Recipient Management"
        "1A9E39D35ABE5747B979FFC0C6E5EA26" = "View-Only Organization Management"
        "45FA417B3574DC4E929BC4B059699792" = "Public Folder Management"
        "E80CDFB75697934981C898B4DBC5A0C6" = "UM Management"
        "B3DDC6BE2A3BE84B97EB2DCE9477E389" = "Help Desk"
        "BEA432C94E1D254EAF99B40573360D5B" = "Records Management"
        "C67FDE2E8339674490FBAFDCA3DFDC95" = "Discovery Management"
        "4DB8E7754EB6C1439565612E69A80A4F" = "Server Management"
        "D1281926D1F55B44866D1D6B5BD87A09" = "Delegated Setup"
        "03B709F451F3BF4388E33495369B6771" = "Hygiene Management"
        "B30A449BA9B420458C4BB22F33C52766" = "Compliance Management"
        "A7D2016C83F003458132789EEB127B84" = "Exchange Servers"
        "EA876A58DB6DD04C9006939818F800EB" = "Exchange Trusted Subsystem"
        "02522ECF9985984A9232056FC704CC8B" = "Managed Availability Servers"
        "4C17D0117EBE6642AFAEE03BC66D381F" = "Exchange Windows Permissions"
        "9C5B963F67F14A4B936CB8EFB19C4784" = "ExchangeLegacyInterop"
        "776B176BD3CB2A4DA7829EA963693013" = "Security Reader"
        "03D7F0316EF4B3498AC434B6E16F09D9" = "Security Administrator"
        "A2A4102E6F676141A2C4AB50F3C102D5" = "PublicFolderMailboxes"
    }

    $exchangeContainer = Get-ExchangeContainer
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($exchangeContainer, "(objectClass=*)", @("otherWellKnownObjects", "distinguishedName"))
    $result = $searcher.FindOne()
    foreach ($val in $result.Properties["otherWellKnownObjects"]) {
        $matchResults = $val | Select-String "^B:32:([^:]+):(.*)$"
        if ($matchResults.Matches.Groups.Count -ne 3) {
            # Only output the raw value of a corrupted entry
            [PSCustomObject]@{
                WellKnownName     = $null
                WellKnownGuid     = $null
                DistinguishedName = $null
                RawValue          = $val
            }

            continue
        }

        $wkGuid = $matchResults.Matches.Groups[1].Value
        $wkName = $otherWellKnownObjectIds[$wkGuid]

        [PSCustomObject]@{
            WellKnownName     = $wkName
            WellKnownGuid     = $wkGuid
            DistinguishedName = $matchResults.Matches.Groups[2].Value
            RawValue          = $val
        }
    }
}
function Get-ExchangeWellKnownSecurityGroups {
    [CmdletBinding()]
    param()
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $exchangeGroups = New-Object 'System.Collections.Generic.List[object]'
    } process {
        try {
            $otherWellKnownObjects = Get-ExchangeOtherWellKnownObjects
        } catch {
            Write-Verbose "Failed to get Get-ExchangeOtherWellKnownObjects"
            Invoke-CatchActions
            return
        }

        foreach ($wkObject in $otherWellKnownObjects) {
            try {
                Write-Verbose "Attempting to get SID from $($wkObject.DistinguishedName)"
                $entry = [ADSI]("LDAP://$($wkObject.DistinguishedName)")
                $wkObject | Add-Member -MemberType NoteProperty -Name SID -Value ((New-Object System.Security.Principal.SecurityIdentifier($entry.objectSid.Value, 0)).Value)
                $exchangeGroups.Add($wkObject)
            } catch {
                Write-Verbose "Failed to find SID"
                Invoke-CatchActions
            }
        }
    } end {
        return $exchangeGroups
    }
}

function Get-SecurityCve-2021-34470 {
    [CmdletBinding()]
    param(
        [object]$MsExchStorageGroup
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $unknown = $null -eq $MsExchStorageGroup
        $isVulnerable = "Unknown"
        try {
            $isVulnerable = (-not $unknown) -and $MsExchStorageGroup.Properties["PossSuperiors"] -eq "computer"
        } catch {
            Write-Verbose "Failed to evaluate MsExchStorageGroup"
            Invoke-CatchActions
        }

        Write-Verbose "Unknown: $unknown IsVulnerable: $isVulnerable"
        return [PSCustomObject]@{
            Unknown      = $unknown
            IsVulnerable = $isVulnerable
        }
    }
}

# Check within each domain if we are vulnerable to CVE-2022-21978
function Get-SecurityCve-2022-21978 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$DomainsAcls,

        [Parameter(Mandatory = $true)]
        [object]$ExchangeWellKnownSecurityGroups,

        [Parameter(Mandatory = $true)]
        [ValidateSet("2013", "2016", "2019")]
        [string]$ExchangeSchemaLevel,

        [Parameter(Mandatory = $true)]
        [bool]$SplitADPermissions
    )
    begin {

        function NewMatchingEntry {
            param(
                [ValidateSet("Domain", "AdminSDHolder")]
                [string]$TargetObject,
                [string]$ObjectTypeGuid,
                [string]$InheritedObjectType
            )

            return [PSCustomObject]@{
                TargetObject        = $TargetObject
                ObjectTypeGuid      = $ObjectTypeGuid
                InheritedObjectType = $InheritedObjectType
            }
        }

        function NewGroupEntry {
            param(
                [string]$Name,
                [object[]]$MatchingEntries
            )

            return [PSCustomObject]@{
                Name     = $Name
                Sid      = $null
                AceEntry = $MatchingEntries
            }
        }

        # Computer Class GUID
        $computerClassGUID = "bf967a86-0de6-11d0-a285-00aa003049e2"

        # userCertificate GUID
        $userCertificateGUID = "bf967a7f-0de6-11d0-a285-00aa003049e2"

        # managedBy GUID
        $managedByGUID = "0296c120-40da-11d1-a9c0-0000f80367c1"

        $writePropertyRight = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
        $denyType = [System.Security.AccessControl.AccessControlType]::Deny
        $inheritanceAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::All

        if ($ExchangeSchemaLevel -eq "2013") {
            $objectVersionSchemaValueMin = 13238
        } else {
            # Exchange 2016 and 2019 have the same MESO objectVersion Value
            $objectVersionSchemaValueMin = 13243
        }

        $groupLists = @(
        (NewGroupEntry "Exchange Servers" @(
            (NewMatchingEntry -TargetObject "Domain" -ObjectTypeGuid $userCertificateGUID -InheritedObjectType $computerClassGUID)
            )),

        (NewGroupEntry "Exchange Windows Permissions" @(
            (NewMatchingEntry -TargetObject "Domain" -ObjectTypeGuid $managedByGUID -InheritedObjectType $computerClassGUID)
            )))

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $domainResults = New-Object 'System.Collections.Generic.List[object]'
    } process {
        # Set the SID on the GroupList
        foreach ($group in $groupLists) {
            $wkObject = $ExchangeWellKnownSecurityGroups | Where-Object { $_.WellKnownName -eq $group.Name }

            if ($null -ne $wkObject) {
                $group.Sid = $wkObject.Sid
            }
        }

        # Loop through each domain and determine if they are secure or not.
        foreach ($domain in $DomainsAcls) {

            $domainName = $domain.DomainName
            Write-Verbose "Checking Domain $domainName"
            $domainAcl = ($domain.Permissions | Where-Object { $_.DN -eq $domain.DomainDN }).Acl
            $adminSdHolderAcl = ($domain.Permissions | Where-Object { $_.DN -eq "CN=AdminSDHolder,CN=System,$($domain.DomainDN)" }).Acl
            $unknownDomain = $domain.DomainDN -eq "Unknown"

            $fallbackLogic = $null -eq $domainAcl -or
            $null -eq $domainAcl.Access -or
            $null -eq $adminSdHolderAcl -or
            $null -eq $adminSdHolderAcl.Access -or
            $unknownDomain -or
            $SplitADPermissions

            $aceListResults = New-Object 'System.Collections.Generic.List[object]'

            if (-not ($fallbackLogic)) {
                # Truly check the ACE in the ACL

                foreach ($group in $groupLists) {
                    Write-Verbose "Looking Ace Entries for the group: $($group.Name)"

                    foreach ($entry in $group.AceEntry) {
                        Write-Verbose "Trying to find the entry GUID: $($entry.ObjectTypeGuid)"
                        if ($entry.TargetObject -eq "AdminSDHolder") {
                            $objectAcl = $adminSdHolderAcl
                            $objectDN = $adminSdHolderDN
                        } else {
                            $objectAcl = $domainAcl
                            $objectDN = $domainDN
                        }
                        Write-Verbose "ObjectDN: $objectDN"

                        try {

                            # We need to pass an IdentityReference object to the constructor
                            $groupIdentityRef = New-Object System.Security.Principal.SecurityIdentifier($group.Sid)

                            $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($groupIdentityRef, $writePropertyRight, $denyType, $entry.ObjectTypeGuid, $inheritanceAll, $entry.InheritedObjectType)

                            $checkAce = $objectAcl.Access.Where({
                            ($_.ActiveDirectoryRights -eq $ace.ActiveDirectoryRights) -and
                            ($_.InheritanceType -eq $ace.InheritanceType) -and
                            ($_.ObjectType -eq $ace.ObjectType) -and
                            ($_.InheritedObjectType -eq $ace.InheritedObjectType) -and
                            ($_.ObjectFlags -eq $ace.ObjectFlags) -and
                            ($_.AccessControlType -eq $ace.AccessControlType) -and
                            ($_.IsInherited -eq $ace.IsInherited) -and
                            ($_.InheritanceFlags -eq $ace.InheritanceFlags) -and
                            ($_.PropagationFlags -eq $ace.PropagationFlags) -and
                            ($_.IdentityReference -eq $ace.IdentityReference.Translate([System.Security.Principal.NTAccount]))
                                })

                            $checkPass = $checkAce.Count -gt 0
                            Write-Verbose "Ace Result Check Passed: $checkPass"

                            $aceListResults.Add([PSCustomObject]@{
                                    ObjectDN  = $objectDN
                                    CheckPass = $checkPass
                                })
                        } catch {
                            Write-Verbose "Failed to do ACE comparison"
                            Invoke-CatchActions
                        }
                    }
                }
            }

            # should be true in fallback or all ace exists
            $allAcePass = ($aceListResults | Where-Object { $_.CheckPass -eq $false }).Count -eq 0
            $MesoUpdated = $domain.MesoObject.ObjectVersion -ge $objectVersionSchemaValueMin
            $domainPassed = $allAcePass -and $MesoUpdated -and (-not $unknownDomain)
            $domainResults.Add([PSCustomObject]@{
                    DomainName    = $domainName
                    DomainPassed  = $domainPassed
                    MesoUpdated   = $MesoUpdated
                    AllAcePass    = $allAcePass
                    UnknownDomain = $unknownDomain
                })
        }
    } end {
        return $domainResults
    }
}


function Get-ExchangeADSplitPermissionsEnabled {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [ScriptBlock]$CatchActionFunction
    )

    <#
        The following bullets are AD split permissions indicators:
        - An organizational unit (OU) named Microsoft 'Exchange Protected Groups' is created
        - The 'Exchange Windows Permissions' security group is created/moved in/to the 'Microsoft Exchange Protected Groups' OU
        - The 'Exchange Trusted Subsystem' security group isn't member of the 'Exchange Windows Permissions' security group
        - ACEs that would have been assigned to the 'Exchange Windows Permissions' security group aren't added to the Active Directory domain object
        See: https://learn.microsoft.com/exchange/permissions/split-permissions/split-permissions?view=exchserver-2019#active-directory-split-permissions
    #>

    $isADSplitPermissionsEnabled = $false
    try {
        $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
        $exchangeTrustedSubsystemDN = ("CN=Exchange Trusted Subsystem,OU=Microsoft Exchange Security Groups," + $rootDSE.rootDomainNamingContext)
        $adSearcher = New-Object DirectoryServices.DirectorySearcher
        $adSearcher.Filter = '(&(objectCategory=group)(cn=Exchange Windows Permissions))'
        $adSearcher.SearchRoot = ("LDAP://OU=Microsoft Exchange Protected Groups," + $rootDSE.rootDomainNamingContext)
        $adSearcherResult = $adSearcher.FindOne()

        if ($null -ne $adSearcherResult) {
            Write-Verbose "'Exchange Windows Permissions' in 'Microsoft Exchange Protected Groups' OU detected"
            # AD split permissions is enabled if 'Exchange Trusted Subsystem' isn't a member of the 'Exchange Windows Permissions' security group
            $isADSplitPermissionsEnabled = (($null -eq $adSearcherResult.Properties.member) -or
            (-not($adSearcherResult.Properties.member).ToLower().Contains($exchangeTrustedSubsystemDN.ToLower())))
        }
    } catch {
        Write-Verbose "OU 'Microsoft Exchange Protected Groups' was not found - AD split permissions not enabled"
        Invoke-CatchActionError $CatchActionFunction
    }

    return $isADSplitPermissionsEnabled
}
function Get-OrganizationInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [bool]$EdgeServer
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $organizationConfig = $null
        $domainsAclPermissions = $null
        $wellKnownSecurityGroups = $null
        $adSchemaInformation = $null
        $getHybridConfiguration = $null
        $enableDownloadDomains = "Unknown" # Set to unknown by default.
        $getAcceptedDomain = $null
        $mapiHttpEnabled = $false
        $securityResults = $null
        $isSplitADPermissions = $false
        $adSiteCount = 0
        $getSettingOverride = $null
    } process {
        try {
            $organizationConfig = Get-OrganizationConfig -ErrorAction Stop
        } catch {
            Write-Yellow "Failed to run Get-OrganizationConfig."
            Invoke-CatchActions
        }

        # Pull out information from OrganizationConfig
        # This is done incase Get-OrganizationConfig and we set a true boolean value of false
        if ($null -ne $organizationConfig) {
            $mapiHttpEnabled = $organizationConfig.MapiHttpEnabled
            # Enabled Download Domains will not be there if running EMS from Exchange 2013.
            # By default, EnableDownloadDomains is set to Unknown in case this is run on 2013 server.
            if ($null -ne $organizationConfig.EnableDownloadDomains) {
                $enableDownloadDomains = $organizationConfig.EnableDownloadDomains
            } else {
                Write-Verbose "No EnableDownloadDomains detected on Get-OrganizationConfig"
            }
        } else {
            Write-Verbose "MAPI HTTP Enabled and Download Domains Enabled results not accurate"
        }

        try {
            $getAcceptedDomain = Get-AcceptedDomain -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to run Get-AcceptedDomain"
            $getAcceptedDomain = "Unknown"
            Invoke-CatchActions
        }

        # NO Edge Server Collection
        if (-not ($EdgeServer)) {

            $adSchemaInformation = Get-ExchangeAdSchemaInformation
            $domainsAclPermissions = Get-ExchangeDomainsAclPermissions
            $wellKnownSecurityGroups = Get-ExchangeWellKnownSecurityGroups
            $isSplitADPermissions = Get-ExchangeADSplitPermissionsEnabled -CatchActionFunction ${Function:Invoke-CatchActions}

            try {
                $getIrmConfiguration = Get-IRMConfiguration -ErrorAction Stop
            } catch {
                Write-Verbose "Failed to get the IRM Configuration"
                Invoke-CatchActions
            }

            try {
                # It was reported that this isn't getting thrown to the catch action when failing. As a quick fix, handling this by looping over errors.
                $currentErrors = $Error.Count
                $getDdgPublicFolders = @(Get-DynamicDistributionGroup "PublicFolderMailboxes*" -IncludeSystemObjects -ErrorAction "Stop")
                Invoke-CatchActionErrorLoop $currentErrors ${Function:Invoke-CatchActions}
            } catch {
                Write-Verbose "Failed to get the dynamic distribution group for public folder mailboxes."
                Invoke-CatchActions
            }

            try {
                $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
                $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher
                $directorySearcher.SearchScope = "Subtree"
                $directorySearcher.SearchRoot = [ADSI]("LDAP://" + $rootDSE.configurationNamingContext.ToString())
                $directorySearcher.Filter = "(objectCategory=site)"
                $directorySearcher.PageSize = 100
                $adSiteCount = ($directorySearcher.FindAll()).Count
            } catch {
                Write-Verbose "Failed to collect AD Site Count information"
                Invoke-CatchActions
            }

            $schemaRangeUpper = (
                ($adSchemaInformation.msExchSchemaVersionPt.Properties["RangeUpper"])[0]).ToInt32([System.Globalization.NumberFormatInfo]::InvariantInfo)

            if ($schemaRangeUpper -lt 15323) {
                $schemaLevel = "2013"
            } elseif ($schemaRangeUpper -lt 17000) {
                $schemaLevel = "2016"
            } else {
                $schemaLevel = "2019"
            }

            $cve21978Params = @{
                DomainsAcls                     = $domainsAclPermissions
                ExchangeWellKnownSecurityGroups = $wellKnownSecurityGroups
                ExchangeSchemaLevel             = $schemaLevel
                SplitADPermissions              = $isSplitADPermissions
            }

            $cve34470Params = @{
                MsExchStorageGroup = $adSchemaInformation.MsExchStorageGroup
            }

            $securityResults = [PSCustomObject]@{
                CVE202221978 = (Get-SecurityCve-2022-21978 @cve21978Params)
                CVE202134470 = (Get-SecurityCve-2021-34470 @cve34470Params)
            }

            try {
                $getHybridConfiguration = Get-HybridConfiguration -ErrorAction Stop
            } catch {
                Write-Yellow "Failed to run Get-HybridConfiguration"
                Invoke-CatchActions
            }

            try {
                $getSettingOverride = Get-SettingOverride -ErrorAction Stop
            } catch {
                Write-Verbose "Failed to run Get-SettingOverride"
                $getSettingOverride = "Unknown"
                Invoke-CatchActions
            }
        }
    } end {
        return [PSCustomObject]@{
            GetOrganizationConfig             = $organizationConfig
            DomainsAclPermissions             = $domainsAclPermissions
            WellKnownSecurityGroups           = $wellKnownSecurityGroups
            AdSchemaInformation               = $adSchemaInformation
            GetHybridConfiguration            = $getHybridConfiguration
            EnableDownloadDomains             = $enableDownloadDomains
            GetAcceptedDomain                 = $getAcceptedDomain
            MapiHttpEnabled                   = $mapiHttpEnabled
            SecurityResults                   = $securityResults
            IsSplitADPermissions              = $isSplitADPermissions
            ADSiteCount                       = $adSiteCount
            GetSettingOverride                = $getSettingOverride
            GetDynamicDgPublicFolderMailboxes = $getDdgPublicFolders
            GetIrmConfiguration               = $getIrmConfiguration
        }
    }
}

# Collects the data required for Health Checker
function Get-HealthCheckerData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ServerNames,

        [Parameter(Mandatory = $true)]
        [bool]$EdgeServer,

        [Parameter(Mandatory = $false)]
        [bool]$ReturnDataCollectionOnly = $false #TODO Remove this an display somewhere else. This function should only do data collection. Once it is optimized to do so.
    )

    function TestComputerName {
        [CmdletBinding()]
        [OutputType([bool])]
        param(
            [string]$ComputerName
        )
        try {
            Write-Verbose "Testing $ComputerName"

            # If local computer, we should just assume that it should work.
            if ($ComputerName -eq $env:COMPUTERNAME) {
                Write-Verbose "Local computer, returning true"
                return $true
            }

            Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-Date } -ErrorAction Stop | Out-Null
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $ComputerName)
            $reg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion") | Out-Null
            Write-Verbose "Returning true back"
            return $true
        } catch {
            Write-Verbose "Failed to run against $ComputerName"
            Invoke-CatchActions
        }
        return $false
    }

    function ExportHealthCheckerXml {
        [CmdletBinding()]
        [OutputType([bool])]
        param(
            [Parameter(Mandatory = $true)]
            [object]$SaveDataObject,

            [Parameter(Mandatory = $true)]
            [hashtable]$ProgressParams
        )
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $dataExported = $false

        try {
            $currentErrors = $Error.Count
            $ProgressParams.Status = "Exporting Data"
            Write-Progress @ProgressParams
            $SaveDataObject | Export-Clixml -Path $Script:OutXmlFullPath -Encoding UTF8 -Depth 2 -ErrorAction Stop -Force
            Write-Verbose "Successfully export out the data"
            $dataExported = $true
        } catch {
            try {
                Write-Verbose "Failed to Export-Clixml. Inner Exception: $_"
                Write-Verbose "Converting HealthCheckerExchangeServer to json."
                $outputXml = [PSCustomObject]@{
                    HealthCheckerExchangeServer = $null
                    HtmlServerValues            = $null
                    DisplayResults              = $null
                }

                if ($null -ne $SaveDataObject.HealthCheckerExchangeServer) {
                    $jsonHealthChecker = $SaveDataObject.HealthCheckerExchangeServer | ConvertTo-Json -Depth 6 -ErrorAction Stop
                    $outputXml.HtmlServerValues = $SaveDataObject.HtmlServerValues
                    $outputXml.DisplayResults = $SaveDataObject.DisplayResults
                } else {
                    $jsonHealthChecker = $SaveDataObject | ConvertTo-Json -Depth 6 -ErrorAction Stop
                }

                $outputXml.HealthCheckerExchangeServer = $jsonHealthChecker | ConvertFrom-Json -ErrorAction Stop
                $outputXml | Export-Clixml -Path $Script:OutXmlFullPath -Encoding UTF8 -Depth 2 -ErrorAction Stop -Force
                Write-Verbose "Successfully export out the data after the convert"
                $dataExported = $true
            } catch {
                Write-Red "Failed to Export-Clixml. Unable to export the data."
            }
        } finally {
            # This prevents the need to call Invoke-CatchActions
            Invoke-ErrorCatchActionLoopFromIndex $currentErrors
        }
        return $dataExported
    }

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $paramWriteProgress = @{
        Id       = 1
        Activity = "Organization Information"
        Status   = "Data Collection"
    }

    Write-Progress @paramWriteProgress
    $organizationInformation = Get-OrganizationInformation -EdgeServer $EdgeServer

    $failedServerList = New-Object "System.Collections.Generic.List[string]"
    $returnDataList = New-Object "System.Collections.Generic.List[object]"
    $serverCount = 0

    foreach ($serverName in $ServerNames) {

        # Set serverName to be not be FQDN if that is what is passed.
        $serverName = $serverName.Split(".")[0]

        $paramWriteProgress.Activity = "Server: $serverName"
        $paramWriteProgress.Status = "Data Collection"
        $paramWriteProgress.PercentComplete = (($serverCount / $ServerNames.Count) * 100)
        Write-Progress @paramWriteProgress
        $serverCount++

        try {
            $fqdn = (Get-ExchangeServer $serverName -ErrorAction Stop).FQDN
            Write-Verbose "Set FQDN to $fqdn"
        } catch {
            Write-Host "Unable to find server: $serverName" -ForegroundColor Yellow
            Invoke-CatchActions
            continue
        }

        # Test out serverName and FQDN to determine if we can properly reach the server.
        # It appears in some environments, you can't do both.
        $serverNameParam = $fqdn

        if (-not (TestComputerName $fqdn)) {
            if (-not (TestComputerName $serverName)) {
                $line = "Unable to connect to server $serverName. Please run locally"
                Write-Verbose $line
                Write-Host $line -ForegroundColor Yellow
                continue
            }
            Write-Verbose "Set serverNameParam to $serverName"
            $serverNameParam = $serverName
        }

        try {
            Invoke-SetOutputInstanceLocation -Server $serverName -FileName "HealthChecker" -IncludeServerName $true

            if (-not $Script:VulnerabilityReport) {
                # avoid having vulnerability report having a txt file with nothing in it besides the Exchange Health Checker Version
                Write-HostLog "Exchange Health Checker version $BuildVersion"
            }

            $HealthObject = $null
            $HealthObject = Get-HealthCheckerExchangeServer -ServerName $serverNameParam
            $HealthObject.OrganizationInformation = $organizationInformation

            # If we successfully got the data, we want to export it out right away.
            # This then allows if an exception does occur in the analysis stage,
            # we then have the data output that is reproducing a problem in that section of code that we can debug.
            $dataExported = ExportHealthCheckerXml -SaveDataObject $HealthObject -ProgressParams $paramWriteProgress
            $paramWriteProgress.Status = "Analyzing Data"
            Write-Progress @paramWriteProgress
            $analyzedResults = Invoke-AnalyzerEngine -HealthServerObject $HealthObject

            if (-not $ReturnDataCollectionOnly) {
                Write-Progress @paramWriteProgress -Completed
                Write-ResultsToScreen -ResultsToWrite $analyzedResults.DisplayResults
            } else {
                $returnDataList.Add($analyzedResults)
            }
        } catch {
            Write-Red "Failed to Health Checker against $serverName"
            $failedServerList.Add($serverName)

            if ($null -eq $HealthObject) {
                # Try to handle the issue so we don't get a false positive report.
                Invoke-CatchActions
            }
            continue
        } finally {

            if ($null -ne $analyzedResults) {
                # Export out the analyzed data, as this is needed for Build HTML Report.
                $dataExported = ExportHealthCheckerXml -SaveDataObject $analyzedResults -ProgressParams $paramWriteProgress
            }

            # for now don't want to display that we output the information if ReturnDataCollectionOnly is false
            if ($dataExported -and -not $ReturnDataCollectionOnly) {
                Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
                Write-Grey("Exported Data Object Written to {0} " -f $Script:OutXmlFullPath)
            }
        }
    }
    Write-Verbose "Failed Server List: $([string]::Join(",", $failedServerList))"

    if ($ReturnDataCollectionOnly) {
        return $returnDataList
    }
}
# The main functionality of Exchange Health Checker.
# Collect information and report it to the screen and export out the results.
function Invoke-HealthCheckerMainReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ServerNames,

        [Parameter(Mandatory = $true)]
        [bool]$EdgeServer
    )

    $currentErrors = $Error.Count

    if ((-not $SkipVersionCheck) -and
                (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/HC-VersionsUrl")) {
        Write-Yellow "Script was updated. Please rerun the command."
        return
    } else {
        $Script:DisplayedScriptVersionAlready = $true
        Write-Green "Exchange Health Checker version $BuildVersion"
    }

    Invoke-ErrorCatchActionLoopFromIndex $currentErrors
    Get-HealthCheckerData $ServerNames $EdgeServer
}

# This function is used to collect the data from the entire environment for vulnerabilities.
# This will still use the entire data collection process that is handled within the main HC data collection
# However, it will not display it on the screen and just analyze the data and pull out the data that we want.
function Invoke-VulnerabilityReport {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $currentErrors = $Error.Count

    if ((-not $SkipVersionCheck) -and
                (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/HC-VersionsUrl")) {
        Write-Yellow "Script was updated. Please rerun the command."
        return
    } else {
        $Script:DisplayedScriptVersionAlready = $true
        Write-Green "Exchange Health Checker version $BuildVersion"
    }

    Invoke-ErrorCatchActionLoopFromIndex $currentErrors
    $stopWatch = [System.Diagnostics.Stopwatch]::StartNew()
    Set-ADServerSettings -ViewEntireForest $true
    $exchangeServers = @(Get-ExchangeServer)
    Write-Verbose "Took $($stopWatch.Elapsed.TotalSeconds) seconds to run Get-ExchangeServer"
    Write-Verbose "Found $($exchangeServers.Count) Exchange Servers"

    $healthCheckerData = New-Object System.Collections.Generic.List[object]
    $serverNames = New-Object System.Collections.Generic.List[string]
    $exchangeServers | ForEach-Object { $serverNames.Add($_.Name.ToLower()) }

    # By default, always try to import the local data and see if we can use it.
    # This is determined by the files being the same version as the current version of the script.
    $importData = Get-ExportedHealthCheckerFiles -Directory $XMLDirectoryPath

    if ($null -ne $importData) {
        $importData |
            Where-Object { $_.HealthCheckerExchangeServer.HealthCheckerVersion -eq $BuildVersion } |
            ForEach-Object {
                $name = $_.HealthCheckerExchangeServer.ServerName.Split(".")[0].ToLower()
                Write-Verbose "Server $name is being imported"
                $healthCheckerData.Add($_)
                # Remove it from the list
                $serverNames.Remove($name) | Out-Null
            }
        Write-Verbose "Imported $($healthCheckerData.Count) server data to save time"
    } else {
        Write-Verbose "No data was imported."
    }

    # We may have imported all the files required, therefore, we can skip attempting to import the data.
    if ($serverNames.Count -gt 0) {
        Get-HealthCheckerData -ServerNames $serverNames -EdgeServer $false -ReturnDataCollectionOnly $true |
            ForEach-Object {
                $healthCheckerData.Add($_)
            }
    }

    Write-Verbose "Took $($stopWatch.Elapsed.TotalSeconds) seconds to get all the Health Checker data"
    $serverVulnerabilityReport = New-Object System.Collections.Generic.List[object]

    foreach ($exchServer in $exchangeServers) {
        $serverName = $exchServer.Name
        Write-Verbose "Working on server $($serverName)"
        $serverData = $healthCheckerData | Where-Object { $_.HealthCheckerExchangeServer.ExchangeInformation.GetExchangeServer.Name -eq $serverName }
        $vulnerabilityList = New-Object System.Collections.Generic.List[object]
        $buildVersionInfo = Get-ExchangeBuildVersionInformation -AdminDisplayVersion $exchServer.AdminDisplayVersion

        if ($null -ne $serverData) {
            $hc = $serverData.HealthCheckerExchangeServer
            $serverName = $hc.ServerName
            $buildVersionInfo = $hc.ExchangeInformation.BuildInformation.VersionInformation
            $securityKey = [array]@($serverData.DisplayResults.Keys) | Where-Object { $_.Name -eq "Security Vulnerability" }

            foreach ($vulnerability in $serverData.DisplayResults[$securityKey]) {
                $securityVulnerability = "Security Vulnerability"
                $iisModule = "IIS module anomalies detected"
                $cveName = [string]::Empty

                if ($vulnerability.Name -eq $securityVulnerability) {
                    if ($vulnerability.CustomName -ne $securityVulnerability) {
                        $cveName = $vulnerability.CustomName
                    } else {
                        $cveName = $vulnerability.CustomValue
                    }
                } elseif (($vulnerability.Name -eq $iisModule -and
                        $vulnerability.CustomValue -eq $true) -or
                    ($null -ne $vulnerability.Name -and
                    $vulnerability.Name -ne $iisModule)) {
                    $cveName = $vulnerability.Name
                } else {
                    Write-Verbose "Failed to determine Security Vulnerability match"
                }

                if ($cveName -ne [string]::Empty) {
                    $vulnerabilityList.Add($cveName)
                }
            }
        } else {
            Write-Verbose "Didn't get Health Checker Data"
        }

        $hasVulnerability = "Unknown"
        $online = $null -ne $serverData

        if ($online) {
            $hasVulnerability = $vulnerabilityList.Count -gt 0
        }

        $serverVulnerabilityReport.Add([PSCustomObject]@{
                Name             = $serverName
                BuildVersion     = $buildVersionInfo.BuildVersion.ToString()
                HasVulnerability = $hasVulnerability
                Vulnerabilities  = $vulnerabilityList
                Online           = $online
            })
    }

    $vulnerabilityReport = [PSCustomObject]@{
        VersionReport = $BuildVersion
        ReportDate    = (Get-Date).ToString()
        Organization  = $healthCheckerData[0].HealthCheckerExchangeServer.OrganizationInformation.GetOrganizationConfig.DistinguishedName
        Servers       = $serverVulnerabilityReport
    }

    $outputFile = "$($Script:OutputFilePath)\HealthChecker-VulnerabilityReport-$($Script:dateTimeStringFormat).json"
    $vulnerabilityReport | ConvertTo-Json -Depth 3 | Out-File $outputFile
    Write-Verbose "Took $($stopWatch.Elapsed.TotalSeconds) seconds to complete vulnerability report."
    Write-Host "Report written to $outputFile"
}


function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
}

function Get-NewLoggerInstance {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = (Get-Location).Path,

        [ValidateNotNullOrEmpty()]
        [string]$LogName = "Script_Logging",

        [bool]$AppendDateTime = $true,

        [bool]$AppendDateTimeToFileName = $true,

        [int]$MaxFileSizeMB = 10,

        [int]$CheckSizeIntervalMinutes = 10,

        [int]$NumberOfLogsToKeep = 10
    )

    $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
    $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)

    if (-not (Test-Path $LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
        } catch {
            throw "Failed to create Log Directory: $LogDirectory. Inner Exception: $_"
        }
    }

    return [PSCustomObject]@{
        FullPath                 = $fullFilePath
        AppendDateTime           = $AppendDateTime
        MaxFileSizeMB            = $MaxFileSizeMB
        CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
        NumberOfLogsToKeep       = $NumberOfLogsToKeep
        BaseInstanceFileName     = $fileName.Replace(".txt", "")
        Instance                 = 1
        NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
        PreventLogCleanup        = $false
        LoggerDisabled           = $false
    } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
}

function Write-LoggerInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]$Object
    )
    process {
        if ($LoggerInstance.LoggerDisabled) { return }

        if ($LoggerInstance.AppendDateTime -and
            $Object.GetType().Name -eq "string") {
            $Object = "[$([System.DateTime]::Now)] : $Object"
        }

        # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
        $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false

        #Upkeep of the logger information
        if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }

        #Set next update time to avoid issues so we can log things
        $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
        $item = Get-ChildItem $LoggerInstance.FullPath

        if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
            $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
            $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
            $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
            $LoggerInstance.Instance++
            $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)

            $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"

            if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                $item | Remove-Item -Force
            }
        }
    }
    end {
        return $LoggerInstance
    }
}

function Invoke-LoggerInstanceCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance
    )
    process {
        if ($LoggerInstance.LoggerDisabled -or
            $LoggerInstance.PreventLogCleanup) {
            return
        }

        Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
            Remove-Item -Force
    }
}

function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    process {
        $consoleHost = $host.Name -eq "ConsoleHost"

        if ($null -ne $Script:WriteHostManipulateObjectAction) {
            $Object = & $Script:WriteHostManipulateObjectAction $Object
        }

        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params

        if ($null -ne $Script:WriteHostDebugAction -and
            $null -ne $Object) {
            &$Script:WriteHostDebugAction $Object
        }
    }
}

function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

function SetWriteHostAction ($DebugAction) {
    $Script:WriteHostDebugAction = $DebugAction
}

function SetWriteHostManipulateObjectAction ($ManipulateObject) {
    $Script:WriteHostManipulateObjectAction = $ManipulateObject
}

function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {

        if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
            $Message = & $Script:WriteVerboseManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Verbose $Message

        if ($null -ne $Script:WriteVerboseDebugAction) {
            & $Script:WriteVerboseDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteVerboseDebugAction) {
            & $Script:WriteRemoteVerboseDebugAction $Message
        }
    }
}

function SetWriteVerboseAction ($DebugAction) {
    $Script:WriteVerboseDebugAction = $DebugAction
}

function SetWriteRemoteVerboseAction ($DebugAction) {
    $Script:WriteRemoteVerboseDebugAction = $DebugAction
}

function SetWriteVerboseManipulateMessageAction ($DebugAction) {
    $Script:WriteVerboseManipulateMessageAction = $DebugAction
}

function Write-Warning {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Warning from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )
    process {

        if ($null -ne $Script:WriteWarningManipulateMessageAction) {
            $Message = & $Script:WriteWarningManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Warning $Message

        # Add WARNING to beginning of the message by default.
        $Message = "WARNING: $Message"

        if ($null -ne $Script:WriteWarningDebugAction) {
            & $Script:WriteWarningDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteWarningDebugAction) {
            & $Script:WriteRemoteWarningDebugAction $Message
        }
    }
}

function SetWriteWarningAction ($DebugAction) {
    $Script:WriteWarningDebugAction = $DebugAction
}

function SetWriteRemoteWarningAction ($DebugAction) {
    $Script:WriteRemoteWarningDebugAction = $DebugAction
}

function SetWriteWarningManipulateMessageAction ($DebugAction) {
    $Script:WriteWarningManipulateMessageAction = $DebugAction
}




function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "24.06.13.2304"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path $env:TEMP $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}

    $BuildVersion = "24.06.13.2304"

    $Script:VerboseEnabled = $false
    #this is to set the verbose information to a different color
    if ($PSBoundParameters["Verbose"]) {
        #Write verbose output in cyan since we already use yellow for warnings
        $Script:VerboseEnabled = $true
        $VerboseForeground = $Host.PrivateData.VerboseForegroundColor
        $Host.PrivateData.VerboseForegroundColor = "Cyan"
    }

    $Script:ServerNameList = New-Object System.Collections.Generic.List[string]
    $Script:Logger = Get-NewLoggerInstance -LogName "HealthChecker-Debug" `
        -LogDirectory $Script:OutputFilePath `
        -AppendDateTime $false `
        -ErrorAction SilentlyContinue
    SetProperForegroundColor
    SetWriteVerboseAction ${Function:Write-DebugLog}
    SetWriteWarningAction ${Function:Write-DebugLog}
} process {
    $Server | ForEach-Object { $Script:ServerNameList.Add($_.ToUpper()) }
} end {
    try {

        if (-not (Confirm-Administrator) -and
            (-not $AnalyzeDataOnly -and
            -not $BuildHtmlServersReport -and
            -not $ScriptUpdateOnly)) {
            Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator."
            $Error.Clear()
            Start-Sleep -Seconds 2
            exit
        }

        Invoke-ErrorMonitoring
        $Script:date = (Get-Date)
        $Script:dateTimeStringFormat = $date.ToString("yyyyMMddHHmmss")

        # Some companies might already provide a full path for HtmlReportFile
        # Detect if it is just a name, if it is, then append OutputFilePath to it.
        # Otherwise, keep it as is
        if ($HtmlReportFile.Contains("\")) {
            $htmlOutFilePath = $HtmlReportFile
        } else {
            $htmlOutFilePath = [System.IO.Path]::Combine($OutputFilePath, $HtmlReportFile)
        }

        # Features that doesn't require Exchange Shell
        if ($BuildHtmlServersReport) {
            Invoke-SetOutputInstanceLocation -FileName "HealthChecker-HTMLServerReport"
            $importData = Get-ExportedHealthCheckerFiles -Directory $XMLDirectoryPath

            if ($null -eq $importData) {
                Write-Host "Doesn't appear to be any Health Check XML files here....stopping the script"
                exit
            }
            Get-HtmlServerReport -AnalyzedHtmlServerValues $importData.HtmlServerValues -HtmlOutFilePath $htmlOutFilePath
            Start-Sleep 2
            return
        }

        if ($AnalyzeDataOnly) {
            Invoke-SetOutputInstanceLocation -FileName "HealthChecker-Analyzer"
            $importData = Get-ExportedHealthCheckerFiles -Directory $XMLDirectoryPath

            if ($null -eq $importData) {
                Write-Host "Doesn't appear to be any Health Check XML files here....stopping the script"
                exit
            }

            $analyzedResults = @()
            foreach ($serverData in $importData) {
                $analyzedServerResults = Invoke-AnalyzerEngine -HealthServerObject $serverData.HealthCheckerExchangeServer
                Write-ResultsToScreen -ResultsToWrite $analyzedServerResults.DisplayResults
                $analyzedResults += $analyzedServerResults
            }

            Get-HtmlServerReport -AnalyzedHtmlServerValues $analyzedResults.HtmlServerValues -HtmlOutFilePath $htmlOutFilePath
            return
        }

        if ($ScriptUpdateOnly) {
            Invoke-SetOutputInstanceLocation -FileName "HealthChecker-ScriptUpdateOnly"
            switch (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/HC-VersionsUrl" -Confirm:$false) {
                ($true) { Write-Green("Script was successfully updated.") }
                ($false) { Write-Yellow("No update of the script performed.") }
                default { Write-Red("Unable to perform ScriptUpdateOnly operation.") }
            }
            return
        }

        # Features that do require Exchange Shell
        if ($LoadBalancingReport) {
            Invoke-SetOutputInstanceLocation -FileName "HealthChecker-LoadBalancingReport"
            Invoke-ConfirmExchangeShell
            Write-Grey "Script Version: $BuildVersion"
            Write-Green("Load Balancing Report on " + $date)
            Get-LoadBalancingReport
            Write-Grey("Output file written to " + $Script:OutputFullPath)
            Write-Break
            Write-Break
            return
        }

        if ($DCCoreRatio) {
            $oldErrorAction = $ErrorActionPreference
            $ErrorActionPreference = "Stop"
            try {
                Get-ExchangeDCCoreRatio
                return
            } finally {
                $ErrorActionPreference = $oldErrorAction
            }
        }

        if ($MailboxReport) {
            Invoke-ConfirmExchangeShell

            foreach ($serverName in $Script:ServerNameList) {
                Invoke-SetOutputInstanceLocation -Server $serverName -FileName "HealthChecker-MailboxReport" -IncludeServerName $true
                Get-MailboxDatabaseAndMailboxStatistics -Server $serverName
                Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
            }
            return
        }

        if ($VulnerabilityReport) {
            Invoke-ConfirmExchangeShell
            Invoke-VulnerabilityReport
            return
        }

        # Main Feature of Health Checker
        Invoke-ConfirmExchangeShell
        Invoke-HealthCheckerMainReport -ServerNames $Script:ServerNameList -EdgeServer $Script:ExchangeShellComputer.EdgeServer
    } finally {
        Get-ErrorsThatOccurred
        if ($Script:VerboseEnabled) {
            $Host.PrivateData.VerboseForegroundColor = $VerboseForeground
        }
        $Script:Logger | Invoke-LoggerInstanceCleanup
        if ($Script:Logger.PreventLogCleanup) {
            Write-Host("Output Debug file written to {0}" -f $Script:Logger.FullPath)
        }
        if (((Get-Date).Ticks % 2) -eq 1) {
            Write-Host("Do you like the script? Visit https://aka.ms/HC-Feedback to rate it and to provide feedback.") -ForegroundColor Green
            Write-Host
        }
        RevertProperForegroundColor
    }
}

# SIG # Begin signature block
# MIIn2AYJKoZIhvcNAQcCoIInyTCCJ8UCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDSjO3QrDe6ZrZa
# g/diX0reS9uD/Z8lFvSchbCE0va0qKCCDXYwggX0MIID3KADAgECAhMzAAADrzBA
# DkyjTQVBAAAAAAOvMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjMxMTE2MTkwOTAwWhcNMjQxMTE0MTkwOTAwWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOS8s1ra6f0YGtg0OhEaQa/t3Q+q1MEHhWJhqQVuO5amYXQpy8MDPNoJYk+FWA
# hePP5LxwcSge5aen+f5Q6WNPd6EDxGzotvVpNi5ve0H97S3F7C/axDfKxyNh21MG
# 0W8Sb0vxi/vorcLHOL9i+t2D6yvvDzLlEefUCbQV/zGCBjXGlYJcUj6RAzXyeNAN
# xSpKXAGd7Fh+ocGHPPphcD9LQTOJgG7Y7aYztHqBLJiQQ4eAgZNU4ac6+8LnEGAL
# go1ydC5BJEuJQjYKbNTy959HrKSu7LO3Ws0w8jw6pYdC1IMpdTkk2puTgY2PDNzB
# tLM4evG7FYer3WX+8t1UMYNTAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQURxxxNPIEPGSO8kqz+bgCAQWGXsEw
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMTgyNjAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAISxFt/zR2frTFPB45Yd
# mhZpB2nNJoOoi+qlgcTlnO4QwlYN1w/vYwbDy/oFJolD5r6FMJd0RGcgEM8q9TgQ
# 2OC7gQEmhweVJ7yuKJlQBH7P7Pg5RiqgV3cSonJ+OM4kFHbP3gPLiyzssSQdRuPY
# 1mIWoGg9i7Y4ZC8ST7WhpSyc0pns2XsUe1XsIjaUcGu7zd7gg97eCUiLRdVklPmp
# XobH9CEAWakRUGNICYN2AgjhRTC4j3KJfqMkU04R6Toyh4/Toswm1uoDcGr5laYn
# TfcX3u5WnJqJLhuPe8Uj9kGAOcyo0O1mNwDa+LhFEzB6CB32+wfJMumfr6degvLT
# e8x55urQLeTjimBQgS49BSUkhFN7ois3cZyNpnrMca5AZaC7pLI72vuqSsSlLalG
# OcZmPHZGYJqZ0BacN274OZ80Q8B11iNokns9Od348bMb5Z4fihxaBWebl8kWEi2O
# PvQImOAeq3nt7UWJBzJYLAGEpfasaA3ZQgIcEXdD+uwo6ymMzDY6UamFOfYqYWXk
# ntxDGu7ngD2ugKUuccYKJJRiiz+LAUcj90BVcSHRLQop9N8zoALr/1sJuwPrVAtx
# HNEgSW+AKBqIxYWM4Ev32l6agSUAezLMbq5f3d8x9qzT031jMDT+sUAoCw0M5wVt
# CUQcqINPuYjbS1WgJyZIiEkBMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGbgwghm0AgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAOvMEAOTKNNBUEAAAAAA68wDQYJYIZIAWUDBAIB
# BQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILoK3TScL/eMMUmuZL0e2chl
# NYoXn38wxTvz5CfMGx49MFoGCisGAQQBgjcCAQwxTDBKoBqAGABDAFMAUwAgAEUA
# eABjAGgAYQBuAGcAZaEsgCpodHRwczovL2dpdGh1Yi5jb20vbWljcm9zb2Z0L0NT
# Uy1FeGNoYW5nZSAwDQYJKoZIhvcNAQEBBQAEggEAVwLkB6I4pHfdvotxVTrDGdDH
# 4p9+quSMf9V6IYqduIpFxd4PhY8Vl8CstcztZT7ENKX/tp9z3wrK8webQWsoH+p1
# PvIm59eBW9lNLe+Nv7o3vjP4hosTzEJEJIvKyC7xFm5rkbNk/vXAYfshWY+sltLG
# 25xElUbeUlayv+i042uF6ORH6+s/fTnd/O5QvyTIAyN+jK3cMkAslxnSrUonCUgb
# f64YEdGidyuYln2+yib24QC90ANdJWIGfWxZ9u/mX6DuOaKSTCsbfXtoG6/x81Iu
# CdqFe35qCHWesrAuwuUEAULPnMmXFD9JnaFW/eykWZgnx98my9WhoiPc4eO9sKGC
# FyowghcmBgorBgEEAYI3AwMBMYIXFjCCFxIGCSqGSIb3DQEHAqCCFwMwghb/AgED
# MQ8wDQYJYIZIAWUDBAIBBQAwggFZBgsqhkiG9w0BCRABBKCCAUgEggFEMIIBQAIB
# AQYKKwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCIfIkCIRM4ZA/McBhCEv1c
# aRq+oNjAJPezKHH8TraRlQIGZldT8pHYGBMyMDI0MDYxNDA2MDMxMy45NjlaMASA
# AgH0oIHYpIHVMIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQx
# JjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOkZDNDEtNEJENC1EMjIwMSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloIIReTCCBycwggUPoAMCAQIC
# EzMAAAHimZmV8dzjIOsAAQAAAeIwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTAwHhcNMjMxMDEyMTkwNzI1WhcNMjUwMTEwMTkwNzI1
# WjCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UE
# CxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYwJAYDVQQL
# Ex1UaGFsZXMgVFNTIEVTTjpGQzQxLTRCRDQtRDIyMDElMCMGA1UEAxMcTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBALVjtZhV+kFmb8cKQpg2mzisDlRI978Gb2amGvbAmCd04JVGeTe/QGzM
# 8KbQrMDol7DC7jS03JkcrPsWi9WpVwsIckRQ8AkX1idBG9HhyCspAavfuvz55khl
# 7brPQx7H99UJbsE3wMmpmJasPWpgF05zZlvpWQDULDcIYyl5lXI4HVZ5N6MSxWO8
# zwWr4r9xkMmUXs7ICxDJr5a39SSePAJRIyznaIc0WzZ6MFcTRzLLNyPBE4KrVv1L
# Fd96FNxAzwnetSePg88EmRezr2T3HTFElneJXyQYd6YQ7eCIc7yllWoY03CEg9gh
# orp9qUKcBUfFcS4XElf3GSERnlzJsK7s/ZGPU4daHT2jWGoYha2QCOmkgjOmBFCq
# QFFwFmsPrZj4eQszYxq4c4HqPnUu4hT4aqpvUZ3qIOXbdyU42pNL93cn0rPTTleO
# UsOQbgvlRdthFCBepxfb6nbsp3fcZaPBfTbtXVa8nLQuMCBqyfsebuqnbwj+lHQf
# qKpivpyd7KCWACoj78XUwYqy1HyYnStTme4T9vK6u2O/KThfROeJHiSg44ymFj+3
# 4IcFEhPogaKvNNsTVm4QbqphCyknrwByqorBCLH6bllRtJMJwmu7GRdTQsIx2HMK
# qphEtpSm1z3ufASdPrgPhsQIRFkHZGuihL1Jjj4Lu3CbAmha0lOrAgMBAAGjggFJ
# MIIBRTAdBgNVHQ4EFgQURIQOEdq+7QdslptJiCRNpXgJ2gUwHwYDVR0jBBgwFoAU
# n6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFt
# cCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcw
# AoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3Nv
# ZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIw
# ADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZI
# hvcNAQELBQADggIBAORURDGrVRTbnulfsg2cTsyyh7YXvhVU7NZMkITAQYsFEPVg
# vSviCylr5ap3ka76Yz0t/6lxuczI6w7tXq8n4WxUUgcj5wAhnNorhnD8ljYqbck3
# 7fggYK3+wEwLhP1PGC5tvXK0xYomU1nU+lXOy9ZRnShI/HZdFrw2srgtsbWow9OM
# uADS5lg7okrXa2daCOGnxuaD1IO+65E7qv2O0W0sGj7AWdOjNdpexPrspL2KEcOM
# eJVmkk/O0ganhFzzHAnWjtNWneU11WQ6Bxv8OpN1fY9wzQoiycgvOOJM93od55EG
# eXxfF8bofLVlUE3zIikoSed+8s61NDP+x9RMya2mwK/Ys1xdvDlZTHndIKssfmu3
# vu/a+BFf2uIoycVTvBQpv/drRJD68eo401mkCRFkmy/+BmQlRrx2rapqAu5k0Nev
# +iUdBUKmX/iOaKZ75vuQg7hCiBA5xIm5ZIXDSlX47wwFar3/BgTwntMq9ra6QRAe
# S/o/uYWkmvqvE8Aq38QmKgTiBnWSS/uVPcaHEyArnyFh5G+qeCGmL44MfEnFEhxc
# 3saPmXhe6MhSgCIGJUZDA7336nQD8fn4y6534Lel+LuT5F5bFt0mLwd+H5GxGzOb
# Zmm/c3pEWtHv1ug7dS/Dfrcd1sn2E4gk4W1L1jdRBbK9xwkMmwY+CHZeMSvBMIIH
# cTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCB
# iDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMp
# TWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEw
# OTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAOThpkzntHIh
# C3miy9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNx
# WuJ+Slr+uDZnhUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFc
# UTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAc
# nVL+tuhiJdxqD89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUo
# veO0hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyzi
# YrLNueKNiOSWrAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9
# fvzZnkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdH
# GO2n6Jl8P0zbr17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7X
# KHYC4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiE
# R9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/
# eKtFtvUeh17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3
# FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAd
# BgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEE
# AYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMI
# MBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMB
# Af8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1Ud
# HwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3By
# b2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQRO
# MEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2Vy
# dHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4IC
# AQCdVX38Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pk
# bHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gng
# ugnue99qb74py27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3
# lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHC
# gRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6
# MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEU
# BHG/ZPkkvnNtyo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvsh
# VGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+
# fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrp
# NPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHI
# qzqKOghif9lwY1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGCAtUwggI+AgEBMIIB
# AKGB2KSB1TCB0jELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
# BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEt
# MCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9ucyBMaW1pdGVkMSYw
# JAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpGQzQxLTRCRDQtRDIyMDElMCMGA1UEAxMc
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUAFpuZ
# afp0bnpJdIhfiB1d8pTohm+ggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOoWPwAwIhgPMjAyNDA2MTQxMjAxMDRa
# GA8yMDI0MDYxNTEyMDEwNFowdTA7BgorBgEEAYRZCgQBMS0wKzAKAgUA6hY/AAIB
# ADAHAgEAAgIIXTAIAgEAAgMBRe8wCgIFAOoXkIACAQAwNgYKKwYBBAGEWQoEAjEo
# MCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG
# 9w0BAQUFAAOBgQAZ384eohJy03wdNre1TVu+bgzQJrSWU1VjjnQonX55jtwSm3F2
# AQYVu8FuQOeolZSBt4WCOyPFYPJgzaBTut+22hZ2v/+OBsdfomYfRm8StaUvUgdx
# DUOoZybmTjOsx3hw2Tt0LO6Y/ScQKnSj2qBlubNEJV2vGiaJ6uO+yYylmzGCBA0w
# ggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB4pmZ
# lfHc4yDrAAEAAAHiMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYL
# KoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIM2ExGYBvxGy8IVg+SYibC3YJzlz
# II1r6NXsXvYSi80cMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgK4kqShD9
# JrjGwVBEzg6C+HeS1OiP247nCGZDiQiPf/8wgZgwgYCkfjB8MQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGlt
# ZS1TdGFtcCBQQ0EgMjAxMAITMwAAAeKZmZXx3OMg6wABAAAB4jAiBCC1eX+iD4dc
# IoWqYMaOprwtmM4lw9MS4CwEELLEmzZUXDANBgkqhkiG9w0BAQsFAASCAgAvhXbx
# wbmuxacAwwJqi8O/rFQjxoJoBp6a2x8wQjjanneps/9s9kajX1LJoEfhjxPuZhuy
# hbLfMtguueoxoDJ+hoOnPzx3bYh4cn8Rtz0KMXfelYJIKBMNVUV4N8EmvP/+8sA6
# yLLCbIqSypximpkoAYe9vaiRJcJvTO75j9imqia+TP1gxhhNQ0mmmGkW9y/930PW
# ZUsjk8Wunlo4sqfSzG5VIQER0YzOOBgwH4eDtDq2bz890HZAQPDzSDZhEvKVF1m6
# 032QN6bgP09iiDau8jfV2+RwDO/lDwdN4Lu/oLcCZ5jZ5OHnBzLa6EPgp4eDB0cn
# yVKpQfKlmEq9YY99MGCCCMDakMdfy0IVg3HALtM2jFNqa7vXKF+f3j2DzNZ6+ry3
# ZqYNbjipsXPAlEte8EWklo+R1/AFyafoC7PaFIBnC3ukWrVpcrC3RovOm3T2+m+k
# MsOlyA+pBFvhEVh0sMfqY+FhrQ3TF4+gT27O8ZKIhIOa3QgXKjYzyjELhT2dzsQq
# RUDwjCx5vQeferRsj3/8qru3Zody7Kz3d8hqXTF1SpWz53v0aGT6Ym8q6gWhaWuZ
# F7jCCpaJZbuJiJ7XW6t7us6vvt6WvMK5aZm0TGDBtyLAuIl+Ntnc9JyYzNxcItFx
# g8VDScOCvk26dgjHnSIxBOwCwi6EkOJcuuUrbQ==
# SIG # End signature block
