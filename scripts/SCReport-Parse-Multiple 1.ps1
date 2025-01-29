#import file(s)
Write-Host "Importing Raw Data" -ForegroundColor Cyan
$headers = "Plugin","Severity", "IP", "MAC","DNS", "Plugin Number", "LastObserved", "Text", "Description"
$directory = "."
$children = Get-ChildItem -Path $directory -Filter "*.csv"

#loops through each CSV file in the $directory matching "TDL - Enumerable*.csv"
foreach($child in $children){
    if($child.Exists){
        $raw = Import-Csv (Join-Path $directory $child.Name) -Header $headers #import action
    }
    #establish lists for tracking results
    $uniqueDNS = New-Object "System.Collections.ArrayList"
    $KeptItems = New-Object "System.Collections.ArrayList"
    $outputList = New-Object "System.Collections.ArrayList"
    #$OSList = Import-Csv invList.txt -Header "Name", "OS"

    #establish hash tables for easy sorts, duplicate checks, and quick counts
    $checksPerDNS = @{}
    $failsPerDNS = @{}
    $inconclusivePerDNS = @{}

    #begin main loop
    $progressNum = 1
    Write-Host "Extracting and Reformatting..." -ForegroundColor Cyan
    Foreach($item in $raw){

        #extract DNS name for asset, ignore column header
        $DNSName = $item.DNS.Split('.')[0]
        if($DNSName -notmatch ".*DNS Name.*"){

            #ensure unique name, never adds same asset twice. Add to all lists
            if(-not ($uniqueDNS -contains $DNSName)){
                $quiet = $uniqueDNS.Add($DNSName)
                $checksPerDNS.Add($DNSName, 0)
                $failsPerDNS.Add($DNSName, 0)
                $inconclusivePerDNS.Add($DNSName, 0)
            }

            #collect check, ignore the nessus defaults
            if (($item.plugin -notmatch "Netstat") -AND ($item.Plugin -notmatch "Nessus")){# -AND ($item.Plugin -notmatch "Plugin")){

                #extract policy from plugin name
                #$policy = $item.Plugin.Split(" V-")[0]
                #$item | add-member -MemberType NoteProperty -Name "Policy" -value $policyNames[$policy]

                #shorten plugin name
                #$item.Plugin = $item.Plugin.Replace($policy + ' ', "")

                #find OS
                #$DNSName = $item.DNS.Split('.')[0]
                #$item | add-member -MemberType NoteProperty -Name "Operating System" -value $OSHash[$DNSName]

                #pass or fail result
                if($item.Severity -match "High"){
                    $item.Severity = "Noncompliant"
                    $quiet = $KeptItems.Add($item)
                    $failsPerDNS[$DNSName] = $failsPerDNS[$DNSName] + 1
                }
                elseif($item.Severity -match "Medium"){
                    $item.Severity = "Inconclusive"
                    $quiet = $KeptItems.Add($item)
                }
                elseif($item.Severity -match "Info"){
                    $item.Severity = "Passed" #we mostly ignore passing results. They're not important
                }

                #$quiet = $KeptItems.Add($item)

                #this if statement can be used to select specific checks by plugin name. Include or exclude as necessary
                if($item.plugin -match ".*"){
                    $checksPerDNS[$DNSName] = $checksPerDNS[$DNSName] + 1

                    #nessus has "stub" checks that provide no value. They're not useful for statistics
                    if($item.Severity -match "Inconclusive"){
                        if($item.Text -notmatch ".*NOTE\: Nessus has not performed this check\.*" -and $item.Text -notmatch ".*NOTE\: Nessus has provided the target output.*"){
                            $inconclusivePerDNS[$DNSName] = $inconclusivePerDNS[$DNSName] + 1
                        }#not a report
                    } #severity

                }#within scope
            }

            #this literally just spits out little progress dots while the script runs
            if(($progressNum % 10000) -eq 0){
                Write-Host . -NoNewLine
            }
            $progressNum++

        }#dns check
    }#for

    Write-Host "`nReformatting Complete" -ForegroundColor Green
    #Write-Host "Exporting..." -ForegroundColor Cyan

    #loop through checks, create our CSV output rows
    $totalChecks = $checksPerDNS.GetEnumerator() | Select-Object Key
    foreach($c in $totalChecks){
        #Write-Host $c.Key"-"$checksPerDNS[$c.Key]"checks performed"
        $out = @{
            DNSName = $c.Key
            Checks = $checksPerDNS[$c.key]
            Fails = $failsPerDNS[$c.Key]
            Inconclusives = $inconclusivePerDNS[$c.Key]
            }
        $outputList += New-Object psobject -Property $out
    }

    #export and output
    $exportName = $child[0].Name.Trim(".csv") + "_MetaData.csv"
    $outputlist | Export-Csv $exportName -NoTypeInformation
    Write-Host "Exporting "$exportName
    $exportName = $child[0].Name.Trim(".csv") + "_FailedChecks.csv"
    $KeptItems | Export-Csv $exportName -NoTypeInformation
    Write-Host "Exporting "$exportName
}
