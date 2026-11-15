} -ErrorAction Stop
                    foreach ($cert in @($remoteCerts)) {
                        if ($cert.NotAfter -lt (Get-Date)) {
                            Add-Finding PKI Critical 'Expired certificate in domain controller machine store' $dc.HostName "$($cert.Subject); expires $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Determine certificate purpose and renew or replace it; verify dependent services.'
                        } elseif ($cert.NotAfter -lt (Get-Date).AddDays($CertificateWarningDays)) {
                            Add-Finding PKI High 'Domain controller certificate expiring soon' $dc.HostName "$($cert.Subject); expires $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Renew the certificate and verify auto-enrollment, template permissions, and chain/revocation availability.'
                        }
                    }
                }
            } catch {
                Add-AssessmentError PKI $dc.HostName "Remote certificate-store query unavailable: $($_.Exception.Message)"
            }
        }
    } catch { Add-AssessmentError PKI $forest.Name $_.Exception.Message }

    # Score categories and calculate weighted overall score.
    $weights = [ordered]@{
        Users = 0.15
        Computers = 0.15
        'Privileged Groups' = 0.20
        'Domain Controllers' = 0.15
        DNS = 0.10
        Replication = 0.10
        PKI = 0.15
    }
    $categoryRows = foreach ($category in $weights.Keys) {
        $score = Get-CategoryScore $category
        $script:CategoryScores[$category] = $score
        [pscustomobject]@{
            Category = $category
            Score = $score
            WeightPercent = [math]::Round($weights[$category] * 100,0)
            WeightedContribution = [math]::Round($score * $weights[$category],2)
            Status = Get-Status $score
            FindingCount = @($script:Findings | Where-Object Category -eq $category).Count
        }
    }
    $overall = [math]::Round((($categoryRows | Measure-Object WeightedContribution -Sum).Sum),2)

    # Critical overrides: serious findings cap the overall score. These are deliberately conservative
    # starter rules and should be tuned against approved organizational baselines.
    $criticalCount = @($script:Findings | Where-Object Severity -eq 'Critical').Count
    $scoreCapReason = $null
    if (@($script:Findings | Where-Object { $_.Category -eq 'Privileged Groups' -and $_.Severity -eq 'Critical' }).Count -gt 0) {
        $overall = [math]::Min($overall,49); $scoreCapReason = 'Critical privileged-access finding'
    }
    if (@($script:Findings | Where-Object { $_.Category -eq 'Replication' -and $_.Severity -eq 'Critical' }).Count -ge 2) {
        $overall = [math]::Min($overall,49); $scoreCapReason = 'Multiple critical replication failures'
    }
    $overallStatus = Get-Status $overall

    $summary = [pscustomobject]@{
        RunId = $script:RunId
        Started = $script:RunStarted
        Completed = Get-Date
        Forest = $forest.Name
        RootDomain = $forest.RootDomain
        CurrentDomain = $currentDomain.DNSRoot
        Domains = ($domains -join '; ')
        DomainControllerCount = $allDCs.Count
        UserCount = $userCount
        ComputerCount = $computerCount
        OverallScore = $overall
        OverallStatus = $overallStatus
        ScoreCapReason = $scoreCapReason
        CriticalFindings = $criticalCount
        HighFindings = @($script:Findings | Where-Object Severity -eq 'High').Count
        MediumFindings = @($script:Findings | Where-Object Severity -eq 'Medium').Count
        LowFindings = @($script:Findings | Where-Object Severity -eq 'Low').Count
        AssessmentErrors = $script:AssessmentErrors.Count
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $jsonPath = Join-Path $OutputPath "ADSHA_$stamp.json"
    $csvPath = Join-Path $OutputPath "ADSHA_Findings_$stamp.csv"
    $htmlPath = Join-Path $OutputPath "ADSHA_Dashboard_$stamp.html"
    $report = [pscustomobject]@{
        Summary = $summary
        Categories = @($categoryRows)
        Findings = @($script:Findings)
        AssessmentErrors = @($script:AssessmentErrors)
        DomainControllers = @($allDCs | Select-Object HostName,Domain,Site,IsGlobalCatalog,OperationMasterRoles,IPv4Address)
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    @($script:Findings) | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

    $categoryHtml = ($categoryRows | ForEach-Object {
        $cls = if ($_.Score -ge 90) {'good'} elseif ($_.Score -ge 80) {'warn'} else {'bad'}
        "<tr><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$($_.Score)</td><td>$($_.WeightPercent)%</td><td>$($_.FindingCount)</td><td class='$cls'>$(ConvertTo-HtmlSafe $_.Status)</td></tr>"
    }) -join "`n"
    $findingHtml = (@($script:Findings | Sort-Object @{Expression={switch ($_.Severity) {'Critical'{0};'High'{1};'Medium'{2};'Low'{3};default{4}}}},Category) | ForEach-Object {
        "<tr><td>$(ConvertTo-HtmlSafe $_.Severity)</td><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$(ConvertTo-HtmlSafe $_.Title)</td><td>$(ConvertTo-HtmlSafe $_.Object)</td><td>$(ConvertTo-HtmlSafe $_.Details)</td><td>$(ConvertTo-HtmlSafe $_.Recommendation)</td></tr>"
    }) -join "`n"
    $errorHtml = (@($script:AssessmentErrors) | ForEach-Object {
        "<tr><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$(ConvertTo-HtmlSafe $_.Target)</td><td>$(ConvertTo-HtmlSafe $_.Message)</td></tr>"
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>ADSHA - Active Directory Health & Security</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#101827;color:#e5e7eb;margin:0;padding:28px}
h1,h2{color:#f8fafc}.sub{color:#9ca3af}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin:22px 0}
.card{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:18px}.score{font-size:38px;font-weight:700}
.good{color:#4ade80}.warn{color:#facc15}.bad{color:#f87171}
table{width:100%;border-collapse:collapse;background:#1e293b;margin:12px 0 28px;font-size:13px}
th,td{padding:10px;border-bottom:1px solid #334155;text-align:left;vertical-align:top}th{background:#273449;color:#cbd5e1}
section{margin-top:30px}.pill{padding:4px 8px;border-radius:5px;background:#334155}
</style></head><body>
<h1>Active Directory Security &amp; Health Analyzer</h1>
<div class="sub">Forest: $(ConvertTo-HtmlSafe $forest.Name) | Assessment: $(ConvertTo-HtmlSafe $script:RunStarted.ToString('yyyy-MM-dd HH:mm:ss')) | Run ID: $($script:RunId)</div>
<div class="grid">
<div class="card"><div class="sub">Overall Score</div><div class="score $(if($overall -ge 80){'good'}elseif($overall -ge 65){'warn'}else{'bad'})">$overall / 100</div><div>$(ConvertTo-HtmlSafe $overallStatus)</div></div>
<div class="card"><div class="sub">Domain Controllers</div><div class="score">$($allDCs.Count)</div></div>
<div class="card"><div class="sub">User Accounts</div><div class="score">$userCount</div></div>
<div class="card"><div class="sub">Computer Accounts</div><div class="score">$computerCount</div></div>
<div class="card"><div class="sub">Critical Findings</div><div class="score bad">$criticalCount</div></div>
<div class="card"><div class="sub">Assessment Errors</div><div class="score"> $($script:AssessmentErrors.Count)</div></div>
</div>
$(if($scoreCapReason){"<p class='bad'><strong>Score cap applied:</strong> $(ConvertTo-HtmlSafe $scoreCapReason)</p>"}else{''})
<section><h2>Category Scores</h2><table><tr><th>Category</th><th>Score</th><th>Weight</th><th>Findings</th><th>Status</th></tr>$categoryHtml</table></section>
<section><h2>Security &amp; Health Findings</h2><table><tr><th>Severity</th><th>Category</th><th>Finding</th><th>Object</th><th>Details</th><th>Recommendation</th></tr>$findingHtml</table></section>
<section><h2>Assessment Errors / Incomplete Checks</h2><p class="sub">Errors indicate checks that could not be completed; they are not automatically treated as healthy.</p><table><tr><th>Category</th><th>Target</th><th>Message</th></tr>$errorHtml</table></section>
</body></html>
"@
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8

    Write-Host ''
    Write-Host "ADSHA assessment complete." -ForegroundColor Green
    Write-Host "Forest: $($forest.Name)"
    Write-Host "Overall score: $overall ($overallStatus)"
    Write-Host "Findings: $($script:Findings.Count) | Collection errors: $($script:AssessmentErrors.Count)"
    Write-Host "HTML: $htmlPath"
    Write-Host "CSV:  $csvPath"
    Write-Host "JSON: $jsonPath"
    Write-Host ''
    Write-Host 'This is an initial baseline model. Review findings and collection errors before treating the score as an authoritative security rating.' -ForegroundColor Yellow
    Start-Process $htmlPath
}
catch {
    Write-Error "ADSHA failed: $($_.Exception.Message)"
    throw
}
