Add-Type -AssemblyName System.Drawing

$shell = New-Object -ComObject Shell.Application
$appsFolder = $shell.NameSpace("shell:AppsFolder")
$csvPath = "app_list.csv"
$htmlPath = "app_list.html"
$results = @()
$index = 1

# ====================================================================
# ====================================================================
$iconDir = "app_icons"
if (-not (Test-Path -LiteralPath $iconDir)) {
    New-Item -ItemType Directory -Path $iconDir | Out-Null
}

foreach ($item in $appsFolder.Items()) {
    # Write-Host "`n=================================================="
    # Write-Host "App : $($item.Name)" -ForegroundColor Cyan

    $entry = [PSCustomObject][ordered]@{
        "No."       = ""
        Icon        = ""
        AppName     = $item.Name
        Location    = "---"
        Type        = "---"
        Target      = "---"
        Desc        = "---"
        ProductName = "---"   
        SizeMB      = "---"
        create_time = "---"   
        fix_time    = "---"   
        visit_time  = "---"   
    }
    $entry."No." = $index++

    try {
        $target = $null
        $extractedScript = $null

        foreach ($v in $item.Verbs()) {
            if ($v.Name -match "Location: " -or $v.Name -match "Open file location") {
                $target = $item.Path
                break
            }
        }

        if (-not $target -or $target -eq "") {
            $target = $item.ExtendedProperty("System.TargetPath")
        }
        if (-not $target -or $target -eq "") {
            $target = $item.ExtendedProperty("System.Link.TargetParsingPath")
        }
        if (-not $target -or $target -eq "") {
            $target = $item.Path
        }

        if (-not $target -or $target -eq "") {
            $results += $entry
            # Write-Host "Info: System/UWP App (no file location)" -ForegroundColor Gray
            continue
        }

        $arguments = $item.ExtendedProperty("System.Link.Arguments")
        
        # ====================================================================
        # ====================================================================
        
        if ($target -match "(?i)(cmd\.exe|%comspec%)$" -and $arguments) {
            $entry.Type = "Shell Command"
            $entry.Location = "$target $arguments"
            
            if ($arguments -match '(?i)([a-zA-Z]:\\[^"<>|?*]+\.(?:bat|cmd|exe|ps1|vbs))') {
                $extractedScript = $matches[1].Trim()
            } 
            elseif ($arguments -match '(?i)([a-zA-Z]:\\[^;"<>|?*]+)' -and $arguments -match '(?i)([a-zA-Z0-9_.-]+\.(?:bat|cmd|exe|ps1|vbs))') {
                $possibleDir = $matches[1].TrimEnd('\')
                $possibleFile = $matches[2]
                $testPath = Join-Path $possibleDir $possibleFile
                if (Test-Path -LiteralPath $testPath -PathType Leaf -ErrorAction SilentlyContinue) {
                    $extractedScript = $testPath
                }
            }
            # Write-Host "Type   : Shell Command (Intercepted!)" -ForegroundColor DarkYellow
            # Write-Host "Command: $($entry.Location)" -ForegroundColor DarkYellow
        }
        elseif ($target -match "(?i)\.lnk$") {
            $lnk = $shell.CreateShortcut($target)
            if ($lnk.TargetPath -match "^(https?|file)://") {
                $entry.Type = "URL"
                $entry.Target = $lnk.TargetPath
                # Write-Host "Type   : URL" -ForegroundColor Cyan
                # Write-Host "Target : $($entry.Target)"
            } else {
                $entry.Type = "Shortcut"
                $entry.Target = $lnk.TargetPath
                $entry.Location = $target
                # Write-Host "Type   : Shortcut" -ForegroundColor Yellow
                # Write-Host "Target : $($lnk.TargetPath)"
            }
        }
        elseif ($target -match "(?i)\.exe$") {
            $entry.Type = "Executable"
            $entry.Location = $target
            # Write-Host "Type   : Executable" -ForegroundColor Magenta
            # Write-Host "Location: $target" -ForegroundColor Green
        }
        elseif ($target -match "(?i)\.url$") {
            $entry.Type = "URL"
            $entry.Location = $target
            $urlContent = Get-Content -LiteralPath $target -Raw -ErrorAction SilentlyContinue
            if ($urlContent -match '(?i)URL\s*=\s*([^\r\n]+)') {
                $entry.Target = $matches[1].Trim()
            }
            # Write-Host "Type   : URL" -ForegroundColor Cyan
            # Write-Host "Target : $($entry.Target)"
        }
        elseif ($target -match "(?i)\.(bat|cmd)$") {
            $entry.Type = "bat-Script"
            $entry.Location = $target
            # Write-Host "Type   : bat-Script" -ForegroundColor Green
            # Write-Host "Location: $target" -ForegroundColor Green
        }
        elseif ($target -match "^(https?|file)://") {
            $entry.Type = "URL"
            $entry.Target = $target
            $entry.Location = $target
            # Write-Host "Type   : URL (Direct Link)" -ForegroundColor Cyan
            # Write-Host "Target : $target"
        }
        elseif ($target -match "(?i)\.html?$") {
            $entry.Type = "URL"
            $entry.Location = $target
            $entry.Target = $target
            if (Test-Path -LiteralPath $target -PathType Leaf) {
                $htmlContent = Get-Content -LiteralPath $target -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                if ($htmlContent -match '(?i)<base\s+href\s*=\s*["''](https?://[^"'']+)["'']') {
                    $entry.Target = $matches[1].Trim()
                } elseif ($htmlContent -match '(?i)(https?://[^\s"''<>]+)') {
                    $entry.Target = $matches[1].Trim()
                }
            }
            # Write-Host "Type   : URL (Local Web Document)" -ForegroundColor Cyan
            # Write-Host "Target : $($entry.Target)"
            # Write-Host "Location: $target" -ForegroundColor Green
        }
        else {
            $entry.Type = "Other"
            $entry.Location = $target
            # Write-Host "Type   : Other" -ForegroundColor Gray
            # Write-Host "Location: $target" -ForegroundColor Green
        }

        $finalTarget = if ($entry.Type -eq "Shell Command") {
                           if ($extractedScript -match "(?i)\.exe$") { $extractedScript } else { $null }
                       } 
                       elseif ($entry.Target -and ($entry.Target -match "(?i)\.exe$")) { $entry.Target } 
                       elseif ($target -match "(?i)\.exe$") { $target }
                       else { $null }

        if ($finalTarget -and (Test-Path -LiteralPath $finalTarget -ErrorAction SilentlyContinue)) {
            $fileInfo = Get-Item -LiteralPath $finalTarget
            $versionInfo = $fileInfo.VersionInfo
            
            $entry.ProductName = if ([string]::IsNullOrWhiteSpace($versionInfo.ProductName)) { "---none---" } 
                                 else { $versionInfo.ProductName.Trim() }

            $tempDesc = if (-not [string]::IsNullOrWhiteSpace($versionInfo.FileDescription)) { $versionInfo.FileDescription } 
                        else { $versionInfo.ProductName }
            $entry.Desc = if ([string]::IsNullOrWhiteSpace($tempDesc)) { "---none---" } 
                          else { $tempDesc.Trim() }

            $entry.SizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        }

        # ====================================================================
        # ====================================================================
        $timeTarget = $null
        
        if ($entry.Type -eq "Shell Command" -and $extractedScript) {
            $timeTarget = $extractedScript
        } 
        elseif ($finalTarget) {
            $timeTarget = $finalTarget
        } 
        elseif ($entry.Target -and (Test-Path -LiteralPath $entry.Target -PathType Leaf -ErrorAction SilentlyContinue)) {
            $timeTarget = $entry.Target
        } 
        elseif ($entry.Location -and (Test-Path -LiteralPath $entry.Location -PathType Leaf -ErrorAction SilentlyContinue)) {
            $timeTarget = $entry.Location
        } 
        elseif ($target -and (Test-Path -LiteralPath $target -PathType Leaf -ErrorAction SilentlyContinue)) {
            $timeTarget = $target
        }

        if (-not $timeTarget -and ($item.Path -and (Test-Path -LiteralPath $item.Path -PathType Leaf -ErrorAction SilentlyContinue))) {
            $timeTarget = $item.Path
            # Write-Host "Time   : Fallback to shortcut file ($item.Path)" -ForegroundColor DarkGray
        }

        if ($timeTarget -and (Test-Path -LiteralPath $timeTarget -ErrorAction SilentlyContinue)) {
            $timeInfo = Get-Item -LiteralPath $timeTarget
            $entry.create_time = $timeInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
            $entry.fix_time    = $timeInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            $entry.visit_time  = $timeInfo.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
        } else {
            if ($entry.Type -eq "URL" -and $item.Path -notmatch "(?i)\.url$" -and $target -notmatch "(?i)\.url$") {
                $entry.create_time = "no actual file"
                $entry.fix_time    = "no actual file"
                $entry.visit_time  = "no actual file"
            }
        }

        if ($entry.Type -eq "Shell Command" -and $entry.create_time -eq "---") {
            $specialMsg = "no readily identifiable .exe/.bat file."
            $entry.create_time = $specialMsg
            $entry.fix_time    = $specialMsg
            $entry.visit_time  = $specialMsg
            # Write-Host "Time   : $specialMsg" -ForegroundColor Gray
        }

         if ($entry.create_time -eq "---") {
            $entry.create_time = $entry.fix_time = $entry.visit_time = "system component or other reason"
        }

        # ====================================================================
        # ====================================================================
        $iconSource = $null
        
        if ($entry.Type -eq "Shell Command") {
            if ($extractedScript -match "(?i)\.exe$" -and (Test-Path -LiteralPath $extractedScript -PathType Leaf -ErrorAction SilentlyContinue)) {
                $iconSource = $extractedScript
            } else {
                $iconSource = "$env:windir\System32\cmd.exe"
            }
        } 
        elseif ($entry.Type -eq "bat-Script") {
            $iconSource = "$env:windir\System32\cmd.exe"
        } 
        elseif ($finalTarget -and (Test-Path -LiteralPath $finalTarget -PathType Leaf -ErrorAction SilentlyContinue)) {
            $iconSource = $finalTarget
        } 
        elseif ($entry.Location -and (Test-Path -LiteralPath $entry.Location -PathType Leaf -ErrorAction SilentlyContinue) -and $entry.Location -match "(?i)\.exe$") {
            $iconSource = $entry.Location
        } 
        elseif ($target -and (Test-Path -LiteralPath $target -PathType Leaf -ErrorAction SilentlyContinue) -and $target -match "(?i)\.exe$") {
            $iconSource = $target
        }

        if (-not $iconSource -and ($item.Path -and (Test-Path -LiteralPath $item.Path -PathType Leaf -ErrorAction SilentlyContinue))) {
            $iconSource = $item.Path
        }

        if ($iconSource) {
            $safeName = $item.Name -replace '[\\/:*?"<>|]', '_'
            $saveIconPath = Join-Path $iconDir "$safeName.png"

            if (-not (Test-Path -LiteralPath $saveIconPath)) {
                try {
                    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconSource)
                    if ($icon) {
                        $bitmap = $icon.ToBitmap()
                        $bitmap.Save($saveIconPath, [System.Drawing.Imaging.ImageFormat]::Png)
                        $bitmap.Dispose()
                        $icon.Dispose()
                    }
                } catch {}
            }

            if (Test-Path -LiteralPath $saveIconPath) {
                $entry.Icon = "$iconDir\$safeName.png"
            }
        }

        if ([string]::IsNullOrWhiteSpace($entry.Icon)) {
            $defaultIconName = "WindowsDefaultIcon"
            $defaultIconPath = Join-Path $iconDir "$defaultIconName.png"

            if (-not (Test-Path -LiteralPath $defaultIconPath)) {
                try {
                    $defaultIcon = [System.Drawing.Icon]::ExtractAssociatedIcon("$env:windir\System32\shell32.dll")
                    if ($defaultIcon) {
                        $defaultBitmap = $defaultIcon.ToBitmap()
                        $defaultBitmap.Save($defaultIconPath, [System.Drawing.Imaging.ImageFormat]::Png)
                        $defaultBitmap.Dispose()
                        $defaultIcon.Dispose()
                    }
                } catch {}
            }

            if (Test-Path -LiteralPath $defaultIconPath) {
                $entry.Icon = $defaultIconPath
            }
        }

    }
    catch {
        $entry.Target = $_.Exception.Message
        # Write-Host "Error  : $_" -ForegroundColor Red
    }

    $results += $entry
}

# ====================================================================
# ====================================================================
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$htmlHeader = @"
<style>
table { border-collapse:collapse; width:100%; font-family:Microsoft YaHei, sans-serif; }
td { padding:8px; border:1px solid #ccc; text-align: left; vertical-align: middle; white-space: nowrap; }
th { padding:8px; border:1px solid #ccc; background-color:#f4f4f4; text-align: left; font-weight: bold; white-space: nowrap; } 
img { border-radius: 4px; }
</style>
<table>
"@

$htmlHeader | Out-File $htmlPath -Encoding UTF8

if ($results.Count -gt 0) {
    $ths = ($results[0].PSObject.Properties | ForEach-Object { "<th>$($_.Name)</th>" }) -join ''
    "<tr>$ths</tr>" | Out-File $htmlPath -Encoding UTF8 -Append
}

foreach ($row in $results) {
    $tds = ($row.PSObject.Properties | ForEach-Object { 
        if ($_.Name -eq "Icon") {
            if ($_.Value) {
                "<td><img src='$($_.Value)' width='32' height='32'/></td>"
            } else {
                "<td></td>"
            }
        } else {
            $safeValue = [string]$_.Value -replace '<', '&lt;' -replace '>', '&gt;'
            "<td>$safeValue</td>" 
        }
    }) -join ''
    
    "<tr>$tds</tr>" | Out-File $htmlPath -Encoding UTF8 -Append
}

"</table>" | Out-File $htmlPath -Encoding UTF8 -Append


Write-Host "ok!!!" -ForegroundColor Green
# Write-Host "CSV  : $csvPath"
# Write-Host "HTML : $htmlPath" -ForegroundColor Green
# Write-Host "Icons: $iconDir folder" -ForegroundColor Green
# perfect version 350 line