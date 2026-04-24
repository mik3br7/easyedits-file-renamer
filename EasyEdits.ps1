param(
    [string]$Folder = ".",
    [string]$Pattern = "*",
    [string]$Prefix = "",
    [string]$Suffix = "",
    [string]$Find = "",
    [string]$Replace = "",
    [Nullable[int]]$NumberFrom = $null,
    [switch]$Recursive,
    [switch]$Apply
)

function Get-RenamePlan {
    param(
        [Parameter(Mandatory)]
        [string]$Folder,
        [string]$Pattern = "*",
        [string]$Prefix = "",
        [string]$Suffix = "",
        [string]$Find = "",
        [string]$Replace = "",
        [Nullable[int]]$NumberFrom = $null,
        [bool]$Recursive = $false,
        [bool]$IncludeDirectories = $false
    )

    $targetFolder = Resolve-Path -LiteralPath $Folder -ErrorAction SilentlyContinue
    if (-not $targetFolder) {
        throw "Folder not found: $Folder"
    }

    if ($IncludeDirectories) {
        $items = Get-ChildItem -LiteralPath $targetFolder.Path -Filter $Pattern -Recurse:$Recursive |
            Sort-Object FullName
    } else {
        $items = Get-ChildItem -LiteralPath $targetFolder.Path -File -Filter $Pattern -Recurse:$Recursive |
            Sort-Object FullName
    }

    $planned = @()

    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items[$i]
        $extension = if ($item.PSIsContainer) { "" } else { $item.Extension }
        $stem = if ($item.PSIsContainer) {
            $item.Name
        } else {
            [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
        }

        if ($Find -ne "") {
            $stem = $stem.Replace($Find, $Replace)
        }

        if ($null -ne $NumberFrom) {
            $stem = "{0:D3}_{1}" -f ($NumberFrom + $i), $stem
        }

        $newName = "{0}{1}{2}{3}" -f $Prefix, $stem, $Suffix, $extension
        $targetPath = Join-Path (Split-Path -Parent $item.FullName) $newName

        $planned += [PSCustomObject]@{
            Source = $item.FullName
            SourceName = $item.Name
            Target = $targetPath
            TargetName = $newName
            WillChange = $item.FullName -ne $targetPath
            IsDirectory = $item.PSIsContainer
            Extension = $extension
            CreationTime = $item.CreationTime
            LastWriteTime = $item.LastWriteTime
            Length = if ($item.PSIsContainer) { $null } else { $item.Length }
            DirectoryPath = Split-Path -Parent $item.FullName
        }
    }

    return $planned
}

function Test-RenamePlan {
    param(
        [Parameter(Mandatory)]
        [object[]]$Plan
    )

    $duplicateTargets = $Plan | Group-Object TargetName | Where-Object Count -gt 1
    if ($duplicateTargets) {
        $names = $duplicateTargets | ForEach-Object { $_.Name }
        throw "Rename stopped because multiple files would end up with the same name:`n  - $($names -join "`n  - ")"
    }

    $conflicts = $Plan | Where-Object { $_.Target -ne $_.Source -and (Test-Path -LiteralPath $_.Target) }
    if ($conflicts) {
        $names = $conflicts | ForEach-Object { $_.TargetName }
        throw "Rename stopped because these target files already exist:`n  - $($names -join "`n  - ")"
    }
}

function Invoke-RenamePlan {
    param(
        [Parameter(Mandatory)]
        [object[]]$Plan,
        [switch]$Apply
    )

    Test-RenamePlan -Plan $Plan

    $changedCount = 0
    $appliedItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Plan) {
        Write-Host "$($item.SourceName) -> $($item.TargetName)"
        if ($item.WillChange) {
            $changedCount++
            if ($Apply) {
                try {
                    Rename-Item -LiteralPath $item.Source -NewName $item.TargetName -ErrorAction Stop
                    [void]$appliedItems.Add($item)
                } catch {
                    $rollbackFailures = New-Object System.Collections.Generic.List[string]

                    for ($i = $appliedItems.Count - 1; $i -ge 0; $i--) {
                        $appliedItem = $appliedItems[$i]

                        try {
                            Rename-Item -LiteralPath $appliedItem.Target -NewName $appliedItem.SourceName -ErrorAction Stop
                        } catch {
                            $rollbackFailures.Add(
                                ("{0} -> {1}: {2}" -f $appliedItem.TargetName, $appliedItem.SourceName, $_.Exception.Message)
                            )
                        }
                    }

                    $failureMessage = "Rename failed for '$($item.SourceName)' -> '$($item.TargetName)'."
                    if ($appliedItems.Count -gt 0) {
                        $failureMessage += " Earlier renames were rolled back."
                    }

                    if ($rollbackFailures.Count -gt 0) {
                        $failureMessage += " Some rollback steps also failed:`n  - $($rollbackFailures -join "`n  - ")"
                    } else {
                        $failureMessage += " No changes were left partially applied."
                    }

                    $failureMessage += "`nOriginal error: $($_.Exception.Message)"
                    throw $failureMessage
                }
            }
        }
    }

    return $changedCount
}

function Invoke-FileRenamerCli {
    try {
        $plan = Get-RenamePlan `
            -Folder $Folder `
            -Pattern $Pattern `
            -Prefix $Prefix `
            -Suffix $Suffix `
            -Find $Find `
            -Replace $Replace `
            -NumberFrom $NumberFrom `
            -Recursive $Recursive.IsPresent

        if (-not $plan) {
            Write-Host "No files matched the given pattern."
            return 0
        }

        $changedCount = Invoke-RenamePlan -Plan $plan -Apply:$Apply.IsPresent

        if ($Apply) {
            Write-Host ""
            Write-Host "Renamed $changedCount file(s)."
        } else {
            Write-Host ""
            Write-Host "Previewed $changedCount file(s)."
            Write-Host "Run again with -Apply to make the changes."
        }

        return 0
    } catch {
        Write-Host $_.Exception.Message
        return 1
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    exit (Invoke-FileRenamerCli)
}
