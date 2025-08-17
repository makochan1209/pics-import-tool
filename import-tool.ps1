param(
    [string]$Prefix = "7CM2",
    [int]$StartNumber = 1
)

# 実行ディレクトリをこのスクリプトファイルのある場所に設定
Set-Location -Path $PSScriptRoot

if (-not ([System.Management.Automation.PSTypeName]'NaturalSortComparer').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections;

public class NaturalSortComparer : IComparer {
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
    public static extern int StrCmpLogicalW(string x, string y);

    public int Compare(object a, object b) {
        return StrCmpLogicalW(((FileInfo)a).Name, ((FileInfo)b).Name);
    }
}
"@
}

$comparer = New-Object NaturalSortComparer

$files = [System.IO.FileInfo[]](Get-ChildItem -File -Path (Get-Location) | Where-Object {
    $_.Extension -notin ".ps1", ".bat"
})

if (-not $files -or $files.Count -eq 0) {
    Write-Host "⚠️ 対象ファイルが見つかりませんでした。" -ForegroundColor Red
    exit
}

[Array]::Sort($files, $comparer)

$previousDate = ""
$counter = $StartNumber
$firstDateEncountered = $false

foreach ($file in $files) {
    $dateYMD = & exiftool -d "%y%m%d" -DateTimeOriginal -s3 $file.FullName
    $dateYMFolder = & exiftool -d "%Y_%m" -DateTimeOriginal -s3 $file.FullName

    if ([string]::IsNullOrWhiteSpace($dateYMD) -or [string]::IsNullOrWhiteSpace($dateYMFolder)) {
        Write-Host "撮影日時なし → スキップ: $($file.Name)" -ForegroundColor Yellow
        continue
    }

    $datePart = $dateYMD.Trim()
    $folderPart = $dateYMFolder.Trim()

    $targetDir = Join-Path (Get-Location) $folderPart
    if (!(Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }

    if (-not $firstDateEncountered) {
        $previousDate = $datePart
        $counter = $StartNumber
        $firstDateEncountered = $true
    } elseif ($datePart -ne $previousDate) {
        $counter = 1
        $previousDate = $datePart
    }

    $attempts = 0
    while ($true) {
        $numberPart = if ($counter -lt 100000) {
            "{0:D5}" -f $counter
        } else {
            "$counter"
        }

        $newFileName = "$Prefix-$datePart-$numberPart$($file.Extension)"
        $newPath = Join-Path $targetDir $newFileName

        if (-not (Test-Path $newPath)) {
            Move-Item -Path $file.FullName -Destination $newPath
            if ($attempts -eq 0) {
                Write-Host "$($file.Name) → $newFileName  （→ $folderPart）"
            } else {
                Write-Host "$($file.Name) → $newFileName  （→ $folderPart、重複のため連番を $numberPart に変更）" -ForegroundColor Cyan
            }
            $counter++
            break
        } else {
            $attempts++
            $counter++
        }
    }
}

Write-Host "✅ 完了: すべてのファイルを正しくリネーム・移動しました。" -ForegroundColor Green
