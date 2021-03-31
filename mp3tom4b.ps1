Clear-Host

<#
# Get count of files in directories.
set-location d:\books
Remove-Item filecounts.txt
$dirs = Get-ChildItem -recurse |  ?{ $_.PSIsContainer } 
foreach ($dir in $dirs) {
    $count = (Get-ChildItem $dir.FullName | Measure-Object).Count
    if ($count -gt 99) {
        Write-Host "$count $dir"
        "$count $($dir.FullName)" | Out-File filecounts.txt -Append
    }
}
#>

$START_PATH = "D:\books\Vince Flynn"
$TEMP_PATH = "F:\temp"
$DELETE_OLD_MP3 = $true

#region Taglib
if ($env:PSModulePath) {
    $modules = $env:PSModulePath
    if ($modules.Contains(";")) {
        $modules = $modules.Split(";")
        $MODULE_PATH = $modules[0]
    } else {
        $MODULE_PATH = $env:PSModulePath
    }
}

$TAGLIB_PATH = "$MODULE_PATH\taglib"

if ((Test-Path $TAGLIB_PATH) -eq $false) {
    try {
        Import-Module taglib -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Host "Need to install the taglib module."
    } finally {
        if ((Test-Path $MODULE_PATH) -eq $false) {
            New-Item -Path $MODULE_PATH -ItemType Directory
        }
        Set-Location $MODULE_PATH
        $command = 'git clone https://github.com/danielmbond/powershell-taglib.git taglib'
        Invoke-Expression -Command:$command -ErrorAction SilentlyContinue | Out-Null
        Import-Module taglib | Out-Null
    }
}

Import-Module taglib | Out-Null
#endregion

Function Remove-InvalidFileNameChars {
  param(
    [Parameter(Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true,
      ValueFromPipelineByPropertyName=$true)]
    [String]$Name
  )

  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [RegEx]::Escape($invalidChars)
  return ($Name -replace $re)
}

Function Check-FileNameLengths($files) {
    $length_previous = $files[0].FullName.Length
    foreach ($file in $files) {
        $length = $file.FullName.Length 
        if ($length -ne $length_previous) {
            return "Filenames are different lengths and might not sort and combine correctly."
        }
    }
    return $true
}

Function Get-MP3Duration($file) {
    # This doesn't work.  Was going to do it to create an ffmetadata file to add chapters.
    $command = "ffprobe -v error -select_streams a:0 -show_entries stream=duration -of default=noprint_wrappers=1:nokey=1 `"$file`""
    $duration = Invoke-Command $command
    return $duration
}

# If there are no subfolders use the base path
$bookDirs = Get-ChildItem -Path $START_PATH -Recurse | ?{ $_.PSIsContainer }
if ($bookDirs.count -lt 2) {
    $bookDirs = $START_PATH
}

$files = ""
$bookDirs
$firstRun = $true
$hasCoverArt = $false

foreach ($bookDir in $bookDirs) {
    if ($bookDir.Fullname) {
        $bookDir = $bookDir.FullName
    }

    if ($bookDir.EndsWith("\") -eq $false) {
        $bookDir = "$bookDir\"
    }

    if (!$TEMP_PATH) {
        $TEMP_PATH = $bookDir
    }

    $album = $null
    
    Set-Location $bookDir
    $mp3s = Get-ChildItem $bookDir -Filter *.mp3
    $mp3sNameLength = Check-FileNameLengths $mp3s

    if ($mp3sNameLength -ne $true) {
        Write-Host $mp3sNameLength
        break;
    }

    $album = Remove-InvalidFileNameChars ($mp3s[0] | get-album)
    $albumMp3 = "$($album)-temp.mp3"
    $albumM4b = "$album.m4b"
    $mp3sList = $null

    foreach ($i in $mp3s) {$mp3sList = "$mp3sList|$i"}
    $mp3sList = $mp3sList.Trim("|")

    if (Test-Path "$TEMP_PATH\$albumMp3") {
        Write-Host "File exists @ $TEMP_PATH\$albumMp3"
        break;
    }

    Write-Host "Get cover art."
    $hasCoverArt = $mp3s[0] | save-picture "$TEMP_PATH\cover.jpg"

    Write-Host "Combining files."
    $command = ("ffmpeg -i `"concat:$mp3sList`" -c:a copy `"$TEMP_PATH\$albumMp3`"").Replace("\\","\")
    $command
    Invoke-Expression $Command

    Write-Host "Converting to m4b."
    $command = ("ffmpeg -fflags +igndts -i `"$TEMP_PATH\$albumMp3`" -vn -c:a aac -q:a 1.2 -y `"$TEMP_PATH\$albumM4b`"").Replace("\\","\")
    $command
    Invoke-Expression $Command

    $finalBook = Get-ChildItem "$TEMP_PATH\$albumM4b" 
    if ($hasCoverArt) {
        write-host "Adding Cover Art"
        Get-ChildItem $finalBook | set-picture $hasCoverArt
    }

    $finalBook | set-title $album
    $finalBook | set-track 1 1
    $finalBook | set-disc 1 1

# Clean up
    if (Test-Path "$TEMP_PATH\$albumM4b") {
        Remove-Item "$TEMP_PATH\$albumMp3" 
        if ($DELETE_OLD_MP3) {
            foreach ($mp3 in $mp3s) {
                Remove-Item $mp3.FullName
            }
        }
        Move-Item -Path "$TEMP_PATH\$albumM4b" -Destination "$bookDir\$albumM4b"
    }
    
    if ($hasCoverArt -and (Test-Path $hasCoverArt)) {
        Remove-Item $hasCoverArt
    } else {
        write-host "Cover art is missing for $bookDir\$albumM4b"
    }
}
