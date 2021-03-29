Clear-Host
# ffmpeg needs to be in your path environment variable or you'll need to put the full path to it
<#
Get count of files in directories.
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

$start_path = "D:\books\Vince Flynn\Executive Power"
$temp_path = "F:\temp"

#region Taglib
$MODULE_PATH = "C:\Program Files\WindowsPowerShell\Modules\"
$TAGLIB_PATH = $MODULE_PATH + "taglib"

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

# If there are no subfolders use the base path
$bookDirs = Get-ChildItem -Path $start_path -Recurse | ?{ $_.PSIsContainer }
if ($bookDirs.count -lt 2) {
    $bookDirs = $start_path
}

$files = ""
$bookDirs
foreach ($bookDir in $bookDirs) {
    if ($bookDir.Fullname) {
        $bookDir = $bookDir.FullName
    }
    if ($bookDir.EndsWith("\") -eq $false) {
        $bookDir = "$bookDir\"
    }

    if (!$temp_path) {
        $temp_path = $bookDir
    }

    $album = $null
    
    Set-Location $bookDir
    $mp3s = Get-ChildItem $bookDir -Filter *.mp3
    $album = Remove-InvalidFileNameChars ($mp3s[0] | get-album)
    $albumMp3 = "$($album)-temp.mp3"
    $albumM4b = "$album.m4b"
    $mp3sList = $null
    foreach ($i in $mp3s) {$mp3sList = "$mp3sList|$i"}
    $mp3sList = $mp3sList.Trim("|")

    if (Test-Path "$temp_path\$albumMp3") {
        Write-Host "File exists @ $temp_path\$albumMp3"
        break;
    }
    $command = "ffmpeg -i `"concat:$mp3sList`" -c:a copy `"$temp_path\$albumMp3`""
    $command
    Invoke-Expression $Command
    
    $command = "ffmpeg -fflags +igndts -i `"$temp_path\$albumMp3`" -vn -c:a aac -q:a 1 -y `"$temp_path\$albumM4b`""
    $command = $Command.Replace("\\","\")
    $command
    Invoke-Expression $Command
    
    # Clean up
    if (Test-Path "$temp_path\$albumM4b") {
        Remove-Item "$temp_path\$albumMp3" 
        foreach ($mp3 in $mp3s) {
            Remove-Item $mp3.FullName
        }
        Move-Item -Path "$temp_path\$albumM4b" -Destination "$bookDir\$albumM4b"
    }
}
