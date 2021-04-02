Clear-Host

# ffmpeg needs to be in your path environment variable or you'll need to put the full path to it

$START_PATH = "D:\books"
$TEMP_PATH = "F:\temp"
$DELETE_OLD_MP3 = $true
$FFMPEG_ERROR_LEVEL = "-nostats -loglevel error"
$GRAB_MISSING__ALBUM_ART_FROM_AUDIBLE = $true

# If this is set to true and your files don't sort correctly your could get an out of 
# order m4b.
$IGNORE_FILE_LENGTHS = $false

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

Function Get-File-Counts($location) {
    # Get count of files in directories.
    # $location = "d:\books"
    set-location $location
    Remove-Item "$location\filecounts.txt"
    $dirs = Get-ChildItem -recurse | ?{ $_.PSIsContainer } 
    foreach ($dir in $dirs) {
        $count = (Get-ChildItem $dir.FullName | Measure-Object).Count
        if ($count -gt 5) {
            Write-Host "$count $dir"
            "$count $($dir.FullName)" | Out-File "$location\filecounts.txt" -Append
        }
    }
}

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
            Set-Clipboard -Value $file.DirectoryName
            return "Filenames in `"$($file.DirectoryName)`" are different lengths and might not sort and combine correctly."
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

Function Download-Pic-From-Audible([string]$searchterm) {
    $url = "https://www.audible.com/search?keywords=$searchterm&ref=a_hp_t1_header_search"
    $browserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.114 Safari/537.36'
    $page = Invoke-WebRequest -Uri $url -UserAgent $browserAgent
    foreach ($line in $page.Content.Split("`r`n")) {
        if ($line.Contains("m.media-amazon.com") -and $line.Contains("bc-image-inset-border")) {
            $start = $line.IndexOf("data-lazy=`"")
            $end = $line.Length - $start - 1 
            $image = $line.Substring($start, $end)
            $url = $image.Remove(0,11)
            break
        }
    }    
    $Filename = [System.IO.Path]::GetFileName($url)
    $ext = $Filename.Remove(0, ($Filename.Length - 3))
    Invoke-WebRequest -Uri $url -OutFile "$TEMP_PATH\cover.$ext"
    if (Test-Path "$TEMP_PATH\cover.$ext") {
        return "$TEMP_PATH\cover.$ext"
    } else {
        return $false
    }
    #break
}


# If there are no subfolders use the base path
$bookDirs = Get-ChildItem -Path $START_PATH -Recurse | ?{ $_.PSIsContainer }
if ($bookDirs.count -lt 2) {
    $bookDirs = $START_PATH
}

$files = ""
$firstRun = $true
$hasCoverArt = $false

write-host "Processing`r`n$bookDirs"

foreach ($bookDir in $bookDirs) {
    if ($mp3s) {
        $mp3s.Clear()
    }

    if ($bookDir.Fullname) {
        $bookDir = $bookDir.FullName
    }

    if ($bookDir.EndsWith("\") -eq $true) {
        $bookDir = "$bookDir"
    }

    if (!$TEMP_PATH) {
        $TEMP_PATH = $bookDir
    }

    $album = $null
    
    Set-Location $bookDir
    $mp3s = Get-ChildItem $bookDir -Filter *.mp3
#    $mp3s = Get-ChildItem $bookDir -Filter *.mp3 | Sort-Object -Property {$_.Name -as [int]}
    if ($mp3s) {
        $mp3sNameLength = Check-FileNameLengths $mp3s

        if ($mp3sNameLength -ne $true) {
            Write-Host $mp3sNameLength
            if ($IGNORE_FILE_LENGTHS -eq $false) {
                break;
            }
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
        $hasCoverArt = ($mp3s[0] | save-picture "$TEMP_PATH\cover.jpg")
        if($GRAB_MISSING__ALBUM_ART_FROM_AUDIBLE -and ($hasCoverArt -eq $false)) {
            Write-Host "Getting cover art from Audible."
            $hasCoverArt = Download-Pic-From-Audible $album
        }

        Write-Host "Combining files."
        $command = ("ffmpeg -i `"concat:$mp3sList`" -c:a copy `"$TEMP_PATH\$albumMp3`" $FFMPEG_ERROR_LEVEL").Replace("\\","\")
        $command
        Invoke-Expression $Command

        Write-Host "Converting to m4b."
        $command = ("ffmpeg -fflags +igndts -i `"$TEMP_PATH\$albumMp3`" -vn -c:a aac -q:a 1.2 -y `"$TEMP_PATH\$albumM4b`" $FFMPEG_ERROR_LEVEL").Replace("\\","\")
        $command
        Invoke-Expression $Command | Out-Null

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
    
        if ($hasCoverArt) {
            while (Test-Path $hasCoverArt) {
                Start-Sleep 5
                try {
                    Remove-Item $hasCoverArt -ErrorAction Stop
                } catch {
                    [System.GC]::Collect()
                    Write-Host "$hasCoverArt is locked."
                }
            }
        } else {
            Write-Host "Cover art is missing for $bookDir\$albumM4b"
        }
    } else {
        Write-Host "Nothing to process in $bookDir."
    }
}
