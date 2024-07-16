# Set the path to the FFmpeg executable
$ffmpegPath = "C:\ProgramData\chocolatey\lib\ffmpeg-full\tools\ffmpeg\bin\ffmpeg.exe"

# Set the directory to search in
$dir = "Z:\Movies"

# Define the list of video file extensions
$videoExtensions = @("*.mp4", "*.avi", "*.mkv")

# Recursively search for all video files
$videos = Get-ChildItem -Path $dir -Recurse -Include $videoExtensions | Where-Object { $_.Mode -notmatch "d" }

foreach($video in $videos)
{
    $videoPath = $video.FullName

    # Run FFmpeg and capture the output
    $output = & $ffmpegPath -i $videoPath 2>&1

    # Search the output for the resolution
    $resolutionPattern = "Stream.*Video.* ([0-9]+)x([0-9]+)"
    $resolutionMatch = $output | Select-String -Pattern $resolutionPattern

    # Extract the resolution from the match
    $resolution = "$($resolutionMatch.Matches.Groups[1].Value)x$($resolutionMatch.Matches.Groups[2].Value)"

    $res = "";
    # Determine if the video is 1080p, 720p, or SD
    if ($resolution -match "1920.*" -or $resolution -match ".*x1080") {
        Write-Host "$videoPath is 1080p"
        $res = " [1080p]"
    }
    elseif ($resolution -match "1280.*" -or $resolution -match ".*x720") {
        Write-Host "$videoPath is 720p"
        $res = " [720p]"
    }
    else {
        Write-Host "$videoPath is SD"
        $videoPath | Out-File "Z:\SD-check.txt" -Append 
        $res = " [DVDRip]"
    }

    # Get the file extension
    $extension = [System.IO.Path]::GetExtension($videoPath)

    # Remove the extension from the file name
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($videoPath)

    # Append the string to the file name
    $newFileName = $fileName + $res + $extension

    # Rename the file
    Rename-Item $videoPath -NewName $newFileName
    #break;
}