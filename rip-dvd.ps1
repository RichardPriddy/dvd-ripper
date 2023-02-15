# Import the Windows Media Player COM object
$wmp = New-Object -ComObject "WMPlayer.OCX.7"

# Specify the output directory for the transcoded videos
$outputDir = "D:\Rip"

# Loop indefinitely
while ($true) {
    # Wait for a DVD to be inserted
    Write-Host "Waiting for DVD..."
    while ($true) {
        # Check for a DVD
        $diskLabel = $null
        foreach ($drive in $wmp.cdromCollection) {
            if ($drive.IsMediaLoaded) {
                $diskLabel = $drive.VolumeName
                break
            }
        }
        
        # If a DVD is found, break out of the loop
        if ($diskLabel) {
            break
        }
        
        # Wait a bit before checking again
        Start-Sleep -Seconds 5
    }
    
    # Clean up the DVD title for use in the output path
    $dvdTitle = $diskLabel -replace '[\\/]','_' -replace ':',' -'
    
    # Build the command line arguments for VLC
    $vlcArgs = "-I dummy"  # run VLC in dummy interface mode
    $vlcArgs += " --no-nav"  # disable DVD menus
    $vlcArgs += " --dvd-device $($drive.cdromcollection.item(0).driveletter) dvd://1"  # specify the DVD drive and title to rip
    $vlcArgs += " --sout " # Transcode options
    $vlcArgs += " :file{dst=`"$outputDir\$dvdTitle.mp4`"}"  # specify the output file path for the transcoded video
    $vlcArgs += " --no-sout-subtitles"  # disable built-in subtitle transcoding
    $vlcArgs += " --add-sout-subtitle=eng:file{dst=`"$outputDir\$dvdTitle.srt`"}"  # extract and save the English subtitle track
    
    # Transcode the DVD to an MP4 using VLC
    Write-Host "Ripping DVD: $diskLabel"
    Start-Process "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" -ArgumentList $vlcArgs -Wait
    
    # Eject the DVD
    $drive.Eject()
    Write-Host "Ejected DVD: $diskLabel"
}
