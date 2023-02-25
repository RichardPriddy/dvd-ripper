# Import the Windows Media Player COM object
$wmp = New-Object -ComObject "WMPlayer.OCX.7"

# Specify the output directory for the transcoded videos
$outputDir = "Z:\Movies"
$drive = $wmp.cdromCollection.Item(0)
# Loop indefinitely
while ($true) {
    $diskLabel = Read-Host "DVD Title"
    
    # Clean up the DVD title for use in the output path
    $dvdTitle = $diskLabel -replace '\s','_' -replace ':',' -'

    $makeMKV = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

    New-Item -Path $outputDir -Name $diskLabel -ItemType "directory"

    $makeMKVArgs = "mkv disc:0 all --minlength=3000 $outputDir\$diskLabel"

    # Transcode the DVD to an MP4 using VLC
    Write-Host "Ripping DVD: $diskLabel"

    & "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" mkv disc:0 all --minlength=3000 "$outputDir\$diskLabel"

    $child = Get-ChildItem "$outputDir\$diskLabel"
    Move-Item "$outputDir\$diskLabel\$($child[0])" "$outputDir\$diskLabel\$dvdTitle.mkv"
    
    # Eject the DVD
    $drive.Eject()
    Write-Host "Ejected DVD: $diskLabel"
}
