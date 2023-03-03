# Import the Windows Media Player COM object
$wmp = New-Object -ComObject "WMPlayer.OCX.7"

# Specify the output directory for the transcoded videos
$outputDir = "Z:\TV"
$drive = $wmp.cdromCollection.Item(0)
# Loop indefinitely
while ($true) {
    $diskLabel = Read-Host "Series Title"
    New-Item -Path $outputDir -Name $diskLabel -ItemType "directory"
    
    $continue = "y"
    
    while($continue -eq "y") {
      # Clean up the DVD title for use in the output path
      $dvdTitle = $diskLabel -replace '\s','_' -replace ':',' -'
      
      $seriesNumber = Read-Host "Series Number"
      $seriesDir = "Series $seriesNumber"
    
      $makeMKV = "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe"

      New-Item -Path "$outputDir\$diskLabel" -Name $seriesDir -ItemType "directory"

      # Transcode the DVD to an MP4 using VLC
      Write-Host "Ripping DVD: $diskLabel"

      & "C:\Program Files (x86)\MakeMKV\makemkvcon64.exe" mkv disc:0 all --minlength=600 "$outputDir\$diskLabel\$seriesDir"

      $children = Get-ChildItem "$outputDir\$diskLabel\$seriesDir" | sort lastwritetime
      
      $episodeNumebr = 1;
      foreach ($child in $children)
      {
        $s = "S{0:d2}E{1:d2}" -f $seriesNumber, $episodeNumebr
        Move-Item "$outputDir\$diskLabel\$seriesDir\$child" "$outputDir\$diskLabel\$seriesDir\$s.mkv"
        $episodeNumebr = $episodeNumebr + 1
      }      

      # Eject the DVD
      $drive.Eject()
      Write-Host "Ejected DVD: $diskLabel"
      
      $continue = Read-Host "Continue series? [y/n]"
    }
}
