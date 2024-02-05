$dpi = 300
$deviceManager = new-object -ComObject WIA.DeviceManager
$wiaFormatBMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatGIF = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
$wiaFormatTIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
foreach($di in $deviceManager.DeviceInfos){
    $device = $di.Connect()
    foreach ($item in $device.Items) {
        $item.Properties.Item("Horizontal Resolution").Value = [string]$dpi #6147
        $item.Properties.Item("Vertical Resolution").Value   = [string]$dpi #6148
#        $item.Properties.Item("Horizontal Start Position").Value = 0 #6149
#        $item.Properties.Item("Vertical Start Position").Value = 0 #6150
#        $item.Properties.Item("Horizontal Extent").Value = 2480 #6151
#        $item.Properties.Item("Vertical Extent").Value = 3507 #6152
#        $item.Properties.Item("Bits Per Pixcel").Value   = 24 #4110
        
        $image = $item.Transfer($wiaFormatPNG) 
    }
    $did = $device.Properties.Item("Unique Device ID").Value
    $did = $did.Substring($did.Length-4,4)
    if($image.FormatID -ne $wiaFormatPNG)
    {
        $imageProcess = new-object -ComObject WIA.ImageProcess
        $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
        $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatPNG
        $image = $imageProcess.Apply($image)
    }
    $yyyy = get-date -Format "yyyy"
    $mmdd = get-date -Format "MMdd"
    $hhmmss = get-date -Format "HHmmss"
    $outd =$yyyy+"/"+$yyyy+$mmdd
    if(-not (Test-Path -Path $outd -PathType Container)){
        mkdir("$(pwd)/$outd")
    }
    $fname = $yyyy+$mmdd+"T"+$hhmmss+"_"+$did+".png"
    $image.SaveFile("$outd/$fname")
}