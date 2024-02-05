var output_folder = "D:/PROTOCOL/scan";
var dpi = 300;//default
var fso = new ActiveXObject( "Scripting.FileSystemObject" );
var fn =  fso.GetFileName(WScript.ScriptFullName);
var m = fn.match(/_(\d+)dpi/);
if(m.length > 0){
  m = parseInt(m[1],10);
  if(0 < m && m <= 4800) dpi=m;
}
var format = 'jpg';
function mkdirs(path){
   var parent = fso.GetParentFolderName(path);
   if(!fso.FolderExists(parent)){
      mkdirs(parent);
   }
   if(!fso.FolderExists(path)){
      fso.CreateFolder(path);
   }
}
var FormatID = {
  'bmp':"{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}",
  'png':"{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}",
  'gif':"{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}",
  'jpg':"{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}",
  'tiff':"{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
}
var wia = new ActiveXObject("Wia.DeviceManager");
for(var i=1;i <= wia.DeviceInfos.Count;++i){
  var device = wia.DeviceInfos(i).Connect();
  if(device.Type != 1) continue;
  var item = device.Items(1);
  item.Properties.Item("Horizontal Resolution").Value = dpi;
  item.Properties.Item("Vertical Resolution").Value   = dpi;
//item.Properties.Item("Horizontal Start Position").Value = 0; //6149
//item.Properties.Item("Vertical Start Position").Value = 0; //6150
//item.Properties.Item("Horizontal Extent").Value = 2480; //6151
//item.Properties.Item("Vertical Extent").Value = 3507; //6152
//item.Properties.Item("Bits Per Pixcel").Value   = 24; //4110
  var image = item.Transfer(FormatID[format]);
  if(image.FormatID != FormatID[format]){
    imageProcess =  new ActiveXObject("WIA.ImageProcess");
    imageProcess.Filters.Add(imageProcess.FilterInfos.Item("Convert").FilterID);
    imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = FormatID[format];
    imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 80;
    image = imageProcess.Apply(image);
  }
  var did = device.Properties.Item("Unique Device ID").Value.slice(-4);
  var now = new Date();
  var yyyy = ""+now.getFullYear();
  var mmdd = ("00"+(now.getMonth()+1)).slice(-2)+("00"+now.getDate()).slice(-2);
  var hhmmss = ("00"+now.getHours()).slice(-2)+("00"+now.getMinutes()).slice(-2)+("00"+now.getSeconds()).slice(-2);
  mkdirs(output_folder + "/"+yyyy+"/"+yyyy+mmdd);
  image.SaveFile(output_folder+"/"+yyyy+"/"+yyyy+mmdd+"/"+did+"_"+yyyy+mmdd+"T"+hhmmss+"."+format);
}
