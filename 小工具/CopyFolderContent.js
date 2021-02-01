function start() {

  var sourceFolder = "source folder id";
  var targetFolder = "target folder id";

  var source = DriveApp.getFolderById(sourceFolder);
  Logger.log(source);
  var target = DriveApp.getFolderById(targetFolder);

    copyFolder(source, target);
    Logger.log(123)
  
}

function copyFolder(source, target) {

  var folders = source.getFolders();
  var files   = source.getFiles();

  while(files.hasNext()) {
    var file = files.next();
    file.makeCopy(file.getName(), target);
  }

  while(folders.hasNext()) {
    var subFolder = folders.next();
    var folderName = subFolder.getName();
    var targetFolder = target.createFolder(folderName);
    copyFolder(subFolder, targetFolder);
  }

}