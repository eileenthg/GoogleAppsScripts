//Folder IDs
const lessonPlanID = 'REDACTED';


// Change which folder to iterate here.
const target = lessonPlanID;



//** CODE */
const filePath = []
const emptyFolders = []

//Check file existence
function checkFiles() {

  let rootfolder=DriveApp.getFolderById(target); 
  
  
  searchFolder(rootfolder);
  
  Logger.log('');
  Logger.log('');

  // List empty folders.
  Logger.log("Empty folders:");
  emptyFolders.forEach(function(value) {
   Logger.log(value);
  });

}

// Seach folders and subfolders for files using Depth-First-Search algorithm.
function searchFolder(rootfolder){
  filePath.push(rootfolder.getName());

  //Seach subfolders
  let folders = rootfolder.getFolders();
  if(folders.hasNext()){
    var flagNoFolder = false;
    while(folders.hasNext()){
      let folder = folders.next();
      searchFolder(folder);
    }
  } else {
    var flagNoFolder = true;
  }

  // State the filePath
  Logger.log("********* Accessing "+ filePath.join() + "********************************");

  //List all files in folders.
  let files = rootfolder.getFiles();

  if(files.hasNext()){
    var flagNoFiles = false;
    while(files.hasNext()){
      let file = files.next();
      Logger.log(file.getName());
    }
  } else {
    var flagNoFiles = true;
    
  }

  if(flagNoFolder && flagNoFiles){
    //Logger.log("Empty folder.");
    emptyFolders.push(filePath.join());
  }
  
  filePath.pop()  
  Logger.log("**************End of list ***************************************************************")

}
