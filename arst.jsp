//Location of folder where config file is located
var homeFolder = "C:\\Users\\lukas\\Documents\\Github\\Plans_ARST_Macro\\files";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

//Initializing configuration object
var config = iniConfigConstructor( homeFolder, fso );

//Loading kdm model file and trying to save it as temporary binary file
ReadDataKDM( config.modelPath + "\\" + config.modelName + ".kdm" );

//If safe mode is turn on in config file then try to create temporary file with original state of model
if( config.safeMode == 1 ){

  var tmpFile = config.homeFolder + "\\tmp.bin";

  if( SaveTempBIN( tmpFile ) < 1 ) errorThrower( "Unable to create temporary file" );
}

//Setting power flow calculation settings with settings from config file
setPowerFlowSettings( config );

//Calculate power flow, if fails throw error 
CPF();

//Try to read file from location specified in configuration file, then make array from file and close the file
var inputFile = readFile( config, fso );
var inputArray = getInputArray( inputFile );
inputFile.close();

//Create file for logging
var resultFile = createFile( config, fso );

//Write all detected transformers from input file with taps and coresponding attribute to log file
resultFile.WriteLine( "Base data" );
logBaseInfo( resultFile, inputArray );

//Get array with transformers that are near their max voltage/reactive power limit
var delayArray = getDelayArray( inputArray );

//Sort delay array, so transformer with the least delay is first in array
sortArray( delayArray );

//While there are any transformers in delay array check if the transformer with the least delay is near it's maximum voltage/reactive power limit.
//If it's near limit then change transformer's tap depending on taps position. After that recalulate power flow and get new delay array.
while( delayArray.length > 0 ){

  var transformer = delayArray[ 0 ][ 0 ], node = delayArray[ 0 ][ 1 ];
  //Reactive power setpoint is required to do tap changes dependent on reactive power 
  //other calculations setpoint are directry taken from plans built in functions 
  var epsilon = parseFloat( delayArray[ 0 ][ 3 ] ), reactivePowerSetpoint = parseFloat( delayArray[ 0 ][ 4 ] );
  
  if( 
      ( 
        reactivePowerSetpoint && 
        ( 
          ( node.Qend > reactivePowerSetpoint + epsilon && transformer.TapLoc === 1 ) || 
          ( node.Qend < reactivePowerSetpoint - epsilon && transformer.TapLoc !== 1 ) 
        )    
      ) 
      || 
      ( 
        ( node.Vi > node.Vs + epsilon && transformer.TapLoc === 1 ) || 
        ( node.Vi < node.Vs - epsilon && transformer.TapLoc !== 1 )  
      ) 
  ){
    
    if( switchTap( transformer, -1, resultFile ) === 0 ){

      delayArray.shift();

      continue;
    }  
      
    switchTapOnParallelTransformers( transformer, -1, resultFile );  
  }
  
  else if( 
    ( reactivePowerSetpoint && 
      ( 
        ( node.Qend > reactivePowerSetpoint + epsilon && transformer.TapLoc !== 1 ) || 
        ( node.Qend < reactivePowerSetpoint - epsilon && transformer.TapLoc === 1 ) 
      )    
    ) 
    || 
    ( 
      ( node.Vi > node.Vs + epsilon && transformer.TapLoc !== 1 ) || 
      ( node.Vi < node.Vs - epsilon && transformer.TapLoc === 1 )  
    )
  ){

    if( switchTap( transformer, 1, resultFile ) === 0 ){

      delayArray.shift();

      continue;
    }  

    switchTapOnParallelTransformers( transformer, 1, resultFile );
  } 
  
  else{
  
    delayArray.shift();
    
    continue; 
  }

  CPF();
  
  delayArray = getDelayArray( inputArray );

  sortArray( delayArray );
}

//Write all detected transformers from input file with taps and coresponding attribute to log file after all tap changes are done
resultFile.WriteLine( "Final data" );
logBaseInfo( resultFile, inputArray );
resultFile.close();

//If safe mode is active, load back original model, meaning that all changes done by this macro is reversed to original state
if( config.safeMode == 1 ){

  //Loading original model
  ReadTempBIN( tmpFile );

  //Removing temporary binary file
  fso.DeleteFile( tmpFile );
}

//Functon takes file to log data into and array with found transformers from input file and
//transformers details are written to logging file
function logBaseInfo( file, inputArray ){

  for( i in inputArray ){

    var transformer = TrfArray.Find( inputArray[ i ][ 0 ] );
    var node = getTransformerNode( inputArray[ i ], transformer );

    file.Write( transformer.Name + ", Tap: " + transformer.Stp0 + "\\" + transformer.Lstp + ", Node: " + node.Name );
    
    if( inputArray[ i ][ 1 ] == "Q" ) file.WriteLine( ", React Pow: " + node.Qend + "\\" + inputArray[ i ][ 3 ] );
    
    else file.WriteLine(  ", Volt: " + node.Vi + "\\" + node.Vs );
  }

  file.WriteLine( " " );
}

function sortArray( array ){

  for( i in array ){
  
    var element = array[ i ], tmp = null;
    
    for( j in array ){
    
      if( element[ 2 ] < array[ j ][ 2 ] ){
      
        tmp = array[ j ];
        
        array[ j ] = element;
        
        array[ i ] = tmp;
        
        break;
      }
    
    }
  
  }
    
}

//Function returns specific tau value depending on it's node declared voltage or if reactive power is used in calculation   
function getTau( basePower, usingReactivePower ){

  if( usingReactivePower ) return 9000;

  switch( basePower ){
  
    case 400: return 3600;
    
    case 220: return 2400;
    
    case 110: return 1200;
  }  
  
}

//Function return transformer's node/branch depending on category in an input file  
function getTransformerNode( inputArrayElement, transformer ){

  var n = null;

  switch( inputArrayElement[ 1 ] ){
    
    case "G":
      
      n = nodArray.Find( transformer.BegName );
      
      break;
     
    case "D":
    
      n = nodArray.Find( transformer.EndName );
    
      break;
      
    case "Q":  
      
      n = braArray.Find( transformer.Name );
      
      break;
  }

  return n;
}

//Function takes input array and log file, returns array with delays of each transformator which falls into checks
function getDelayArray( inputArray ){

  var delayArray = [];
  
  for( i in inputArray ){

    var transformer = TrfArray.Find( inputArray[ i ][ 0 ] );

    var node = getTransformerNode( inputArray[ i ], transformer );
     
    var delay = u = q = 0, reactivePowerSetpoint = null, epsilon = parseFloat( inputArray[ i ][ 2 ] );
    
    //Depending on criteria specified in input file for each transfomers check if conditions are true then calculate delay and push it to delay array

    if( inputArray[ i ][ 1 ] === "G" || inputArray[ i ][ 1 ] === "D" ){

      //If calculated voltage out of range of voltage setpoint +/- epsilon then calculate delta of these values 
      if( node.Vi > ( node.Vs + epsilon ) ) u = node.Vi - ( node.Vs + epsilon );
      
      else if( node.Vi < ( node.Vs - epsilon ) ) u = ( node.Vs - epsilon ) - node.Vi;
    
      else continue;
      
      //Get Tau multiply it by 0.1 and divide it by delta u
      delay = ( getTau( node.Vn, false ) * 0.1 ) / u;
    }  
    
    else if( inputArray[ i ][ 1 ] === "Q" ){
    
      reactivePowerSetpoint = parseFloat( inputArray[ i ][ 3 ] );
            
      //If transformer's end node reactive power is out of range of reactive power setpoint +/- epsilon then calculate delta of these values   
      if( node.Qend > reactivePowerSetpoint + epsilon ) q = node.Qend - ( reactivePowerSetpoint + epsilon );
    
      else if( node.Qend < reactivePowerSetpoint - epsilon ) q = ( reactivePowerSetpoint - epsilon ) - node.Qend; 
      
      else continue;
      
      //Get Tau multiply it by 0.1 and divide it by delta u
      delay = ( getTau( 0, true ) * 0.1 ) / q;
    }

    delayArray.push( [ transformer, node, delay, epsilon, reactivePowerSetpoint ] ); 
  }

  return delayArray;
}

function switchTap( transfomer, value, resultFile ){

  if( transfomer.Stp0 + value < 1 || transfomer.Stp0 + value > transfomer.Lstp ){
    
    resultFile.WriteLine( transformer.Name + " reached max/min tap " + transformer.Stp0 + "\\" + transformer.Lstp );
  
    return 0;
  }

  else { 
    
    transfomer.Stp0 += value;

    return 1;
  }
}

//Function checks if transformer have any parallel transformers, if so then try to change theirs tap to match changes.
//If any found transformer is at it's tap limit then log that in file instead doing any changes.
function switchTapOnParallelTransformers( transformer, value, resultFile ){

  for( var i = 1; i < Data.N_Trf; i++ ){

    nextTransformer = TrfArray.get( i );

    if( transformer.Name === nextTransformer.Name ) continue; 
    
    if( 
      ( transformer.BegName !== nextTransformer.BegName || transformer.EndName !== nextTransformer.EndName ) && 
      ( transformer.BegName !== nextTransformer.EndName || transformer.EndName !== nextTransformer.BegName )  
    ) continue;
    
    if( transformer.BegName === nextTransformer.BegName ) switchTap( nextTransformer, value, resultFile ); 
        
    else switchTap( nextTransformer, -value, resultFile);
  }

}

//Function uses JS Math.round, takes value and returns rounded value to specified decimals 
function roundTo( value, precision ){

  return Math.round( value * ( 10 * precision ) ) / ( 10 * precision ) ;
}

//Set power flow settings using config file
function setPowerFlowSettings( config ){

  Calc.Itmax = config.maxIterations;
  Calc.EPS10 = config.startingPrecision;
  Calc.Eps = config.precision;
  Calc.Met = config.method;
}

//Basic error thrower with error message window
function errorThrower( message ){
  
  MsgBox( message, 16, "Error" );
  throw message;
}

//Calls built in power flow calculate function, throws error when it fails
function CPF(){

  if( CalcLF() != 1 ) errorThrower( "Power Flow calculation failed" );
}

//Function takes config object and depending on it's config creates folder in specified location. 
//Throws error if config object is null and when folder can't be created
function createFolder( config, fso ){
  
  if( !config ) errorThrower( "Unable to load configuration" );
  
  var folder = config.folderName;
  var folderPath = config.homeFolder + folder;
  
  if( !fso.FolderExists( folderPath ) ){
    
    try{ fso.CreateFolder( folderPath ); }
    
    catch( err ){ errorThrower( "Unable to create folder" ); }
  }
  
  folder += "\\";

  return folder;
}

//Function takes config object and depending on it's config creates file in specified location.
//Also can create folder where results are located depending on configuration file 
//Throws error if config object is null and when file can't be created
function createFile( config, fso ){
 
  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;
  
  var folder = ( config.createResultsFolder == 1 ) ? createFolder( config, fso ) : "";
  var timeStamp = ( config.addTimestampToFile == 1 ) ? getCurrentDate() + "--" : "";
  var fileLocation = config.homeFolder + folder + timeStamp + config.resultFileName + ".txt";
  
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( err ){ errorThrower( "File already exists or unable to create it" ); }

  return file;
} 

//Function takes config object and depending on it reads file from specified location.
//Throws error if config object is null or when file can't be read 
function readFile( config, fso ){

  if( !config ) errorThrower( "Unable to load configuration" );

  var file = null;

  var fileLocation = config.inputFileLocation + config.inputFileName + "." + config.inputFileFormat;
  
  try{ file = fso.OpenTextFile( fileLocation, 1, false, 0 ); }

  catch( err ){ errorThrower( "Unable to find or open file" ); }

  return file;
}

//Function uses built in .ini function to get it's settings from config file.
//Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ){
  
  var configFile = iniPath + "\\config.ini";

  if( !fso.FileExists( configFile ) ) errorThrower( "config.ini file not found" );

  //Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( configFile );

  var hFolder = ini.GetString( "main", "homeFolder", Main.WorkDir );
  
  //Declaring conf object and trying to fill it with config.ini configuration
  var conf = {
  
    //Main
    homeFolder: hFolder,
    modelName: ini.GetString( "main", "modelName", "model" ),
    modelPath: ini.GetString( "main", "modelPath", hFolder ),  
    safeMode: ini.GetBool( "main", "safeMode", 1 ),

    //Folder
    createResultsFolder: ini.GetBool( "folder", "createResultsFolder", 0 ),
    folderName: ini.GetString( "folder", "folderName", "folder" ),
    
    //Files
    addTimestampToFile: ini.GetBool( "files", "addTimestampToFile", 1 ),
    inputFileLocation: ini.GetString( "files", "inputFileLocation", hFolder ),
    inputFileName: ini.GetString( "files", "inputFileName", "input" ),
    inputFileFormat: ini.GetString( "files", "inputFileFormat", "txt" ),
    resultFileName: ini.GetString( "files", "rsultFileName", "log" ),
    roundingPrecision: ini.GetInt( "files", "roundingPrecision", 2 ),
    
    //Power Flow
    maxIterations: ini.GetInt( "power flow", "maxIterations", 300 ),
    startingPrecision: ini.GetDouble( "power flow", "startingPrecision", 10.00 ),
    precision: ini.GetDouble( "power flow", "precision", 1.00 ),
    method: ini.GetInt( "power flow", "method", 1 )
  };
  
  //Overwriting config.ini file
  //Main
  ini.WriteString( "main", "homeFolder", conf.homeFolder );
  ini.WriteString( "main", "modelName", conf.modelName );
  ini.WriteString( "main", "modelPath", conf.modelPath );
  ini.WriteBool( "main", "safeMode", conf.safeMode );
  
  //Folder
  ini.WriteBool( "folder", "createResultsFolder", conf.createResultsFolder );
  ini.WriteString( "folder", "folderName", conf.folderName );
    
  //Files
  ini.WriteBool( "files", "addTimestampToFile", conf.addTimestampToFile );
  ini.WriteString( "files", "inputFileLocation", conf.inputFileLocation );
  ini.WriteString( "files", "inputFileName", conf.inputFileName );
  ini.WriteString( "files", "inputFileFormat", conf.inputFileFormat );
  ini.WriteString( "files", "resultFileName", conf.resultFileName );
  ini.WriteInt( "file", "roundingPrecision", conf.roundingPrecision );
    
  //Power Flow
  ini.WriteInt( "power flow", "maxIterations", conf.maxIterations );
  ini.WriteDouble( "power flow", "startingPrecision", conf.startingPrecision );
  ini.WriteDouble( "power flow", "precision", conf.precision );
  ini.WriteInt( "power flow", "method", conf.method );
 
  return conf;
}

//Function gets file and takes each line into a array and after finding whiteline pushes array into other array
function getInputArray( file ){

  var array = [];

  while(!file.AtEndOfStream){

    var tmp = [], line = file.ReadLine();
      
    while( line != "" ){
     
      tmp.push( line.replace(/(^\s+|\s+$)/g, '') );
    
      if( !file.AtEndOfStream ) line = file.ReadLine();
      
      else break; 
    }
    
    array.push( tmp );
  }

  return array;
}

//Function takes current date and returns it in file safe format  
function getCurrentDate(){
  
  var current = new Date();
  
  var formatedDateArray = [ ( '0' + ( current.getMonth() + 1 ) ).slice( -2 ), ( '0' + current.getDate() ).slice( -2 ), 
  ( '0' + current.getHours() ).slice( -2 ), ( '0' + current.getMinutes() ).slice( -2 ), ( '0' + current.getSeconds() ).slice( -2 ) ];
  
  return current.getFullYear() + "-" + formatedDateArray[ 0 ] + "-" + formatedDateArray[ 1 ] + "--" + formatedDateArray[ 2 ] + "-" + formatedDateArray[ 3 ] + "-" + formatedDateArray[ 4 ];
}
