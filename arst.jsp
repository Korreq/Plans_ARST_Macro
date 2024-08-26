//Location of folder where config file is located
var homeFolder = "C:\\Users\\lukas\\Documents\\Github\\Plans_ARST_Macro\\files";

//Creating file operation object
var fso = new ActiveXObject( "Scripting.FileSystemObject" );

//Initializing configuration object
var conf = iniConfigConstructor( homeFolder, fso );

//Loading kdm model file and trying to save it as temporary binary file
ReadDataKDM( conf.modelPath + "\\" + conf.modelName + ".kdm" );

if( conf.safeMode == 1 ){

  var tmpFile = conf.homeFolder + "\\tmp.bin";

  if( SaveTempBIN( tmpFile ) < 1 ) 
    
    errorThrower( "Unable to create temporary file", "Unable to create temporary file, check if you are able to create files in homeFolder location" );
}

//Setting power flow calculation settings with settings from config file
setPowerFlowSettings( conf );

//Calculate power flow, if fails throw error 
CPF();

var inputFile = readFile( conf, fso );
var inputArr = getInputArray( inputFile );
inputFile.close();

var resultFile = createFile( conf, fso );

resultFile.WriteLine( "Base data" );
logBaseInfo( resultFile, inputArr );

var delArr = getDelayArray( inputArr, resultFile );

sortArray( delArr );

//while( check > 0 && delArr.length > 0 ){
while( delArr.length > 0 ){

  var t = delArr[ 0 ][ 0 ], n = delArr[ 0 ][ 1 ], eps = delArr[ 0 ][ 3 ];
  
  var qVs = ( delArr[ 0 ][ 4 ] ) ? delArr[ 0 ][ 4 ] : null ;
  
  if( 
      ( 
        qVs && 
        ( 
          ( n.Qend > qVs + eps && t.TapLoc === 1 ) || 
          ( n.Qend < qVs - eps && t.TapLoc !== 1 ) 
        )    
      ) 
      || 
      ( 
        ( n.Vi > n.Vs + eps && t.TapLoc === 1 ) || 
        ( n.Vi < n.Vs - eps && t.TapLoc !== 1 )  
      )
    ){
    
    t.Stp0--;
    
    switchTapOnParallelTransformers( t, -1, resultFile );  
  }
  
  else if( 
      ( 
        qVs && 
        ( 
          ( n.Qend > qVs + eps && t.TapLoc !== 1 ) || 
          ( n.Qend < qVs - eps && t.TapLoc === 1 ) 
        )    
      ) 
      || 
      ( 
        ( n.Vi > n.Vs + eps && t.TapLoc !== 1 ) || 
        ( n.Vi < n.Vs - eps && t.TapLoc === 1 )  
      )
    ){
  
    t.Stp0++;
    
    switchTapOnParallelTransformers( t, 1, resultFile );
  }
  
  else{
  
    delArr.shift();
    
    continue; 
  }

  CPF();
  
  delArr = getDelayArray( inputArr, resultFile );

  sortArray( delArr );
}

resultFile.WriteLine( "Final data" );
logBaseInfo( resultFile, inputArr );

resultFile.close();

if( conf.safeMode == 1 ){

  //Loading original model
  ReadTempBIN( tmpFile );

  //Removing temporary binary file
  fso.DeleteFile( tmpFile );
}

function logBaseInfo( file, inputArr ){

  for( i in inputArr ){

    var t = TrfArray.Find( inputArr[ i ][ 0 ] );
    var n = getTransformerNode( inputArr[ i ], t );

    file.Write( t.Name + ", Tap: " + t.Stp0 + "\\" + t.Lstp + ", Node: " + n.Name );
    
    if( inputArr[ i ][ 1 ] == "Q" ){
    
      var bra = BraArray.Find( inputArr[ i ][ 0 ] );
    
      file.WriteLine( ", React Pow: " + bra.Qend + "\\" + inputArr[ i ][ 3 ] );
    }
    
    else{
    
      file.WriteLine(  ", Volt: " + n.Vi + "\\" + n.Vs );
    }
    
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
        
        array [ i ] = tmp;
        
        break;
      }
    
    }
  
  }
    
}

function getTau( basePower, usingReactivePower ){

  if( usingReactivePower ) return 9000;

  switch( basePower ){
  
    case 400: return 3600;
    
    case 220: return 2400;
    
    case 110: return 1200;
  }  
  
}

function getTransformerNode( inputArrElement, t ){

  var n = null;

  switch( inputArrElement[ 1 ] ){
    
    case "G":
      
      n = nodArray.Find( t.BegName );
      
      break;
     
    case "D":
    
      n = nodArray.Find( t.EndName );
    
      break;
      
    case "Q":  
      
      n = braArray.Find( t.Name );
      
      break;
  }

  return n;
}

function getDelayArray( inputArr, resultFile ){

  var delArr = [];
  
  for( i in inputArr ){

    var t = TrfArray.Find( inputArr[ i ][ 0 ] );

    var n = getTransformerNode( inputArr[ i ], t );
     
    var tm = u = q = 0;
      
    var eps = inputArr[ i ][ 2 ];
    
    if( ( t.TapLoc === 1 && t.Stp0 < 2 ) || ( t.TapLoc === 0 && t.Stp0 >= t.Lstp ) ){ 
    
      resultFile.WriteLine( t.Name + " reached max tap " + t.Stp0 + "\\" + t.Lstp );

      continue; 
    }
   
    if( inputArr[ i ][ 1 ] === "G" || inputArr[ i ][ 1 ] === "D" ){
         
      if( n.Vi > ( n.Vs + eps ) ) u = n.Vi - ( n.Vs + eps );
      
      else if( n.Vi < ( n.Vs - eps ) ) u = ( n.Vs - eps ) - n.Vi;
    
      else continue;
      
      tm = ( getTau( n.Vn, false ) * 0.1 ) / u;
      
      delArr.push( [ t, n, tm, eps ] ); 
    }  
    
    else if( inputArr[ i ][ 1 ] === "Q" ){
    
      var qVs = inputArr[ i ][ 3 ] ;
    
      var bra = n;
              
      if( bra.Qend > qVs + eps ) q = bra.Qend - ( qVs + eps );
    
      else if( bra.Qend < qVs - eps ) q = ( qVs - eps ) - bra.Qend; 
       
      else continue;
      
      tm = ( getTau( 0, true ) * 0.1 ) / q;
      
      delArr.push( [ t, n, tm, eps, qVs ] ); 
    }

    //delArr.push( [ t, n, tm, eps ] ); 
  }

  return delArr;
}

function switchTapOnParallelTransformers( t, value, resultFile ){

  var string;

  for( var i = 1; i < Data.N_Trf; i++ ){

    nt = TrfArray.get( i );

    if( t.Name === nt.Name ) continue; 
    
    if( ( t.BegName !== nt.BegName || t.EndName !== nt.EndName ) && ( t.BegName !== nt.EndName || t.EndName !== nt.BegName )  ) continue;
    
    string = nt.Name + ": " + nt.Stp0 + " >>> ";

    if( ( value < 0 && nt.Stp0 > 1 ) || ( value > 0 && nt.Stp0 < nt.Lstp ) ){
    
      if( t.BegName === nt.BegName ) nt.Stp0 += value;
  
      else nt.Stp0 -= value; 
    }
    
    else resultFile.WriteLine( nt.Name + " reached max tap " + nt.Stp0 + "\\" + nt.Lstp );
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

//Basic error thrower
function errorThrower( message, error ){
  
  MsgBox( message, 16, "Error" );
  throw error;
}

//Function adds loading bin file before throwing an error
function saveErrorThrower( message, error, binPath ){

  try{ ReadTempBIN( binPath ); }
  
  catch( e ){ MsgBox( "Couldn't load original model", 16, "Error" ) }

  errorThrower( message, error );
}

//Calls built in power flow calculate function, throws error when it fails
function CPF(){

  if( CalcLF() != 1 ) errorThrower( "Power Flow calculation failed", -1 );
}

//Function takes conf object and depending on it's config creates folder in specified location. 
//Throws error if conf object is null and when folder can't be created
function createFolder( conf, fso ){

  var message = "Unable to load configuration";
  
  if( !conf ) errorThrower( message, message );
  
  var folder = conf.folderName;
  var folderPath = conf.homeFolder + folder;
  
  if( !fso.FolderExists( folderPath ) ){
    
    try{ fso.CreateFolder( folderPath ); }
    
    catch( err ){ 
    
      errorThrower( "Unable to create folder", "Unable to create folder, check if you are able to create folders in that location" );
    }

  }
  
  folder += "\\";

  return folder;
}

//Function takes conf object and depending on it's config creates file in specified location.
//Also can create folder where results are located depending on configuration file 
//Throws error if conf object is null and when file can't be created
function createFile( conf, fso ){
  
  var message = "Unable to load configuration";
  if( !conf ) errorThrower( message, message );

  var file = null;
  
  var folder = ( conf.createResultsFolder == 1 ) ? createFolder( conf, fso ) : "";
  var timeStamp = ( conf.addTimestampToFile == 1 ) ? getCurrentDate() + "--" : "";
  var fileLocation = conf.homeFolder + folder + timeStamp + conf.resultFileName + ".txt";
  
  try{ file = fso.CreateTextFile( fileLocation ); }
  
  catch( err ){ 
    
    errorThrower( "File arleady exists or unable to create it", "File arleady exists or unable to create it, check if you are able to create files in that location" );
  }

  return file;
} 

function readFile( conf, fso ){

  var message = "Unable to load configuration";
  if( !conf ) errorThrower( message, message );

  var file = null;

  var fileLocation = conf.inputFileLocation + conf.inputFileName + "." + conf.inputFileFormat;
  
  try{ file = fso.OpenTextFile( fileLocation, 1, false, 0 ); }

  catch( err ){ errorThrower( "Unable to open file", -1 ); }

  return file;
}

//Function uses built in .ini function to get it's settings from config file.
//Returns conf object with settings taken from file. If file isn't found error is throwed instead.
function iniConfigConstructor( iniPath, fso ){
  
  var confFile = iniPath + "\\config.ini";

  if( !fso.FileExists( confFile ) ) errorThrower( "config.ini file not found", "Config file error, make sure your file location has config.ini file" );

  //Initializing plans built in ini manager
  var ini = CreateIniObject();
  ini.Open( confFile );

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
 
function getInputArray( file ){

  var arr = [];

  while(!file.AtEndOfStream){

    var tmp = [];
  
    var line = file.ReadLine();
      
    while( line != "" ){
     
      tmp.push( line.replace(/(^\s+|\s+$)/g, '') );
    
      if( !file.AtEndOfStream ) line = file.ReadLine();
      
      else break; 
    }
    
    arr.push( tmp );
  }

  return arr;
}

//Function takes current date and returns it in file safe format  
function getCurrentDate(){
  
  var current = new Date();
  
  var formatedDateArray = [ ( '0' + ( current.getMonth() + 1 ) ).slice( -2 ), ( '0' + current.getDate() ).slice( -2 ), 
  ( '0' + current.getHours() ).slice( -2 ), ( '0' + current.getMinutes() ).slice( -2 ), ( '0' + current.getSeconds() ).slice( -2 ) ];
  
  return current.getFullYear() + "-" + formatedDateArray[ 0 ] + "-" + formatedDateArray[ 1 ] + "--" + formatedDateArray[ 2 ] + "-" + formatedDateArray[ 3 ] + "-" + formatedDateArray[ 4 ];
}
