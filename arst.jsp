ReadDataKDM( "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\MODEL_ZW.kdm" );
SaveTempBin( "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\arst\\tmp.bin" );

CalcLF();

var inputArr = [ [ "MIL-A3", "Q", 10, 70 ], [ "MIL-A1B", "D", 1 ], [ "MIL-A5", "D", 0.5 ], [ "MIL-T4B", "D", 0.5 ] ];

printInfo( inputArr );

var delArr = getDelayArray( inputArr );

sortArray( delArr );

cprintf( "0 pass:" );

printPass( delArr );

var check = delArr.length;

var counter = 1;

while( check > 0 ){
  
  var t = delArr[ 0 ][ 0 ], n = delArr[ 0 ][ 1 ], eps = delArr[ 0 ][ 3 ];
  
  CalcLF();
  
  cprintf( "Vi: " + n.Vi + ", Vs: " + n.Vs + ", Eps: " + eps );
  
  if( n.Vi > n.Vs + eps ){
  
    if( t.TapLoc === 1 ) t.Stp0--; 
      
    else t.Stp0++;
  }

  else if( n.Vi < n.Vs - eps ){
  
    if( t.TapLoc === 1 ) t.Stp0++;
    
    else t.Stp0--;
  }
  
  else{ 
    
    check--;
    
    delArr.shift();
    
    continue; 
  }
    
  delArr = getDelayArray( inputArr );

  sortArray( delArr );
    
  cprintf( counter + " pass:"  );
  
  counter++;
  
  printPass( delArr );
}

printInfo( inputArr );

ReadTempBin( "C:\\Users\\lukas\\Documents\\Codes\\Plans JS macros\\arst\\tmp.bin" );

function printPass( delArr ){

  for( i in delArr ){ cprintf( delArr[ i ][ 0 ].Name + ": " + delArr[ i ][ 2 ] + ", Volt: " + delArr[ i ][ 1 ].Vi + "/" + delArr[ i ][ 1 ].Vs ); }
}

function printInfo( inputArr ){

  cprintf("");

  for( i in inputArr ){

    var t = TrfArray.Find( inputArr[ i ][ 0 ] );

    cprintf( "N: " + t.Name + ", Current Tap: " + t.Stp0 + ", Max Tap:" + t.Lstp );
  }

  cprintf("");
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

function getDelayArray( inputArr ){

  var delArr = [];
  
  for( i in inputArr ){

    var t = TrfArray.Find( inputArr[ i ][ 0 ] );

    var n = null;
    
    switch( inputArr[ i ][ 1 ] ){
    
      case "G":
        
        n = nodArray.Find( t.BegName );
        
        break;
       
      case "D":
      case "Q":
      
        n = nodArray.Find( t.EndName );
        
        break;
    }
     
    var tm = u = q = 0;
      
    var eps = inputArr[ i ][ 2 ];
    
    if( ( t.TapLoc === 1 && t.Stp0 < 2 ) || ( t.TapLoc === 0 && t.Stp0 >= t.Lstp ) ){ continue; }
   
    if( inputArr[ i ][ 1 ] === "G" || inputArr[ i ][ 1 ] === "D" ){
         
      if( n.Vi > ( n.Vs + eps ) ) u = n.Vi - ( n.Vs + eps );
      
      else if( n.Vi < ( n.Vs - eps ) ) u = ( n.Vs - eps ) - n.Vi;
    
      else continue;
      
      tm = ( getTau( n.Vn, false ) * 0.1 ) / u;
    }  
    
    else if( inputArr[ i ][ 1 ] === "Q" ){
    
      var qVs = inputArr[ i ][ 3 ] ;
    
      var bra = BraArray.Find( inputArr[ i ][ 0 ] );
         
      if( bra.Qend > qVs + eps ) q = bra.Qend - ( qVs + eps );
    
      else if( bra.Qend < qVs - eps ) q = ( qVs - eps ) - bra.Qend;
            
      else continue;
      
      tm = ( getTau( 0, true ) * 0.1 ) / q;
    }

    delArr.push( [ t, n, tm, eps ] ); 
  }

  return delArr;

}