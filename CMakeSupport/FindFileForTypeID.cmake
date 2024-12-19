#*******************************************************************************
#
#  SYNOPSYS CONFIDENTIAL - This is an unpublished, proprietary work of
#  Synopsys, Inc., and is fully protected under copyright and trade
#  secret laws. You may not view, use, disclose, copy, or distribute this
#  file or any information contained herein except pursuant to a valid
#  written license from Synopsys.
#
#*******************************************************************************
#*******************************************************************************

MACRO(FileForTypeID typeID prefix )
    set( ${prefix}_PATH FALSE)
    
    string(CONCAT regPathBase "HKEY_LOCAL_MACHINE\\Software\\Classes\\TypeLib\\"  ${typeID} )
    #message( STATUS "regPathBase=${regPathBase}" )
    cmake_host_system_information(RESULT codes QUERY WINDOWS_REGISTRY ${regPathBase} SUBKEYS SEPARATOR ";")

    #MESSAGE( STATUS "codes=${codes}" )
    foreach( code ${codes} )
        string(CONCAT regPath ${regPathBase} "\\" ${code} "\\0\\win32" )
        #MESSAGE( STATUS "regPath=${regPath}" )
        cmake_host_system_information(RESULT path QUERY WINDOWS_REGISTRY ${regPath} VALUE "" )
        #MESSAGE( STATUS "path=${path}" )
        if ( EXISTS ${path} )
            #MESSAGE( STATUS "prefix=${prefix}" )
            set( ${prefix}_PATH ${path} )
            #message( STATUS "${prefix}_PATH = ${${prefix}_PATH}" )
            break()
        endif()
    endforeach()
ENDMACRO()

    #    QSettings settings( QLatin1String( "HKEY_LOCAL_MACHINE\\Software\\Classes\\TypeLib\\" ) + typeLib, QSettings::NativeFormat );
   #     typeLib.clear();
  #      QStringList codes = settings.childGroups();
 #       for ( int c = 0; c < codes.count(); ++c )
 #       {
 #           typeLib = settings.value( QLatin1Char( '/' ) + codes.at( c ) + QLatin1String( "/0/win32/." ) ).toString();
 #           if ( QFile::exists( typeLib ) )
 #               break;
#        }
#
#        if ( !typeLib.isEmpty() )
#            fprintf( stdout, "\"%s\"\n", qPrintable( typeLib ) );
#        return 0;
