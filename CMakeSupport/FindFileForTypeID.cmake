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
        string(CONCAT regPathZero ${regPathBase} "\\" ${code} "\\0" )
        #MESSAGE( STATUS "regPathZero=${regPathZero}" )
        cmake_host_system_information(RESULT oses QUERY WINDOWS_REGISTRY ${regPathZero} SUBKEYS SEPARATOR ";")
        foreach( os ${oses} )
            string(CONCAT regPath ${regPathZero} "\\" ${os} )
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
    endforeach()
ENDMACRO()

