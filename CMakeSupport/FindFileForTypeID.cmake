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

if( DUMPCPP_EXECUTABLE AND NOT EXISTS ${DUMPCPP_EXECUTABLE} )
        MESSAGE( WARNING "'${DUMPCPP_EXECUTABLE}' no longer exists" )
        UNSET( DUMPCPP_EXECUTABLE CACHE )
        UNSET( DUMPCPP_VERSION CACHE )
endif()

if( NOT DUMPCPP_EXECUTABLE )
    find_program(DUMPCPP_EXECUTABLE NAMES sab_dumpcpp
      PATHS ${CMAKE_INSTALL_PREFIX} ${CMAKE_BINARY_DIR}/dumpcpp/RelWithDebInfo ${CMAKE_BINARY_DIR}/dumpcpp/Release ${CMAKE_BINARY_DIR}/dumpcpp/Debug
        DOC "path to the dumpcpp executable (from build area)" 
        NO_DEFAULT_PATH
#        NO_CACHE
        )
    if( NOT DUMPCPP_EXECUTABLE )
        MESSAGE( WARNING "Could not find build area dumpcpp.  Re-run cmake after initial build" )
    else()
        file( REAL_PATH ${DUMPCPP_EXECUTABLE} DUMPCPP_EXECUTABLE EXPAND_TILDE)
        file( TO_CMAKE_PATH ${DUMPCPP_EXECUTABLE} DUMPCPP_EXECUTABLE)
        mark_as_advanced(DUMPCPP_EXECUTABLE)
    endif()

    if( DUMPCPP_EXECUTABLE )
        execute_process( COMMAND ${DUMPCPP_EXECUTABLE} --version OUTPUT_VARIABLE DUMPCPP_VERSION OUTPUT_STRIP_TRAILING_WHITESPACE)
        mark_as_advanced(DUMPCPP_VERSION)

        MESSAGE( STATUS "Using dumpcpp: '${DUMPCPP_EXECUTABLE}' - ${DUMPCPP_VERSION}" )
    endif()
endif()

MACRO(FileForTypeID typeID prefix )
    set( ${prefix}_OLBPATH FALSE)
    
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
                set( ${prefix}_OLBPATH ${path} )
                #message( STATUS "${prefix}_OLBPATH = ${${prefix}_OLBPATH}" )
                break()
            endif()
        endforeach()
    endforeach()
ENDMACRO()


MACRO( GenerateCPPFromFileID fileID prefix enumPrefix )
    if( NOT DUMPCPP_EXECUTABLE )
        MESSAGE( FATAL_ERROR "Could not find sab_dumpcpp" )
    endif()

    if ( NOT EXISTS ${DUMPCPP_EXECUTABLE} )
        message( FATAL_ERROR "${DUMPCPP_EXECUTABLE} does not exist" )
    endif()

    FileForTypeID( ${fileID} ${prefix} )

    #message( STATUS "${prefix}_OLBPATH=${${prefix}_OLBPATH}" )
    if ( NOT EXISTS ${${prefix}_OLBPATH} )
        message( FATAL_ERROR "Could not find OLB file '${${prefix}_OLBPATH}' for file id ${fileID}" )
    endif()

    set( ${prefix}_CPP ${CMAKE_CURRENT_BINARY_DIR}/${prefix}.cpp )
    set( ${prefix}_H ${CMAKE_CURRENT_BINARY_DIR}/${prefix}.h )

    #message( STATUS "${prefix}_CPP=${${prefix}_CPP}" )
    #message( STATUS "${prefix}_H=${${prefix}_H}" )
    #message( STATUS "DUMPCPP_EXECUTABLE=${DUMPCPP_EXECUTABLE} - ${DUMPCPP_VERSION}" )

    ADD_CUSTOM_COMMAND( 
        OUTPUT 
            ${${prefix}_CPP} ${${prefix}_H}
        COMMAND ${CMAKE_COMMAND} -E env "PATH=${QT_MSVCDIR}/bin;$ENV{PATH}" "${DUMPCPP_EXECUTABLE}" "${fileID}" -o ${prefix} -e ${enumPrefix}
        COMMENT "[DUMPCPP] Generating ${prefix}.cpp and ${prefix}.h from '${${prefix}_OLBPATH}' using '${DUMPCPP_EXECUTABLE}'"
        VERBATIM
        DEPENDS
           ${${prefix}_OLBPATH}
           ${DUMPCPP_EXECUTABLE}
    )
endmacro()
   

