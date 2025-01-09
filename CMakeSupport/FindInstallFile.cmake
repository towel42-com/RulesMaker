#include(CMakeParseArguments)

FUNCTION (InstallFile inFile outFile)
    set( options REMOVE_ORIG )
    set( oneValueArgs )
    set( multiValueArgs )

    cmake_parse_arguments( "" "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )

    #MESSAGE( STATUS "inFile=${inFile}" )
    #MESSAGE( STATUS "outFile=${outFile}" )

    if( IS_DIRECTORY ${outFile} )
        #MESSAGE( STATUS "${outFile} IS DIR" )
        get_filename_component( baseName ${inFile} NAME)
        #MESSAGE( STATUS "basename=${baseName}" )
        SET( outFile ${outFile}/${baseName})
    endif()

    #MESSAGE( STATUS "inFile=${inFile}" )
    #MESSAGE( STATUS "outFile=${outFile}" )

    get_filename_component( baseName ${outFile} NAME)

    EXECUTE_PROCESS( 
        COMMAND ${CMAKE_COMMAND} -E copy_if_different "${inFile}" "${outFile}"
        RESULT_VARIABLE copyResult
        OUTPUT_QUIET 
        ERROR_QUIET
    )
    IF ( NOT copyResult EQUAL 0 )
        message( FATAL_ERROR "Problem copying ${inFile} to ${outFile}" )
    ENDIF()

    if ( _REMOVE_ORIG )
        file(REMOVE ${inFile})
    ENDIF()
ENDFUNCTION()

FUNCTION(InstallFilePostBuild)
    set( options )
    set( oneValueArgs TARGET INFILE TARGET_DIR)
    set( multiValueArgs CONFIGURATIONS )

    cmake_parse_arguments( "" "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )

    if( NOT _TARGET )
        MESSAGE( FATAL_ERROR "TARGET argument not set" )
    ENDIF()

    if( NOT _INFILE )
        MESSAGE( FATAL_ERROR "INFILE argument not set" )
    ENDIF()

    if( NOT _TARGET_DIR )
        MESSAGE( FATAL_ERROR "TARGET_DIR argument not set" )
    ENDIF()
    
    SET( _CALL_FUNC_CMD
        ${CMAKE_COMMAND} -P ${CMAKE_SOURCE_DIR}/SABUtils/Modules/InstallFile.cmake
        --
        ${_INFILE}
        ${_TARGET_DIR}
    )
    
    #MESSAGE( STATUS "===============================" )
    #MESSAGE( STATUS " InstallFilePostBuild" )
    #MESSAGE( STATUS " TARGET=${_TARGET}" )
    #MESSAGE( STATUS " INFILE=${_INFILE}" )
    #MESSAGE( STATUS " TARGET_DIR=${_TARGET_DIR}" )
    #MESSAGE( STATUS " CALL_FUNC_CMD=${_CALL_FUNC_CMD}" )
    
    if ( NOT _CONFIGURATIONS )
        if ( _INFILE MATCHES ".*\.pdb" )
            SET( _CONFIGURATIONS
                Debug
                RelWithDebInfo
            )
        endif()
    endif()

    if ( _CONFIGURATIONS )
        foreach( currConfig ${_CONFIGURATIONS} )
            add_custom_command( TARGET ${_TARGET} POST_BUILD
                COMMAND "$<$<CONFIG:${currConfig}>:${_CALL_FUNC_CMD}>"
                COMMAND_EXPAND_LISTS
            )
        endforeach()
    else()
        add_custom_command( TARGET ${_TARGET} POST_BUILD
            COMMAND ${_CALL_FUNC_CMD}
            COMMAND_EXPAND_LISTS
        )
     endif()
        
ENDFUNCTION()

FUNCTION(InstallFilesPostBuild)
    set( options )
    set( oneValueArgs TARGET TARGET_DIR)
    set( multiValueArgs INFILES CONFIGURATIONS )

    cmake_parse_arguments( "" "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )

    if( NOT _TARGET )
        MESSAGE( FATAL_ERROR "TARGET argument not set" )
    ENDIF()

    if( NOT _INFILES )
        MESSAGE( FATAL_ERROR "INFILES argument not set" )
    ENDIF()

    if( NOT _TARGET_DIR )
        MESSAGE( FATAL_ERROR "TARGET_DIR argument not set" )
    ENDIF()

    #MESSAGE( STATUS "===============================" )
    #MESSAGE( STATUS " InstallFilesPostBuild" )
    #MESSAGE( STATUS " TARGET=${_TARGET}" )
    #MESSAGE( STATUS " INFILES=${_INFILES}" )
    #MESSAGE( STATUS " TARGET_DIR=${_TARGET_DIR}" )


    foreach( curr ${_INFILES} )
        IF( NOT _CONFIGURATIONS )
            if ( curr MATCHES ".*\.pdb" )
                SET( _CONFIGS
                    Debug
                    RelWithDebInfo
                )
            endif()        
        ELSE()
            SET( _CONFIGS ${_CONFIGURATIONS} )
        ENDIF()
        InstallFilePostBuild( TARGET ${_TARGET} INFILE ${curr} TARGET_DIR ${_TARGET_DIR} CONFIGURATIONS ${_CONFIGS} )
    endforeach()
    endfunction()
