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

MACRO(IncludeProjectSettings)
    set( options  )
    set( oneValueArgs QT )
    set( multiValueArgs )

    cmake_parse_arguments( _INCLUDE_PROJECT_SETTINGS "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )

    #MESSAGE( STATUS "" )
    SET( _INCLUDE_PROJECT_STTINGS_QT 1)

    #MESSAGE( STATUS "IncludeProjectSettings CMAKE_CURRENT_LIST_DIR=${CMAKE_CURRENT_LIST_DIR}" )
    #MESSAGE( STATUS "IncludeProjectSettings _INCLUDE_PROJECT_SETTINGS_QT=${_INCLUDE_PROJECT_SETTINGS_QT}" )

    SET( CURR_DIR ${CMAKE_CURRENT_LIST_DIR} )
    get_filename_component(STOP_DIR ${CMAKE_SOURCE_DIR} DIRECTORY)
    #MESSAGE( STATUS "CURR_DIR=${CURR_DIR}" )
    #MESSAGE( STATUS "STOP_DIR=${STOP_DIR}" )
    
    SET( PROJECT_FOUND 0 )
    while( NOT ${CURR_DIR} STREQUAL ${STOP_DIR} )
        #MESSAGE( STATUS "Checking ${CURR_DIR} for QtProject.cmake" )
        SET( currProjectFile "${CURR_DIR}/QtProject.cmake" )
        
        if ( EXISTS "${currProjectFile}" )
            #MESSAGE( STATUS "Found ${currProjectFile}" )
            include( ${currProjectFile} )
            SET( PROJECT_FOUND 1 )
            break()
        endif()
        
        get_filename_component(CURR_DIR ${CURR_DIR} DIRECTORY)
    endwhile()

    if ( NOT PROJECT_FOUND )
	    SET( currProjectFile "${CMAKE_SOURCE_DIR}/CMakeSupport/QtProject.cmake" )
        
        if ( EXISTS "${currProjectFile}" )
            #MESSAGE( STATUS "Found ${currProjectFile}" )
            include( ${currProjectFile} )
            SET( PROJECT_FOUND 1 )
        endif()
    endif()
    
    if ( NOT PROJECT_FOUND )
        if ( _INCLUDE_PROJECT_SETTINGS_QT )
            MESSAGE( FATAL_ERROR "No QtProject.cmake file found" )
        else()
            MESSAGE( FATAL_ERROR "No Project.cmake file found" )
        endif()
    endif()

ENDMACRO()
