SET(GIT_FOUND FALSE)
SET(GIT_EXE_FOUND FALSE)

FIND_PROGRAM(GIT_EXE_EXECUTABLE git
	DOC "GIT command line client")
MARK_AS_ADVANCED(GIT_EXE_EXECUTABLE)

find_package(InstallFile REQUIRED)

IF(GIT_EXE_EXECUTABLE)
    SET(GIT_EXE_FOUND TRUE)
    SET(GIT_FOUND TRUE)

    IF( NOT EXISTS "${GIT_EXE_EXECUTABLE}" )
        MESSAGE( "GIT Does not exist: '${GIT_EXE_EXECTUABLE}'" )
        UNSET( GIT_EXE_EXECUTABLE CACHE )
        UNSET( GIT_EXE_FOUND CACHE )
        UNSET( GIT_FOUND CACHE )
        UNSET( GIT_EXE_EXECUTABLE )
        UNSET( GIT_EXE_FOUND )
        UNSET( GIT_FOUND )
        MESSAGE( FATAL_ERROR "GIT Does not exist" )
    ENDIF()

    MACRO(GetGitInfo dir prefix)
        #sets the following variables
        # ${prefix}_REV -> The current git revision to 8 characters (32 bits)
        # ${prefix}_REV_DIFF -> If the git repo has been modified
        # ${prefix}_REV_TAG -> The current git tag
        # ${prefix}_REV_BRANCH -> The current git branch

        SET(_GIT_SAVED_LC_ALL "$ENV{LC_ALL}")
        SET(ENV{LC_ALL} C)

        MESSAGE( STATUS "Using GIT: '${GIT_EXE_EXECUTABLE}'" )
        MESSAGE( STATUS "Getting GIT info on '${dir}'" )
        EXECUTE_PROCESS(
            COMMAND 
                ${GIT_EXE_EXECUTABLE} -C "${dir}" 
                    describe --abbrev=8 "--dirty=;TRUE" --always
                    OUTPUT_VARIABLE _FULL_GIT_VERSION 
                    ERROR_VARIABLE ${prefix}_ERROR
                    OUTPUT_STRIP_TRAILING_WHITESPACE
                    ERROR_STRIP_TRAILING_WHITESPACE
        )
        
        if ( "${_FULL_GIT_VERSION}" STREQUAL "" )
            MESSAGE( FATAL_ERROR "Could not get GIT info on directory '${dir}'\r     '${${prefix}_ERROR}'" )
            SET(${prefix}_REV "N/A")
            SET(${prefix}_DIFF "")
            SET(${prefix}_TAG "N/A")
            SET(${prefix}_BRANCH "N/A")
        else()
            LIST( LENGTH _FULL_GIT_VERSION _LEN )
            if ( _LEN GREATER 2)
                MESSAGE( FATAL_ERROR "Invalid GIT format for version returned '${_FULL_GIT_VERSION}'" )
            ENDIF()

            LIST( GET _FULL_GIT_VERSION 0 ${prefix}_REV )
            if ( _LEN EQUAL 2 )
                LIST( GET _FULL_GIT_VERSION 1 ${prefix}_DIFF )
            ELSE()
                SET( ${prefix}_DIFF FALSE )
            ENDIF()

            string(REPLACE "-g" ";"  ${prefix}_REV ${${prefix}_REV} )
            LIST( LENGTH ${prefix}_REV _LEN )
            if ( _LEN GREATER 2)
                MESSAGE( FATAL_ERROR "Invalid GIT format for version returned '${_FULL_GIT_VERSION}'" )
            ENDIF()
            
            if ( _LEN EQUAL 2)
                LIST( GET ${prefix}_REV 1 ${prefix}_REV )
            ENDIF()
                      
            execute_process(
                COMMAND ${GIT_EXE_EXECUTABLE} -C "${dir}" 
                    describe --exact-match --tags
                    WORKING_DIRECTORY "${dir}"
                    OUTPUT_VARIABLE ${prefix}_TAG
                    ERROR_QUIET
                    OUTPUT_STRIP_TRAILING_WHITESPACE
            )

            execute_process(
                COMMAND ${GIT_EXE_EXECUTABLE} -C "${dir}" 
                    rev-parse --abbrev-ref HEAD
                    WORKING_DIRECTORY "${dir}"
                    OUTPUT_VARIABLE ${prefix}_BRANCH
                    ERROR_QUIET
                    OUTPUT_STRIP_TRAILING_WHITESPACE
            )
        endif()
        SET(ENV{LC_ALL} ${_GIT_SAVED_LC_ALL})
    ENDMACRO()
ENDIF(GIT_EXE_EXECUTABLE)

IF(NOT GIT_FOUND)
  IF(NOT GIT_FIND_QUIETLY)
    MESSAGE(STATUS "GIT executable was not found.")
  ELSE(NOT GIT_FIND_QUIETLY)
    IF(GIT_FIND_REQUIRED)
      MESSAGE(FATAL_ERROR "GIT executable was not found.")
    ENDIF(GIT_FIND_REQUIRED)
  ENDIF(NOT GIT_FIND_QUIETLY)
ENDIF(NOT GIT_FOUND)

# FindGIT.cmake ends here.
