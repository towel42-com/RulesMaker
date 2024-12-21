set(_PROJECT_NAME OutlookAPI)
set(FOLDER_NAME libs)

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
endif()


if( DUMPCPP_EXECUTABLE )
    MESSAGE( STATUS "Using dumpcpp: '${DUMPCPP_EXECUTABLE}'" )
endif()

find_package( FileForTypeID  )
FileForTypeID( "{00062FFF-0000-0000-C000-000000000046}" MSOUTL_OLB )
#message( STATUS "MSOUTL_OLB_PATH=${MSOUTL_OLB_PATH}" )
if ( NOT EXISTS ${MSOUTL_OLB_PATH} )
    message( FATAL_ERROR "Could not find OLB file for MS Outlook" )
endif()

ADD_CUSTOM_COMMAND( 
    OUTPUT 
        ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    COMMAND ${CMAKE_COMMAND} -E env "PATH=${QT_MSVCDIR}/bin;$ENV{PATH}" "${DUMPCPP_EXECUTABLE}" "${MSOUTL_OLB_PATH}" -o MSOUTL
    COMMENT "[DUMPCPP] Generating MSOUTL.cpp and MSOUT.h from '${MSOUTL_OLB_PATH}' using '${DUMPCPP_EXECUTABLE}'"
    VERBATIM
    DEPENDS
       ${MSOUTL_OLB}
)

set(qtproject_UIS
)

set(project_SRCS
    OutlookAPI.cpp
    OutlookAPI_account.cpp
    OutlookAPI_dump.cpp
    OutlookAPI_email.cpp
    OutlookAPI_folders.cpp
    OutlookAPI_rules.cpp
    OutlookAPI_settings.cpp
    OutlookAPI_tools.cpp
    OutlookAPI_toString.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)
 
set(qtproject_H
    OutlookAPI.h
)

set(project_H
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS    
    Qt5::Widgets Qt5::AxContainer 
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
)

set(qtproject_QRC
)
