cmake_minimum_required(VERSION 3.22)

project(root-RulesMaker 
    VERSION 1.0 
    DESCRIPTION "Tool for generating Outlook Rules"
    HOMEPAGE_URL http://towel42.com
    ) 

SET( T42_GLOBAL_USE_QT true )
SET(CMAKE_MODULE_PATH ${CMAKE_SOURCE_DIR}/T42-CMakeUtils/Modules ${CMAKE_MODULE_PATH})

find_package(Qt6 COMPONENTS Core Widgets AxContainer REQUIRED)
find_package( T42Utils REQUIRED )

#PrintAllVariables()

if ( ( "${Qt6_VERSION}" STREQUAL "" ) OR ( NOT EXISTS ${QT6_INSTALL_PREFIX} ) )
    if( ( "${Qt6_VERSION}" STREQUAL "" ) AND ( NOT EXISTS ${QT6_INSTALL_PREFIX} ) )
        MESSAGE( FATAL_ERROR "Could not determine Qt6_VERSION OR QT6_INSTALL_PREFIX" )
    endif()
    
    if ( "${Qt6_VERSION}" STREQUAL "" )
        MESSAGE( FATAL_ERROR "Could not determine Qt6_VERSION" )
    endif()

    if ( NOT EXISTS ${QT6_INSTALL_PREFIX} )
        MESSAGE( FATAL_ERROR "Could not determine QT6_INSTALL_PREFIX" )
    endif()
endif()
    

MESSAGE( STATUS "Qt6_VERSION=${Qt6_VERSION}" )
MESSAGE( STATUS "QT6_INSTALL_PREFIX=${QT6_INSTALL_PREFIX}" )

include( ${CMAKE_SOURCE_DIR}/VersionInfo.cmake )

file( REAL_PATH ~/bin/RulesMaker CMAKE_INSTALL_PREFIX EXPAND_TILDE)
file( REAL_PATH ~/bin HOME_BIN_DIR EXPAND_TILDE)

cmake_policy(SET CMP0144 NEW)
cmake_policy(SET CMP0139 NEW)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED true)

set_property( GLOBAL PROPERTY USE_FOLDERS ON )

add_subdirectory( gui )
add_subdirectory( cli )
add_subdirectory( MainWindow )
add_subdirectory( Models )
add_subdirectory( OutlookAPI )
add_subdirectory( OutlookLib )
add_subdirectory( dumpcpp )

include( InstallerInfo.cmake )
