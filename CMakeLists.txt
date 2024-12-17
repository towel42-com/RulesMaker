cmake_minimum_required(VERSION 3.22)

project(root-RulesMaker 
    VERSION 1.0 
    DESCRIPTION "Tool for generating Outlook Rules"
    HOMEPAGE_URL http://towel42.com
    ) 

find_package(Qt5 COMPONENTS Core Widgets AxContainer REQUIRED)

set( QT_VERSION ${Qt5Core_VERSION_STRING})
set( QT_MSVCDIR ${_qt5Core_install_prefix} )

SET(CMAKE_MODULE_PATH ${CMAKE_SOURCE_DIR}/CMakeSupport ${CMAKE_MODULE_PATH})

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
add_subdirectory( OutlookAPI )
add_subdirectory( dumpcpp )
add_subdirectory( CMakeSupport )

include( InstallerInfo.cmake )
