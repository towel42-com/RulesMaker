cmake_minimum_required(VERSION 3.22)

project(root-RulesMaker 
    VERSION 1.0 
    DESCRIPTION "Tool for generating outlook rules"
    HOMEPAGE_URL http://towel42.com
    ) 

SET(CMAKE_MODULE_PATH ${CMAKE_SOURCE_DIR}/CMakeSupport ${CMAKE_MODULE_PATH})

file( REAL_PATH ~/bin/RulesMaker CMAKE_INSTALL_PREFIX EXPAND_TILDE)
file( REAL_PATH ~/bin HOME_BIN_DIR EXPAND_TILDE)

cmake_policy(SET CMP0144 NEW)
cmake_policy(SET CMP0139 NEW)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED true)

set_property( GLOBAL PROPERTY USE_FOLDERS ON )

add_subdirectory( main )
add_subdirectory( dumpcpp )
add_subdirectory( CMakeSupport )

include( InstallerInfo.cmake )
