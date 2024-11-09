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

cmake_minimum_required(VERSION 3.22)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED true)
find_package(Threads REQUIRED)

set_property(GLOBAL PROPERTY USE_FOLDERS ON)

SET( THIRD_PARTY_INCLUDE_DIRS ${THIRD_PARTY_DIR}/include )
if( WIN32 )
    SET( THIRD_PARTY_INCLUDE_DIRS ${THIRD_PARTY_DIR}/win_flex ${THIRD_PARTY_INCLUDE_DIRS} )
endif()


find_package(Qt5 5.12 COMPONENTS Core Widgets AxContainer REQUIRED )
SET( _PROJECT_INCLUDE_DIRECTORIES
    ${CMAKE_SOURCE_DIR}
    ${CMAKE_BINARY_DIR}
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_CURRENT_BINARY_DIR}
    ${PROJECT_INCLUDE_DIRS}
    ${Qt5Core_INCLUDE_DIRS} ${Qt5Xml_INCLUDE_DIRS}
)

source_group("Header Files" FILES ${project_H} )
source_group("Source Files" FILES ${project_SRCS} )

SET( _CMAKE_FILES CMakeLists.txt include.cmake ${EXTRA_CMAKE_FILES} )
source_group("CMake Files" FILES ${_CMAKE_FILES} )
FILE(GLOB _CMAKE_MODULE_FILES "${CMAKE_CURRENT_SOURCE_DIR}/CMakeSupport/*")
source_group("CMake Files\\Modules" FILES ${_CMAKE_MODULE_FILES} )

if( EXISTS "${CMAKE_SOURCE_DIR}/CMakeSupport/CompilerSettings.cmake" )
	include( ${CMAKE_SOURCE_DIR}/CMakeSupport/CompilerSettings.cmake )
endif()

SET( _PROJECT_DEPENDENCIES
    ${project_SRCS} 
    ${project_H}  
    ${qtproject_H}  
    ${_CMAKE_FILES}
    ${_CMAKE_MODULE_FILES}
)

SET( project_pub_DEPS
    # insert and "global default" public depends here
     ${project_pub_DEPS}
)

SET( project_pri_DEPS
    # insert and "global default" private depends here
    ${project_pri_DEPS}
)
