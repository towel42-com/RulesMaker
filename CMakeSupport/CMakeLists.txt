cmake_minimum_required(VERSION 3.22)

project( CMakeSupport )

FILE(GLOB _CMAKE_SUPPORT_FILES "${CMAKE_SOURCE_DIR}/CMakeSupport/*")
SET(_CMAKE_SUPPORT_FILES ${_CMAKE_SUPPORT_FILES} "${CMAKE_SOURCE_DIR}/CMakeLists.txt" "${CMAKE_SOURCE_DIR}/InstallerInfo.cmake"  "${CMAKE_SOURCE_DIR}/LICENSE.txt")

#MESSAGE( STATUS ${_CMAKE_SUPPORT_FILES} )

source_group("CMake Files" FILES ${_CMAKE_SUPPORT_FILES} )

add_library( CMakeSupport INTERFACE ${_CMAKE_SUPPORT_FILES} )
    
