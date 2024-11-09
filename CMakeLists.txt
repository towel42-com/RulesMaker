cmake_minimum_required(VERSION 3.22)

project(root-RulesMaker) 

SET(CMAKE_MODULE_PATH ${CMAKE_SOURCE_DIR}/CMakeSupport ${CMAKE_MODULE_PATH})

cmake_policy(SET CMP0144 NEW)
cmake_policy(SET CMP0139 NEW)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED true)

set_property( GLOBAL PROPERTY USE_FOLDERS ON )

add_subdirectory( main )
add_subdirectory( dumpcpp )

