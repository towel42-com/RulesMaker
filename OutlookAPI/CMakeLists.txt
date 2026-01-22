cmake_minimum_required(VERSION 3.22)

find_package(IncludeProjectSettings REQUIRED)
find_package(Deploy REQUIRED)

include( ${CMAKE_CURRENT_LIST_DIR}/include.cmake )
project( ${_PROJECT_NAME} )
IncludeProjectSettings()

add_library( ${PROJECT_NAME} 
                ${_PROJECT_DEPENDENCIES} 
                ${_CMAKE_MODULE_FILES}
          )

set_target_properties( ${PROJECT_NAME} PROPERTIES FOLDER ${FOLDER_NAME} )

target_include_directories( ${PROJECT_NAME} 
    PUBLIC 
        ${_PROJECT_INCLUDE_DIRECTORIES} 
        ${project_pub_LIB_DIRS}
    PRIVATE 
        ${project_pri_DEPS} 
        ${project_pri_LIB_DIRS}
)

target_link_libraries( ${PROJECT_NAME}
    PUBLIC
       ${project_pub_DEPS}
    PRIVATE 
       ${project_pri_DEPS}
)

#add_dependencies(${PROJECT_NAME} OutlookLib)