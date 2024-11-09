cmake_minimum_required(VERSION 3.22)

find_package(IncludeProjectSettings REQUIRED)
include( ${CMAKE_CURRENT_LIST_DIR}/include.cmake )
project( ${_PROJECT_NAME} )
IncludeProjectSettings()

add_executable( ${PROJECT_NAME}
                ${_PROJECT_DEPENDENCIES} 
                ${_CMAKE_MODULE_FILES}
          )
set ( DEBUG_PATH 
        "%PATH%;C:/Qt/Qt5.12.12/5.12.12/msvc2017_64/bin"
        )

set_target_properties( ${PROJECT_NAME} PROPERTIES FOLDER ${FOLDER_NAME} 
                                    VS_DEBUGGER_WORKING_DIRECTORY "$<TARGET_FILE_DIR:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_COMMAND "$<TARGET_FILE:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_ENVIRONMENT "PATH=${DEBUG_PATH}" 
                                    )

target_include_directories( ${PROJECT_NAME} PUBLIC ${_PROJECT_INCLUDE_DIRECTORIES} ${CMAKE_BINARY_DIR}/src/vpe )
target_include_directories( ${PROJECT_NAME} PUBLIC C:/Qt/Qt5.12.12/5.12.12/./msvc2017_64/include/QtCore/5.12.12/ )
target_include_directories( ${PROJECT_NAME} PUBLIC C:/Qt/Qt5.12.12/5.12.12/./msvc2017_64/include/QtCore/5.12.12/QtCore/ )

target_link_directories( ${PROJECT_NAME}
    PUBLIC
        ${project_pub_LIB_DIRS}
    PRIVATE
        ${project_pri_LIB_DIRS}
)

target_link_libraries( ${PROJECT_NAME}
    PUBLIC
       ${project_pub_DEPS} Qt5::Widgets
    PRIVATE 
       ${project_pri_DEPS}
)


