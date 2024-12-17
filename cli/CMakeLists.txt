cmake_minimum_required(VERSION 3.22)

find_package(IncludeProjectSettings REQUIRED)
find_package(Deploy REQUIRED)

include( ${CMAKE_CURRENT_LIST_DIR}/include.cmake )
project( ${_PROJECT_NAME} )
IncludeProjectSettings()

add_executable( ${PROJECT_NAME}
                ${_PROJECT_DEPENDENCIES} 
                ${_CMAKE_MODULE_FILES}
                resources/application.rc 
          )

set ( DEBUG_PATH 
        "%PATH%;${QT_MSVCDIR}/bin"
        )

set_target_properties( ${PROJECT_NAME} PROPERTIES FOLDER ${FOLDER_NAME} 
                                    VS_DEBUGGER_WORKING_DIRECTORY "$<TARGET_FILE_DIR:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_COMMAND "$<TARGET_FILE:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_ENVIRONMENT "PATH=${DEBUG_PATH}" 
                                    )

target_include_directories( ${PROJECT_NAME} PUBLIC ${_PROJECT_INCLUDE_DIRECTORIES} )

target_link_directories( ${PROJECT_NAME}
    PUBLIC
        ${project_pub_LIB_DIRS}
    PRIVATE
        ${project_pri_LIB_DIRS}
)

target_link_libraries( ${PROJECT_NAME}
    PUBLIC
       ${project_pub_DEPS}
    PRIVATE 
       ${project_pri_DEPS}
)

DeploySystem( ${PROJECT_NAME} . INSTALL_ONLY 1 )
DeployQt( ${PROJECT_NAME} .)
INSTALL( TARGETS ${PROJECT_NAME} RUNTIME DESTINATION . )
INSTALL( FILES ${CMAKE_CURRENT_BINARY_DIR}/$<CONFIG>/${PROJECT_NAME}.pdb DESTINATION . CONFIGURATIONS Debug RelWithDebInfo )
