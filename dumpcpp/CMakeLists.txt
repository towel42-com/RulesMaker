cmake_minimum_required(VERSION 3.22)

find_package(IncludeProjectSettings REQUIRED)
find_package(Deploy REQUIRED)

include( ${CMAKE_CURRENT_LIST_DIR}/include.cmake )
project( ${_PROJECT_NAME} )
IncludeProjectSettings()

add_executable( ${PROJECT_NAME}
                ${_PROJECT_DEPENDENCIES} 
                ${_CMAKE_MODULE_FILES}
          )
set ( DEBUG_PATH 
        "%PATH%;${QT6_INSTALL_PREFIX}/bin"
        )

set_target_properties( ${PROJECT_NAME} PROPERTIES FOLDER ${FOLDER_NAME} 
                                    VS_DEBUGGER_WORKING_DIRECTORY "$<TARGET_FILE_DIR:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_COMMAND "$<TARGET_FILE:${PROJECT_NAME}>" 
                                    VS_DEBUGGER_ENVIRONMENT "PATH=${DEBUG_PATH}" 
                                    )

target_compile_definitions(${PROJECT_NAME} 
    PRIVATE 
        QT_ASCII_CAST_WARNINGS QT_NO_CAST_TO_ASCII)

if ( EXISTS ${QT6_INSTALL_PREFIX}/include/QtCore/${Qt6_VERSION}/QtCore/private/qmetaobject_p.h )
    target_include_directories( ${PROJECT_NAME} PUBLIC ${QT6_INSTALL_PREFIX}/include/QtCore/${Qt6_VERSION}/ )
    target_include_directories( ${PROJECT_NAME} PUBLIC ${QT6_INSTALL_PREFIX}/include/QtCore/${Qt6_VERSION}/QtCore/ )
else()
     MESSAGE( FATAL_ERROR 
"Could not find qmetaobject_p.h at
    '${QT6_INSTALL_PREFIX}/include/QtCore/${Qt6_VERSION}/QtCore/private/qmetaobject_p.h'" )
endif()

target_include_directories( ${PROJECT_NAME} PUBLIC ${CMAKE_BINARY_DIR} )

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

