set(_PROJECT_NAME RulesMaker-gui)
set(FOLDER_NAME Apps)

set(qtproject_UIS
)

set(project_SRCS
    main.cpp
)
 
set(qtproject_H
)

set(project_H
    ${CMAKE_BINARY_DIR}/Version.h
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
    Qt5::Widgets
    Qt5::AxContainer 
    MainWindow 
    Models 
    OutlookAPI 
    OutlookLib
)

set(qtproject_QRC
)
