set(_PROJECT_NAME sab_dumpcpp)
set(FOLDER_NAME Apps)


set(qtproject_UIS
)

set(project_SRCS
    MetaUtils.cpp
    main.cpp
    moc.cpp
    utils.cpp
)

set(qtproject_H
)

set(project_H
    moc.h
    utils.h
    MetaUtils.h
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS
     Qt6::Widgets 
     Qt6::AxContainer 
     Qt6::Gui
     Qt6::CorePrivate
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
)

