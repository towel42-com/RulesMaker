set(_PROJECT_NAME OutlookAPI)
set(FOLDER_NAME libs)

set(qtproject_UIS
)

set(project_SRCS
    OutlookAPI.cpp
    OutlookAPI_account.cpp
    OutlookAPI_dump.cpp
    OutlookAPI_email.cpp
    OutlookAPI_folders.cpp
    OutlookAPI_rules.cpp
    OutlookAPI_settings.cpp
    OutlookAPI_tools.cpp
    OutlookAPI_utils.cpp
)
 
set(qtproject_H
    OutlookAPI.h
)

set(project_H
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS    
    Qt5::Widgets Qt5::AxContainer OutlookLib
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
)

set(qtproject_QRC
)
