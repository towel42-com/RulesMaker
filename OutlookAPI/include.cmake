set(_PROJECT_NAME OutlookAPI)
set(FOLDER_NAME libs)

set(qtproject_UIS
    SelectAccount.ui
    ShowRule.ui
)

set(project_SRCS
    EmailAddress.cpp
    ExceptionHandler.cpp
    OutlookAPI.cpp
    OutlookObj.cpp
    OutlookAPI_account.cpp
    OutlookAPI_dump.cpp
    OutlookAPI_email.cpp
    OutlookAPI_emailAddresses.cpp
    OutlookAPI_folders.cpp
    OutlookAPI_rules.cpp
    OutlookAPI_copyRules.cpp
    OutlookAPI_equalRules.cpp
    OutlookAPI_loadRules.cpp
    OutlookAPI_nameForRules.cpp
    OutlookAPI_mergeRules.cpp
    OutlookAPI_settings.cpp
    OutlookAPI_tools.cpp
    OutlookAPI_utils.cpp
    ShowRule.cpp
    SelectAccount.cpp
)
 
set(qtproject_H
    OutlookAPI.h
    ExceptionHandler.h
    SelectAccount.h
    ShowRule.h
)

set(project_H
    OutlookObj.h
    EmailAddress.h
    OutlookAPI_pri.h
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
