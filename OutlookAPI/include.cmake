set(_PROJECT_NAME OutlookAPI)
set(FOLDER_NAME libs)

set(qtproject_UIS
    DelayDlg.ui
    SelectAccount.ui
    SelectFolders.ui
    ShowRule.ui
)

set(project_SRCS
    DelayDlg.cpp
    EmailAddress.cpp
    OutlookAPI.cpp
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
    SelectFolders.cpp
)
 
set(qtproject_H
    DelayDlg.h
    OutlookAPI.h
    SelectAccount.h
    SelectFolders.h
    ShowRule.h
)

set(project_H
    EmailAddress.h
    OutlookAPI_pri.h
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS    
    Qt6::Widgets 
    Qt6::AxContainer 
    OutlookLib
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
    ${CMAKE_BINARY_DIR}
)

set( project_pri_DEPS
)

set(qtproject_QRC
)
