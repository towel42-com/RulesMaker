set(_PROJECT_NAME OutlookAPI)
set(FOLDER_NAME libs)

find_package( FileForTypeID  )

GenerateCPPFromFileID( "{00062FFF-0000-0000-C000-000000000046}" MSOUTL )

#message( "DUMPCPP=${DUMPCPP_EXECUTABLE}" )
#message( "MSOUTL_OLBPATH=${MSOUTL_OLBPATH}" )
#message( "MSOUTL_CPP=${MSOUTL_CPP}" )
#message( "MSOUTL_H=${MSOUTL_H}" )

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
    OutlookAPI_toString.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)
 
set(qtproject_H
    OutlookAPI.h
)

set(project_H
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS    
    Qt5::Widgets Qt5::AxContainer 
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
)

set(qtproject_QRC
)
