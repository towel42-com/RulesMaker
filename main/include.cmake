set(_PROJECT_NAME RulesMaker)
set(FOLDER_NAME Apps)


if( NOT DUMPCPP_EXECUTABLE )
    find_program(DUMPCPP_EXECUTABLE NAMES dumpcpp
      PATHS ${CMAKE_BINARY_DIR}/dumpcpp/RelWithDebInfo  ${CMAKE_BINARY_DIR}/dumpcpp/Release ${CMAKE_BINARY_DIR}/dumpcpp/Debug
        DOC "path to the dumpcpp executable (from build area)" 
        NO_DEFAULT_PATH
        )
    if( NOT DUMPCPP_EXECUTABLE )
        MESSAGE( WARNING "Could not find build area dumpcpp.  Re-run cmake after initial build" )
    else()
        mark_as_advanced(DUMPCPP_EXECUTABLE)
        MESSAGE( STATUS "Found dumpcpp: ${DUMPCPP_EXECUTABLE}" )
    endif()
endif()

set( MSOUTL_OLB "C:/Program Files (x86)/Microsoft Office/Root/Office16/MSOUTL.OLB" )

ADD_CUSTOM_COMMAND( 
    OUTPUT 
        ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    COMMAND "${DUMPCPP_EXECUTABLE}" "${MSOUTL_OLB}" -o MSOUTL
    COMMENT "[DUMPCPP] Generating MSOUTL.cpp and MSOUT.h from ${MSOUTL_OLB}"
    VERBATIM
    DEPENDS
       ${MSOUTL_OLB}
)

set(qtproject_UIS
    ContactsView.ui
    FoldersView.ui
    RulesView.ui
    EmailView.ui
    OutlookSetup.ui
)

set(project_SRCS
    ContactsView.cpp
    EmailView.cpp
    EmailGroupingModel.cpp
    FoldersView.cpp
    RulesView.cpp
    OutlookHelpers.cpp
    ContactsModel.cpp
    EmailModel.cpp
    FoldersModel.cpp
    OutlookSetup.cpp
    RulesModel.cpp
    main.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)

set(qtproject_H
    ContactsView.h
    EmailModel.h
    EmailView.h
    FoldersView.h
    RulesView.h
    OutlookSetup.h
)

set(project_H
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    ContactsModel.h
    EmailGroupingModel.h
    FoldersModel.h
    RulesModel.h
    OutlookHelpers.h
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
)

