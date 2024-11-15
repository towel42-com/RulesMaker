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
    EmailView.ui
    FoldersView.ui
    FoldersDlg.ui
    MainWindow.ui
    OutlookSetup.ui
    RulesView.ui
)

set(project_SRCS
    GroupedEmailModel.cpp
    EmailView.cpp
    FoldersDlg.cpp
    FoldersView.cpp
    FoldersModel.cpp
    main.cpp
    MainWindow.cpp
    OutlookHelpers.cpp
    OutlookSetup.cpp
    RulesModel.cpp
    RulesView.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)

set(qtproject_H
    GroupedEmailModel.h
    EmailView.h
    FoldersDlg.h
    FoldersModel.h
    FoldersView.h
    MainWindow.h
    OutlookHelpers.h
    OutlookSetup.h
    RulesModel.h
    RulesView.h
)

set(project_H
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
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

