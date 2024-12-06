set(_PROJECT_NAME RulesMaker)
set(FOLDER_NAME Apps)


if( NOT DUMPCPP_EXECUTABLE )
    find_program(DUMPCPP_EXECUTABLE NAMES sab_dumpcpp
      PATHS ${CMAKE_INSTALL_PREFIX} ${CMAKE_BINARY_DIR}/dumpcpp/RelWithDebInfo ${CMAKE_BINARY_DIR}/dumpcpp/Release ${CMAKE_BINARY_DIR}/dumpcpp/Debug
        DOC "path to the dumpcpp executable (from build area)" 
        NO_DEFAULT_PATH
#        NO_CACHE
        )
    if( NOT DUMPCPP_EXECUTABLE )
        MESSAGE( WARNING "Could not find build area dumpcpp.  Re-run cmake after initial build" )
    else()
        file( REAL_PATH ${DUMPCPP_EXECUTABLE} DUMPCPP_EXECUTABLE EXPAND_TILDE)
        file( TO_CMAKE_PATH ${DUMPCPP_EXECUTABLE} DUMPCPP_EXECUTABLE)
        mark_as_advanced(DUMPCPP_EXECUTABLE)
    endif()
endif()

if( DUMPCPP_EXECUTABLE )
    MESSAGE( STATUS "Using dumpcpp: '${DUMPCPP_EXECUTABLE}'" )
endif()

set( MSOUTL_OLB "C:/Program Files (x86)/Microsoft Office/Root/Office16/MSOUTL.OLB" )

ADD_CUSTOM_COMMAND( 
    OUTPUT 
        ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    COMMAND ${CMAKE_COMMAND} -E env "PATH=C:/Qt/Qt5.12.12/5.12.12/msvc2017_64/bin;$ENV{PATH}" "${DUMPCPP_EXECUTABLE}" "${MSOUTL_OLB}" -o MSOUTL
    COMMENT "[DUMPCPP] Generating MSOUTL.cpp and MSOUT.h from ${MSOUTL_OLB}"
    VERBATIM
    DEPENDS
       ${MSOUTL_OLB}
)

set(qtproject_UIS
    EmailView.ui
    FoldersView.ui
    MainWindow.ui
    RulesView.ui
    StatusProgress.ui
)

set(project_SRCS
    EmailModel.cpp
    EmailView.cpp
    FoldersView.cpp
    FoldersModel.cpp
    ListFilterModel.cpp
    main.cpp
    MainWindow.cpp
    OutlookAPI.cpp
    OutlookAPI_getEmailAddresses.cpp
    RulesModel.cpp
    RulesView.cpp
    StatusProgress.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)
 
set(qtproject_H
    EmailModel.h
    EmailView.h
    FoldersModel.h
    ListFilterModel.h
    FoldersView.h
    MainWindow.h
    OutlookAPI.h
    RulesModel.h
    RulesView.h
    StatusProgress.h
    WidgetWithStatus.h
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

set(qtproject_QRC
    app.qrc
)
