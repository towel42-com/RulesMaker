set(_PROJECT_NAME RulesMaker)
set(FOLDER_NAME Apps)


if( NOT DUMPCPP_EXECUTABLE )
    find_program(DUMPCPP_EXECUTABLE NAMES dumpcpp
        DOC "path to the dumpcpp executable (from Qt)" 
        REQUIRED 
        )
    mark_as_advanced(DUMPCPP_EXECUTABLE)
    MESSAGE( STATUS "Found dumpcpp: ${DUMPCPP_EXECUTABLE}" )
endif()

set( MSOUTL_OLB "C:/Program Files (x86)/Microsoft Office/Root/Office16/MSOUTL.OLB" )

ADD_CUSTOM_COMMAND( 
    OUTPUT 
        ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    COMMAND ${DUMPCPP_EXECUTABLE} ${MSOUTL_OLB} -o MSOUTL
    COMMENT "[DUMPCPP] Generating MSOUTL.cpp and MSOUT.h from ${MSOUTL_OLB}"
    VERBATIM
    DEPENDS
       ${MSOUTL_OLB}
)

set(qtproject_UIS
    ContactsView.ui
    FoldersView.ui
    RulesView.ui
)

set(project_SRCS
    ContactsView.cpp
    FoldersView.cpp
    RulesView.cpp
    OutlookHelpers.cpp
    ContactsModel.cpp
    FoldersModel.cpp
    RulesModel.cpp
    main.cpp
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.cpp
    ${MSOUTL_OLB}
)

set(qtproject_H
    ContactsView.h
    FoldersView.h
    RulesView.h
)

set(project_H
    ${CMAKE_CURRENT_BINARY_DIR}/MSOUTL.h
    ContactsModel.h
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

