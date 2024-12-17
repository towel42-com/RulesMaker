set(_PROJECT_NAME MainWindow)
set(FOLDER_NAME libs)

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
    MainWindow.cpp
    RulesModel.cpp
    RulesView.cpp
    StatusProgress.cpp
)
 
set(qtproject_H
    EmailModel.h
    EmailView.h
    FoldersModel.h
    ListFilterModel.h
    FoldersView.h
    MainWindow.h
    RulesModel.h
    RulesView.h
    StatusProgress.h
    WidgetWithStatus.h
)

set(project_H
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
    MainWindow.qrc
)
