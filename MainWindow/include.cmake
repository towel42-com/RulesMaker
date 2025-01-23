set(_PROJECT_NAME MainWindow)
set(FOLDER_NAME libs)

set(qtproject_UIS
    FilterFromEmailView.ui
    FoldersView.ui
    MainWindow.ui
    RulesView.ui
    Settings.ui
)

set(project_SRCS
    FilterFromEmailView.cpp
    FoldersView.cpp
    MainWindow.cpp
    RulesView.cpp
    Settings.cpp
    StatusProgress.cpp
)
 
set(qtproject_H
    FilterFromEmailView.h
    FoldersView.h
    MainWindow.h
    RulesView.h
    Settings.h
    StatusProgress.h
    WidgetWithStatus.h
)

set(project_H
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS
    Qt5::Widgets Qt5::AxContainer OutlookAPI
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
