set(_PROJECT_NAME OutlookLib)
set(FOLDER_NAME libs)

#message( "DUMPCPP=${DUMPCPP_EXECUTABLE}" )
find_package( FileForTypeID  )


GenerateCPPFromFileID( "{00062FFF-0000-0000-C000-000000000046}" MSOUTL ol )

#message( "MSOUTL_TYPEID_FILEPATH=${MSOUTL_TYPEID_FILEPATH}" )
#message( "MSOUTL_CPP=${MSOUTL_CPP}" )
#message( "MSOUTL_H=${MSOUTL_H}" )
#GenerateCPPFromFileID( "{2735412F-7F64-5B0F-8F00-5D77AFBE261E}" IMAPI2 im )
#message( "IMAPI2_TYPEID_FILEPATH=${IMAPI2_TYPEID_FILEPATH}" )
#message( "IMAPI2_CPP=${IMAPI2_CPP}" )
#message( "IMAPI2_H=${IMAPI2_H}" )


set(qtproject_UIS
)

set(project_SRCS
    ${MSOUTL_CPP}
    ${IMAPI2_CPP}
)
 
set(qtproject_H
)

set(project_H
    ${MSOUTL_H}
    ${IMAPI2_H}
)

set( project_pub_LIB_DIRS 
)

set( project_pub_DEPS    
    Qt6::Widgets Qt6::AxContainer 
)

set( EXTRA_CMAKE_FILES
)

set( project_pri_LIB_DIRS 
)

set( project_pri_DEPS
)

set(qtproject_QRC
)
