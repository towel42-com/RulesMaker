SET( CPACK_PACKAGE_VERSION_MAJOR ${MAJOR_VERSION} )
    
SET( CPACK_GENERATOR ZIP )

SET( CPACK_NSIS_ENABLE_UNINSTALL_BEFORE_INSTALL ON )
SET( CPACK_NSIS_EXECUTABLES_DIRECTORY "." )

include( CPack )
