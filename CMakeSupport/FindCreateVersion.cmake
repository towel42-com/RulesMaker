find_package(InstallFile REQUIRED)

MACRO(CreateVersion dir versionTemplate )

    set( options )
    set( oneValueArgs MAJOR MINOR PATCH DIFF APP_NAME VENDOR HOMEPAGE PRODUCT_HOMEPAGE EMAIL BUILD_DATE BUILD_TIME COPYRIGHT)
    set( multiValueArgs )
    cmake_parse_arguments(
        _CREATE_VERSION
        "${options}"
        "${oneValueArgs}"
        "${multiValueArgs}"
        ${ARGN}
        )

    set( IN_FILE ${versionTemplate} )

    set(OUTFILE "${CMAKE_BINARY_DIR}/Version.h")
    set(TMP_OUTFILE ${OUTFILE}.tmp)
    
    #message( STATUS "_CREATE_VERSION_MAJOR=${_CREATE_VERSION_MAJOR}" )
    #message( STATUS "_CREATE_VERSION_MINOR=${_CREATE_VERSION_MINOR}" )
    #message( STATUS "_CREATE_VERSION_PATCH=${_CREATE_VERSION_PATCH}" )
    #message( STATUS "_CREATE_VERSION_DIFF=${_CREATE_VERSION_DIFF}" )
    message( STATUS "Generating (or updating) version file '${OUTFILE}'" )
    
    set(VERSION_FILE_MAJOR_VERSION ${_CREATE_VERSION_MAJOR})
    set(VERSION_FILE_MINOR_VERSION ${_CREATE_VERSION_MINOR})
    set(VERSION_FILE_PATCH_VERSION ${_CREATE_VERSION_PATCH})
    MATH( EXPR VERSION_FILE_PATCH_VERSION_LOW  "0x${_CREATE_VERSION_PATCH} & 0x0000FFFF" OUTPUT_FORMAT DECIMAL)
    MATH( EXPR VERSION_FILE_PATCH_VERSION_HIGH "(0x${_CREATE_VERSION_PATCH} & 0xFFFF0000)>>16" OUTPUT_FORMAT DECIMAL)
    STRING(TOLOWER ${_CREATE_VERSION_DIFF} VERSION_FILE_DIFF)
    set(VERSION_FILE_APP_NAME      ${_CREATE_VERSION_APP_NAME})
    set(VERSION_FILE_VENDOR        ${_CREATE_VERSION_VENDOR})
    set(VERSION_FILE_HOMEPAGE      ${_CREATE_VERSION_HOMEPAGE})
    set(VERSION_FILE_PRODUCT_HOMEPAGE ${_CREATE_VERSION_PRODUCT_HOMEPAGE})
    set(VERSION_FILE_EMAIL         ${_CREATE_VERSION_EMAIL})
    set(VERSION_FILE_BUILD_DATE    ${_CREATE_VERSION_BUILD_DATE})
    set(VERSION_FILE_BUILD_TIME    ${_CREATE_VERSION_BUILD_TIME})
    set(VERSION_FILE_COPYRIGHT    ${_CREATE_VERSION_COPYRIGHT})
    set(VERSION_FILE_START_YEAR ${_CREATE_VERSION_START_YEAR})

    configure_file(
        "${IN_FILE}"
        "${TMP_OUTFILE}"
    )

    InstallFile( ${TMP_OUTFILE} ${OUTFILE} REMOVE_ORIG ) # creates a dependency on TMP_OUTFILE
    set_property( 
        DIRECTORY ${dir} 
        APPEND
        PROPERTY CMAKE_CONFIGURE_DEPENDS
        ${OUTFILE}
        )

ENDMACRO()
