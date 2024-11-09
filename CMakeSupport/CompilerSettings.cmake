#*******************************************************************************
#
#  SYNOPSYS CONFIDENTIAL - This is an unpublished, proprietary work of
#  Synopsys, Inc., and is fully protected under copyright and trade
#  secret laws. You may not view, use, disclose, copy, or distribute this
#  file or any information contained herein except pursuant to a valid
#  written license from Synopsys.
#
#*******************************************************************************
#*******************************************************************************

IF(CMAKE_SIZEOF_VOID_P EQUAL 4) #32 bit
    SET( BITSIZE 32 )
ELSEIF(CMAKE_SIZEOF_VOID_P EQUAL 8) #64 bit
    SET( BITSIZE 64 )
ELSE () 
    MESSAGE( STATUS "Unknown Bitsize - CMAKE_SIZEOF_VOID_P not set to 4 or 8" )
    MESSAGE( STATUS "-DCMAKE_SIZEOF_VOID_P=4 for 32 bit" )
    MESSAGE( FATAL_ERROR "-DCMAKE_SIZEOF_VOID_P=8 for 64 bit" )
ENDIF() 

IF(NOT WIN32)
    IF( NOT WARN_ALL )
        MESSAGE( STATUS  "Warning all enabled - ${CMAKE_PROJECT_NAME}" )
    ENDIF()
ENDIF()

IF(${NO_EXTENSIVE_WARNINGS})
    MESSAGE( WARNING "Extensive warnings disabled - ${PROJECT_NAME}" )
ENDIF()

#IF( NO_WARNING_AS_ERROR)
#    MESSAGE( WARNING "Warning as an Error disabled - ${PROJECT_NAME}" )
#ENDIF()

IF(${NO_EXTENSIVE_WARNINGS})
    MESSAGE( WARNING "Extensive warnings disabled - ${PROJECT_NAME}" )
ENDIF()

set(CMAKE_INCLUDE_CURRENT_DIR ON)

add_compile_definitions( 
    _SILENCE_CXX17_ITERATOR_BASE_CLASS_DEPRECATION_WARNING 
    _SILENCE_ALL_CXX17_DEPRECATION_WARNINGS  
    UNICODE 

    YY_NO_UNISTD_H
    DWBBCT_DEBUG
    $<IF:$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Debug>>,_DEBUG,NDEBUG>

    $<$<CXX_COMPILER_ID:MSVC>:NO_LICENSE_CHECK>

    $<$<CXX_COMPILER_ID:MSVC>:_CRT_SECURE_NO_WARNINGS>
    $<$<CXX_COMPILER_ID:MSVC>:_CRT_SECURE_NO_DEPRECATE>
    $<$<CXX_COMPILER_ID:MSVC>:_CRT_NONSTDC_NO_WARNINGS>
    $<$<CXX_COMPILER_ID:MSVC>:_SCL_SECURE_NO_WARNINGS>
    
    $<IF:$<CXX_COMPILER_ID:MSVC>,Synopsys_Win32,Synopsys_linux>
    $<IF:$<CXX_COMPILER_ID:MSVC>,WIN32,LINUX>
    $<$<AND:$<NOT:$<CXX_COMPILER_ID:MSVC>>,$<EQUAL:${BITSIZE},64>>:Synopsys_amd64>
    $<$<AND:$<NOT:$<CXX_COMPILER_ID:MSVC>>,$<EQUAL:${BITSIZE},64>>:Synopsys_linux64>

#    $<$<CONFIG:Debug>:DWBBCT_DEBUG>
)

add_link_options(
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<EQUAL:${BITSIZE},64>>:/STACK:18388608>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<EQUAL:${BITSIZE},64>>:/HIGHENTROPYVA:NO>
)

add_compile_options(
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<EQUAL:${BITSIZE},32>,$<VERSION_GREATER_EQUAL:${MSVC_VERSION},1800>>:/SAFESEH:NO>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<EQUAL:${BITSIZE},64>>:/bigobj>

    $<$<CXX_COMPILER_ID:MSVC>:/EHsc>
    $<$<CXX_COMPILER_ID:MSVC>:/MP>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4005>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4013>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4024>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4028>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4047>
    $<$<CXX_COMPILER_ID:MSVC>:/w34062> 
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4065>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4090>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<NOT:$<BOOL:${NO_WARNING_AS_ERROR}>>,$<NOT:$<BOOL:${SOFT_COMPILE}>>>:/w34100>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4100>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4101>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4133>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4143>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4146>

    $<$<CXX_COMPILER_ID:MSVC>:/wd4231> 
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4244>
    $<$<CXX_COMPILER_ID:MSVC>:/wd4251> 
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4267>
    $<$<CXX_COMPILER_ID:MSVC>:/wd4273> 

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4311>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4312>

    $<$<CXX_COMPILER_ID:MSVC>:/wd4503>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4606>
    $<$<CXX_COMPILER_ID:MSVC>:/w34700> 
    $<$<CXX_COMPILER_ID:MSVC>:/w34701> 
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4715>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<NOT:$<BOOL:${SOFT_COMPILE}>>>:/w34715>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4716>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<NOT:$<BOOL:${SOFT_COMPILE}>>>:/w34716>

    $<$<CXX_COMPILER_ID:MSVC>:/w34717> 
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${SOFT_COMPILE}>>:/wd4739>

    $<$<CXX_COMPILER_ID:MSVC>:/w44800> 



    #$<$<AND:$<CXX_COMPILER_ID:MSVC>,$<NOT:$<BOOL:${NO_WARNING_AS_ERROR}>>>:/WX>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<BOOL:${WARN_ALL}>>:/W4>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Debug>>:/W3>

    $<$<CXX_COMPILER_ID:MSVC>:/Zc:__cplusplus>
    $<$<CXX_COMPILER_ID:MSVC>:/Zc:__STDC__>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Debug>>:/MDd>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Debug>>:/ZI>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Debug>>:/Od>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/MD>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/Zi>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/O2>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:Release>>:/Ob1>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:RelWithDebInfo>>:/MD>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:RelWithDebInfo>>:/Zi>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:RelWithDebInfo>>:/O2>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:RelWithDebInfo>>:/Ob1>

    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:MinSizeRel>>:/MD>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:MinSizeRel>>:/Zi>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:MinSizeRel>>:/O2>
    $<$<AND:$<CXX_COMPILER_ID:MSVC>,$<CONFIG:MinSizeRel>>:/Ob1>
    
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Waddress>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wchar-subscripts>

    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wcomment>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wformat>

    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wmaybe-uninitialized>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wmissing-braces>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wnonnull>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wparentheses>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wreturn-type>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wsequence-point>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wsign-compare> # big changes

    # we dont use -fstrict aliasing or -fstrict-overflow the following 2 are worthless
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wstrict-aliasing>
    # $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wstrict-overflow=1>

    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wswitch>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wtrigraphs>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wuninitialized>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunknown-pragmas> # have to figure out a way with windows
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-function>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-label>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-value>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-variable>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wvolatile-register-var>

    #the following are form -Wextra
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wclobbered>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wempty-body>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wignored-qualifiers>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wmissing-field-initializers>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wtype-limits>
    # $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wshift-negative-value> #not available on our compiler
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-parameter>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wunused-but-set-parameter>

    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wno-return-local-addr>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wreorder>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wc++11-compat>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Werror=suggest-override>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_EXTENSIVE_WARNINGS}>>>:-Wc++14-compat> # not available on our compiler

    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<NOT:$<BOOL:${NO_WARNING_AS_ERROR}>>>:-Werror>
    $<$<AND:$<CXX_COMPILER_ID:GNU>,$<BOOL:${WARN_ALL}$>>:-Wall>
    $<$<CXX_COMPILER_ID:GNU>:-fPIC>
)

