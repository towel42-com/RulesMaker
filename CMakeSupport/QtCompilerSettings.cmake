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

find_package(Qt5 COMPONENTS Core REQUIRED) # always need this at a miniumum

include( ${CMAKE_CURRENT_LIST_DIR}/CompilerSettings.cmake )

add_compile_definitions(
    $<$<CONFIG:Debug>:QT_DEBUG>
    $<$<CONFIG:Release>:QT_NO_DEBUG>
    $<$<CONFIG:Release>:QT_NO_NDEBUG>
    $<$<CONFIG:Release>:QT_NO_DEBUG_OUTPUT>
    
    $<$<CONFIG:RelWithDebInfo>:QT_NO_DEBUG>
    $<$<CONFIG:RelWithDebInfo>:QT_NO_NDEBUG>
    $<$<CONFIG:RelWithDebInfo>:QT_NO_DEBUG_OUTPUT>
    
    $<$<CONFIG:MinSizeRel>:QT_NO_DEBUG>
    $<$<CONFIG:MinSizeRel>:QT_NO_NDEBUG>
    $<$<CONFIG:MinSizeRel>:QT_NO_DEBUG_OUTPUT>

    QT_STRICT_ITERATORS 
    QT_CC_WARNINGS 
    QT_NO_WARNINGS
)

