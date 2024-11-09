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

MACRO(FindDumpCPP)
    #unset(BISON_EXECUTABLE CACHE)
    #unset(SED_EXECUTABLE CACHE)
    unset(DUMPCPP_EXECUTABLE CACHE)

    if( NOT DUMPCPP_EXECUTABLE )
        find_program(DUMPCPP_EXECUTABLE NAMES dumpcpp
            DOC "path to the dumpcpp executable (from Qt)" 
            REQUIRED 
            )
        mark_as_advanced(DUMPCPP_EXECUTABLE)
        MESSAGE( STATUS "Found dumpcpp: ${DUMPCPP_EXECUTABLE}" )
    endif()
ENDMACRO()


MACRO( POSTPROCESS_INLINE file guard prefix )
    set( OUTPUT_FILE "${CMAKE_CURRENT_BINARY_DIR}/${file}" )
    ADD_CUSTOM_COMMAND( 
        OUTPUT 
            ${OUTPUT_FILE}
        COMMAND "${SED_EXECUTABLE}" -i "'s/${guard}/${prefix}${guard}/g'" "${OUTPUT_FILE}"
        COMMENT "[SED] Post Processing ${OUTPUT_FILE}" 
        DEPENDS 
            ${BISON_VPEParser_OUTPUT_HEADER}
    )
endmacro()

MACRO( POSTPROCESS infile outfile guard prefix )
    set( INPUT_FILE "${CMAKE_CURRENT_BINARY_DIR}/${infile}" )
    set( OUTPUT_FILE "${CMAKE_CURRENT_BINARY_DIR}/${outfile}" )
    ADD_CUSTOM_COMMAND( 
        OUTPUT 
            ${OUTPUT_FILE}
        COMMAND "${SED_EXECUTABLE}" "'s/${guard}/${prefix}${guard}/g'" "${INPUT_FILE}" > "${OUTPUT_FILE}"
        COMMENT "[SED] Post Processing ${INPUT_FILE} to ${OUTPUT_FILE}" 
        DEPENDS 
            ${INPUT_FILE}
    )
endmacro()
