# The MIT License (MIT)
#
# Copyright (c) 2017 Nathan Osman
# Copyright (c) 2020-2021 Scott Aron Bloom - Work on linux and mac
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sub-license, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

function( DeploySystem target directory)
    set( options )
    set( oneValueArgs INSTALL_ONLY )
    set( multiValueArgs )

    cmake_parse_arguments( "" "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )
    
    #message( STATUS "Deploy System ${target}" )
    if ( WIN32 )
        set(CMAKE_INSTALL_UCRT_LIBRARIES FALSE)
        #set(CMAKE_INSTALL_DEBUG_LIBRARIES TRUE ) 
    ENDIF()

    # deployqt doesn't work correctly with the system runtime libraries,
    # so we fall back to one of CMake's own modules for copying them over
    SET(CMAKE_INSTALL_SYSTEM_RUNTIME_DESTINATION .)
    include(InstallRequiredSystemLibraries)

    if ( NOT _INSTALL_ONLY )
        #message( STATUS "${CMAKE_INSTALL_SYSTEM_RUNTIME_LIBS}" )
        foreach(lib ${CMAKE_INSTALL_SYSTEM_RUNTIME_LIBS})
            get_filename_component(filename "${lib}" NAME)
            add_custom_command(TARGET ${target} POST_BUILD
                COMMAND "${CMAKE_COMMAND}" -E echo "Deploying System Library '${filename}' for '${target}'"
                COMMAND "${CMAKE_COMMAND}" -E copy_if_different "${lib}" \"$<TARGET_FILE_DIR:${target}>\"
            )
        endforeach()
    endif()
endfunction()

