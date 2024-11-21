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

find_package(DeploySystem REQUIRED)

if( NOT DEFINED USE_QT OR USE_QT )
	find_package(Qt5Core REQUIRED)

	if( NOT DEFINED DEPLOYQT_EXECUTABLE )
		# Retrieve the absolute path to qmake and then use that path to find
		# the <os>deployqt binaries
		get_target_property(_qmake_executable Qt5::qmake IMPORTED_LOCATION)
		get_filename_component(_qt_bin_dir "${_qmake_executable}" DIRECTORY)

		if( WIN32 )
			find_program(DEPLOYQT_EXECUTABLE windeployqt HINTS "${_qt_bin_dir}")
			if(NOT DEPLOYQT_EXECUTABLE)
				message(FATAL_ERROR "windeployqt not found")
			endif()
			message(STATUS "Found windeployqt: ${DEPLOYQT_EXECUTABLE}")
            # Doing this with MSVC 2015 requires CMake 3.6+
            if( (MSVC_VERSION VERSION_EQUAL 1900 OR MSVC_VERSION VERSION_GREATER 1900) AND CMAKE_VERSION VERSION_LESS "3.6")
                message(WARNING "Deploying with MSVC 2015+ requires CMake 3.6+")
		endif()
        ELSEIF( APPLE )
            find_program(DEPLOYQT_EXECUTABLE macdeployqt HINTS "${_qt_bin_dir}")
            if(NOT DEPLOYQT_EXECUTABLE)
                message(FATAL_ERROR "macdeployqt not found")
            endif()
            message(STATUS "Found macdeployqt: ${DEPLOYQT_EXECUTABLE}")
        ELSEIF( UNIX )
            find_program(DEPLOYQT_EXECUTABLE linuxdeployqt HINTS "${_qt_bin_dir}")
            if(NOT DEPLOYQT_EXECUTABLE)
                message(WARNING "linuxdeployqt not found")
            else()
                message(STATUS "Found macdeployqt: ${DEPLOYQT_EXECUTABLE}")
            endif()
        ENDIF()
	mark_as_advanced(DEPLOYQT_EXECUTABLE)
endif()


	# Add commands that copy the required Qt files to the same directory as the
	# target after being built as well as including them in final installation
	function(DeployQt target directory)
		if(NOT DEPLOYQT_EXECUTABLE)
			IF( UNIX )
				return()
			ENDIF()

			message(FATAL_ERROR "deployqt not found")
		endif()

		set( options )
		set( oneValueArgs INSTALL_ONLY NON_INSTALL_ONLY NO_TRANSLATIONS )
		set( multiValueArgs )

		cmake_parse_arguments( "" "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN} )
		if ( _NO_TRANSLATIONS )
			SET( NO_TRANSLATIONS_OPT "--no-translations")
		endif()

		if ( NOT _INSTALL_ONLY )
			SET(_QTDEPLOY_TARGET_DIR "$<TARGET_FILE:${target}>" )
			IF( WIN32 )
				SET(_QTDEPLOY_OPTIONS_LCL "--verbose=1;--no-compiler-runtime;--no-angle;--no-opengl-sw;--pdb;${NO_TRANSLATIONS_OPT}" )
			ELSEIF( APPLE )
				SET(_QTDEPLOY_TARGET_DIR "$<TARGET_FILE:${target}>/../.." )
				SET(_QTDEPLOY_OPTIONS_LCL "--verbose=1;--no-compiler-runtime;--always-overwrite;${NO_TRANSLATIONS_OPT}" )
			ELSEIF( UNIX )
				SET(_QTDEPLOY_OPTIONS_LCL "-verbose=1;${NO_TRANSLATIONS_OPT}" )
				return()
			ENDIF()

			# Run deployqt immediately after build to make the build area "complete"
			add_custom_command(TARGET ${target} POST_BUILD
				COMMAND "${CMAKE_COMMAND}" -E echo "Deploying Qt to Build Area for Project '${target}' using '${DEPLOYQT_EXECUTABLE}' ..."
				COMMAND "${CMAKE_COMMAND}" -E
					env PATH="${_qt_bin_dir}" "${DEPLOYQT_EXECUTABLE}"
						${_QTDEPLOY_OPTIONS_LCL}
						${_QTDEPLOY_TARGET_DIR}
			)
		endif()
		
		# install(CODE ...) doesn't support generator expressions, but
		# file(GENERATE ...) does - store the path in a file
		file(GENERATE 
			OUTPUT "${CMAKE_CURRENT_BINARY_DIR}/${target}_$<CONFIG>_path"
			CONTENT "$<TARGET_FILE:${target}>"
		)

		# Before installation, run a series of commands that copy each of the Qt
		# runtime files to the appropriate directory for installation
		if ( NOT _NON_INSTALL_ONLY )
			install(CODE
				"
				file(READ \"${CMAKE_CURRENT_BINARY_DIR}/${target}_\${CMAKE_INSTALL_CONFIG_NAME}_path\" _file)
				IF( WIN32 )
					SET(_QTDEPLOY_OPTIONS \"--dry-run;--list;mapping;--no-compiler-runtime;--no-angle;--no-opengl-sw;${NO_TRANSLATIONS_OPT}\" )
				ELSEIF( APPLE )
					SET(_QTDEPLOY_OPTIONS \"--dry-run;--list;mapping;${NO_TRANSLATIONS_OPT}\" )
				ELSEIF( UNIX )
					SET(_QTDEPLOY_OPTIONS \"--dry-run;--list;mapping;${NO_TRANSLATIONS_OPT}\" )
				ENDIF()

				MESSAGE( STATUS \"Deploying Qt to the Install Area '\${CMAKE_INSTALL_PREFIX}/${directory}' for Project '${target}' using '${DEPLOYQT_EXECUTABLE}' ...\" )
				execute_process(
					COMMAND \"${CMAKE_COMMAND}\" -E
						env PATH=\"${_qt_bin_dir}\" \"${DEPLOYQT_EXECUTABLE}\"
							\${_QTDEPLOY_OPTIONS}
							\${_file}
					OUTPUT_VARIABLE _output
					OUTPUT_STRIP_TRAILING_WHITESPACE
				)
				separate_arguments(_files NATIVE_COMMAND \${_output})
				while(_files)
					list(GET _files 0 _src)
					list(GET _files 1 _dest)
					execute_process(
						COMMAND \"${CMAKE_COMMAND}\" -E
							compare_files \"\${_src}\" \"\${CMAKE_INSTALL_PREFIX}/\${directory}/\${_dest}\"
							OUTPUT_VARIABLE _outvar
							ERROR_VARIABLE _errvar
							RESULT_VARIABLE _result_code
					)
					if( \${_result_code} )
						MESSAGE( STATUS \"Installing: \${CMAKE_INSTALL_PREFIX}/\${directory}/\${_dest}\" )
						execute_process(
							COMMAND \"${CMAKE_COMMAND}\" -E
								copy \${_src} \"\${CMAKE_INSTALL_PREFIX}/\${directory}/\${_dest}\"
						)
					ELSE()
						MESSAGE( STATUS \"Up-to-date: \${CMAKE_INSTALL_PREFIX}/${directory}/\${_dest}\" )
					ENDIF()
					list(REMOVE_AT _files 0 1)
				endwhile()
				MESSAGE( STATUS \"Finished deploying Qt\" )
				"
			)
		endif()
	endfunction()
endif()

function (PrintList listVar)
    MESSAGE( STATUS "List -> ${listVar}:" )
    foreach(curr ${${listVar}})
        MESSAGE( STATUS "    ${curr}" )
    endforeach()
endfunction()

