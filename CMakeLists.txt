cmake_minimum_required(VERSION 3.24)

project(
    ExcelCompareProfessional
    VERSION 2.0.4
    LANGUAGES C CXX
)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

find_package(
    Qt6 6.5 REQUIRED
    COMPONENTS
        Core
        Gui
        Widgets
        Concurrent
        PrintSupport
)

include(FetchContent)

if(POLICY CMP0169)
    cmake_policy(SET CMP0169 OLD)
endif()

set(BUILD_TESTING OFF CACHE BOOL "" FORCE)
set(QXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(QXLSX_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)

FetchContent_Declare(
    QXlsx
    GIT_REPOSITORY https://github.com/QtExcel/QXlsx.git
    GIT_TAG v1.5.0
    SOURCE_SUBDIR QXlsx
)

FetchContent_MakeAvailable(QXlsx)

# Native legacy Excel (.xls) reader: libxls 1.6.3
# libxls does not ship a CMake project, so its five library sources are
# compiled directly as a static dependency.
FetchContent_Declare(
    libxls
    GIT_REPOSITORY https://github.com/libxls/libxls.git
    GIT_TAG v1.6.3
    GIT_SHALLOW TRUE
)

FetchContent_GetProperties(libxls)
if(NOT libxls_POPULATED)
    FetchContent_Populate(libxls)
endif()

add_library(libxlsreader STATIC
    ${libxls_SOURCE_DIR}/src/xlstool.c
    ${libxls_SOURCE_DIR}/src/endian.c
    ${libxls_SOURCE_DIR}/src/locale.c
    ${libxls_SOURCE_DIR}/src/ole.c
    ${libxls_SOURCE_DIR}/src/xls.c
)

target_include_directories(libxlsreader
    PUBLIC
        ${libxls_SOURCE_DIR}/include
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/cmake/libxls
)

target_compile_definitions(libxlsreader PRIVATE
    _CRT_SECURE_NO_WARNINGS
    LIBXLS_STATIC
)

if(MSVC)
    target_compile_options(libxlsreader PRIVATE /W3)
endif()


set(APP_SOURCES
    src/main.cpp
    src/MainWindow.h
    src/MainWindow.cpp
    src/ExcelEngine.h
    src/ExcelEngine.cpp
    resources.qrc
)

if(WIN32)
    list(APPEND APP_SOURCES windows_version.rc)
endif()

add_executable(
    ExcelCompareProfessional
    ${APP_SOURCES}
)

target_compile_definitions(
    ExcelCompareProfessional
    PRIVATE
        AAT_COMPANY_NAME="AAT-Tech Ltd"
        AAT_PRODUCT_NAME="Excel Compare Professional"
)

if(MSVC)
    target_compile_options(ExcelCompareProfessional PRIVATE /utf-8)
endif()

target_link_libraries(
    ExcelCompareProfessional
    PRIVATE
        Qt6::Core
        Qt6::Gui
        Qt6::Widgets
        Qt6::Concurrent
        Qt6::PrintSupport
        QXlsx::QXlsx
        libxlsreader
)

if(WIN32)
    set_target_properties(
        ExcelCompareProfessional
        PROPERTIES
            WIN32_EXECUTABLE TRUE
            OUTPUT_NAME "ExcelCompareProfessional"
    )
endif()

install(
    TARGETS ExcelCompareProfessional
    RUNTIME DESTINATION .
)