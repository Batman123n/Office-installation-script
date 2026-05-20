cmake_minimum_required(VERSION 3.16)
project(Dario_Installer LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Enable LTO
set(CMAKE_INTERPROCEDURAL_OPTIMIZATION ON)

# Point CMake to the static Qt installation
set(CMAKE_PREFIX_PATH "D:/QtStatic")

find_package(Qt6 COMPONENTS Widgets REQUIRED)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

add_executable(Dario_Installer WIN32 main.cpp mainwindow.cpp mainwindow.h resources.qrc)

target_link_libraries(Dario_Installer PRIVATE Qt6::Widgets)

# Ensure static linking flags are handled if necessary
if(WIN32)
    target_compile_definitions(Dario_Installer PRIVATE QT_STATIC)
endif()
