cmake_minimum_required(VERSION 3.29)
project(InputMethodSwitching)

set(CMAKE_CXX_STANDARD 20)

add_executable(InputMethodSwitching main.cpp
        constants.h)
#include_directories("C:\\Program Files (x86)\\Windows Kits\\10\\Include\\10.0.22621.0\\um")