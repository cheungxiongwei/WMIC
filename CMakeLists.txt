cmake_minimum_required(VERSION 3.28.0)
project(WMIC)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(SOURCE_FILES main.cpp wmic.cpp)
set(HEADER_FILES wmic.h)

add_executable(${PROJECT_NAME} ${SOURCE_FILES} ${HEADER_FILES})