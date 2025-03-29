cmake_minimum_required(VERSION 3.30)
project(WMIC
        VERSION 1.0.0.0
        DESCRIPTION "WMIC retrieves hardware info, generates fingerprints, and manages storage."
        HOMEPAGE_URL "https://github.com/cheungxiongwei/WMIC"
        LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 23)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_subdirectory(src)