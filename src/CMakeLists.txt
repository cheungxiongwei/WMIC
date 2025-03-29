set(WMIC_HEADERS wmic.h)
set(WMIC_SOURCES wmic.cpp)

# 添加静态库
add_library(wmic_lib STATIC ${WMIC_HEADERS} ${WMIC_SOURCES})
target_include_directories(wmic_lib PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}>
        $<INSTALL_INTERFACE:include>
)

# 添加可执行文件
if ("${CMAKE_CURRENT_LIST_DIR}" STREQUAL "${CMAKE_CURRENT_SOURCE_DIR}")
    add_executable(WMIC main.cpp)
    target_link_libraries(WMIC PRIVATE wmic_lib)
endif ()

# 安装配置
install(TARGETS wmic_lib
        EXPORT wmic_targets
        LIBRARY DESTINATION lib
        ARCHIVE DESTINATION lib
        RUNTIME DESTINATION bin
        INCLUDES DESTINATION include
)

install(FILES ${WMIC_HEADERS} DESTINATION include/wmic)

# 导出目标
install(EXPORT wmic_targets
        FILE wmic-targets.cmake
        NAMESPACE wmic::
        DESTINATION lib/cmake/wmic
)

# 创建并安装配置文件
include(CMakePackageConfigHelpers)
configure_package_config_file(
        ${CMAKE_CURRENT_SOURCE_DIR}/config.cmake.in
        ${CMAKE_CURRENT_BINARY_DIR}/wmic-config.cmake
        INSTALL_DESTINATION lib/cmake/wmic
)

install(FILES ${CMAKE_CURRENT_BINARY_DIR}/wmic-config.cmake
        DESTINATION lib/cmake/wmic
)