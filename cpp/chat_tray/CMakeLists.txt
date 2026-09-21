cmake_minimum_required(VERSION 3.20)
project(otd421_chat_tray LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

add_executable(otd421_chat_tray WIN32 main.cpp)
target_compile_definitions(otd421_chat_tray PRIVATE UNICODE _UNICODE WIN32_LEAN_AND_MEAN NOMINMAX)
target_compile_options(otd421_chat_tray PRIVATE "$<$<CXX_COMPILER_ID:MSVC>:/utf-8>")
target_link_libraries(otd421_chat_tray PRIVATE winhttp comctl32 shell32 user32 gdi32 winsqlite3)
