cmake_minimum_required(VERSION 3.10)
project(email_sender C)

set(CMAKE_C_STANDARD 99)

# Find libcurl
find_package(CURL REQUIRED)

# Define the executable
add_executable(email-sender email-sender.c)

# Link with libcurl
target_link_libraries(email-sender CURL::libcurl)
