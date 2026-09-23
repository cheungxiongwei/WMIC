# AGENTS.md

Windows CLI tool. C++23 / MSVC / Windows SDK. CMake + CMakePresets. No third-party deps.

## Stack constraints

- Language: C++23, compiler: MSVC only (`cl.exe`). Do not use GCC/Clang-only extensions.
- Dependencies: none beyond the Windows SDK. Do not add vcpkg/Conan/FetchContent libraries without asking.
- Source files are UTF-8; add `/utf-8` to compiler flags so MSVC doesn't misinterpret literals as the ANSI codepage.

## Build / test commands

Presets are the source of truth; read `CMakePresets.json` before changing the build.

```
cmake --preset <configure-preset>
cmake --build --preset <build-preset>
ctest --preset <test-preset>
```

- Build artifacts go under `build/`; never commit them.
- Run from a Developer Command Prompt (or a shell with `vcvarsall.bat` sourced) so `cl.exe` and the Windows SDK are on `PATH`. A missing compiler is a shell problem, not a CMake problem.

## Conventions

- Do not add comments unless asked.
- Enable `/W4` and treat warnings as errors in the project's CMake flags; fix warnings rather than suppressing them.