#pragma once

#ifdef DRAGDROPLIB_EXPORTS
#define DRAGDROPLIB_API __declspec(dllexport)
#else
#define DRAGDROPLIB_API __declspec(dllimport)
#endif

#include <windows.h>

extern "C"
  {
  DRAGDROPLIB_API LRESULT CALLBACK DragDrop(int code, WPARAM wParam, LPARAM lParam);
  }
