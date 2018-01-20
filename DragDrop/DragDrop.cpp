#include <windows.h>
#include <shlobj.h>
#include <shellapi.h>
#include <utility>
#include <algorithm>
#include <string>
#include <fstream>
#include <map>
#include <vector>

#undef max

std::wstring application_path;
std::wstring file_name;


std::wstring ToLower(std::wstring i_string)
  {
  for (wchar_t & c : i_string)
    c = tolower(c);
  return i_string;
  }

std::pair<double, HWND> best_window(std::numeric_limits<double>::lowest(), nullptr);
std::vector<wchar_t> buffer, buffer2;

std::wstring RemoveSubstring(std::wstring i_string, const std::wstring & i_substring)
  {
  auto pos = i_string.find(i_substring.c_str());
  if (pos != std::wstring::npos)
    i_string.erase(pos, i_substring.size());
  return i_string;
  }

//////////////////////////////////////////////////////////////////////////

bool IsMainWindow(HWND handle)
  {
  return GetWindow(handle, GW_OWNER) == (HWND)nullptr && IsWindowVisible(handle);
  }

void CopyNameToBuffer(HANDLE hProcess)
  {
  DWORD value = static_cast<unsigned long>(buffer.size());
  QueryFullProcessImageName(hProcess, 0, buffer.data(), &value);
  if (value >= buffer.size() - 1)
    {
    buffer.resize(buffer.size() * 2);
    CopyNameToBuffer(hProcess);
    }
  }

bool CompareApplicationPath(HWND hwnd, const std::wstring & i_application_path)
  {
  if (IsMainWindow(hwnd) == false)
    return false;

  DWORD process_id;
  GetWindowThreadProcessId(hwnd, &process_id);

  static std::map<DWORD, HWND> process_to_main_window;

  if (process_to_main_window.find(process_id) != process_to_main_window.end())
    return false;

  process_to_main_window[process_id] = hwnd;

  HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, process_id);
  if (hProcess == nullptr)
    return false;

  CopyNameToBuffer(hProcess);
  size_t length = GetLongPathName(buffer.data(), buffer2.data(), static_cast<DWORD>(buffer2.size()));
  if (length > buffer2.size())
    {
    buffer2.resize(length);
    GetLongPathName(buffer.data(), buffer2.data(), static_cast<DWORD>(buffer2.size()));
    }

  CloseHandle(hProcess);

  return ToLower(buffer2.data()) == i_application_path;
  }

//////////////////////////////////////////////////////////////////////////

BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM /*lParam*/)
  {  
  if (CompareApplicationPath(hwnd, application_path) == true)
    best_window = std::make_pair(0., hwnd);
  return true;
  }

//////////////////////////////////////////////////////////////////////////

double GetScore_VisualStudio(HWND i_hwnd, const std::wstring & i_file_name)
  {
  int index = 0;
  for (HWND hwnd = GetTopWindow(nullptr); hwnd != i_hwnd; hwnd = GetNextWindow(hwnd, GW_HWNDNEXT))
    --index;

  return index;
  }

BOOL CALLBACK EnumWindowsProc_VisualStudio(HWND hwnd, LPARAM /*lParam*/)
  {
  if (CompareApplicationPath(hwnd, application_path) == true)
    best_window = std::max(best_window, std::make_pair(GetScore_VisualStudio(hwnd, file_name), hwnd));

  return true;
  }

//////////////////////////////////////////////////////////////////////////

void Show(HWND hwnd)
  {
  WINDOWPLACEMENT place;
  memset(&place, 0, sizeof(WINDOWPLACEMENT));
  place.length = sizeof(WINDOWPLACEMENT);
  GetWindowPlacement(hwnd, &place);
  switch (place.showCmd)
    {
    case SW_SHOWMAXIMIZED:
      ShowWindow(hwnd, SW_SHOWMAXIMIZED);
      break;
    case SW_SHOWMINIMIZED:
      ShowWindow(hwnd, SW_RESTORE);
      break;
    default:
      ShowWindow(hwnd, SW_NORMAL);
      break;
    }
  SetForegroundWindow(hwnd);
  }

int __stdcall wWinMain(HINSTANCE /*hInstance*/,
             HINSTANCE /*hPrevInstance*/,
             LPWSTR    lpCmdLine,
             int       /*nCmdShow*/)
  {
  buffer.resize(20);
  buffer2.resize(20);
  int argc;
  LPWSTR * args;
  args = CommandLineToArgvW(lpCmdLine, &argc);

  if (argc < 2)
    return 0;

#ifdef _WIN64
  UINT message_id = RegisterWindowMessage(L"MyDragDropMessage64");
#else
  UINT message_id = RegisterWindowMessage(L"MyDragDropMessage32");
#endif

  application_path = ToLower(args[0]);
  file_name = ToLower(args[1]);
  std::wstring mode = (argc == 3 ? ToLower(args[2]) : L"");
  std::wstring file_name_normal = args[1];
  LocalFree(args);

  if (mode == L"vs")
    { // for Visual Studio
    EnumWindows(&EnumWindowsProc_VisualStudio, 0);
    if (true)
      {
      std::wofstream file(std::wstring(application_path) + L".buffer");
      file << file_name_normal;
      }
    HWND hwnd = best_window.second;
    Show(hwnd);
    Sleep(5);
    SendMessage(hwnd, message_id, 0, 0);
    return 0;
    }

  EnumWindows(&EnumWindowsProc, 0);
  if (best_window.second != nullptr)
    {
    if (true)
      {
      std::wofstream file(std::wstring(application_path) + L".buffer");      
      file << file_name_normal;
      }

#ifdef _WIN64
    HMODULE dll = LoadLibrary(L"DragDropLib64.dll");
    HOOKPROC addr = (HOOKPROC)GetProcAddress(dll, "DragDrop");
#else
    HMODULE dll = LoadLibrary(L"DragDropLib32.dll");
    HOOKPROC addr = (HOOKPROC)GetProcAddress(dll, "_DragDrop@12");
#endif

    auto hook = SetWindowsHookEx(WH_CALLWNDPROC, addr, dll, 0);
    HWND hwnd = best_window.second;
    Show(hwnd);
    Sleep(5);
    SendMessage(hwnd, message_id, 0, 0);

    UnhookWindowsHookEx(hook);
    }
  else 
    {
    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    std::wstring args = 
      std::wstring() + L"\"" + application_path + L"\" \"" + file_name_normal + L"\"";

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    CreateProcess(
      const_cast<wchar_t*>(application_path.c_str()),
      const_cast<wchar_t*>(args.c_str()), nullptr, nullptr, FALSE,
      0,
      nullptr, nullptr,
      &si,
      &pi);    
    }

  return 0;
  }

