#include <Windows.h>
#include <tlhelp32.h>

#include <algorithm>
#include <fstream>
#include <iostream>
#include <map>
#include <string>
#include <utility>
#include <vector>
#include <memory>
#include <shlobj.h>

#undef max

struct RawString
  {
    RawString()
      {
      mp_buffer = std::make_unique<wchar_t[]>(1);
      mp_buffer.get()[0] = '\0';
      m_capacity = 0;
      }
    RawString(const RawString& i_other) = delete;
    RawString(RawString&& i_other) noexcept = delete;
    auto operator=(const RawString& i_other) -> RawString& = delete;
    auto operator=(RawString&& i_other) noexcept -> RawString& = delete;

    wchar_t* data(){ return mp_buffer.get(); }
    size_t capacity()
      {
      return m_capacity; 
      }
    void reserve(const size_t i_size)
      {
      if (m_capacity >= i_size)
        return;
      auto new_buffer = std::make_unique<wchar_t[]>(i_size + 1);
      wcscpy(new_buffer.get(), mp_buffer.get());
      m_capacity = i_size;
      mp_buffer = std::move(new_buffer);
      }

    void to_lower()
      {
      wchar_t* ptr = mp_buffer.get();
      for (; *ptr != '\0'; ++ptr)
        {
        *ptr = tolower(*ptr);
        }
      }

    RawString& operator=(const wchar_t* i_string)
      {
      size_t length = wcslen(i_string);
      reserve(length);
      wcscpy(mp_buffer.get(), i_string);
      return *this;
      }

    friend auto operator==(const RawString& i_lhs, const RawString& i_rhs) -> bool
      {
      return wcscmp(i_lhs.mp_buffer.get(), i_rhs.mp_buffer.get()) == 0;
      }

    friend auto operator==(const RawString& i_lhs, const wchar_t* i_rhs) -> bool
      {
      return wcscmp(i_lhs.mp_buffer.get(), i_rhs) == 0;
      }

    size_t m_capacity;
    std::unique_ptr<wchar_t[]> mp_buffer;
  };

RawString application_path;
RawString file_name;
RawString buffer, buffer2;
std::pair<double, HWND> best_window(std::numeric_limits<double>::lowest(), nullptr);

std::wstring RemoveSubstring(std::wstring i_string, const std::wstring & i_substring)
  {
  const auto pos = i_string.find(i_substring.c_str());
  if (pos != std::wstring::npos)
    i_string.erase(pos, i_substring.size());
  return i_string;
  }

//////////////////////////////////////////////////////////////////////////

IVirtualDesktopManager* g_virtual_desktop_manager;

bool IsMainWindow(const HWND i_handle)
  {
  return GetWindow(i_handle, GW_OWNER) == nullptr && IsWindowVisible(i_handle);
  }

void CopyNameToBuffer(const HANDLE i_h_process)
  {
  DWORD value = static_cast<unsigned long>(buffer.capacity());
  QueryFullProcessImageName(i_h_process, 0, buffer.data(), &value);
  if (value == buffer.capacity())
    {
    buffer.reserve(value * 2 + 1);
    CopyNameToBuffer(i_h_process);
    }
  }

bool CompareApplicationPath(const HWND i_hwnd)
  {
  if (IsMainWindow(i_hwnd) == false)
    return false;

  DWORD process_id;
  GetWindowThreadProcessId(i_hwnd, &process_id);

  static std::map<DWORD, HWND> process_to_main_window;

  if (process_to_main_window.find(process_id) != process_to_main_window.end())
    return false;

  process_to_main_window[process_id] = i_hwnd;

  const HANDLE h_process = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, process_id);
  if (h_process == nullptr)
    return false;

  CopyNameToBuffer(h_process);
  const size_t length = GetLongPathName(buffer.data(), buffer2.data(), static_cast<DWORD>(buffer2.capacity()));
  if (length > buffer2.capacity())
    {
    buffer2.reserve(length);  
    GetLongPathName(buffer.data(), buffer2.data(), static_cast<DWORD>(buffer2.capacity()));
    }

  CloseHandle(h_process);
  buffer2.to_lower();

  return buffer2 == application_path;
  }

//////////////////////////////////////////////////////////////////////////

double GetWindowScore(const HWND i_hwnd)
  {
  int index = 0;
  for (HWND hwnd = GetTopWindow(nullptr); hwnd != i_hwnd; hwnd = GetNextWindow(hwnd, GW_HWNDNEXT))
    --index;

  return index;
  }

BOOL CALLBACK EnumWindowsProc(HWND i_hwnd, LPARAM /*lParam*/)
  {
  BOOL is_on_current;
  if (g_virtual_desktop_manager && SUCCEEDED(g_virtual_desktop_manager->IsWindowOnCurrentVirtualDesktop(i_hwnd, &is_on_current)) && is_on_current == false)
    return true;

  if (IsWindowVisible(i_hwnd) == false)
    return true;

  if (GetWindowTextLengthW(i_hwnd) == 0)
    return true;

  if (CompareApplicationPath(i_hwnd) == true)
    best_window = std::max(best_window, std::make_pair(GetWindowScore(i_hwnd), i_hwnd));

  return true;
  }

//////////////////////////////////////////////////////////////////////////

void Show(const HWND i_hwnd)
  {
  WINDOWPLACEMENT place;
  memset(&place, 0, sizeof(WINDOWPLACEMENT));
  place.length = sizeof(WINDOWPLACEMENT);
  GetWindowPlacement(i_hwnd, &place);
  switch (place.showCmd)
    {
    case SW_SHOWMINIMIZED:
      ShowWindow(i_hwnd, SW_RESTORE);
      break;
    default:
      break;
    }
  SetForegroundWindow(i_hwnd);
  }

DWORD GetParentProcessID(DWORD dwProcessID)
{
	DWORD dwParentProcessID = -1 ;
	HANDLE			hProcessSnapshot ;
	PROCESSENTRY32	processEntry32 ;
	
	hProcessSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) ;
	if(hProcessSnapshot != INVALID_HANDLE_VALUE)
	{
		processEntry32.dwSize = sizeof(PROCESSENTRY32) ;
		if(Process32First(hProcessSnapshot, &processEntry32))
		{
			do
			{
				if (dwProcessID == processEntry32.th32ProcessID)
				{
					dwParentProcessID = processEntry32.th32ParentProcessID ;
					break ;
				}
			}
			while(Process32Next(hProcessSnapshot, &processEntry32)) ;
			
			CloseHandle(hProcessSnapshot) ;
		}
	}

	return dwParentProcessID ;
}

int __stdcall wWinMain(
  HINSTANCE /*hInstance*/,
  HINSTANCE /*hPrevInstance*/,
  const LPWSTR lpCmdLine,
  int /*nCmdShow*/)
  {
  buffer.reserve(20);
  buffer2.reserve(20);
  int i_argc;
  LPWSTR * i_args;
  i_args = CommandLineToArgvW(lpCmdLine, &i_argc);

  if (i_argc == 0)
    {
    AttachConsole(GetParentProcessID(GetCurrentProcessId()));
    if (freopen("CONOUT$", "w", stdout) == nullptr)
      {
      AllocConsole();
      freopen("CONOUT$", "w", stdout);
      }
    std::cout << "  First argument - path to application\n";
    std::cout << "  Second argument - file to be drag&dropped into application\n";
    std::cout << "  Third argument (optional)\n";
    std::cout << "  VS - special mode for drag&drop into Visual Studio";
    FreeConsole();
    return 0;
    }

#ifdef _WIN64
  const UINT message_id = RegisterWindowMessage(L"MyDragDropMessage64");
#else
  const UINT message_id = RegisterWindowMessage(L"MyDragDropMessage32");
#endif


  application_path = i_args[0];
  application_path.to_lower();
  file_name = i_args[1];
  file_name.to_lower();
  const std::wstring file_name_normal = i_args[1];

  RawString mode;
  if (i_argc == 3)
    {
    mode = i_args[2];
    mode.to_lower();
    }

  LocalFree(i_args);

  // g_virtual_desktop_manager
  if (SUCCEEDED(CoInitialize(nullptr)))
    {
    CoCreateInstance(CLSID_VirtualDesktopManager, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&g_virtual_desktop_manager));
    }

  EnumWindows(&EnumWindowsProc, 0);

  if (mode == L"vs")
    { // for Visual Studio
    if (true)
      {
      std::wofstream file(std::wstring(application_path.data()) + L".buffer");
      file << file_name_normal;
      }
    const HWND hwnd = best_window.second;
    Show(hwnd);
    Sleep(5);
    SendMessage(hwnd, message_id, 0, 0);
    }
  else if (best_window.second != nullptr)
    {
    if (true)
      {
      std::wofstream file(std::wstring(application_path.data()) + L".buffer");      
      file << file_name_normal;
      }

#ifdef _WIN64
    const HMODULE dll = LoadLibrary(L"DragDropLib64.dll");
    const HOOKPROC address = (HOOKPROC)GetProcAddress(dll, "DragDrop");
#else
    const HMODULE dll = LoadLibrary(L"DragDropLib32.dll");
    const HOOKPROC address = (HOOKPROC)GetProcAddress(dll, "_DragDrop@12");
#endif

    const auto hook = SetWindowsHookEx(WH_CALLWNDPROC, address, dll, 0);
    const HWND hwnd = best_window.second;
    Show(hwnd);
    Sleep(5);
    SendMessage(hwnd, message_id, 0, 0);

    UnhookWindowsHookEx(hook);
    }
  else 
    {
    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    const std::wstring new_args = std::wstring() + L"\"" + std::wstring(application_path.data()) + L"\" \"" + file_name_normal + L"\"";

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    CreateProcess(
      const_cast<wchar_t*>(application_path.data()),
      const_cast<wchar_t*>(new_args.c_str()), nullptr, nullptr, FALSE,
      0,
      nullptr, nullptr,
      &si,
      &pi);    
    }

  if (g_virtual_desktop_manager)
    g_virtual_desktop_manager->Release();

  return 0;
  }

