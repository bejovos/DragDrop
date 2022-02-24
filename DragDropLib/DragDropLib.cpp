#include "DragDropLib.h"

#include <fstream>
#include <unknwnbase.h>
#include <string>
#include <shlobj.h>
#include <Strsafe.h>

UINT message_id;

BOOL APIENTRY DllMain(HMODULE hModule,
                      DWORD ul_reason_for_call,
                      LPVOID lpReserved
  )
{
  switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
#ifdef _WIN64
      message_id = RegisterWindowMessage(L"MyDragDropMessage64");
#else
      message_id = RegisterWindowMessage(L"MyDragDropMessage32");
#endif
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
      break;
  }
  return TRUE;
}

HGLOBAL DupGlobalMem(HGLOBAL hMem)
{
  SIZE_T len = GlobalSize(hMem);
  PVOID source = GlobalLock(hMem);
  PVOID dest = GlobalAlloc(GMEM_FIXED, len);
  memcpy_s(dest, len, source, len);
  GlobalUnlock(hMem);
  return dest;
}

class CDataObject : public IDataObject
{
public:
  LONG m_lRefCount;
  FORMATETC* fmtetc;
  STGMEDIUM* stgmed;

  CDataObject(FORMATETC* fmtetc, STGMEDIUM* stgmed)
  {
    m_lRefCount = 1;
    this->fmtetc = fmtetc;
    this->stgmed = stgmed;
  }

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void __RPC_FAR *__RPC_FAR * ppvObject) override
  {
    return 0;
  }

  ULONG STDMETHODCALLTYPE AddRef() override
  {
    return ++m_lRefCount;
  }

  ULONG STDMETHODCALLTYPE Release() override
  {
    return --m_lRefCount;
  }

  HRESULT STDMETHODCALLTYPE GetData(FORMATETC* pformatetcIn, STGMEDIUM* pMedium) override
  {
    if (pformatetcIn->tymed != 1)
      return DV_E_FORMATETC;
    if (pformatetcIn->cfFormat != CF_HDROP)
      return DV_E_FORMATETC;

    pMedium->tymed = this->stgmed->tymed;
    pMedium->pUnkForRelease = 0;

    if (pMedium->tymed == TYMED_HGLOBAL)
      pMedium->hGlobal = DupGlobalMem(this->stgmed->hGlobal);

    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* pformatetc, STGMEDIUM* pmedium) override
  {
    return 0;
  }

  HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* pformatetcIn) override
  {
    if (pformatetcIn->tymed != 1)
      return DV_E_FORMATETC;
    if (pformatetcIn->cfFormat != CF_HDROP)
      return DV_E_FORMATETC;

    return 0;
  }

  HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* pformatectIn, FORMATETC* pformatetcOut) override
  {
    return 0;
  }

  HRESULT STDMETHODCALLTYPE SetData(FORMATETC* pformatetc, STGMEDIUM* pmedium, BOOL fRelease) override
  {
    return E_NOTIMPL;
  }

  HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC** ppenumFormatEtc) override
  {
    return 0;
  }

  HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC* pformatetc, DWORD advf, IAdviseSink* pAdvSink, DWORD* pdwConnection) override
  {
    return 0;
  }

  HRESULT STDMETHODCALLTYPE DUnadvise(DWORD dwConnection) override
  {
    return 0;
  }

  HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** ppenumAdvise) override
  {
    return 0;
  }
};

extern "C" {
LRESULT CALLBACK DragDrop(int code, WPARAM wParam, LPARAM lParam)
{
  if (code < 0)
    return CallNextHookEx(0, code, wParam, lParam);

  if (((CWPSTRUCT*)lParam)->message == message_id) {
    RECT rect;
    GetWindowRect(((CWPSTRUCT*)lParam)->hwnd, &rect);
    POINTL point;
    point.x = (rect.left + rect.right) / 2;// +400;
    point.y = (rect.top + rect.bottom) / 2;// +400;

    HGLOBAL hfgd = GlobalAlloc(GMEM_ZEROINIT | GMEM_MOVEABLE | GHND | GMEM_SHARE,
                               (DWORD)(sizeof(DROPFILES) + 2 * (MAX_PATH) * sizeof(wchar_t)));
    LPDROPFILES pDropFiles = (LPDROPFILES)GlobalLock(hfgd);
    pDropFiles->pFiles = sizeof(DROPFILES);
    pDropFiles->fWide = true;

    wchar_t buff[MAX_PATH];
    GetModuleFileName(NULL, buff, MAX_PATH);
    std::wstring buffer_filename = std::wstring(buff) + L".buffer";

    std::wifstream file(buffer_filename);
    size_t offset = 0;
    while (!file.eof()) {
      file.getline(buff, MAX_PATH);
      std::wstring file_name = buff;
      lstrcpy((LPWSTR)(&pDropFiles[1]) + offset, file_name.c_str());
      offset += file_name.size() + 1;
    }
    file.close();

    pDropFiles->pt.x = point.x;
    pDropFiles->pt.y = point.y;
    GlobalUnlock(hfgd);

    IUnknown* ptr = (IUnknown*)GetProp(((CWPSTRUCT*)lParam)->hwnd, L"OleDropTargetInterface");

    IDropTarget* drop_target = nullptr;
    if (ptr)
      ptr->QueryInterface(IID_IDropTarget, (void**)&drop_target);

    if (drop_target == nullptr) // use old approach
    {
      std::wifstream file(buffer_filename);
      file.getline(buff, MAX_PATH);
      std::wstring file_name = buff;
      file.close();

      SendMessage(((CWPSTRUCT*)lParam)->hwnd, WM_DROPFILES, (WPARAM)hfgd, 0);
    }
    else { // use new approach
      FORMATETC fmtetc = { CF_HDROP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
      STGMEDIUM stgmed = { TYMED_HGLOBAL, { 0 }, 0 };
      stgmed.hGlobal = hfgd;
      IDataObject* pDataObject = new CDataObject(&fmtetc, &stgmed);
      DWORD ret = DROPEFFECT_COPY;
      auto result1 = drop_target->DragEnter(pDataObject, 0, point, &ret);
      auto result2 = drop_target->Drop(pDataObject, 0, point, &ret);
      delete pDataObject;
    }

    GlobalFree(hfgd);
  }

  return CallNextHookEx(0, code, wParam, lParam);
}
}
