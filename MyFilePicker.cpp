#include <Windows.h>
#include <Shobjidl.h> // Contains the Common Item Dialog API


// Custom event handler class implementing IFileDialogEvents
class FileDialogEventHandler : public IFileDialogEvents
{
public:
  // IUnknown methods (unused)
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) { return E_NOTIMPL; }
  ULONG STDMETHODCALLTYPE AddRef() { return 1; }
  ULONG STDMETHODCALLTYPE Release() { return 1; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnFileOk(IFileDialog*) { return S_OK; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnFolderChange(IFileDialog*) { return S_OK; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnFolderChanging(IFileDialog*, IShellItem*) { return S_OK; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnSelectionChange(IFileDialog*)
  {
    // Get the selected items from the dialog
    IShellItemArray* pItemArray;
    ////m_pFileDialog->GetCurrentSelection(pItemArray);
    //HRESULT hr = m_pFileDialog->GetResults(&pItemArray);
    //IShellItem* pItemArray;
    IFileOpenDialog* pfod;
    HRESULT hr = m_pFileDialog->QueryInterface(IID_PPV_ARGS(&pfod));
    //m_pFileDialog->QueryInterface<IFileOpen>
    if (FAILED(hr)) {
      return hr;
    }
    hr = pfod->GetSelectedItems(&pItemArray);
    if (SUCCEEDED(hr))
    {
      // Get the count of selected items (files)
      DWORD count = 0;
      hr = pItemArray->GetCount(&count);
      if (SUCCEEDED(hr))
      {
        // Update the file count
        m_fileCount = count;
      }

      IFileDialogCustomize* pCustomize;
      hr = m_pFileDialog->QueryInterface(IID_IFileDialogCustomize, reinterpret_cast<void**>(&pCustomize));
      if (SUCCEEDED(hr)) {
        if (m_fileCount > 1) {
          pCustomize->SetControlState(10, CDCS_ENABLEDVISIBLE);
        }
        else {
          pCustomize->SetControlState(10, CDCS_INACTIVE|CDCS_VISIBLE);
        }
      }

      // Release the shell item array
      pItemArray->Release();
    }

    return S_OK;
  }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnShareViolation(IFileDialog*, IShellItem*, FDE_SHAREVIOLATION_RESPONSE*) { return S_OK; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnTypeChange(IFileDialog*) { return S_OK; }

  // IFileDialogEvents method
  HRESULT STDMETHODCALLTYPE OnOverwrite(IFileDialog*, IShellItem*, FDE_OVERWRITE_RESPONSE*) { return S_OK; }

  // Setter method to associate the IFileDialog object with the event handler
  void SetFileDialog(IFileOpenDialog* pFileDialog)
  {
    m_pFileDialog = pFileDialog;
    m_pFileDialog->Advise(this, &m_dwCookie);
  }

  // Getter method to retrieve the number of selected files
  DWORD GetFileCount() const { return m_fileCount; }

private:
  IFileDialog* m_pFileDialog;
  DWORD m_dwCookie;
  DWORD m_fileCount;
};




// Function to show the file picker dialog
HRESULT ShowFilePicker(HWND hwndOwner, LPWSTR* filePath)
{
  // Initialize the COM library
  HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  if (FAILED(hr))
  {
    return hr;
  }

  // Create the File Open Dialog object
  IFileOpenDialog* pFileOpenDialog;
  hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL, IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpenDialog));
  if (SUCCEEDED(hr))
  {// Set the dialog options to allow multiple file selection
    DWORD dwOptions;
    hr = pFileOpenDialog->GetOptions(&dwOptions);
    //set option to select multiple files
    hr = pFileOpenDialog->SetOptions(dwOptions | FOS_ALLOWMULTISELECT);

    // Get the IFileDialogCustomize interface
    IFileDialogCustomize* pCustomize;
    hr = pFileOpenDialog->QueryInterface(IID_IFileDialogCustomize, reinterpret_cast<void**>(&pCustomize));
    if (SUCCEEDED(hr))
    {
      // Add a custom button to the dialog
      hr = pCustomize->AddPushButton(10, L"My Custom Button");
      pCustomize->AddText(11, L"");
      //pCustomize->MakeProminent(10);
      if (SUCCEEDED(hr))
      {
        // Handle the custom button
        pCustomize->SetControlLabel(1, L"Custom Button");
        //pCustomize->SetCallback(this);
        // Release the IFileDialogCustomize interface
        pCustomize->Release();
      }

      // Create the event handler
      FileDialogEventHandler eventHandler;

      // Set the event handler for dynamic updates
      eventHandler.SetFileDialog(pFileOpenDialog);


      // Show the dialog
      hr = pFileOpenDialog->Show(hwndOwner);
      if (SUCCEEDED(hr))
      {
        // Get the file path from the dialog
        IShellItem* pItem;
        hr = pFileOpenDialog->GetResult(&pItem);
        if (SUCCEEDED(hr))
        {
          // Get the file path as a display name
          PWSTR pszFilePath;
          hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
          if (SUCCEEDED(hr))
          {
            // Allocate memory for the file path and copy it
            size_t len = wcslen(pszFilePath) + 1;
            *filePath = new WCHAR[len];
            wcscpy_s(*filePath, len, pszFilePath);

            // Release the string memory
            CoTaskMemFree(pszFilePath);
          }

          // Release the shell item
          pItem->Release();
        }
      }

      // Release the file open dialog
      pFileOpenDialog->Release();
    }
  }
  // Uninitialize the COM library
  CoUninitialize();

  return hr;
}


// Entry point of the application
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
  // Create a window
  HWND hwnd = CreateWindowEx(0, L"STATIC", L"File Picker Demo", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 800, 600, NULL, NULL, hInstance, NULL);

  // Show the file picker dialog
  LPWSTR filePath = nullptr;
  if (SUCCEEDED(ShowFilePicker(hwnd, &filePath)))
  {
    // Do something with the file path
    MessageBox(hwnd, filePath, L"Selected File", MB_OK);

    // Free the file path memory
    delete[] filePath;
  }

  // Destroy the window
  DestroyWindow(hwnd);

  return 0;
}
