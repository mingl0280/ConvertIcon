// ConvertIcon.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <Windows.h>
#include <fstream>
#include <filesystem>
#include <vector>
#include <cstdint>
#include <olectl.h>
#include <cctype>
#include <algorithm>
#include "cxxopts.hpp"
#pragma comment(lib, "oleaut32.lib")


using namespace std;

void ConvertFile(std::wstring file_path);
void ConvertDir(std::wstring dir, bool recurse = false);
HRESULT SaveIcon(HICON hIcon, const wchar_t* path);
std::wstring get_wstring(const std::string& s);

struct ICOv1Desc
{
    uint32_t FIELD1;
    uint16_t width;
    uint16_t height;
    uint16_t width_bytes;
    uint16_t FIELD2;
};

int main(int argc, char* argv[])
{
    cxxopts::Options opts("MS ico 1.0 convert app", "MS ico 1.0 to 3.0 icon converter, works on device independent 1.0 images (include \"Both Format\".");
    opts.add_options()("f,file", "specified file", cxxopts::value<std::string>())
        ("d,directory", "specific directory", cxxopts::value<std::string>())
        ("r,recursive","recurse search all child directories")
        ("h,help", "Show Help");

    const auto result = opts.parse(argc, argv);

    if (result.count("file") != 0 && result.count("directory") != 0 || result.count("help"))
    {
        cout << opts.help() << endl;
        return 0;
    }

    if (result.count("file") == 0 && result.count("directory") == 0)
    {
        if (result.count("recursive"))
            ConvertDir(TEXT("."), result["recursive"].as<bool>());
        else
            ConvertDir(TEXT("."));
        return 0;
    }

    if (result.count("file") != 0)
    {
        ConvertFile(get_wstring(result["file"].as<std::string>()));
    }
    if (result.count("directory"))
    {
        if (result.count("recursive"))
            ConvertDir(get_wstring(result["directory"].as<std::string>()), result["recursive"].as<bool>());
        else
            ConvertDir(get_wstring(result["directory"].as<std::string>()));
    }
    return 0;    
}

std::wstring get_wstring(const std::string& s)
{
    //https://stackoverflow.com/a/6588525/8298647
    const char* cs = s.c_str();
    const size_t wn = mbsrtowcs(NULL, &cs, 0, NULL);

    if (wn == size_t(-1))
    {
        std::cout << "Error in mbsrtowcs(): " << errno << std::endl;
        return L"";
    }

    std::vector<wchar_t> buf(wn + 1);
    const size_t wn_again = mbsrtowcs(&buf[0], &cs, wn + 1, NULL);

    if (wn_again == size_t(-1))
    {
        std::cout << "Error in mbsrtowcs(): " << errno << std::endl;
        return L"";
    }

    return std::wstring(&buf[0], wn);
}


HRESULT SaveIcon(HICON hIcon, const wchar_t* path)
{
    // Thanks to https://stackoverflow.com/a/4338491/8298647
    // Create the IPicture interface
    PICTDESC desc = { sizeof(PICTDESC) };
    desc.picType = PICTYPE_ICON;
    desc.icon.hicon = hIcon;
    IPicture* pPicture = 0;
    HRESULT hr = OleCreatePictureIndirect(&desc, IID_IPicture, FALSE, (void**)&pPicture);
    if (FAILED(hr)) return hr;

    // Create a stream and save the image
    IStream* pStream = 0;
    CreateStreamOnHGlobal(0, TRUE, &pStream);
    LONG cbSize = 0;
    hr = pPicture->SaveAsFile(pStream, TRUE, &cbSize);

    // Write the stream content to the file
    if (!FAILED(hr)) {
        HGLOBAL hBuf = 0;
        GetHGlobalFromStream(pStream, &hBuf);
        void* buffer = GlobalLock(hBuf);
        HANDLE hFile = CreateFile(path, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, 0, 0);
        if (!hFile) hr = HRESULT_FROM_WIN32(GetLastError());
        else {
            DWORD written = 0;
            WriteFile(hFile, buffer, cbSize, &written, 0);
            CloseHandle(hFile);
        }
        GlobalUnlock(buffer);
    }
    // Cleanup
    pStream->Release();
    pPicture->Release();
    return hr;
}

void ConvertDir(std::wstring dir, bool recurse)
{
    namespace fs = std::filesystem;
    fs::directory_entry dir_entry(dir);
    if (!dir_entry.exists())
    {
        return;
    }

    vector<fs::path> process_items;

    if (recurse)
    {
        for (auto& item : fs::recursive_directory_iterator(dir))
        {
            auto extension = item.path().extension().string();
            std::transform(extension.begin(), extension.end(), extension.begin(), [](auto c) { return std::tolower(c); });

            if (extension == ".ico")
                process_items.emplace_back(fs::absolute(item.path()));
        }
    }
    else
    {
        for (auto& item : fs::directory_iterator(dir))
        {
            auto extension = item.path().extension().string();
            std::transform(extension.begin(), extension.end(), extension.begin(), [](auto c) { return std::tolower(c); });

            if (extension == ".ico")
                process_items.emplace_back(fs::absolute(item.path()));
        }
    }

    for(auto &item : process_items)
    {
        ConvertFile(item.wstring());
    }

}

void ConvertFile(std::wstring file_path)
{
    ifstream input_file(file_path, ios::in | ios::binary);
    if (!input_file)
    {
        wcout << TEXT("Error: ") << file_path << TEXT(" cannot be opened.") << endl;
        return;
    }
    uint16_t header;
    input_file.read((char*)&header, 2);
    if (header != 0x1 && header != 0x201)
    {
        wcout << TEXT("Error: ") << file_path << TEXT("is not a device-independent, v1.0 ICO image file!") << endl;
        return;
    }
    ICOv1Desc ico_desc {};
    input_file.read((char*)&ico_desc, sizeof(ICOv1Desc));
    auto bitmap_bytes = (uint32_t)ico_desc.width_bytes * (uint32_t)ico_desc.height;
    uint8_t* AndBytes = new uint8_t[bitmap_bytes]{};
    uint8_t* XorBytes = new uint8_t[bitmap_bytes]{};
    input_file.read((char*)AndBytes, bitmap_bytes);
    input_file.read((char*)XorBytes, bitmap_bytes);
    input_file.close();
    HICON icon_new {};
    HINSTANCE hinstance {};
    icon_new = CreateIcon(hinstance, ico_desc.width, ico_desc.height, 1, 1, AndBytes, XorBytes);

    auto result_path = filesystem::path(file_path);
    result_path.replace_extension(TEXT("New") + result_path.extension().wstring());
    SaveIcon(icon_new, result_path.wstring().c_str());
    delete[] AndBytes;
    delete[] XorBytes;
    AndBytes = nullptr;
    XorBytes = nullptr;
}


// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
