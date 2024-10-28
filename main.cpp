#include <windows.h>
#include <msctf.h>
#include <iostream>
#include "constants.h"

#pragma comment(lib, "msctf.lib")


struct LanguageProfile {
    CLSID clsid;
    LANGID langid;
    std::wstring description;
};

enum class Flag {
    NOW = -1,
    EN = 0,
    SG = 1,
    WR = 2,
    HI = 3,
};


/*
*
微软拼音
******************************************
CLSID: 81D4E9C9-1D3B-41BC-9E-6C-4B-40-BF-79-E3-5E
LangID: 2052
Profile GUID: FA550B04-5AD7-411F-A5-AC-CA-3-8E-C5-15-D7
******************************************
搜狗拼音
******************************************
CLSID: E7EA138E-69F8-11D7-A6-EA-0-6-5B-84-43-10
LangID: 2052
Profile GUID: E7EA138F-69F8-11D7-A6-EA-0-6-5B-84-43-11
******************************************
英语键盘
******************************************
CLSID: 0-0-0-0-0-0-0-0-0-0-0
LangID: 1033
Profile GUID: 0-0-0-0-0-0-0-0-0-0-0
******************************************
Hi 英文输入法
*******************************************
CLSID: E7EA138F-69F8-11D7-EE-EE-0-6-5B-84-43-10
LangID: 2052
Profile GUID: E7EA138F-69F8-11D7-EE-EE-0-6-5B-84-43-11
******************************************
 */
// 定义搜狗输入法的CLSID
constexpr CLSID SG_CLSID = {0xE7EA138E, 0x69F8, 0x11D7, {0xA6, 0xEA, 0x00, 0x06, 0x5B, 0x84, 0x43, 0x10}};
// 定义英语键盘的CLSID
constexpr CLSID EN_CLSID = {0x00000000, 0x0000, 0x0000, {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}};
// 定义微软输入法的CLSID
constexpr CLSID WR_CLSID = {0x81D4E9C9, 0x1D3B, 0x41BC, {0x9E, 0x6C, 0x4B, 0x40, 0xBF, 0x79, 0xE3, 0x5E}};
// 定义Hi英文输入法的CLSID
constexpr CLSID HI_CLSID = {0xE7EA138F, 0x69F8, 0x11D7, {0xEE, 0xEE, 0x00, 0x06, 0x5B, 0x84, 0x43, 0x10}};

// 定义搜狗拼音的语言配置文件
constexpr LanguageProfile SG_LG_PROFILE = {
    .clsid = SG_CLSID,
    .langid = 2052,
    .description = L"搜狗拼音"
};

// 定义英语键盘的语言配置文件
constexpr LanguageProfile EN_LG_PROFILE = {
    .clsid = EN_CLSID,
    .langid = 1033,
    .description = L"英语键盘"
};

// 定义微软拼音的语言配置文件
constexpr LanguageProfile WR_LG_PROFILE = {
    .clsid = WR_CLSID,
    .langid = 2052,
    .description = L"微软拼音"
};

// 定义Hi英文输入法的语言配置文件
constexpr LanguageProfile HI_LG_PROFILE = {
    .clsid = HI_CLSID,
    .langid = 2052,
    .description = L"Hi英文输入法"
};

// 根据标志获取对应的语言配置文件
const LanguageProfile *getLanguageProfile(const Flag flag) {
    switch (flag) {
        case Flag::EN: return &EN_LG_PROFILE;
        case Flag::SG: return &SG_LG_PROFILE;
        case Flag::WR: return &WR_LG_PROFILE;
        case Flag::HI: return &HI_LG_PROFILE;
        default: return &EN_LG_PROFILE;
    }
}

// 解析命令行参数，返回输入法标志
Flag parseFlag(const int argc, char *argv[]) {
    auto flag = Flag::NOW; // 默认值
    if (argc > 1) {
        try {
            if (int flagValue = std::stoi(argv[1]); flagValue >= -1 && flagValue <= 3) {
                flag = static_cast<Flag>(flagValue);
            } else {
                std::cerr << "Invalid flag value: " << flagValue << ", defaulting to EN" << std::endl;
                flag = Flag::EN;
            }
        } catch ([[maybe_unused]] const std::exception &e) {
            std::cerr << "Invalid flag value: " << argv[1] << ", defaulting to EN" << std::endl;
            flag = Flag::EN;
        }
    }
    return flag;
}

// 初始化COM库
bool initializeCOM() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        std::cerr << "CoInitializeEx failed: " << hr << std::endl;
        return false;
    }
    return true;
}

// 创建输入处理器配置文件管理器实例
bool createProfileManager(ITfInputProcessorProfileMgr **ppProfileMgr) {
    const HRESULT hr = CoCreateInstance(CLSID_TF_InputProcessorProfiles, nullptr, CLSCTX_INPROC_SERVER,
                                        IID_ITfInputProcessorProfileMgr, reinterpret_cast<void **>(ppProfileMgr));
    if (FAILED(hr)) {
        std::cerr << "CoCreateInstance failed: " << hr << std::endl;
        return false;
    }
    return true;
}

// 枚举输入处理器配置文件
bool enumerateProfiles(ITfInputProcessorProfileMgr *pProfileMgr, IEnumTfInputProcessorProfiles **ppEnumProfiles) {
    if (const HRESULT hr = pProfileMgr->EnumProfiles(0, ppEnumProfiles); FAILED(hr)) {
        std::cerr << "EnumProfiles failed: " << hr << std::endl;
        return false;
    }
    return true;
}

// 输出当前激活文件信息
bool printCurrentProfile(ITfInputProcessorProfileMgr *pProfileMgr) {
    TF_INPUTPROCESSORPROFILE profileTmp;
    pProfileMgr->GetActiveProfile(GUID_TFCAT_TIP_KEYBOARD, &profileTmp);
    std::wcout << L"******************************************" << std::endl;
    std::wcout << L"CLSID: " << std::hex << std::uppercase;
    std::wcout << profileTmp.clsid.Data1 << L"-";
    std::wcout << profileTmp.clsid.Data2 << L"-";
    std::wcout << profileTmp.clsid.Data3 << L"-";
    for (int i = 0; i < 8; ++i) {
        std::wcout << std::hex << std::uppercase << (int) profileTmp.clsid.Data4[i];
        if (i < 7) std::wcout << L"-";
    }
    std::wcout << std::dec << std::nouppercase << std::endl; // 切换回十进制输出并关闭大写

    std::wcout << L"LangID: " << profileTmp.langid << std::endl;


    std::wcout << L"Profile GUID: " << std::hex << std::uppercase;
    std::wcout << profileTmp.guidProfile.Data1 << L"-";
    std::wcout << profileTmp.guidProfile.Data2 << L"-";
    std::wcout << profileTmp.guidProfile.Data3 << L"-";
    for (int i = 0; i < 8; ++i) {
        std::wcout << std::hex << std::uppercase << (int) profileTmp.guidProfile.Data4[i];
        if (i < 7) std::wcout << L"-";
    }
    std::wcout << std::dec << std::nouppercase << std::endl; // 切换回十进制输出并关闭大写
    std::wcout << L"******************************************" << std::endl;

    return true;
}

// 激活指定的输入处理器配置文件
bool activateProfile(ITfInputProcessorProfileMgr *pProfileMgr, const LanguageProfile &selectLanguageProfile,
                     IEnumTfInputProcessorProfiles *pEnumProfiles) {
    TF_INPUTPROCESSORPROFILE profile;
    ULONG fetched = 0;
    while (pEnumProfiles->Next(1, &profile, &fetched) == S_OK) {
        if (profile.clsid == selectLanguageProfile.clsid && profile.langid == selectLanguageProfile.langid) {
            const HRESULT hr = pProfileMgr->ActivateProfile(
                TF_PROFILETYPE_INPUTPROCESSOR,
                profile.langid,
                profile.clsid,
                profile.guidProfile,
                nullptr,
                TF_IPPMF_FORSESSION | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE | TF_IPPMF_FORPROCESS |
                TF_IPPMF_ENABLEPROFILE);

            if (FAILED(hr)) {
                if (E_FAIL == hr) {
                    std::cout << "ActivateProfile E_FAIL" << std::endl;
                } else {
                    std::cout << "ActivateProfile Unknown" << std::endl;
                }
                std::cerr << "ActivateProfile failed: " << hr << std::endl;
                return false;
            } else {
                std::cout << "ActivateProfile success" << std::endl;
            }
            return true;
        }
    }
    return true;
}

// 释放资源
void releaseResources(IEnumTfInputProcessorProfiles *pEnumProfiles, ITfInputProcessorProfileMgr *pProfileMgr) {
    if (pEnumProfiles) {
        pEnumProfiles->Release();
    }
    if (pProfileMgr) {
        pProfileMgr->Release();
    }
    CoUninitialize();
}

// 主函数
int main(const int argc, char *argv[]) {
    const Flag flag = parseFlag(argc, argv);


    if (!initializeCOM()) {
        return 1;
    }

    ITfInputProcessorProfileMgr *pProfileMgr = nullptr;
    if (!createProfileManager(&pProfileMgr)) {
        CoUninitialize();
        return 1;
    }

    if (flag == Flag::NOW) {
        printCurrentProfile(pProfileMgr);
        return 0;
    }

    IEnumTfInputProcessorProfiles *pEnumProfiles = nullptr;
    if (!enumerateProfiles(pProfileMgr, &pEnumProfiles)) {
        pProfileMgr->Release();
        CoUninitialize();
        return 1;
    }

    const LanguageProfile selectLanguageProfile = *getLanguageProfile(flag);

    if (!activateProfile(pProfileMgr, selectLanguageProfile, pEnumProfiles)) {
        pEnumProfiles->Release();
        pProfileMgr->Release();
        CoUninitialize();
        return 1;
    }

    releaseResources(pEnumProfiles, pProfileMgr);

    return 0;
}
