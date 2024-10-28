#include <windows.h>
#include <msctf.h>
#include <iostream>

#pragma comment(lib, "msctf.lib")


#define TF_IPPMF_FORPROCESS                     0x10000000
#define TF_IPPMF_FORSESSION                     0x20000000
#define TF_IPPMF_FORSYSTEMALL                   0x40000000
#define TF_IPPMF_ENABLEPROFILE                  0x00000001
#define TF_IPPMF_DISABLEPROFILE                 0x00000002
#define TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE   0x00000004


struct LanguageProfile {
    CLSID clsid;
    LANGID langid;
    std::wstring description;
};

enum class Flag {
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
// 搜狗输入法
constexpr CLSID SG_CLSID = {0xE7EA138E, 0x69F8, 0x11D7, {0xA6, 0xEA, 0x00, 0x06, 0x5B, 0x84, 0x43, 0x10}};
// 英语键盘
constexpr CLSID EN_CLSID = {0x00000000, 0x0000, 0x0000, {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}};
// 微软输入法
constexpr CLSID WR_CLSID = {0x81D4E9C9, 0x1D3B, 0x41BC, {0x9E, 0x6C, 0x4B, 0x40, 0xBF, 0x79, 0xE3, 0x5E}};
// Hi 英文输入法
constexpr CLSID HI_CLSID = {0xE7EA138F, 0x69F8, 0x11D7, {0xEE, 0xEE, 0x00, 0x06, 0x5B, 0x84, 0x43, 0x10}};

constexpr LanguageProfile SG_LG_PROFILE = {
    .clsid = SG_CLSID,
    .langid = 2052,
    .description = L"搜狗拼音"
};

constexpr LanguageProfile EN_LG_PROFILE = {
    .clsid = EN_CLSID,
    .langid = 1033,
    .description = L"英语键盘"
};

constexpr LanguageProfile WR_LG_PROFILE = {
    .clsid = WR_CLSID,
    .langid = 2052,
    .description = L"微软拼音"
};

constexpr LanguageProfile HI_LG_PROFILE = {
    .clsid = HI_CLSID,
    .langid = 2052,
    .description = L"Hi英文输入法"
};


const LanguageProfile *getLanguageProfile(const Flag flag) {
    switch (flag) {
        case Flag::EN: return &EN_LG_PROFILE;
        case Flag::SG: return &SG_LG_PROFILE;
        case Flag::WR: return &WR_LG_PROFILE;
        case Flag::HI: return &HI_LG_PROFILE;
        default: return &EN_LG_PROFILE;
    }
}


int main(int argc, char *argv[]) {
    auto flag = Flag::HI; // 默认值

    if (argc > 1) {
        try {
            int flagValue = std::stoi(argv[1]);
            flag = static_cast<Flag>(flagValue);
            if (flag != Flag::EN && flag != Flag::SG && flag != Flag::WR && flag != Flag::HI) {
                std::cerr << "Invalid flag value: " << flagValue << ", defaulting to EN" << std::endl;
                flag = Flag::EN;
            }
        } catch ([[maybe_unused]] const std::exception &e) {
            std::cerr << "Invalid flag value: " << argv[1] << ", defaulting to EN" << std::endl;
            flag = Flag::EN;
        }
    }

    const LanguageProfile selectLanguageProfile = *getLanguageProfile(flag);
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        std::cerr << "CoInitialize failed: " << hr << std::endl;
        return 1;
    }

    ITfInputProcessorProfileMgr *pProfileMgr = nullptr;
    hr = CoCreateInstance(CLSID_TF_InputProcessorProfiles, nullptr, CLSCTX_INPROC_SERVER,
                          IID_ITfInputProcessorProfileMgr, reinterpret_cast<void **>(&pProfileMgr));
    if (FAILED(hr)) {
        std::cerr << "CoCreateInstance failed: " << hr << std::endl;
        CoUninitialize();
        return 1;
    }

    IEnumTfInputProcessorProfiles *pEnumProfiles = nullptr;
    hr = pProfileMgr->EnumProfiles(0, &pEnumProfiles); // 传入 0 枚举所有配置文件
    if (FAILED(hr)) {
        std::cerr << "EnumProfiles failed: " << hr << std::endl;
        pProfileMgr->Release();
        CoUninitialize();
        return 1;
    }


    TF_INPUTPROCESSORPROFILE profile;
    ULONG fetched = 0;


    while (pEnumProfiles->Next(1, &profile, &fetched) == S_OK) {

        if (profile.clsid == selectLanguageProfile.clsid) {
            if (profile.langid == selectLanguageProfile.langid) {
                hr = pProfileMgr->ActivateProfile(
                    TF_PROFILETYPE_INPUTPROCESSOR,
                    profile.langid,
                    profile.clsid,
                    profile.guidProfile,
                    nullptr,
                    TF_IPPMF_FORSESSION | TF_IPPMF_DONTCARECURRENTINPUTLANGUAGE | TF_IPPMF_FORPROCESS | TF_IPPMF_ENABLEPROFILE);

                if (FAILED(hr)) {
                     if (E_FAIL == hr) {
                        std::cout << "ActivateProfile E_FAIL" << std::endl;
                    } else {
                        std::cout << "ActivateProfile Unknown" << std::endl;
                    }
                    std::cerr << "ActivateProfile failed: " << hr << std::endl;
                    return 1;
                } else {
                    std::cout << "ActivateProfile success" << std::endl;
                }
            }
        }

    }

    pEnumProfiles->Release();
    pProfileMgr->Release();
    CoUninitialize();

    return 0;
}
