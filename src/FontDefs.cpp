/**
 * @file FontDefs.cpp
 * @brief 폰트 프리셋 정의 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "FontDefs.h"

namespace cpyhwpx {

// 정적 멤버 초기화
bool FontDefs::s_bInitialized = false;
std::map<std::wstring, FontPreset> FontDefs::s_presets;

//=============================================================================
// 헬퍼 메서드
//=============================================================================

FontPreset FontDefs::CreateUniform(const std::wstring& font_name, int font_type)
{
    FontPreset preset;
    preset.FaceNameHangul = font_name;
    preset.FaceNameLatin = font_name;
    preset.FaceNameHanja = font_name;
    preset.FaceNameJapanese = font_name;
    preset.FaceNameOther = font_name;
    preset.FaceNameSymbol = font_name;
    preset.FaceNameUser = font_name;

    preset.FontTypeHangul = font_type;
    preset.FontTypeLatin = font_type;
    preset.FontTypeHanja = font_type;
    preset.FontTypeJapanese = font_type;
    preset.FontTypeOther = font_type;
    preset.FontTypeSymbol = font_type;
    preset.FontTypeUser = font_type;

    return preset;
}

FontPreset FontDefs::CreateDual(const std::wstring& hangul_name, const std::wstring& latin_name)
{
    FontPreset preset;
    preset.FaceNameHangul = hangul_name;
    preset.FaceNameLatin = latin_name;
    preset.FaceNameHanja = hangul_name;
    preset.FaceNameJapanese = hangul_name;
    preset.FaceNameOther = hangul_name;
    preset.FaceNameSymbol = latin_name;
    preset.FaceNameUser = hangul_name;

    preset.FontTypeHangul = FontType::TTF;
    preset.FontTypeLatin = FontType::TTF;
    preset.FontTypeHanja = FontType::TTF;
    preset.FontTypeJapanese = FontType::TTF;
    preset.FontTypeOther = FontType::TTF;
    preset.FontTypeSymbol = FontType::TTF;
    preset.FontTypeUser = FontType::TTF;

    return preset;
}

//=============================================================================
// 개별 폰트 프리셋
//=============================================================================

// 맑은 고딕 계열
FontPreset FontDefs::MalgunGothic() { return CreateUniform(L"맑은 고딕"); }
FontPreset FontDefs::MalgunGothicBold() { return CreateUniform(L"맑은 고딕"); }

// 굴림 계열
FontPreset FontDefs::Gulim() { return CreateUniform(L"굴림"); }
FontPreset FontDefs::GulimChe() { return CreateUniform(L"굴림체"); }
FontPreset FontDefs::Dotum() { return CreateUniform(L"돋움"); }
FontPreset FontDefs::DotumChe() { return CreateUniform(L"돋움체"); }
FontPreset FontDefs::Batang() { return CreateUniform(L"바탕"); }
FontPreset FontDefs::BatangChe() { return CreateUniform(L"바탕체"); }
FontPreset FontDefs::Gungsuh() { return CreateUniform(L"궁서"); }
FontPreset FontDefs::GungsuhChe() { return CreateUniform(L"궁서체"); }

// 나눔 계열
FontPreset FontDefs::NanumGothic() { return CreateUniform(L"나눔고딕"); }
FontPreset FontDefs::NanumGothicBold() { return CreateUniform(L"나눔고딕 Bold"); }
FontPreset FontDefs::NanumGothicLight() { return CreateUniform(L"나눔고딕 Light"); }
FontPreset FontDefs::NanumGothicExtraBold() { return CreateUniform(L"나눔고딕 ExtraBold"); }
FontPreset FontDefs::NanumMyeongjo() { return CreateUniform(L"나눔명조"); }
FontPreset FontDefs::NanumMyeongjoBold() { return CreateUniform(L"나눔명조 Bold"); }
FontPreset FontDefs::NanumMyeongjoExtraBold() { return CreateUniform(L"나눔명조 ExtraBold"); }
FontPreset FontDefs::NanumPenScript() { return CreateUniform(L"나눔손글씨 펜"); }
FontPreset FontDefs::NanumBrushScript() { return CreateUniform(L"나눔손글씨 붓"); }
FontPreset FontDefs::NanumBarunGothic() { return CreateUniform(L"나눔바른고딕"); }
FontPreset FontDefs::NanumBarunGothicBold() { return CreateUniform(L"나눔바른고딕 Bold"); }
FontPreset FontDefs::NanumBarunGothicLight() { return CreateUniform(L"나눔바른고딕 Light"); }
FontPreset FontDefs::NanumBarunGothicUltraLight() { return CreateUniform(L"나눔바른고딕 UltraLight"); }
FontPreset FontDefs::NanumBarunPen() { return CreateUniform(L"나눔바른펜"); }
FontPreset FontDefs::NanumSquare() { return CreateUniform(L"나눔스퀘어"); }
FontPreset FontDefs::NanumSquareBold() { return CreateUniform(L"나눔스퀘어 Bold"); }
FontPreset FontDefs::NanumSquareExtraBold() { return CreateUniform(L"나눔스퀘어 ExtraBold"); }
FontPreset FontDefs::NanumSquareLight() { return CreateUniform(L"나눔스퀘어 Light"); }
FontPreset FontDefs::NanumSquareRound() { return CreateUniform(L"나눔스퀘어라운드"); }
FontPreset FontDefs::NanumSquareRoundBold() { return CreateUniform(L"나눔스퀘어라운드 Bold"); }
FontPreset FontDefs::NanumSquareRoundExtraBold() { return CreateUniform(L"나눔스퀘어라운드 ExtraBold"); }
FontPreset FontDefs::NanumSquareRoundLight() { return CreateUniform(L"나눔스퀘어라운드 Light"); }

// 한컴 계열
FontPreset FontDefs::HCRBatang() { return CreateUniform(L"HCR Batang"); }
FontPreset FontDefs::HCRDotum() { return CreateUniform(L"HCR Dotum"); }
FontPreset FontDefs::HYHeadline() { return CreateUniform(L"HY헤드라인M"); }
FontPreset FontDefs::HYGothic() { return CreateUniform(L"HY중고딕"); }
FontPreset FontDefs::HYPost() { return CreateUniform(L"HY견고딕"); }
FontPreset FontDefs::HYGungso() { return CreateUniform(L"HY궁서"); }
FontPreset FontDefs::HYPMokgak() { return CreateUniform(L"HY목각파임B"); }
FontPreset FontDefs::HYGraphic() { return CreateUniform(L"HY그래픽"); }
FontPreset FontDefs::HYYeopseo() { return CreateUniform(L"HY엽서L"); }

// D2Coding
FontPreset FontDefs::D2Coding() { return CreateUniform(L"D2Coding"); }
FontPreset FontDefs::D2CodingBold() { return CreateUniform(L"D2Coding"); }

// 영문 폰트
FontPreset FontDefs::Arial() { return CreateUniform(L"Arial"); }
FontPreset FontDefs::ArialBold() { return CreateUniform(L"Arial"); }
FontPreset FontDefs::TimesNewRoman() { return CreateUniform(L"Times New Roman"); }
FontPreset FontDefs::TimesNewRomanBold() { return CreateUniform(L"Times New Roman"); }
FontPreset FontDefs::CourierNew() { return CreateUniform(L"Courier New"); }
FontPreset FontDefs::CourierNewBold() { return CreateUniform(L"Courier New"); }
FontPreset FontDefs::Verdana() { return CreateUniform(L"Verdana"); }
FontPreset FontDefs::VerdanaBold() { return CreateUniform(L"Verdana"); }
FontPreset FontDefs::Georgia() { return CreateUniform(L"Georgia"); }
FontPreset FontDefs::GeorgiaBold() { return CreateUniform(L"Georgia"); }
FontPreset FontDefs::Tahoma() { return CreateUniform(L"Tahoma"); }
FontPreset FontDefs::TahomaBold() { return CreateUniform(L"Tahoma"); }
FontPreset FontDefs::Consolas() { return CreateUniform(L"Consolas"); }
FontPreset FontDefs::ConsolasBold() { return CreateUniform(L"Consolas"); }

// 본고딕 계열
FontPreset FontDefs::SourceHanSans() { return CreateUniform(L"본고딕"); }
FontPreset FontDefs::SourceHanSansBold() { return CreateUniform(L"본고딕 Bold"); }
FontPreset FontDefs::SourceHanSansLight() { return CreateUniform(L"본고딕 Light"); }
FontPreset FontDefs::SourceHanSansMedium() { return CreateUniform(L"본고딕 Medium"); }
FontPreset FontDefs::SourceHanSansHeavy() { return CreateUniform(L"본고딕 Heavy"); }

// 본명조 계열
FontPreset FontDefs::SourceHanSerif() { return CreateUniform(L"본명조"); }
FontPreset FontDefs::SourceHanSerifBold() { return CreateUniform(L"본명조 Bold"); }
FontPreset FontDefs::SourceHanSerifLight() { return CreateUniform(L"본명조 Light"); }
FontPreset FontDefs::SourceHanSerifMedium() { return CreateUniform(L"본명조 Medium"); }
FontPreset FontDefs::SourceHanSerifHeavy() { return CreateUniform(L"본명조 Heavy"); }

// 스포카 한 산스
FontPreset FontDefs::SpoqaHanSans() { return CreateUniform(L"Spoqa Han Sans"); }
FontPreset FontDefs::SpoqaHanSansBold() { return CreateUniform(L"Spoqa Han Sans Bold"); }
FontPreset FontDefs::SpoqaHanSansLight() { return CreateUniform(L"Spoqa Han Sans Light"); }
FontPreset FontDefs::SpoqaHanSansThin() { return CreateUniform(L"Spoqa Han Sans Thin"); }

// Pretendard
FontPreset FontDefs::Pretendard() { return CreateUniform(L"Pretendard"); }
FontPreset FontDefs::PretendardBold() { return CreateUniform(L"Pretendard Bold"); }
FontPreset FontDefs::PretendardLight() { return CreateUniform(L"Pretendard Light"); }
FontPreset FontDefs::PretendardMedium() { return CreateUniform(L"Pretendard Medium"); }
FontPreset FontDefs::PretendardSemiBold() { return CreateUniform(L"Pretendard SemiBold"); }
FontPreset FontDefs::PretendardExtraBold() { return CreateUniform(L"Pretendard ExtraBold"); }
FontPreset FontDefs::PretendardThin() { return CreateUniform(L"Pretendard Thin"); }
FontPreset FontDefs::PretendardExtraLight() { return CreateUniform(L"Pretendard ExtraLight"); }
FontPreset FontDefs::PretendardBlack() { return CreateUniform(L"Pretendard Black"); }

// SUIT
FontPreset FontDefs::SUIT() { return CreateUniform(L"SUIT"); }
FontPreset FontDefs::SUITBold() { return CreateUniform(L"SUIT Bold"); }
FontPreset FontDefs::SUITLight() { return CreateUniform(L"SUIT Light"); }
FontPreset FontDefs::SUITMedium() { return CreateUniform(L"SUIT Medium"); }
FontPreset FontDefs::SUITSemiBold() { return CreateUniform(L"SUIT SemiBold"); }
FontPreset FontDefs::SUITExtraBold() { return CreateUniform(L"SUIT ExtraBold"); }
FontPreset FontDefs::SUITThin() { return CreateUniform(L"SUIT Thin"); }
FontPreset FontDefs::SUITExtraLight() { return CreateUniform(L"SUIT ExtraLight"); }
FontPreset FontDefs::SUITHeavy() { return CreateUniform(L"SUIT Heavy"); }

// KoPub
FontPreset FontDefs::KoPubBatang() { return CreateUniform(L"KoPub바탕체 Medium"); }
FontPreset FontDefs::KoPubBatangBold() { return CreateUniform(L"KoPub바탕체 Bold"); }
FontPreset FontDefs::KoPubBatangLight() { return CreateUniform(L"KoPub바탕체 Light"); }
FontPreset FontDefs::KoPubDotum() { return CreateUniform(L"KoPub돋움체 Medium"); }
FontPreset FontDefs::KoPubDotumBold() { return CreateUniform(L"KoPub돋움체 Bold"); }
FontPreset FontDefs::KoPubDotumLight() { return CreateUniform(L"KoPub돋움체 Light"); }

// 마루부리
FontPreset FontDefs::MaruBuri() { return CreateUniform(L"마루 부리"); }
FontPreset FontDefs::MaruBuriBold() { return CreateUniform(L"마루 부리 Bold"); }
FontPreset FontDefs::MaruBuriLight() { return CreateUniform(L"마루 부리 Light"); }
FontPreset FontDefs::MaruBuriSemiBold() { return CreateUniform(L"마루 부리 SemiBold"); }
FontPreset FontDefs::MaruBuriExtraLight() { return CreateUniform(L"마루 부리 ExtraLight"); }

//=============================================================================
// 프리셋 관리
//=============================================================================

void FontDefs::InitializePresets()
{
    if (s_bInitialized) return;

    // 맑은 고딕
    s_presets[L"맑은고딕"] = MalgunGothic();
    s_presets[L"맑은 고딕"] = MalgunGothic();
    s_presets[L"MalgunGothic"] = MalgunGothic();
    s_presets[L"Malgun Gothic"] = MalgunGothic();

    // 굴림/돋움/바탕/궁서
    s_presets[L"굴림"] = Gulim();
    s_presets[L"굴림체"] = GulimChe();
    s_presets[L"돋움"] = Dotum();
    s_presets[L"돋움체"] = DotumChe();
    s_presets[L"바탕"] = Batang();
    s_presets[L"바탕체"] = BatangChe();
    s_presets[L"궁서"] = Gungsuh();
    s_presets[L"궁서체"] = GungsuhChe();

    // 나눔 계열
    s_presets[L"나눔고딕"] = NanumGothic();
    s_presets[L"나눔고딕 Bold"] = NanumGothicBold();
    s_presets[L"나눔고딕 Light"] = NanumGothicLight();
    s_presets[L"나눔고딕 ExtraBold"] = NanumGothicExtraBold();
    s_presets[L"나눔명조"] = NanumMyeongjo();
    s_presets[L"나눔명조 Bold"] = NanumMyeongjoBold();
    s_presets[L"나눔명조 ExtraBold"] = NanumMyeongjoExtraBold();
    s_presets[L"나눔손글씨 펜"] = NanumPenScript();
    s_presets[L"나눔손글씨 붓"] = NanumBrushScript();
    s_presets[L"나눔바른고딕"] = NanumBarunGothic();
    s_presets[L"나눔바른고딕 Bold"] = NanumBarunGothicBold();
    s_presets[L"나눔바른고딕 Light"] = NanumBarunGothicLight();
    s_presets[L"나눔바른고딕 UltraLight"] = NanumBarunGothicUltraLight();
    s_presets[L"나눔바른펜"] = NanumBarunPen();
    s_presets[L"나눔스퀘어"] = NanumSquare();
    s_presets[L"나눔스퀘어 Bold"] = NanumSquareBold();
    s_presets[L"나눔스퀘어 ExtraBold"] = NanumSquareExtraBold();
    s_presets[L"나눔스퀘어 Light"] = NanumSquareLight();
    s_presets[L"나눔스퀘어라운드"] = NanumSquareRound();
    s_presets[L"나눔스퀘어라운드 Bold"] = NanumSquareRoundBold();
    s_presets[L"나눔스퀘어라운드 ExtraBold"] = NanumSquareRoundExtraBold();
    s_presets[L"나눔스퀘어라운드 Light"] = NanumSquareRoundLight();

    // 영문 폰트
    s_presets[L"Arial"] = Arial();
    s_presets[L"Times New Roman"] = TimesNewRoman();
    s_presets[L"Courier New"] = CourierNew();
    s_presets[L"Verdana"] = Verdana();
    s_presets[L"Georgia"] = Georgia();
    s_presets[L"Tahoma"] = Tahoma();
    s_presets[L"Consolas"] = Consolas();

    // D2Coding
    s_presets[L"D2Coding"] = D2Coding();
    s_presets[L"D2코딩"] = D2Coding();

    // Pretendard
    s_presets[L"Pretendard"] = Pretendard();
    s_presets[L"프리텐다드"] = Pretendard();
    s_presets[L"Pretendard Bold"] = PretendardBold();
    s_presets[L"Pretendard Light"] = PretendardLight();
    s_presets[L"Pretendard Medium"] = PretendardMedium();
    s_presets[L"Pretendard SemiBold"] = PretendardSemiBold();
    s_presets[L"Pretendard ExtraBold"] = PretendardExtraBold();

    // 본고딕/본명조
    s_presets[L"본고딕"] = SourceHanSans();
    s_presets[L"Source Han Sans"] = SourceHanSans();
    s_presets[L"본명조"] = SourceHanSerif();
    s_presets[L"Source Han Serif"] = SourceHanSerif();

    s_bInitialized = true;
}

FontPreset FontDefs::GetPreset(const std::wstring& preset_name)
{
    InitializePresets();

    auto it = s_presets.find(preset_name);
    if (it != s_presets.end()) {
        return it->second;
    }

    // 기본값: 맑은 고딕
    return MalgunGothic();
}

bool FontDefs::HasPreset(const std::wstring& preset_name)
{
    InitializePresets();
    return s_presets.find(preset_name) != s_presets.end();
}

std::vector<std::wstring> FontDefs::GetPresetNames()
{
    InitializePresets();

    std::vector<std::wstring> names;
    for (const auto& pair : s_presets) {
        names.push_back(pair.first);
    }
    return names;
}

} // namespace cpyhwpx
