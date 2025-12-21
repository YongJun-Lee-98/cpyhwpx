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

// 보한글 레거시 폰트 (20개)
FontPreset FontDefs::BHGothic() { return CreateUniform(L"고딕", FontType::HFT); }
FontPreset FontDefs::BHMyeongjo() { return CreateUniform(L"명조", FontType::HFT); }
FontPreset FontDefs::BHSaemmul() { return CreateUniform(L"샘물", FontType::HFT); }
FontPreset FontDefs::BHPilgi() { return CreateUniform(L"필기", FontType::HFT); }
FontPreset FontDefs::BHSinMyeongjo() { return CreateUniform(L"신명조", FontType::HFT); }
FontPreset FontDefs::BHGyeonMyeongjo() { return CreateUniform(L"견명조", FontType::HFT); }
FontPreset FontDefs::BHJungGothic() { return CreateUniform(L"중고딕", FontType::HFT); }
FontPreset FontDefs::BHGyeonGothic() { return CreateUniform(L"견고딕", FontType::HFT); }
FontPreset FontDefs::BHGraphic() { return CreateUniform(L"그래픽", FontType::HFT); }
FontPreset FontDefs::BHGungseo() { return CreateUniform(L"궁서", FontType::HFT); }
FontPreset FontDefs::BHGaneunGonghan() { return CreateUniform(L"가는공한", FontType::HFT); }
FontPreset FontDefs::BHJungganGonghan() { return CreateUniform(L"중간공한", FontType::HFT); }
FontPreset FontDefs::BHGulgeunGonghan() { return CreateUniform(L"굵은공한", FontType::HFT); }
FontPreset FontDefs::BHGaneunHan() { return CreateUniform(L"가는한", FontType::HFT); }
FontPreset FontDefs::BHJungganHan() { return CreateUniform(L"중간한", FontType::HFT); }
FontPreset FontDefs::BHGulgeunHan() { return CreateUniform(L"굵은한", FontType::HFT); }
FontPreset FontDefs::BHPenHeullim() { return CreateUniform(L"펜흘림", FontType::HFT); }
FontPreset FontDefs::BHHeadline() { return CreateUniform(L"헤드라인", FontType::HFT); }
FontPreset FontDefs::BHGaneunHeadline() { return CreateUniform(L"가는헤드라인", FontType::HFT); }
FontPreset FontDefs::BHTaeNamu() { return CreateUniform(L"태나무", FontType::HFT); }

// 휴먼 폰트 (9개)
FontPreset FontDefs::HumanMyeongjo() { return CreateUniform(L"휴먼명조"); }
FontPreset FontDefs::HumanGothic() { return CreateUniform(L"휴먼고딕"); }
FontPreset FontDefs::HumanYetche() { return CreateUniform(L"휴먼옛체"); }
FontPreset FontDefs::HumanGaneunSaemche() { return CreateUniform(L"휴먼가는샘체"); }
FontPreset FontDefs::HumanJungganSaemche() { return CreateUniform(L"휴먼중간샘체"); }
FontPreset FontDefs::HumanGulgeunSaemche() { return CreateUniform(L"휴먼굵은샘체"); }
FontPreset FontDefs::HumanGaneunPamche() { return CreateUniform(L"휴먼가는팸체"); }
FontPreset FontDefs::HumanJungganPamche() { return CreateUniform(L"휴먼중간팸체"); }
FontPreset FontDefs::HumanGulgeunPamche() { return CreateUniform(L"휴먼굵은팸체"); }

// 양재 폰트 (9개)
FontPreset FontDefs::YangJaeDaunMyeongjoM() { return CreateDual(L"양재다운명조M", L"Times New Roman"); }
FontPreset FontDefs::YangJaeBonmokgakM() { return CreateDual(L"양재본목각M", L"Times New Roman"); }
FontPreset FontDefs::YangJaeSoseul() { return CreateDual(L"양재소슬", L"Times New Roman"); }
FontPreset FontDefs::YangJaeTeunteunB() { return CreateDual(L"양재튼튼B", L"Times New Roman"); }
FontPreset FontDefs::YangJaeChamsutB() { return CreateDual(L"양재참숯B", L"Times New Roman"); }
FontPreset FontDefs::YangJaeDulgi() { return CreateDual(L"양재둘기", L"Times New Roman"); }
FontPreset FontDefs::YangJaeMaehwa() { return CreateDual(L"양재매화", L"Times New Roman"); }
FontPreset FontDefs::YangJaeShanel() { return CreateDual(L"양재샤넬", L"Times New Roman"); }
FontPreset FontDefs::YangJaeWadang() { return CreateDual(L"양재와당", L"Times New Roman"); }
FontPreset FontDefs::YangJaeInitial() { return CreateDual(L"양재이니셜", L"Times New Roman"); }

// 신명 폰트 (10개)
FontPreset FontDefs::SMSeMyeongjo() { return CreateUniform(L"신명세명조"); }
FontPreset FontDefs::SMJungMyeongjo() { return CreateUniform(L"신명중명조"); }
FontPreset FontDefs::SMTaeMyeongjo() { return CreateUniform(L"신명태명조"); }
FontPreset FontDefs::SMGyeonMyeongjo() { return CreateUniform(L"신명견명조"); }
FontPreset FontDefs::SMSinmunMyeongjo() { return CreateUniform(L"신명신문명조"); }
FontPreset FontDefs::SMSeGothic() { return CreateUniform(L"신명세고딕"); }
FontPreset FontDefs::SMJungGothic() { return CreateUniform(L"신명중고딕"); }
FontPreset FontDefs::SMTaeGothic() { return CreateUniform(L"신명태고딕"); }
FontPreset FontDefs::SMGyeonGothic() { return CreateUniform(L"신명견고딕"); }
FontPreset FontDefs::SMGungseo() { return CreateUniform(L"신명궁서"); }

// 특수 한자 폰트 (#접두사) (5개)
FontPreset FontDefs::HanjaSeMyeongjo() { return CreateUniform(L"#세명조", FontType::HFT); }
FontPreset FontDefs::HanjaJungMyeongjo() { return CreateUniform(L"#중명조", FontType::HFT); }
FontPreset FontDefs::HanjaTaeMyeongjo() { return CreateUniform(L"#태명조", FontType::HFT); }
FontPreset FontDefs::HanjaGyeonMyeongjo() { return CreateUniform(L"#견명조", FontType::HFT); }
FontPreset FontDefs::HanjaJungGothic() { return CreateUniform(L"#중고딕", FontType::HFT); }

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

    // SUIT
    s_presets[L"SUIT"] = SUIT();
    s_presets[L"SUIT Bold"] = SUITBold();
    s_presets[L"SUIT Light"] = SUITLight();
    s_presets[L"SUIT Medium"] = SUITMedium();
    s_presets[L"SUIT SemiBold"] = SUITSemiBold();
    s_presets[L"SUIT ExtraBold"] = SUITExtraBold();
    s_presets[L"SUIT Thin"] = SUITThin();
    s_presets[L"SUIT ExtraLight"] = SUITExtraLight();
    s_presets[L"SUIT Heavy"] = SUITHeavy();

    // Spoqa Han Sans
    s_presets[L"Spoqa Han Sans"] = SpoqaHanSans();
    s_presets[L"Spoqa Han Sans Bold"] = SpoqaHanSansBold();
    s_presets[L"Spoqa Han Sans Light"] = SpoqaHanSansLight();
    s_presets[L"Spoqa Han Sans Thin"] = SpoqaHanSansThin();

    // KoPub
    s_presets[L"KoPub바탕체 Medium"] = KoPubBatang();
    s_presets[L"KoPub바탕체"] = KoPubBatang();
    s_presets[L"KoPub바탕체 Bold"] = KoPubBatangBold();
    s_presets[L"KoPub바탕체 Light"] = KoPubBatangLight();
    s_presets[L"KoPub돋움체 Medium"] = KoPubDotum();
    s_presets[L"KoPub돋움체"] = KoPubDotum();
    s_presets[L"KoPub돋움체 Bold"] = KoPubDotumBold();
    s_presets[L"KoPub돋움체 Light"] = KoPubDotumLight();

    // 마루부리
    s_presets[L"마루 부리"] = MaruBuri();
    s_presets[L"마루부리"] = MaruBuri();
    s_presets[L"마루 부리 Bold"] = MaruBuriBold();
    s_presets[L"마루 부리 Light"] = MaruBuriLight();
    s_presets[L"마루 부리 SemiBold"] = MaruBuriSemiBold();
    s_presets[L"마루 부리 ExtraLight"] = MaruBuriExtraLight();

    // 보한글 레거시 폰트 (20개)
    s_presets[L"고딕"] = BHGothic();
    s_presets[L"BHGothic"] = BHGothic();
    s_presets[L"명조"] = BHMyeongjo();
    s_presets[L"BHMyeongjo"] = BHMyeongjo();
    s_presets[L"샘물"] = BHSaemmul();
    s_presets[L"BHSaemmul"] = BHSaemmul();
    s_presets[L"필기"] = BHPilgi();
    s_presets[L"BHPilgi"] = BHPilgi();
    s_presets[L"신명조"] = BHSinMyeongjo();
    s_presets[L"BHSinMyeongjo"] = BHSinMyeongjo();
    s_presets[L"견명조"] = BHGyeonMyeongjo();
    s_presets[L"BHGyeonMyeongjo"] = BHGyeonMyeongjo();
    s_presets[L"중고딕"] = BHJungGothic();
    s_presets[L"BHJungGothic"] = BHJungGothic();
    s_presets[L"견고딕"] = BHGyeonGothic();
    s_presets[L"BHGyeonGothic"] = BHGyeonGothic();
    s_presets[L"그래픽"] = BHGraphic();
    s_presets[L"BHGraphic"] = BHGraphic();
    s_presets[L"보한글궁서"] = BHGungseo();  // 윈도우 궁서와 구분
    s_presets[L"BHGungseo"] = BHGungseo();
    s_presets[L"가는공한"] = BHGaneunGonghan();
    s_presets[L"BHGaneunGonghan"] = BHGaneunGonghan();
    s_presets[L"중간공한"] = BHJungganGonghan();
    s_presets[L"BHJungganGonghan"] = BHJungganGonghan();
    s_presets[L"굵은공한"] = BHGulgeunGonghan();
    s_presets[L"BHGulgeunGonghan"] = BHGulgeunGonghan();
    s_presets[L"가는한"] = BHGaneunHan();
    s_presets[L"BHGaneunHan"] = BHGaneunHan();
    s_presets[L"중간한"] = BHJungganHan();
    s_presets[L"BHJungganHan"] = BHJungganHan();
    s_presets[L"굵은한"] = BHGulgeunHan();
    s_presets[L"BHGulgeunHan"] = BHGulgeunHan();
    s_presets[L"펜흘림"] = BHPenHeullim();
    s_presets[L"BHPenHeullim"] = BHPenHeullim();
    s_presets[L"헤드라인"] = BHHeadline();
    s_presets[L"BHHeadline"] = BHHeadline();
    s_presets[L"가는헤드라인"] = BHGaneunHeadline();
    s_presets[L"BHGaneunHeadline"] = BHGaneunHeadline();
    s_presets[L"태나무"] = BHTaeNamu();
    s_presets[L"BHTaeNamu"] = BHTaeNamu();

    // 휴먼 폰트 (9개)
    s_presets[L"휴먼명조"] = HumanMyeongjo();
    s_presets[L"HumanMyeongjo"] = HumanMyeongjo();
    s_presets[L"휴먼고딕"] = HumanGothic();
    s_presets[L"HumanGothic"] = HumanGothic();
    s_presets[L"휴먼옛체"] = HumanYetche();
    s_presets[L"HumanYetche"] = HumanYetche();
    s_presets[L"휴먼가는샘체"] = HumanGaneunSaemche();
    s_presets[L"HumanGaneunSaemche"] = HumanGaneunSaemche();
    s_presets[L"휴먼중간샘체"] = HumanJungganSaemche();
    s_presets[L"HumanJungganSaemche"] = HumanJungganSaemche();
    s_presets[L"휴먼굵은샘체"] = HumanGulgeunSaemche();
    s_presets[L"HumanGulgeunSaemche"] = HumanGulgeunSaemche();
    s_presets[L"휴먼가는팸체"] = HumanGaneunPamche();
    s_presets[L"HumanGaneunPamche"] = HumanGaneunPamche();
    s_presets[L"휴먼중간팸체"] = HumanJungganPamche();
    s_presets[L"HumanJungganPamche"] = HumanJungganPamche();
    s_presets[L"휴먼굵은팸체"] = HumanGulgeunPamche();
    s_presets[L"HumanGulgeunPamche"] = HumanGulgeunPamche();

    // 양재 폰트 (10개)
    s_presets[L"양재다운명조M"] = YangJaeDaunMyeongjoM();
    s_presets[L"YangJaeDaunMyeongjoM"] = YangJaeDaunMyeongjoM();
    s_presets[L"양재본목각M"] = YangJaeBonmokgakM();
    s_presets[L"YangJaeBonmokgakM"] = YangJaeBonmokgakM();
    s_presets[L"양재소슬"] = YangJaeSoseul();
    s_presets[L"YangJaeSoseul"] = YangJaeSoseul();
    s_presets[L"양재튼튼B"] = YangJaeTeunteunB();
    s_presets[L"YangJaeTeunteunB"] = YangJaeTeunteunB();
    s_presets[L"양재참숯B"] = YangJaeChamsutB();
    s_presets[L"YangJaeChamsutB"] = YangJaeChamsutB();
    s_presets[L"양재둘기"] = YangJaeDulgi();
    s_presets[L"YangJaeDulgi"] = YangJaeDulgi();
    s_presets[L"양재매화"] = YangJaeMaehwa();
    s_presets[L"YangJaeMaehwa"] = YangJaeMaehwa();
    s_presets[L"양재샤넬"] = YangJaeShanel();
    s_presets[L"YangJaeShanel"] = YangJaeShanel();
    s_presets[L"양재와당"] = YangJaeWadang();
    s_presets[L"YangJaeWadang"] = YangJaeWadang();
    s_presets[L"양재이니셜"] = YangJaeInitial();
    s_presets[L"YangJaeInitial"] = YangJaeInitial();

    // 신명 폰트 (10개)
    s_presets[L"신명세명조"] = SMSeMyeongjo();
    s_presets[L"SMSeMyeongjo"] = SMSeMyeongjo();
    s_presets[L"신명중명조"] = SMJungMyeongjo();
    s_presets[L"SMJungMyeongjo"] = SMJungMyeongjo();
    s_presets[L"신명태명조"] = SMTaeMyeongjo();
    s_presets[L"SMTaeMyeongjo"] = SMTaeMyeongjo();
    s_presets[L"신명견명조"] = SMGyeonMyeongjo();
    s_presets[L"SMGyeonMyeongjo"] = SMGyeonMyeongjo();
    s_presets[L"신명신문명조"] = SMSinmunMyeongjo();
    s_presets[L"SMSinmunMyeongjo"] = SMSinmunMyeongjo();
    s_presets[L"신명세고딕"] = SMSeGothic();
    s_presets[L"SMSeGothic"] = SMSeGothic();
    s_presets[L"신명중고딕"] = SMJungGothic();
    s_presets[L"SMJungGothic"] = SMJungGothic();
    s_presets[L"신명태고딕"] = SMTaeGothic();
    s_presets[L"SMTaeGothic"] = SMTaeGothic();
    s_presets[L"신명견고딕"] = SMGyeonGothic();
    s_presets[L"SMGyeonGothic"] = SMGyeonGothic();
    s_presets[L"신명궁서"] = SMGungseo();
    s_presets[L"SMGungseo"] = SMGungseo();

    // 특수 한자 폰트 (#접두사) (5개)
    s_presets[L"#세명조"] = HanjaSeMyeongjo();
    s_presets[L"HanjaSeMyeongjo"] = HanjaSeMyeongjo();
    s_presets[L"#중명조"] = HanjaJungMyeongjo();
    s_presets[L"HanjaJungMyeongjo"] = HanjaJungMyeongjo();
    s_presets[L"#태명조"] = HanjaTaeMyeongjo();
    s_presets[L"HanjaTaeMyeongjo"] = HanjaTaeMyeongjo();
    s_presets[L"#견명조"] = HanjaGyeonMyeongjo();
    s_presets[L"HanjaGyeonMyeongjo"] = HanjaGyeonMyeongjo();
    s_presets[L"#중고딕"] = HanjaJungGothic();
    s_presets[L"HanjaJungGothic"] = HanjaJungGothic();

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
