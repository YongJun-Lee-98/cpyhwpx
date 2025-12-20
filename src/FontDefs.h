/**
 * @file FontDefs.h
 * @brief 폰트 프리셋 정의
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * fonts.py의 111개 폰트 프리셋 대응
 */

#pragma once

#include "HwpTypes.h"
#include <string>
#include <map>
#include <vector>

namespace cpyhwpx {

/**
 * @brief 폰트 타입 상수
 */
namespace FontType {
    constexpr int TTF = 2;      // TrueType Font
    constexpr int HFT = 1;      // HWP Font
}

/**
 * @class FontDefs
 * @brief 폰트 프리셋 관리 클래스
 *
 * 자주 사용되는 폰트 조합을 프리셋으로 제공
 */
class FontDefs {
public:
    //=========================================================================
    // 프리셋 가져오기
    //=========================================================================

    /**
     * @brief 프리셋 이름으로 폰트 정보 가져오기
     * @param preset_name 프리셋 이름 (예: "맑은고딕", "굴림", "나눔고딕" 등)
     * @return 폰트 프리셋 (없으면 기본값)
     */
    static FontPreset GetPreset(const std::wstring& preset_name);

    /**
     * @brief 프리셋 존재 여부 확인
     */
    static bool HasPreset(const std::wstring& preset_name);

    /**
     * @brief 모든 프리셋 이름 목록
     */
    static std::vector<std::wstring> GetPresetNames();

    //=========================================================================
    // 개별 폰트 프리셋
    //=========================================================================

    // 맑은 고딕 계열
    static FontPreset MalgunGothic();
    static FontPreset MalgunGothicBold();

    // 굴림 계열
    static FontPreset Gulim();
    static FontPreset GulimChe();
    static FontPreset Dotum();
    static FontPreset DotumChe();
    static FontPreset Batang();
    static FontPreset BatangChe();
    static FontPreset Gungsuh();
    static FontPreset GungsuhChe();

    // 나눔 계열
    static FontPreset NanumGothic();
    static FontPreset NanumGothicBold();
    static FontPreset NanumGothicLight();
    static FontPreset NanumGothicExtraBold();
    static FontPreset NanumMyeongjo();
    static FontPreset NanumMyeongjoBold();
    static FontPreset NanumMyeongjoExtraBold();
    static FontPreset NanumPenScript();
    static FontPreset NanumBrushScript();
    static FontPreset NanumBarunGothic();
    static FontPreset NanumBarunGothicBold();
    static FontPreset NanumBarunGothicLight();
    static FontPreset NanumBarunGothicUltraLight();
    static FontPreset NanumBarunPen();
    static FontPreset NanumSquare();
    static FontPreset NanumSquareBold();
    static FontPreset NanumSquareExtraBold();
    static FontPreset NanumSquareLight();
    static FontPreset NanumSquareRound();
    static FontPreset NanumSquareRoundBold();
    static FontPreset NanumSquareRoundExtraBold();
    static FontPreset NanumSquareRoundLight();

    // 한컴 계열
    static FontPreset HCRBatang();
    static FontPreset HCRDotum();
    static FontPreset HYHeadline();
    static FontPreset HYGothic();
    static FontPreset HYPost();
    static FontPreset HYGungso();
    static FontPreset HYPMokgak();
    static FontPreset HYGraphic();
    static FontPreset HYYeopseo();

    // D2Coding
    static FontPreset D2Coding();
    static FontPreset D2CodingBold();

    // 영문 폰트
    static FontPreset Arial();
    static FontPreset ArialBold();
    static FontPreset TimesNewRoman();
    static FontPreset TimesNewRomanBold();
    static FontPreset CourierNew();
    static FontPreset CourierNewBold();
    static FontPreset Verdana();
    static FontPreset VerdanaBold();
    static FontPreset Georgia();
    static FontPreset GeorgiaBold();
    static FontPreset Tahoma();
    static FontPreset TahomaBold();
    static FontPreset Consolas();
    static FontPreset ConsolasBold();

    // 본고딕 계열
    static FontPreset SourceHanSans();
    static FontPreset SourceHanSansBold();
    static FontPreset SourceHanSansLight();
    static FontPreset SourceHanSansMedium();
    static FontPreset SourceHanSansHeavy();

    // 본명조 계열
    static FontPreset SourceHanSerif();
    static FontPreset SourceHanSerifBold();
    static FontPreset SourceHanSerifLight();
    static FontPreset SourceHanSerifMedium();
    static FontPreset SourceHanSerifHeavy();

    // 스포카 한 산스
    static FontPreset SpoqaHanSans();
    static FontPreset SpoqaHanSansBold();
    static FontPreset SpoqaHanSansLight();
    static FontPreset SpoqaHanSansThin();

    // Pretendard
    static FontPreset Pretendard();
    static FontPreset PretendardBold();
    static FontPreset PretendardLight();
    static FontPreset PretendardMedium();
    static FontPreset PretendardSemiBold();
    static FontPreset PretendardExtraBold();
    static FontPreset PretendardThin();
    static FontPreset PretendardExtraLight();
    static FontPreset PretendardBlack();

    // Suit
    static FontPreset SUIT();
    static FontPreset SUITBold();
    static FontPreset SUITLight();
    static FontPreset SUITMedium();
    static FontPreset SUITSemiBold();
    static FontPreset SUITExtraBold();
    static FontPreset SUITThin();
    static FontPreset SUITExtraLight();
    static FontPreset SUITHeavy();

    // 기타
    static FontPreset KoPubBatang();
    static FontPreset KoPubBatangBold();
    static FontPreset KoPubBatangLight();
    static FontPreset KoPubDotum();
    static FontPreset KoPubDotumBold();
    static FontPreset KoPubDotumLight();
    static FontPreset MaruBuri();
    static FontPreset MaruBuriBold();
    static FontPreset MaruBuriLight();
    static FontPreset MaruBuriSemiBold();
    static FontPreset MaruBuriExtraLight();

    //=========================================================================
    // 헬퍼 메서드
    //=========================================================================

    /**
     * @brief 단일 폰트명으로 프리셋 생성
     * @param font_name 폰트 이름
     * @param font_type 폰트 타입 (기본: TTF)
     * @return 생성된 프리셋
     */
    static FontPreset CreateUniform(const std::wstring& font_name, int font_type = FontType::TTF);

    /**
     * @brief 한글/영문 폰트 분리 프리셋 생성
     * @param hangul_name 한글 폰트
     * @param latin_name 영문 폰트
     * @return 생성된 프리셋
     */
    static FontPreset CreateDual(const std::wstring& hangul_name, const std::wstring& latin_name);

private:
    // 프리셋 초기화
    static void InitializePresets();
    static bool s_bInitialized;
    static std::map<std::wstring, FontPreset> s_presets;
};

} // namespace cpyhwpx
