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

    // 보한글 레거시 폰트 (20개)
    static FontPreset BHGothic();           // 고딕
    static FontPreset BHMyeongjo();         // 명조
    static FontPreset BHSaemmul();          // 샘물
    static FontPreset BHPilgi();            // 필기
    static FontPreset BHSinMyeongjo();      // 신명조
    static FontPreset BHGyeonMyeongjo();    // 견명조
    static FontPreset BHJungGothic();       // 중고딕
    static FontPreset BHGyeonGothic();      // 견고딕
    static FontPreset BHGraphic();          // 그래픽
    static FontPreset BHGungseo();          // 궁서
    static FontPreset BHGaneunGonghan();    // 가는공한
    static FontPreset BHJungganGonghan();   // 중간공한
    static FontPreset BHGulgeunGonghan();   // 굵은공한
    static FontPreset BHGaneunHan();        // 가는한
    static FontPreset BHJungganHan();       // 중간한
    static FontPreset BHGulgeunHan();       // 굵은한
    static FontPreset BHPenHeullim();       // 펜흘림
    static FontPreset BHHeadline();         // 헤드라인
    static FontPreset BHGaneunHeadline();   // 가는헤드라인
    static FontPreset BHTaeNamu();          // 태나무

    // 휴먼 폰트 (9개)
    static FontPreset HumanMyeongjo();      // 휴먼명조
    static FontPreset HumanGothic();        // 휴먼고딕
    static FontPreset HumanYetche();        // 휴먼옛체
    static FontPreset HumanGaneunSaemche(); // 휴먼가는샘체
    static FontPreset HumanJungganSaemche();// 휴먼중간샘체
    static FontPreset HumanGulgeunSaemche();// 휴먼굵은샘체
    static FontPreset HumanGaneunPamche();  // 휴먼가는팸체
    static FontPreset HumanJungganPamche(); // 휴먼중간팸체
    static FontPreset HumanGulgeunPamche(); // 휴먼굵은팸체

    // 양재 폰트 (9개)
    static FontPreset YangJaeDaunMyeongjoM();  // 양재다운명조M
    static FontPreset YangJaeBonmokgakM();     // 양재본목각M
    static FontPreset YangJaeSoseul();         // 양재소슬
    static FontPreset YangJaeTeunteunB();      // 양재튼튼B
    static FontPreset YangJaeChamsutB();       // 양재참숯B
    static FontPreset YangJaeDulgi();          // 양재둘기
    static FontPreset YangJaeMaehwa();         // 양재매화
    static FontPreset YangJaeShanel();         // 양재샤넬
    static FontPreset YangJaeWadang();         // 양재와당
    static FontPreset YangJaeInitial();        // 양재이니셜

    // 신명 폰트 (10개)
    static FontPreset SMSeMyeongjo();       // 신명세명조
    static FontPreset SMJungMyeongjo();     // 신명중명조
    static FontPreset SMTaeMyeongjo();      // 신명태명조
    static FontPreset SMGyeonMyeongjo();    // 신명견명조
    static FontPreset SMSinmunMyeongjo();   // 신명신문명조
    static FontPreset SMSeGothic();         // 신명세고딕
    static FontPreset SMJungGothic();       // 신명중고딕
    static FontPreset SMTaeGothic();        // 신명태고딕
    static FontPreset SMGyeonGothic();      // 신명견고딕
    static FontPreset SMGungseo();          // 신명궁서

    // 특수 한자 폰트 (#접두사) (5개)
    static FontPreset HanjaSeMyeongjo();    // #세명조
    static FontPreset HanjaJungMyeongjo();  // #중명조
    static FontPreset HanjaTaeMyeongjo();   // #태명조
    static FontPreset HanjaGyeonMyeongjo(); // #견명조
    static FontPreset HanjaJungGothic();    // #중고딕

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
