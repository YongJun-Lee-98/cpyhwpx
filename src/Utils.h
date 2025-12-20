/**
 * @file Utils.h
 * @brief 유틸리티 함수 모음
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * 주소 변환, 문자열 처리, 기타 헬퍼 함수
 */

#pragma once

#include "HwpTypes.h"
#include <Windows.h>
#include <OAIdl.h>      // VARIANT
#include <string>
#include <vector>
#include <tuple>
#include <regex>
#include <cwctype>      // iswalpha, towupper, towlower

namespace cpyhwpx {
namespace Utils {

//=============================================================================
// 셀 주소 변환 (엑셀 스타일)
//=============================================================================

/**
 * @brief 엑셀 주소를 튜플로 변환 (1-based)
 * @param address 셀 주소 ("A1", "B2", "AA10" 등)
 * @return (row, col) 튜플 (1-based)
 *
 * @example
 *   AddrToTuple("A1") -> (1, 1)
 *   AddrToTuple("B3") -> (3, 2)
 *   AddrToTuple("AA1") -> (1, 27)
 */
std::tuple<int, int> AddrToTuple(const std::wstring& address);

/**
 * @brief 튜플을 엑셀 주소로 변환 (1-based)
 * @param row 행 (1-based)
 * @param col 열 (1-based)
 * @return 셀 주소 문자열
 *
 * @example
 *   TupleToAddr(1, 1) -> "A1"
 *   TupleToAddr(3, 2) -> "B3"
 *   TupleToAddr(1, 27) -> "AA1"
 */
std::wstring TupleToAddr(int row, int col);

/**
 * @brief 엑셀 주소를 튜플로 변환 (0-based)
 * @param address 셀 주소
 * @return (row, col) 튜플 (0-based)
 */
std::tuple<int, int> AddrToTupleZeroBased(const std::wstring& address);

/**
 * @brief 열 문자를 열 번호로 변환 (1-based)
 * @param col_str 열 문자 ("A", "B", "AA" 등)
 * @return 열 번호 (1-based)
 */
int ColToNum(const std::wstring& col_str);

/**
 * @brief 열 번호를 열 문자로 변환 (1-based)
 * @param col 열 번호 (1-based)
 * @return 열 문자
 */
std::wstring NumToCol(int col);

//=============================================================================
// 셀 범위 파싱
//=============================================================================

/**
 * @brief 셀 범위 문자열을 튜플로 파싱
 * @param range 범위 문자열 ("A1:B3" 형식)
 * @return (start_row, start_col, end_row, end_col) (1-based)
 */
std::tuple<int, int, int, int> ParseRange(const std::wstring& range);

/**
 * @brief 범위 내 셀 주소 목록 생성
 * @param range 범위 문자열
 * @return 셀 주소 벡터
 */
std::vector<std::wstring> ExpandRange(const std::wstring& range);

//=============================================================================
// 문자열 처리
//=============================================================================

/**
 * @brief 문자열 좌우 공백 제거
 */
std::wstring Trim(const std::wstring& str);

/**
 * @brief 문자열 좌측 공백 제거
 */
std::wstring LTrim(const std::wstring& str);

/**
 * @brief 문자열 우측 공백 제거
 */
std::wstring RTrim(const std::wstring& str);

/**
 * @brief 문자열 분할
 * @param str 입력 문자열
 * @param delimiter 구분자
 * @return 분할된 문자열 벡터
 */
std::vector<std::wstring> Split(const std::wstring& str, const std::wstring& delimiter);

/**
 * @brief 문자열 결합
 * @param parts 문자열 벡터
 * @param delimiter 구분자
 * @return 결합된 문자열
 */
std::wstring Join(const std::vector<std::wstring>& parts, const std::wstring& delimiter);

/**
 * @brief 문자열 치환
 * @param str 입력 문자열
 * @param from 찾을 문자열
 * @param to 대체 문자열
 * @return 치환된 문자열
 */
std::wstring Replace(const std::wstring& str, const std::wstring& from, const std::wstring& to);

/**
 * @brief 대소문자 무시 비교
 */
bool EqualsIgnoreCase(const std::wstring& a, const std::wstring& b);

/**
 * @brief 소문자로 변환
 */
std::wstring ToLower(const std::wstring& str);

/**
 * @brief 대문자로 변환
 */
std::wstring ToUpper(const std::wstring& str);

//=============================================================================
// 파일 경로 처리
//=============================================================================

/**
 * @brief 파일 확장자 추출
 */
std::wstring GetExtension(const std::wstring& path);

/**
 * @brief 파일명 추출 (확장자 포함)
 */
std::wstring GetFileName(const std::wstring& path);

/**
 * @brief 파일명 추출 (확장자 제외)
 */
std::wstring GetFileNameWithoutExtension(const std::wstring& path);

/**
 * @brief 디렉토리 경로 추출
 */
std::wstring GetDirectory(const std::wstring& path);

/**
 * @brief 경로 결합
 */
std::wstring CombinePath(const std::wstring& path1, const std::wstring& path2);

/**
 * @brief 파일 존재 여부 확인
 */
bool FileExists(const std::wstring& path);

/**
 * @brief 디렉토리 존재 여부 확인
 */
bool DirectoryExists(const std::wstring& path);

//=============================================================================
// 색상 처리
//=============================================================================

/**
 * @brief RGB 값을 COLORREF로 변환
 */
inline COLORREF RGBToColorRef(int r, int g, int b) {
    return RGB(r, g, b);
}

/**
 * @brief COLORREF를 RGB 튜플로 변환
 */
inline std::tuple<int, int, int> ColorRefToRGB(COLORREF color) {
    return std::make_tuple(GetRValue(color), GetGValue(color), GetBValue(color));
}

/**
 * @brief 16진수 문자열을 COLORREF로 변환
 * @param hex "#RRGGBB" 또는 "RRGGBB" 형식
 */
COLORREF HexToColorRef(const std::wstring& hex);

/**
 * @brief COLORREF를 16진수 문자열로 변환
 */
std::wstring ColorRefToHex(COLORREF color);

//=============================================================================
// 숫자 처리
//=============================================================================

/**
 * @brief 정수 튜플 유효성 검사
 * @param values 검사할 정수 벡터
 * @return 모든 값이 유효한 정수인지
 */
bool CheckTupleOfInts(const std::vector<int>& values);

/**
 * @brief 범위 내 값 제한
 */
template<typename T>
T Clamp(T value, T min_val, T max_val) {
    if (value < min_val) return min_val;
    if (value > max_val) return max_val;
    return value;
}

//=============================================================================
// 데이터 처리
//=============================================================================

/**
 * @brief 중복 파일명 처리
 * @param file_list 파일명 목록
 * @return 중복이 제거/리네임된 목록
 */
std::vector<std::wstring> RenameDuplicatesInList(const std::vector<std::wstring>& file_list);

/**
 * @brief 선택 영역에서 데이터 추출
 * @param data 전체 데이터 (2D 문자열 벡터)
 * @param selection 선택 영역 (start_row, start_col, end_row, end_col)
 * @return 추출된 데이터
 */
std::vector<std::vector<std::wstring>> CropDataFromSelection(
    const std::vector<std::vector<std::wstring>>& data,
    const std::tuple<int, int, int, int>& selection);

//=============================================================================
// 레지스트리 관련
//=============================================================================

/**
 * @brief 보안 모듈 레지스트리 키 확인
 * @param key_name 키 이름
 * @return 키 존재 여부
 */
bool CheckRegistryKey(const std::wstring& key_name);

//=============================================================================
// COM 유틸리티
//=============================================================================

/**
 * @brief VARIANT를 wstring으로 변환
 */
std::wstring VariantToString(const VARIANT& var);

/**
 * @brief VARIANT를 int로 변환
 */
int VariantToInt(const VARIANT& var, int default_value = 0);

/**
 * @brief VARIANT를 double로 변환
 */
double VariantToDouble(const VARIANT& var, double default_value = 0.0);

/**
 * @brief VARIANT를 bool로 변환
 */
bool VariantToBool(const VARIANT& var, bool default_value = false);

} // namespace Utils
} // namespace cpyhwpx
