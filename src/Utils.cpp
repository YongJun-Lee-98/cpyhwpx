/**
 * @file Utils.cpp
 * @brief 유틸리티 함수 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "Utils.h"
#include <algorithm>
#include <cctype>
#include <cwctype>
#include <sstream>
#include <iomanip>
#include <map>

namespace cpyhwpx {
namespace Utils {

//=============================================================================
// 셀 주소 변환
//=============================================================================

std::tuple<int, int> AddrToTuple(const std::wstring& address)
{
    if (address.empty()) {
        return std::make_tuple(0, 0);
    }

    // 열 문자와 행 숫자 분리
    size_t i = 0;
    while (i < address.length() && iswalpha(address[i])) {
        i++;
    }

    std::wstring col_str = address.substr(0, i);
    std::wstring row_str = address.substr(i);

    int col = ColToNum(col_str);
    int row = 0;

    if (!row_str.empty()) {
        try {
            row = std::stoi(row_str);
        } catch (...) {
            row = 0;
        }
    }

    return std::make_tuple(row, col);
}

std::wstring TupleToAddr(int row, int col)
{
    return NumToCol(col) + std::to_wstring(row);
}

std::tuple<int, int> AddrToTupleZeroBased(const std::wstring& address)
{
    auto [row, col] = AddrToTuple(address);
    return std::make_tuple(row - 1, col - 1);
}

int ColToNum(const std::wstring& col_str)
{
    if (col_str.empty()) return 0;

    int result = 0;
    for (wchar_t c : col_str) {
        wchar_t upper = towupper(c);
        if (upper >= L'A' && upper <= L'Z') {
            result = result * 26 + (upper - L'A' + 1);
        }
    }
    return result;
}

std::wstring NumToCol(int col)
{
    if (col <= 0) return L"";

    std::wstring result;
    while (col > 0) {
        col--;  // 1-based to 0-based
        result = static_cast<wchar_t>(L'A' + (col % 26)) + result;
        col /= 26;
    }
    return result;
}

//=============================================================================
// 셀 범위 파싱
//=============================================================================

std::tuple<int, int, int, int> ParseRange(const std::wstring& range)
{
    // "A1:B3" 형식 파싱
    size_t colonPos = range.find(L':');
    if (colonPos == std::wstring::npos) {
        // 단일 셀
        auto [row, col] = AddrToTuple(range);
        return std::make_tuple(row, col, row, col);
    }

    std::wstring start = range.substr(0, colonPos);
    std::wstring end = range.substr(colonPos + 1);

    auto [startRow, startCol] = AddrToTuple(start);
    auto [endRow, endCol] = AddrToTuple(end);

    // 정규화 (start <= end)
    if (startRow > endRow) std::swap(startRow, endRow);
    if (startCol > endCol) std::swap(startCol, endCol);

    return std::make_tuple(startRow, startCol, endRow, endCol);
}

std::vector<std::wstring> ExpandRange(const std::wstring& range)
{
    std::vector<std::wstring> result;

    auto [startRow, startCol, endRow, endCol] = ParseRange(range);

    for (int row = startRow; row <= endRow; row++) {
        for (int col = startCol; col <= endCol; col++) {
            result.push_back(TupleToAddr(row, col));
        }
    }

    return result;
}

//=============================================================================
// 문자열 처리
//=============================================================================

std::wstring Trim(const std::wstring& str)
{
    return LTrim(RTrim(str));
}

std::wstring LTrim(const std::wstring& str)
{
    size_t start = str.find_first_not_of(L" \t\n\r\f\v");
    return (start == std::wstring::npos) ? L"" : str.substr(start);
}

std::wstring RTrim(const std::wstring& str)
{
    size_t end = str.find_last_not_of(L" \t\n\r\f\v");
    return (end == std::wstring::npos) ? L"" : str.substr(0, end + 1);
}

std::vector<std::wstring> Split(const std::wstring& str, const std::wstring& delimiter)
{
    std::vector<std::wstring> result;
    if (str.empty()) return result;

    size_t start = 0;
    size_t end = str.find(delimiter);

    while (end != std::wstring::npos) {
        result.push_back(str.substr(start, end - start));
        start = end + delimiter.length();
        end = str.find(delimiter, start);
    }

    result.push_back(str.substr(start));
    return result;
}

std::wstring Join(const std::vector<std::wstring>& parts, const std::wstring& delimiter)
{
    if (parts.empty()) return L"";

    std::wstring result = parts[0];
    for (size_t i = 1; i < parts.size(); i++) {
        result += delimiter + parts[i];
    }
    return result;
}

std::wstring Replace(const std::wstring& str, const std::wstring& from, const std::wstring& to)
{
    if (from.empty()) return str;

    std::wstring result = str;
    size_t pos = 0;

    while ((pos = result.find(from, pos)) != std::wstring::npos) {
        result.replace(pos, from.length(), to);
        pos += to.length();
    }

    return result;
}

bool EqualsIgnoreCase(const std::wstring& a, const std::wstring& b)
{
    if (a.length() != b.length()) return false;

    for (size_t i = 0; i < a.length(); i++) {
        if (towlower(a[i]) != towlower(b[i])) {
            return false;
        }
    }
    return true;
}

std::wstring ToLower(const std::wstring& str)
{
    std::wstring result = str;
    std::transform(result.begin(), result.end(), result.begin(), towlower);
    return result;
}

std::wstring ToUpper(const std::wstring& str)
{
    std::wstring result = str;
    std::transform(result.begin(), result.end(), result.begin(), towupper);
    return result;
}

//=============================================================================
// 파일 경로 처리
//=============================================================================

std::wstring GetExtension(const std::wstring& path)
{
    size_t dotPos = path.rfind(L'.');
    size_t slashPos = path.rfind(L'\\');
    size_t fwdSlashPos = path.rfind(L'/');

    size_t lastSlash = (slashPos == std::wstring::npos) ? 0 : slashPos;
    if (fwdSlashPos != std::wstring::npos && fwdSlashPos > lastSlash) {
        lastSlash = fwdSlashPos;
    }

    if (dotPos == std::wstring::npos || dotPos < lastSlash) {
        return L"";
    }

    return path.substr(dotPos);
}

std::wstring GetFileName(const std::wstring& path)
{
    size_t slashPos = path.rfind(L'\\');
    size_t fwdSlashPos = path.rfind(L'/');

    size_t lastSlash = 0;
    if (slashPos != std::wstring::npos) lastSlash = slashPos + 1;
    if (fwdSlashPos != std::wstring::npos && fwdSlashPos >= lastSlash) {
        lastSlash = fwdSlashPos + 1;
    }

    return path.substr(lastSlash);
}

std::wstring GetFileNameWithoutExtension(const std::wstring& path)
{
    std::wstring fileName = GetFileName(path);
    size_t dotPos = fileName.rfind(L'.');

    if (dotPos == std::wstring::npos) {
        return fileName;
    }

    return fileName.substr(0, dotPos);
}

std::wstring GetDirectory(const std::wstring& path)
{
    size_t slashPos = path.rfind(L'\\');
    size_t fwdSlashPos = path.rfind(L'/');

    size_t lastSlash = std::wstring::npos;
    if (slashPos != std::wstring::npos) lastSlash = slashPos;
    if (fwdSlashPos != std::wstring::npos) {
        if (lastSlash == std::wstring::npos || fwdSlashPos > lastSlash) {
            lastSlash = fwdSlashPos;
        }
    }

    if (lastSlash == std::wstring::npos) {
        return L"";
    }

    return path.substr(0, lastSlash);
}

std::wstring CombinePath(const std::wstring& path1, const std::wstring& path2)
{
    if (path1.empty()) return path2;
    if (path2.empty()) return path1;

    wchar_t lastChar = path1.back();
    if (lastChar == L'\\' || lastChar == L'/') {
        return path1 + path2;
    }

    return path1 + L"\\" + path2;
}

bool FileExists(const std::wstring& path)
{
    DWORD attr = GetFileAttributesW(path.c_str());
    return (attr != INVALID_FILE_ATTRIBUTES && !(attr & FILE_ATTRIBUTE_DIRECTORY));
}

bool DirectoryExists(const std::wstring& path)
{
    DWORD attr = GetFileAttributesW(path.c_str());
    return (attr != INVALID_FILE_ATTRIBUTES && (attr & FILE_ATTRIBUTE_DIRECTORY));
}

//=============================================================================
// 색상 처리
//=============================================================================

COLORREF HexToColorRef(const std::wstring& hex)
{
    std::wstring cleanHex = hex;
    if (!cleanHex.empty() && cleanHex[0] == L'#') {
        cleanHex = cleanHex.substr(1);
    }

    if (cleanHex.length() != 6) {
        return RGB(0, 0, 0);
    }

    try {
        unsigned int r = std::stoul(cleanHex.substr(0, 2), nullptr, 16);
        unsigned int g = std::stoul(cleanHex.substr(2, 2), nullptr, 16);
        unsigned int b = std::stoul(cleanHex.substr(4, 2), nullptr, 16);
        return RGB(r, g, b);
    } catch (...) {
        return RGB(0, 0, 0);
    }
}

std::wstring ColorRefToHex(COLORREF color)
{
    std::wostringstream oss;
    oss << L"#"
        << std::hex << std::setfill(L'0') << std::setw(2) << GetRValue(color)
        << std::setw(2) << GetGValue(color)
        << std::setw(2) << GetBValue(color);
    return oss.str();
}

//=============================================================================
// 숫자 처리
//=============================================================================

bool CheckTupleOfInts(const std::vector<int>& values)
{
    // 모든 값이 유효한 범위 내인지 확인
    for (int val : values) {
        if (val < INT_MIN || val > INT_MAX) {
            return false;
        }
    }
    return true;
}

//=============================================================================
// 데이터 처리
//=============================================================================

std::vector<std::wstring> RenameDuplicatesInList(const std::vector<std::wstring>& file_list)
{
    std::vector<std::wstring> result;
    std::map<std::wstring, int> counts;

    for (const auto& name : file_list) {
        if (counts.find(name) == counts.end()) {
            counts[name] = 0;
            result.push_back(name);
        } else {
            counts[name]++;
            std::wstring newName = name + L"_" + std::to_wstring(counts[name]);
            result.push_back(newName);
        }
    }

    return result;
}

std::vector<std::vector<std::wstring>> CropDataFromSelection(
    const std::vector<std::vector<std::wstring>>& data,
    const std::tuple<int, int, int, int>& selection)
{
    auto [startRow, startCol, endRow, endCol] = selection;

    std::vector<std::vector<std::wstring>> result;

    for (int row = startRow; row <= endRow && row < static_cast<int>(data.size()); row++) {
        std::vector<std::wstring> rowData;
        for (int col = startCol; col <= endCol && col < static_cast<int>(data[row].size()); col++) {
            rowData.push_back(data[row][col]);
        }
        result.push_back(rowData);
    }

    return result;
}

//=============================================================================
// 레지스트리 관련
//=============================================================================

bool CheckRegistryKey(const std::wstring& key_name)
{
    // HWP 보안 모듈 레지스트리 확인
    // HKEY_CURRENT_USER\Software\HNC\HwpAutomation
    HKEY hKey;
    std::wstring fullPath = L"Software\\HNC\\HwpAutomation\\" + key_name;

    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, fullPath.c_str(), 0, KEY_READ, &hKey);

    if (result == ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return true;
    }

    return false;
}

//=============================================================================
// COM 유틸리티
//=============================================================================

std::wstring VariantToString(const VARIANT& var)
{
    switch (var.vt) {
        case VT_BSTR:
            return var.bstrVal ? std::wstring(var.bstrVal) : L"";
        case VT_I4:
            return std::to_wstring(var.lVal);
        case VT_I2:
            return std::to_wstring(var.iVal);
        case VT_R8:
            return std::to_wstring(var.dblVal);
        case VT_R4:
            return std::to_wstring(var.fltVal);
        case VT_BOOL:
            return var.boolVal ? L"true" : L"false";
        default:
            return L"";
    }
}

int VariantToInt(const VARIANT& var, int default_value)
{
    switch (var.vt) {
        case VT_I4:
            return var.lVal;
        case VT_I2:
            return var.iVal;
        case VT_R8:
            return static_cast<int>(var.dblVal);
        case VT_R4:
            return static_cast<int>(var.fltVal);
        case VT_BOOL:
            return var.boolVal ? 1 : 0;
        case VT_BSTR:
            if (var.bstrVal) {
                try {
                    return std::stoi(var.bstrVal);
                } catch (...) {}
            }
            return default_value;
        default:
            return default_value;
    }
}

double VariantToDouble(const VARIANT& var, double default_value)
{
    switch (var.vt) {
        case VT_R8:
            return var.dblVal;
        case VT_R4:
            return var.fltVal;
        case VT_I4:
            return static_cast<double>(var.lVal);
        case VT_I2:
            return static_cast<double>(var.iVal);
        case VT_BSTR:
            if (var.bstrVal) {
                try {
                    return std::stod(var.bstrVal);
                } catch (...) {}
            }
            return default_value;
        default:
            return default_value;
    }
}

bool VariantToBool(const VARIANT& var, bool default_value)
{
    switch (var.vt) {
        case VT_BOOL:
            return var.boolVal != VARIANT_FALSE;
        case VT_I4:
            return var.lVal != 0;
        case VT_I2:
            return var.iVal != 0;
        case VT_BSTR:
            if (var.bstrVal) {
                std::wstring str(var.bstrVal);
                return (str == L"true" || str == L"1" || str == L"True" || str == L"TRUE");
            }
            return default_value;
        default:
            return default_value;
    }
}

} // namespace Utils
} // namespace cpyhwpx
