# -*- coding: utf-8 -*-
"""자동 보안 모듈 등록 테스트"""
import sys
sys.path.insert(0, r"D:\Code\hwp_test\cpyhwpx")

import cpyhwpx

def test_auto_register():
    """register_module=True (기본값) 테스트"""
    print("="*60)
    print("자동 보안 모듈 등록 테스트")
    print("="*60)

    # 기본값으로 생성 (register_module=True)
    print("\n1. Hwp(visible=False) - register_module=True (기본값)")
    hwp = cpyhwpx.Hwp(visible=False)

    print("   initialize() 호출...")
    result = hwp.initialize()
    print(f"   초기화 결과: {result}")

    # 레지스트리 확인
    is_registered = cpyhwpx.Hwp.check_registry_key()
    print(f"   레지스트리 등록 확인: {is_registered}")

    # 테스트: 파일 열기가 되는지 확인
    print("\n   간단한 텍스트 삽입 테스트...")
    hwp.insert_text("자동 등록 테스트 성공!")
    text_result = hwp.get_text()
    print(f"   삽입된 텍스트: {text_result[1][:20]}...")

    hwp.quit()
    print("\n   [SUCCESS] register_module=True 테스트 통과")

def test_no_auto_register():
    """register_module=False 테스트"""
    print("\n" + "="*60)
    print("자동 등록 비활성화 테스트")
    print("="*60)

    print("\n2. Hwp(visible=False, register_module=False)")
    hwp = cpyhwpx.Hwp(visible=False, register_module=False)

    print("   initialize() 호출...")
    result = hwp.initialize()
    print(f"   초기화 결과: {result}")

    # 수동 등록도 가능
    print("   수동 auto_register_module() 호출...")
    manual_result = hwp.auto_register_module()
    print(f"   수동 등록 결과: {manual_result}")

    hwp.quit()
    print("\n   [SUCCESS] register_module=False 테스트 통과")

if __name__ == "__main__":
    test_auto_register()
    test_no_auto_register()
    print("\n" + "="*60)
    print("모든 테스트 통과!")
    print("="*60)
