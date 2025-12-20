# -*- coding: utf-8 -*-
"""
cpyhwpx 64-bit 브릿지 테스트
64-bit Python에서 32-bit HWP 제어 테스트
"""

import sys
import struct
import os

# cpyhwpx 패키지 경로 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

print("=" * 60)
print("cpyhwpx 64-bit 브릿지 테스트")
print("=" * 60)

# 아키텍처 확인
bits = struct.calcsize("P") * 8
print(f"Python 버전: {sys.version}")
print(f"아키텍처: {bits}-bit")
print()

if bits != 64:
    print("경고: 이 테스트는 64-bit Python에서 실행해야 합니다.")
    print("현재 환경에서는 직접 로드 방식으로 테스트합니다.")
    print()

try:
    import cpyhwpx
    print(f"cpyhwpx 모듈 로드 성공!")
    print(f"버전: {cpyhwpx.__version__}")
    print(f"아키텍처 정보: {cpyhwpx.get_architecture_info()}")
    print()
except ImportError as e:
    print(f"cpyhwpx 모듈 로드 실패: {e}")
    sys.exit(1)

print("=" * 60)
print("HWP 연동 테스트 시작")
print("=" * 60)

try:
    # HWP 인스턴스 생성
    print("\n1. Hwp 인스턴스 생성 중...")
    hwp = cpyhwpx.Hwp(visible=True, new_instance=True)
    print("   Hwp 인스턴스 생성 완료!")

    # 초기화
    print("\n2. HWP 초기화 중...")
    result = hwp.initialize()
    print(f"   초기화 결과: {result}")

    # 보안 모듈 등록
    print("\n3. 보안 모듈 등록 중...")
    result = hwp.register_module("FilePathCheckDLL", "FilePathCheckerModuleExample")
    print(f"   등록 결과: {result}")

    # 새 문서 생성
    print("\n4. 새 문서 생성 중...")
    hwp.clear()
    print("   새 문서 생성 완료!")

    # 텍스트 삽입
    print("\n5. 텍스트 삽입 중...")
    hwp.insert_text("안녕하세요!")
    hwp.insert_text("\n")
    hwp.insert_text(f"64-bit Python에서 브릿지를 통해 작성된 문서입니다.")
    hwp.insert_text("\n")
    hwp.insert_text(f"Python 아키텍처: {bits}-bit")
    print("   텍스트 삽입 완료!")

    # 파일 저장
    print("\n6. 파일 저장 중...")
    save_path = os.path.join(os.path.dirname(__file__), "bridge_test_output.hwp")
    result = hwp.save_as(save_path)
    print(f"   저장 결과: {result}")
    print(f"   저장 경로: {save_path}")

    # 잠시 대기
    import time
    print("\n3초 후 HWP 종료...")
    time.sleep(3)

    # HWP 종료
    print("\n7. HWP 종료 중...")
    hwp.quit()
    print("   HWP 종료 완료!")

    print("\n" + "=" * 60)
    print("테스트 성공!")
    print("=" * 60)

except Exception as e:
    print(f"\n오류 발생: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
