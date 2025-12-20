# -*- coding: utf-8 -*-
"""
InsertFile 기능 테스트
"""
import sys
sys.path.insert(0, 'D:/Code/hwp_test/cpyhwpx')

import cpyhwpx
import time
import os

def test_insert_file():
    print("=" * 60)
    print("cpyhwpx InsertFile Test")
    print("=" * 60)

    # 삽입할 테스트 파일 경로
    test_file = r"D:\Code\hwp_test\cpyhwpx\test_image_output.hwp"

    if not os.path.exists(test_file):
        print(f"테스트 파일 없음: {test_file}")
        print("먼저 test_image_run.py를 실행하여 테스트 파일을 생성하세요.")
        return

    # HWP 초기화
    print("\n[0] HWP 초기화")
    print("-" * 40)
    hwp = cpyhwpx.Hwp(visible=True)
    hwp.initialize()
    hwp.clear()
    time.sleep(0.5)
    print("HWP 초기화 완료")

    print("\n[1] 기본 파일 삽입 테스트")
    print("-" * 40)

    # 텍스트 삽입
    hwp.insert_text("=== 파일 삽입 테스트 ===\n\n")
    hwp.insert_text("아래에 파일이 삽입됩니다:\n\n")
    print("텍스트 삽입 완료")

    # 기본 파일 삽입 (모든 포맷 유지)
    print(f"삽입할 파일: {test_file}")
    result = hwp.insert_file(test_file)
    print(f"insert_file(기본) = {result}")

    print("\n[2] 옵션 테스트")
    print("-" * 40)

    # 문서 끝으로 이동
    hwp.move_pos(3, 0, 0)
    hwp.insert_text("\n\n=== 두 번째 삽입 (문자모양만 유지) ===\n\n")

    # 문자 모양만 유지하고 삽입
    result = hwp.insert_file(test_file,
                              keep_section=0,
                              keep_charshape=1,
                              keep_parashape=0,
                              keep_style=0)
    print(f"insert_file(문자모양만) = {result}")

    print("\n" + "=" * 60)
    print("InsertFile Test Complete!")
    print("=" * 60)

    # 문서 저장
    save_path = "D:/Code/hwp_test/cpyhwpx/test_insert_file_output.hwp"
    hwp.save_as(save_path)
    print(f"\n저장됨: {save_path}")

    print("\n5초 후 HWP 종료...")
    time.sleep(5)
    hwp.quit()
    print("HWP 종료")

if __name__ == "__main__":
    test_insert_file()
