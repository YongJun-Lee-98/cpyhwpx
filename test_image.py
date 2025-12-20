# -*- coding: utf-8 -*-
"""
이미지 삽입 기능 테스트
"""
import sys
sys.path.insert(0, 'D:/Code/hwp_test/cpyhwpx')

import cpyhwpx
import time
import os

def test_insert_picture():
    print("=" * 60)
    print("cpyhwpx InsertPicture Test")
    print("=" * 60)

    # 테스트 이미지 경로 (샘플 이미지가 없으면 스킵)
    test_image = "D:/Code/hwp_test/test_image.jpg"

    # HWP 초기화
    hwp = cpyhwpx.Hwp(visible=True)
    hwp.initialize()
    hwp.clear()
    time.sleep(0.5)

    print("\n[1] 기본 이미지 삽입 테스트")
    print("-" * 40)

    # 텍스트 삽입
    hwp.insert_text("이미지 삽입 테스트\n\n")

    if os.path.exists(test_image):
        # 기본 이미지 삽입 (원본 크기)
        result = hwp.insert_picture(test_image)
        print(f"insert_picture(기본) = {result}")

        hwp.insert_text("\n\n")

        # 지정 크기로 삽입 (50x40mm)
        result = hwp.insert_picture(test_image, sizeoption=1, width=50, height=40)
        print(f"insert_picture(50x40mm) = {result}")

        hwp.insert_text("\n\n")

        # 그레이스케일 효과
        result = hwp.insert_picture(test_image, effect=1)
        print(f"insert_picture(그레이스케일) = {result}")
    else:
        print(f"테스트 이미지 없음: {test_image}")
        print("이미지 없이 API 호출 테스트...")
        # 없는 파일로 호출 (실패 예상)
        result = hwp.insert_picture("D:/nonexistent.jpg")
        print(f"insert_picture(없는 파일) = {result} (False 예상)")

    print("\n[2] 테이블 셀 내 이미지 삽입 테스트")
    print("-" * 40)

    # 새 문단
    hwp.insert_text("\n\n테이블 내 이미지 삽입:\n\n")

    # 테이블 생성
    hwp.create_table(rows=2, cols=2)
    time.sleep(0.3)

    # 테이블 첫 셀로 이동
    hwp.get_into_nth_table(0)
    print(f"is_cell() = {hwp.is_cell()}")

    if os.path.exists(test_image):
        # 셀 맞춤 (종횡비 유지)
        result = hwp.insert_picture(test_image, sizeoption=3)
        print(f"insert_picture(셀맞춤+종횡비) = {result}")

    print("\n" + "=" * 60)
    print("InsertPicture Test Complete!")
    print("=" * 60)

    # 문서 끝으로 이동
    hwp.move_pos(3, 0, 0)  # moveDocEnd = 3

    time.sleep(2)
    hwp.quit()
    print("HWP 종료")

if __name__ == "__main__":
    test_insert_picture()
