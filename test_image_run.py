# -*- coding: utf-8 -*-
"""
이미지 삽입 기능 테스트 (실행용)
"""
import sys
sys.path.insert(0, 'D:/Code/hwp_test/cpyhwpx')

import cpyhwpx
import time

def test_insert_picture():
    print("=" * 60)
    print("cpyhwpx InsertPicture Test")
    print("=" * 60)

    # 테스트 이미지 경로
    test_image = r"D:\Code\hwp_test\.venv\Lib\site-packages\win32\Demos\images\smiley.bmp"

    # HWP 초기화
    print("\n[0] HWP 초기화")
    print("-" * 40)
    hwp = cpyhwpx.Hwp(visible=True)
    hwp.initialize()
    hwp.clear()
    time.sleep(0.5)
    print("HWP 초기화 완료")

    print("\n[1] 기본 이미지 삽입 테스트")
    print("-" * 40)

    # 텍스트 삽입
    hwp.insert_text("이미지 삽입 테스트\n\n")
    print("텍스트 삽입 완료")

    # 기본 이미지 삽입 (원본 크기)
    print(f"테스트 이미지: {test_image}")
    result = hwp.insert_picture(test_image)
    print(f"insert_picture(원본크기) = {result}")

    hwp.insert_text("\n\n")

    # 지정 크기로 삽입 (30x30mm)
    result = hwp.insert_picture(test_image, sizeoption=1, width=30, height=30)
    print(f"insert_picture(30x30mm) = {result}")

    hwp.insert_text("\n\n")

    # 그레이스케일 효과
    result = hwp.insert_picture(test_image, effect=1)
    print(f"insert_picture(그레이스케일) = {result}")

    print("\n[2] 테이블 셀 내 이미지 삽입 테스트")
    print("-" * 40)

    # 새 문단
    hwp.insert_text("\n\n테이블 내 이미지 삽입:\n\n")

    # 테이블 생성
    result = hwp.create_table(rows=2, cols=2)
    print(f"create_table(2,2) = {result}")
    time.sleep(0.3)

    # 테이블 첫 셀로 이동
    result = hwp.get_into_nth_table(0)
    print(f"get_into_nth_table(0) = {result}")
    print(f"is_cell() = {hwp.is_cell()}")

    # 셀 맞춤 (종횡비 유지)
    result = hwp.insert_picture(test_image, sizeoption=3)
    print(f"insert_picture(셀맞춤+종횡비) = {result}")

    # 다음 셀로 이동
    hwp.table_right_cell()
    result = hwp.insert_picture(test_image, sizeoption=2)  # 셀 맞춤 (종횡비 무시)
    print(f"insert_picture(셀맞춤) = {result}")

    print("\n" + "=" * 60)
    print("InsertPicture Test Complete!")
    print("=" * 60)

    # 문서 저장
    save_path = "D:/Code/hwp_test/cpyhwpx/test_image_output.hwp"
    hwp.save_as(save_path)
    print(f"\n저장됨: {save_path}")

    print("\n5초 후 HWP 종료...")
    time.sleep(5)
    hwp.quit()
    print("HWP 종료")

if __name__ == "__main__":
    test_insert_picture()
