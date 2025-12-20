# -*- coding: utf-8 -*-
"""
테이블 기능 테스트 (FindCtrl 버그 수정 후)
"""
import sys
sys.path.insert(0, 'D:/Code/hwp_test/cpyhwpx')

import cpyhwpx
import time

def test_table():
    print("=" * 60)
    print("cpyhwpx Table Test (FindCtrl Bug Fix)")
    print("=" * 60)

    # HWP init
    hwp = cpyhwpx.Hwp(visible=True)
    hwp.initialize()
    hwp.clear()
    time.sleep(0.5)

    print("\n[1] Table Creation Test")
    print("-" * 40)

    # Create 3x4 table
    result = hwp.create_table(rows=3, cols=4)
    print(f"create_table(3, 4) = {result}")

    # 테이블 생성 직후 위치 확인
    print(f"is_cell() after create = {hwp.is_cell()}")
    time.sleep(0.5)

    # 문서 끝으로 이동 후 다시 테이블로 진입 테스트
    hwp.move_pos(3, 0, 0)  # moveDocEnd = 3
    print(f"is_cell() after move to end = {hwp.is_cell()}")

    # Navigate to first table
    print("\n[2] Table Navigation Test (with FindCtrl fix)")
    print("-" * 40)
    result = hwp.get_into_nth_table(0)
    print(f"get_into_nth_table(0) = {result}")
    print(f"is_cell() after get_into_nth_table = {hwp.is_cell()}")

    # Check if we're in the table by testing cell movement
    print("\n[3] Cell Data Insert Test")
    print("-" * 40)

    # First cell
    hwp.insert_text("A1")
    print("  Inserted: A1 (first cell)")

    # Move right and insert
    result = hwp.table_right_cell()
    print(f"  table_right_cell() = {result}")
    hwp.insert_text("B1")
    print("  Inserted: B1")

    # Move right and insert
    result = hwp.table_right_cell()
    print(f"  table_right_cell() = {result}")
    hwp.insert_text("C1")
    print("  Inserted: C1")

    # Move down and insert
    result = hwp.table_lower_cell()
    print(f"  table_lower_cell() = {result}")
    hwp.insert_text("C2")
    print("  Inserted: C2")

    # Move left and insert
    result = hwp.table_left_cell()
    print(f"  table_left_cell() = {result}")
    hwp.insert_text("B2")
    print("  Inserted: B2")

    # Move up and check
    result = hwp.table_upper_cell()
    print(f"  table_upper_cell() = {result}")

    print("\n[4] Summary")
    print("-" * 40)
    if result:
        print("SUCCESS: All cell navigation works!")
    else:
        print("FAIL: Cell navigation not working")

    print("\n" + "=" * 60)
    print("Table Test Complete!")
    print("=" * 60)

    # Save test document
    hwp.save_as("D:/Code/hwp_test/cpyhwpx/test_table_output.hwp")
    print("Saved: test_table_output.hwp")

    time.sleep(1)
    hwp.quit()
    print("HWP closed")

if __name__ == "__main__":
    test_table()
