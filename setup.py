# -*- coding: utf-8 -*-
"""
cpyhwpx - C++ HWP Automation Library

32-bit/64-bit Python 모두 지원하는 HWP 자동화 라이브러리
- 32-bit Python: 직접 cpyhwpx.pyd 로드
- 64-bit Python: subprocess 브릿지를 통해 32-bit 서버 사용
"""

import os
import sys
from setuptools import setup, find_packages

# README 읽기
def read_readme():
    readme_path = os.path.join(os.path.dirname(__file__), "README.md")
    if os.path.exists(readme_path):
        with open(readme_path, encoding="utf-8") as f:
            return f.read()
    return ""

# 패키지 데이터 (pyd 파일 포함)
package_data = {
    "cpyhwpx": [
        "_native/*.pyd",
        "_native/*.py",
    ],
}

setup(
    name="cpyhwpx",
    version="1.0.0",
    author="cpyhwpx Contributors",
    author_email="",
    description="C++ HWP Automation Library - 32-bit/64-bit Python 지원",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/example/cpyhwpx",

    # 패키지 설정
    packages=find_packages(exclude=["tests", "tests.*", "docs", "build", "build32"]),
    package_data=package_data,
    include_package_data=True,

    # Python 버전
    python_requires=">=3.8",

    # 의존성 없음 (표준 라이브러리만 사용)
    install_requires=[],

    # 개발 의존성
    extras_require={
        "dev": [
            "pytest>=7.0",
            "pytest-cov>=4.0",
        ],
    },

    # 분류
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: C++",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],

    keywords="hwp, hangul, word processor, automation, com, pyhwpx",

    # 기타 옵션
    zip_safe=False,

    # 진입점 (필요시)
    entry_points={
        # "console_scripts": [
        #     "cpyhwpx=cpyhwpx.cli:main",
        # ],
    },
)
