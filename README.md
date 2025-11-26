# 📊 Survey Data Analysis Automation (설문 데이터 분석 자동화)

## Project Overview
이 프로젝트는 구글 폼 등을 통해 수집된 설문 데이터(CSV/Excel)를 자동으로 전처리하고, 기술통계량 및 신뢰도 분석(Cronbach's Alpha) 결과를 포함한 엑셀 보고서를 생성하는 파이썬 스크립트입니다.

## Features
- **Data Cleaning**: 한글 텍스트 응답을 5점 척도(1~5) 숫자로 자동 변환
- **Descriptive Statistics**: 주요 변수의 평균, 중위값, 표준편차 자동 산출
- **Reliability Analysis**: 크론바흐 알파(Cronbach's Alpha) 계수 자동 계산 및 수식 포함
- **Excel Report**: `xlsxwriter`를 사용하여 수식(=AVERAGE, =VAR.S 등)이 살아있는 엑셀 파일 생성

## How to Use
1. 필요한 라이브러리 설치
   ```bash
   pip install -r requirements.txt

  - 원본 데이터(input.xlsx)를 폴더에 위치 *이름을 반드시 input.xlsx로 바꿔줘야 함..

스크립트 실행
Bash

python excel.py

결과 파일(Final_Submission.xlsx) 확인
