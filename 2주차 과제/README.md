# 혈액 기부 서비스 센터 데이터 탐색적 분석 (EDA)

> **과목**: Statistics for Machine Learning — 2주차 과제  
> **데이터셋**: [OpenML - blood-transfusion-service-center (ID: 1464)](https://www.openml.org/d/1464)  
> **작성일**: 2026-03-22

---

## 프로젝트 개요

대만 신주시 헌혈 센터의 헌혈자 데이터를 기반으로, **다음 헌혈 캠페인 시 실제로 헌혈할 가능성이 높은 헌혈자를 식별**하기 위한 탐색적 데이터 분석(EDA)을 수행한 프로젝트이다. RFMTC(Recency, Frequency, Monetary, Time, Class) 모델에 기반한 헌혈 행동 데이터를 다각도로 분석하여, 헌혈 센터의 효율적인 헌혈자 관리 및 마케팅 전략 수립에 기여하는 것을 목표로 한다.

---

## 데이터셋 정보

| 항목 | 내용 |
|------|------|
| 출처 | OpenML (ID: 1464), UCI Machine Learning Repository |
| 수집 주체 | 대만 신주시 헌혈 센터 |
| 원본 행 수 | 748건 |
| 분석 행 수 | 533건 (중복 215건 제거 후) |
| 열 수 | 5개 (특성 4개 + 타겟 1개) |

### 변수 설명

| 변수 | 의미 | 설명 |
|------|------|------|
| V1 (Recency) | 마지막 헌혈 이후 경과 개월 | 값이 작을수록 최근 헌혈 |
| V2 (Frequency) | 총 헌혈 횟수 | 헌혈 빈도 |
| V3 (Monetary) | 총 헌혈량 (c.c.) | V2 x 250 |
| V4 (Time) | 첫 헌혈 이후 경과 개월 | 헌혈 이력 기간 |
| Class | 2007.3 헌혈 여부 | 1=미헌혈, 2=헌혈 |

---

## 프로젝트 구조

```
2주차 과제/
├── README.md                              # 프로젝트 설명 문서 (현재 파일)
├── scripts/
│   ├── eda_report.py                      # EDA 분석 전체 파이프라인 스크립트
│   └── create_ppt.py                      # EDA 보고서 → 발표용 PPT 변환 스크립트
└── output/
    ├── blood_transfusion_raw.csv          # 원본 데이터 (CSV)
    ├── docs/
    │   ├── eda_report.md                  # EDA 분석 결과 보고서 (마크다운)
    │   ├── 2455027 천영진.pptx             # 생성된 발표 자료
    │   └── 1_statistics for machine learning.pdf  # 강의 참고 자료
    ├── figures/                           # 시각화 결과물 (21개 PNG 파일)
    │   ├── 01_histogram_kde.png           # 히스토그램 + KDE
    │   ├── 02_boxplot.png                 # 박스플롯
    │   ├── 03_class_distribution.png      # 클래스 분포 (막대/원 그래프)
    │   ├── 04_binning.png                 # 범주 재그룹화 (Binning)
    │   ├── 05_transformations.png         # 변환 비교 (원본/로그/루트/박스-콕스)
    │   ├── 05b_yeojohnson.png             # Yeo-Johnson 변환
    │   ├── 06_outliers.png                # 이상치 시각화
    │   ├── 07_pearson_heatmap.png         # Pearson 상관계수 히트맵
    │   ├── 08_spearman_heatmap.png        # Spearman 상관계수 히트맵
    │   ├── 09_covariance_heatmap.png      # 공분산 행렬 히트맵
    │   ├── 10_pairplot.png                # Pair Plot (산점도 행렬)
    │   ├── 10b_violin_overall.png         # 바이올린 플롯
    │   ├── 10c_scatter_trendline.png      # 산점도 + 추세선
    │   ├── 11_class_comparison.png        # Class별 Box + Violin 비교
    │   ├── 11b_segment_comparison.png     # VIP vs 일반 세그먼트 비교
    │   ├── 12_scaling_comparison.png      # 스케일링 비교
    │   ├── 13_cdf_lineplot.png            # CDF 선 그래프
    │   ├── 14_class_profile_lineplot.png  # Class별 변수 프로파일
    │   ├── pdf_page_48.png                # 강의 자료 캡처 (p.48)
    │   ├── pdf_page_49.png                # 강의 자료 캡처 (p.49)
    │   └── pdf_page_50.png                # 강의 자료 캡처 (p.50)
    └── tables/                            # 통계 테이블 (CSV)
        ├── descriptive_statistics.csv     # 기술 통계량
        ├── segment_comparison.csv         # 세그먼트별 비교
        └── transformation_comparison.csv  # 변환 효과 비교
```

---

## 분석 파이프라인

`scripts/eda_report.py`는 아래 8단계의 분석을 순차적으로 수행한다.

| 단계 | 내용 | 주요 기법 |
|------|------|-----------|
| 1 | 데이터 로드 | 4가지 방식 (openml, fetch_openml, StringIO, requests) |
| 2 | 데이터 프로파일링 | 명세 확인, 결측치 분석, 중복 제거, 기초 통계량 산출 |
| 3 | 단변량 분석 | 왜도/첨도 해석, 정규성 검정(Shapiro-Wilk), 히스토그램+KDE, 박스플롯, 클래스 분포, Binning |
| 4 | 데이터 변환 | 로그/루트/박스-콕스/Yeo-Johnson 변환 비교 |
| 5 | 이상치 탐지 | Z-score, Modified Z-score(MAD), IQR 방식 |
| 6 | 다변량 분석 | Pearson/Spearman 상관계수, 공분산, VIF, Pair Plot, 산점도+추세선, Class별/세그먼트별 비교 |
| 7 | 스케일링 | 표준화(StandardScaler), 정규화(MinMaxScaler) |
| 8 | 추가 시각화 | CDF 선 그래프, Class별 프로파일 플롯 |

---

## 주요 분석 결과

### 가설 검증

| 가설 | 결과 | 근거 |
|------|------|------|
| H1: 최근 헌혈자가 재헌혈 | **채택** | V1-Class Spearman -0.298 |
| H2: 고빈도 헌혈자가 재헌혈 | **채택** | V2-Class Pearson 0.175, VIP 세그먼트 분석 |
| H3: V3 = V2 x 250 | **채택** | Pearson/Spearman 1.000, VIF = inf |

### 핵심 인사이트

1. **V1(Recency)이 가장 강력한 예측 변수** — Class와 Spearman -0.298
2. **V2-V3 완벽한 다중공선성** — VIF = inf, 모델링 시 V3 제거 필수
3. **클래스 불균형 존재** — 헌혈 28.0% vs 미헌혈 72.0%, SMOTE 등 오버샘플링 필요
4. **극도의 양의 왜도** — V1~V3 모두 강한 양의 왜도(2.3~3.0), 박스-콕스 변환 권장
5. **V4의 낮은 예측력** — Class와 Spearman -0.142, 오랜 이력이 재헌혈을 보장하지 않음

---

## 실행 방법

### 사전 요구사항

- Python 3.8 이상
- 필요 라이브러리:

```bash
pip install pandas numpy matplotlib seaborn scipy scikit-learn openml requests statsmodels
```

### EDA 분석 실행

```bash
cd "2주차 과제"
python scripts/eda_report.py
```

실행 시 `output/figures/` 및 `output/tables/` 디렉토리에 시각화/통계 결과물이 자동 생성된다.

### 발표 자료(PPT) 생성

```bash
pip install python-pptx
python scripts/create_ppt.py
```

실행 시 프로젝트 루트에 발표용 `.pptx` 파일이 생성된다. 21장의 슬라이드로 구성되며, EDA 분석 결과를 시각적으로 정리한 발표 자료이다.

---

## 기술 스택

| 구분 | 기술 |
|------|------|
| 언어 | Python 3 |
| 데이터 처리 | pandas, numpy |
| 시각화 | matplotlib, seaborn |
| 통계 분석 | scipy (Shapiro-Wilk, Box-Cox, Yeo-Johnson, Z-score) |
| 머신러닝 유틸 | scikit-learn (StandardScaler, MinMaxScaler, fetch_openml) |
| 다중공선성 | statsmodels (VIF) |
| 데이터 소스 | openml, requests |
| 발표 자료 | python-pptx |
