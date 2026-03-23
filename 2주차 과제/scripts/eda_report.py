# -*- coding: utf-8 -*-
"""
============================================================
  혈액 기부 서비스 센터 데이터 탐색적 분석(EDA) 보고서
  대상: OpenML - blood-transfusion-service-center
  과목: Statistics for Machine Learning (2주차 과제)
============================================================
"""

# ===== 단계 0-4: 공통 import 및 설정 =====
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # 비대화형 백엔드 (plt.show() 대신 파일 저장)
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats
from scipy.stats import shapiro, zscore, boxcox, yeojohnson
from sklearn.datasets import fetch_openml
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from io import StringIO
import requests
import warnings
import os

warnings.filterwarnings('ignore')

# 폰트 설정 (영문 기본 폰트 사용 - 한글 깨짐 방지)
plt.rcParams['axes.unicode_minus'] = False
sns.set_style('whitegrid')

os.makedirs('output/figures', exist_ok=True)
os.makedirs('output/tables', exist_ok=True)

print("=" * 60)
print("EDA 분석 시작")
print("=" * 60)

# ==============================================================
# 단계 1: 데이터 로드 (4가지 방식)
# ==============================================================
print("\n" + "=" * 60)
print("[단계 1] 데이터 로드 (4가지 방식)")
print("=" * 60)

# --- 단계 1-1: openml 모듈 사용 ---
print("\n--- 1-1: openml 모듈 ---")
try:
    import openml
    dataset = openml.datasets.get_dataset(1464)  # blood-transfusion ID=1464
    X_oml, y_oml, _, attr_names = dataset.get_data(
        target=dataset.default_target_attribute
    )
    df_openml = pd.concat([X_oml, y_oml], axis=1)
    print(f"  성공 - 크기: {df_openml.shape}")
    print(f"  컬럼: {df_openml.columns.tolist()}")
except Exception as e:
    print(f"  openml 로드 실패: {e}")
    df_openml = None

# --- 단계 1-2: fetch_openml (scikit-learn) ---
print("\n--- 1-2: fetch_openml ---")
data = fetch_openml(name='blood-transfusion-service-center', version=1, as_frame=True)
df_fetch = data.frame
print(f"  성공 - 크기: {df_fetch.shape}")
print(f"  컬럼: {df_fetch.columns.tolist()}")

# --- 단계 1-3: StringIO 활용 (URL fallback 포함) ---
print("\n--- 1-3: StringIO ---")
URLS = [
    "https://www.openml.org/data/download/53796/php0iVrYT",
    "https://www.openml.org/data/get_csv/53796/blood-transfusion-service-center.arff",
    "https://raw.githubusercontent.com/jbrownlee/Datasets/master/transfusion.csv",
]

df_stringio = None
for url in URLS:
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        df_stringio = pd.read_csv(StringIO(response.text))
        print(f"  성공 URL: {url}")
        break
    except Exception as e:
        print(f"  실패 URL: {url} -> {e}")
        continue

if df_stringio is None:
    csv_text = df_fetch.to_csv(index=False)
    df_stringio = pd.read_csv(StringIO(csv_text))
    print("  Fallback: fetch_openml -> StringIO")
print(f"  크기: {df_stringio.shape}")

# --- 단계 1-4: requests 라이브러리 ---
print("\n--- 1-4: requests ---")
df_requests = None
for url in URLS:
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        with open("output/blood_transfusion_raw.csv", "w", encoding="utf-8") as f:
            f.write(response.text)
        df_requests = pd.read_csv("output/blood_transfusion_raw.csv")
        print(f"  성공 URL: {url}")
        break
    except Exception as e:
        print(f"  실패 URL: {url} -> {e}")
        continue

if df_requests is None:
    df_fetch.to_csv("output/blood_transfusion_raw.csv", index=False)
    df_requests = pd.read_csv("output/blood_transfusion_raw.csv")
    print("  Fallback: fetch_openml -> CSV 저장")
print(f"  크기: {df_requests.shape}")

# --- 단계 1-5: 비교 ---
print("\n--- 1-5: 로드 결과 비교 ---")
load_results = {
    'openml': df_openml.shape if df_openml is not None else "실패",
    'fetch_openml': df_fetch.shape,
    'StringIO': df_stringio.shape,
    'requests': df_requests.shape,
}
for method, shape in load_results.items():
    print(f"  {method}: {shape}")

# 데이터 내용 일관성 검증
dfs_to_compare = [('fetch_openml', df_fetch), ('StringIO', df_stringio), ('requests', df_requests)]
if df_openml is not None:
    dfs_to_compare.insert(0, ('openml', df_openml))
ref_name, ref_df = dfs_to_compare[0]
for name, check_df in dfs_to_compare[1:]:
    if ref_df.shape == check_df.shape:
        print(f"  {ref_name} vs {name}: shape 일치 [OK]")
    else:
        print(f"  {ref_name} vs {name}: shape 불일치 ({ref_df.shape} vs {check_df.shape})")

# ==============================================================
# 단계 2: 데이터 프로파일링
# ==============================================================
print("\n" + "=" * 60)
print("[단계 2] 데이터 프로파일링")
print("=" * 60)

# 이후 분석에는 fetch_openml 데이터 사용
df = df_fetch.copy()

# 수치형으로 변환 (fetch_openml이 category로 반환할 수 있음)
for col in df.columns:
    if df[col].dtype.name == 'category':
        # 카테고리형: 숫자로 파싱 가능하면 변환, 아니면 cat.codes 사용
        try:
            df[col] = pd.to_numeric(df[col].astype(str), errors='raise')
        except (ValueError, TypeError):
            df[col] = df[col].cat.codes
    else:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# --- 2-1: 데이터 명세 ---
print("\n--- 2-1: 데이터 명세 ---")
print(f"  형태: {df.shape}")
print(f"  컬럼: {df.columns.tolist()}")
print(f"  데이터 타입:\n{df.dtypes}")
print(f"\n처음 5행:\n{df.head()}")

# Class 변수 인코딩 동적 확인
class_candidates = [c for c in df.columns if c.lower() == 'class']
if class_candidates:
    class_col = class_candidates[0]
else:
    # fallback: fetch_openml의 target 이름 또는 마지막 컬럼
    class_col = getattr(data, 'target_names', [df.columns[-1]])[0] if hasattr(data, 'target_names') and data.target_names else df.columns[-1]
    print(f"  경고: 'Class' 컬럼 미발견, 대체 사용: {class_col}")
class_unique = sorted(df[class_col].dropna().unique())
print(f"\n  Class 고유값: {class_unique}")

# Majority/Minority 클래스 검증 (라벨 매핑 정확성 확보)
majority_class = df[class_col].value_counts().idxmax()
minority_class = df[class_col].value_counts().idxmin()
print(f"  Majority class: {majority_class} ({df[class_col].value_counts().max()}건)")
print(f"  Minority class: {minority_class} ({df[class_col].value_counts().min()}건)")

if set(class_unique) == {0, 1} or set(class_unique) == {0.0, 1.0}:
    CLASS_LABELS = {str(int(class_unique[0])): 'Not Donated', str(int(class_unique[1])): 'Donated'}
    CLASS_COLORS = {str(int(class_unique[0])): '#2196F3', str(int(class_unique[1])): '#FF5722'}
elif set(class_unique) == {1, 2} or set(class_unique) == {1.0, 2.0}:
    # OpenML blood-transfusion: majority=미헌혈, minority=헌혈 (데이터 기반 매핑)
    CLASS_LABELS = {str(int(majority_class)): 'Not Donated', str(int(minority_class)): 'Donated'}
    CLASS_COLORS = {str(int(majority_class)): '#2196F3', str(int(minority_class)): '#FF5722'}
    print(f"  검증: Majority({int(majority_class)})='Not Donated', Minority({int(minority_class)})='Donated'")
else:
    CLASS_LABELS = {str(v): str(v) for v in class_unique}
    palette = ['#2196F3', '#FF5722', '#4CAF50', '#FFC107', '#9C27B0']
    CLASS_COLORS = {str(int(v)): palette[i % len(palette)] for i, v in enumerate(class_unique)}
print(f"  Label mapping: {CLASS_LABELS}")

# feature_cols: 결측치 처리 및 이후 분석에서 공통으로 사용 (class 제외 수치형 컬럼)
numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
feature_cols = [c for c in numeric_cols if c.lower() != 'class']

# --- 2-2: 데이터 정제 (중복 제거를 결측치 처리보다 먼저 수행) ---
print("\n--- 2-2: 데이터 정제 ---")
duplicates = df.duplicated().sum()
print(f"  중복 행 수: {duplicates}")
if duplicates > 0:
    df = df.drop_duplicates().reset_index(drop=True)
    print(f"  중복 제거 후 행 수: {len(df)}")

# --- 2-3: 결측치 분석 ---
print("\n--- 2-3: 결측치 분석 ---")
missing = df.isnull().sum()
missing_pct = (df.isnull().sum() / len(df)) * 100
missing_report = pd.DataFrame({'결측치 수': missing, '결측치 비율(%)': missing_pct})
print(missing_report)

if missing.sum() > 0:
    df[feature_cols] = df[feature_cols].fillna(df[feature_cols].mean())
    print("  -> 피처 컬럼 평균값으로 결측치 대체 완료 (Class 컬럼 제외)")
else:
    print("  -> 결측치 없음")

# --- 2-4: 기초 통계량 ---
print("\n--- 2-4: 기초 통계량 ---")
print(df.describe())

stats_table = pd.DataFrame()
for col in feature_cols:
    stats_table[col] = pd.Series({
        '평균': df[col].mean(),
        '중앙값': df[col].median(),
        '최빈값': df[col].mode().values[0],  # 다중 최빈값 시 첫 번째 사용
        '최빈값 개수': len(df[col].mode()),
        '범위': df[col].max() - df[col].min(),
        'Q1': df[col].quantile(0.25),
        'Q2': df[col].quantile(0.50),
        'Q3': df[col].quantile(0.75),
        'IQR': df[col].quantile(0.75) - df[col].quantile(0.25),
        '90% 백분위수': df[col].quantile(0.90),
        '분산': df[col].var(),
        '표준편차': df[col].std(),
        '왜도': df[col].skew(),
        '첨도': df[col].kurtosis()
    })

stats_table.to_csv("output/tables/descriptive_statistics.csv", encoding="utf-8-sig")
print("  -> output/tables/descriptive_statistics.csv 저장 완료")

for col in feature_cols:
    print(f"\n  [{col}]")
    print(f"    평균={df[col].mean():.4f}, 중앙값={df[col].median():.4f}, 최빈값={df[col].mode().values[0]}")
    print(f"    범위={df[col].max() - df[col].min()}, Q1={df[col].quantile(0.25):.4f}, Q3={df[col].quantile(0.75):.4f}")
    print(f"    IQR={df[col].quantile(0.75) - df[col].quantile(0.25):.4f}, 90%={df[col].quantile(0.90):.4f}")
    print(f"    분산={df[col].var():.4f}, 표준편차={df[col].std():.4f}")
    print(f"    왜도={df[col].skew():.4f}, 첨도={df[col].kurtosis():.4f}")

# ==============================================================
# 단계 3: 변수별 개별 특성 분석
# ==============================================================
print("\n" + "=" * 60)
print("[단계 3] 변수별 개별 특성 분석")
print("=" * 60)

# --- 3-1: 왜도/첨도 해석 ---
print("\n--- 3-1: 왜도/첨도 해석 ---")
for col in feature_cols:
    skewness = df[col].skew()
    kurt = df[col].kurtosis()
    if skewness > 1: s_i = "강한 양의 왜도"
    elif skewness > 0.5: s_i = "중간 양의 왜도"
    elif skewness > -0.5: s_i = "대칭에 가까움"
    elif skewness > -1: s_i = "중간 음의 왜도"
    else: s_i = "강한 음의 왜도"
    # pandas .kurtosis()는 초과 첨도(excess kurtosis)를 반환 (정규분포 기준 = 0)
    if kurt > 1: k_i = "뾰족(이상치 가능)"
    elif kurt > -1: k_i = "정규분포 유사"
    else: k_i = "편평"
    print(f"  {col}: 왜도={skewness:.3f}({s_i}), 첨도={kurt:.3f}({k_i})")

# --- 3-2: 정규성 검정 ---
print("\n--- 3-2: 정규성 검정 (Shapiro-Wilk) ---")
for col in feature_cols:
    sample = df[col].dropna()
    if len(sample) > 5000:
        sample = sample.sample(5000, random_state=42)
    stat, p_val = shapiro(sample)
    result = "정규분포 따름" if p_val > 0.05 else "정규분포 아님"
    print(f"  {col}: W={stat:.4f}, p={p_val:.6f} -> {result}")

# --- 3-3: 히스토그램 + KDE ---
print("\n--- 3-3: 히스토그램 + KDE ---")
n_features = len(feature_cols)
n_cols_hist = 2
n_rows_hist = (n_features + n_cols_hist - 1) // n_cols_hist
fig, axes = plt.subplots(n_rows_hist, n_cols_hist, figsize=(14, 5 * n_rows_hist))
axes = axes.flatten()
for i, col in enumerate(feature_cols):
    ax = axes[i]
    ax.hist(df[col], bins=30, density=True, alpha=0.7, color='steelblue', edgecolor='white')
    df[col].plot.kde(ax=ax, color='red', linewidth=2)
    ax.axvline(df[col].mean(), color='green', linestyle='--', label=f'Mean: {df[col].mean():.1f}')
    ax.axvline(df[col].median(), color='orange', linestyle='--', label=f'Median: {df[col].median():.1f}')
    mode_val = df[col].mode().values[0]
    ax.axvline(mode_val, color='blue', linestyle=':', linewidth=2, label=f'Mode: {mode_val:.0f}')
    ax.set_title(f'{col} - Histogram + KDE', fontsize=13)
    ax.set_xlabel(col); ax.set_ylabel('Density')
    ax.legend(fontsize=8)
for j in range(n_features, n_rows_hist * n_cols_hist):
    axes[j].set_visible(False)
plt.tight_layout()
plt.savefig("output/figures/01_histogram_kde.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/01_histogram_kde.png")

# --- 3-4: 박스플롯 ---
print("\n--- 3-4: 박스플롯 ---")
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(16, 5))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    sns.boxplot(y=df[col], ax=axes[i], color='lightcoral')
    axes[i].set_title(f'{col} Boxplot', fontsize=12)
for j in range(len(feature_cols), len(axes)):
    axes[j].set_visible(False)
plt.tight_layout()
plt.savefig("output/figures/02_boxplot.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/02_boxplot.png")

# --- 3-5: 범주형(Class) 분석 ---
print("\n--- 3-5: 범주형 변수 분석 (Class) ---")
class_counts = df[class_col].value_counts()
class_pct = df[class_col].value_counts(normalize=True) * 100
print(f"  빈도:\n{class_counts}")
print(f"  비율(%):\n{class_pct}")

fig, axes = plt.subplots(1, 2, figsize=(12, 5))
class_counts.plot(kind='bar', ax=axes[0],
                  color=[CLASS_COLORS.get(str(int(k)), '#999') for k in class_counts.index],
                  edgecolor='white')
axes[0].set_title('Class Frequency (Bar Chart)', fontsize=13)
axes[0].set_xlabel('Class'); axes[0].set_ylabel('Count')
axes[0].set_xticklabels([CLASS_LABELS.get(str(int(k)), str(k)) for k in class_counts.index], rotation=0)

axes[1].pie(class_counts,
            labels=[CLASS_LABELS.get(str(int(k)), str(k)) for k in class_counts.index],
            autopct='%1.1f%%',
            colors=[CLASS_COLORS.get(str(int(k)), '#999') for k in class_counts.index],
            startangle=90)
axes[1].set_title('Class Proportion (Pie Chart)', fontsize=13)
plt.tight_layout()
plt.savefig("output/figures/03_class_distribution.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/03_class_distribution.png")

# --- 3-5b: 희소 클래스(Rare Labels) 분석 ---
print("\n--- 3-5b: 희소 클래스(Rare Labels) 확인 ---")
rare_threshold = 0.05  # 5% 미만이면 희소 클래스
for cls_val in class_counts.index:
    pct = class_counts[cls_val] / len(df) * 100
    is_rare = "희소 클래스" if pct < rare_threshold * 100 else "정상"
    label = CLASS_LABELS.get(str(int(cls_val)), str(cls_val))
    print(f"  Class {int(cls_val)} ({label}): {class_counts[cls_val]}건 ({pct:.1f}%) -> {is_rare}")
print(f"  결론: 최소 비율 클래스 = {class_counts.min()/len(df)*100:.1f}% -> ", end='')
if class_counts.min()/len(df) < rare_threshold:
    print("희소 클래스 존재 -> 오버샘플링 고려 필요")
else:
    print(f"희소 클래스 없음 (기준: {rare_threshold*100}% 미만). 단, 클래스 불균형({class_counts.min()/len(df)*100:.1f}% vs {class_counts.max()/len(df)*100:.1f}%)은 존재.")

# --- 3-6: 파생변수 생성 (Binning) ---
print("\n--- 3-6: 파생변수 (Binning) ---")
recency_col = [c for c in feature_cols if 'recency' in c.lower()][0] if any('recency' in c.lower() for c in feature_cols) else feature_cols[0]
freq_col = [c for c in feature_cols if 'frequency' in c.lower()][0] if any('frequency' in c.lower() for c in feature_cols) else feature_cols[1]

recency_bins = [-1, 5, 10, 20, 50, max(50, df[recency_col].max()) + 1]
recency_labels = ['0-5', '6-10', '11-20', '21-50', f'51-{int(df[recency_col].max())}']
df['Recency_Bin'] = pd.cut(df[recency_col], bins=recency_bins, labels=recency_labels)

freq_bins = [-1, 5, 10, 20, 50, max(50, df[freq_col].max()) + 1]
freq_labels = ['0-5', '6-10', '11-20', '21-50', f'51-{int(df[freq_col].max())}']
df['Frequency_Bin'] = pd.cut(df[freq_col], bins=freq_bins, labels=freq_labels)
print(f"  Recency_Bin NaN: {df['Recency_Bin'].isna().sum()}")
print(f"  Frequency_Bin NaN: {df['Frequency_Bin'].isna().sum()}")

fig, axes = plt.subplots(1, 2, figsize=(14, 5))
df['Recency_Bin'].value_counts().sort_index().plot(kind='bar', ax=axes[0], color='mediumpurple')
axes[0].set_title(f'{recency_col} Binned Frequency', fontsize=12)
axes[0].set_xlabel('Bin'); axes[0].set_ylabel('Count')
df['Frequency_Bin'].value_counts().sort_index().plot(kind='bar', ax=axes[1], color='mediumseagreen')
axes[1].set_title(f'{freq_col} Binned Frequency', fontsize=12)
axes[1].set_xlabel('Bin'); axes[1].set_ylabel('Count')
plt.tight_layout()
plt.savefig("output/figures/04_binning.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/04_binning.png")

# Bin 파생변수 정리 (이후 분석에 영향 주지 않도록)
df.drop(['Recency_Bin', 'Frequency_Bin'], axis=1, inplace=True)

# ==============================================================
# 단계 4: 왜곡 변수 전처리 및 변환
# ==============================================================
print("\n" + "=" * 60)
print("[단계 4] 데이터 변환")
print("=" * 60)

# --- 4-1: 변환 비교 (원본/로그/루트/박스-콕스) ---
print("\n--- 4-1: 변환 비교 ---")
fig, axes = plt.subplots(max(len(feature_cols), 2), 4, figsize=(20, max(len(feature_cols), 2)*4))
if len(feature_cols) == 1:
    axes = [axes[0]]  # 첫 행만 사용
for i, col in enumerate(feature_cols):
    col_data = df[col].dropna()
    axes[i][0].hist(col_data, bins=30, color='steelblue', alpha=0.7, edgecolor='white')
    axes[i][0].set_title(f'{col} - Original (Skew: {col_data.skew():.3f})')

    log_data = np.log1p(col_data)
    axes[i][1].hist(log_data, bins=30, color='coral', alpha=0.7, edgecolor='white')
    axes[i][1].set_title(f'{col} - Log Transform (Skew: {log_data.skew():.3f})')

    sqrt_data = np.sqrt(np.maximum(col_data, 0))
    axes[i][2].hist(sqrt_data, bins=30, color='mediumseagreen', alpha=0.7, edgecolor='white')
    axes[i][2].set_title(f'{col} - Sqrt Transform (Skew: {sqrt_data.skew():.3f})')

    bc_shift = max(0, -col_data.min() + 1) if col_data.min() <= 0 else 0
    bc_data, lam = boxcox(col_data + bc_shift)
    axes[i][3].hist(bc_data, bins=30, color='mediumpurple', alpha=0.7, edgecolor='white')
    axes[i][3].set_title(f'{col} - Box-Cox (lambda={lam:.3f}, Skew: {pd.Series(bc_data).skew():.3f})')
plt.tight_layout()
plt.savefig("output/figures/05_transformations.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/05_transformations.png")

# Yeo-Johnson 변환
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(18, 4))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    col_data = df[col].dropna()
    yj_data, yj_lam = yeojohnson(col_data)
    axes[i].hist(yj_data, bins=30, color='teal', alpha=0.7, edgecolor='white')
    axes[i].set_title(f'{col} Yeo-Johnson\n(lambda={yj_lam:.3f}, Skew: {pd.Series(yj_data).skew():.3f})')
plt.tight_layout()
plt.savefig("output/figures/05b_yeojohnson.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/05b_yeojohnson.png")

# --- 4-2: 변환 효과 비교 테이블 ---
print("\n--- 4-2: 변환 효과 비교 ---")
transform_comparison = []
for col in feature_cols:
    col_data = df[col].dropna()
    row = {
        '변수': col,
        '원본 왜도': col_data.skew(),
        '로그변환 왜도': np.log1p(col_data).skew(),
        '루트변환 왜도': np.sqrt(np.maximum(col_data, 0)).skew(),
    }
    bc_shift2 = max(0, -col_data.min() + 1) if col_data.min() <= 0 else 0
    bc_data, _ = boxcox(col_data + bc_shift2)
    row['박스콕스 왜도'] = pd.Series(bc_data).skew()
    yj_data, _ = yeojohnson(col_data)
    row['Yeo-Johnson 왜도'] = pd.Series(yj_data).skew()
    skews = {k: abs(v) for k, v in row.items() if '왜도' in k and v is not None}
    row['최적 변환'] = min(skews, key=skews.get).replace(' 왜도', '')
    transform_comparison.append(row)

transform_df = pd.DataFrame(transform_comparison)
transform_df.to_csv("output/tables/transformation_comparison.csv", encoding="utf-8-sig", index=False)
print(transform_df.to_string(index=False))

# ==============================================================
# 단계 5: 이상치 탐지
# ==============================================================
print("\n" + "=" * 60)
print("[단계 5] 이상치 탐지")
print("=" * 60)

# --- 5-1: Z-score ---
print("\n--- 5-1: Z-score (|Z| > 3) ---")
print("  [주의] Z-score는 정규분포 가정. 비정규 데이터에서는 Modified Z-score(MAD)가 더 강건합니다.")
for col in feature_cols:
    valid_data = df[col].dropna()
    z = zscore(valid_data)
    outliers = (np.abs(z) > 3).sum()
    print(f"  {col}: 이상치 {outliers}개 ({outliers/len(valid_data)*100:.1f}%)")

# --- 5-1b: Modified Z-score (MAD 기반, 비정규분포에 강건) ---
print("\n--- 5-1b: Modified Z-score (MAD 기반, |M| > 3.5) ---")
for col in feature_cols:
    valid_data = df[col].dropna()
    median_val = valid_data.median()
    mad = np.median(np.abs(valid_data - median_val))
    if mad == 0:
        print(f"  {col}: MAD=0 (값이 중앙값에 집중 - Modified Z-score 계산 불가)")
        continue
    modified_z = 0.6745 * (valid_data - median_val) / mad
    outliers_mz = (np.abs(modified_z) > 3.5).sum()
    # 기존 Z-score와 비교
    z = zscore(valid_data)
    outliers_z = (np.abs(z) > 3).sum()
    diff_note = f" (차이: {abs(outliers_mz - outliers_z)}개)" if outliers_mz != outliers_z else ""
    print(f"  {col}: Modified Z-score 이상치={outliers_mz}개 vs Z-score={outliers_z}개{diff_note}")

# --- 5-2: IQR ---
print("\n--- 5-2: IQR ---")
for col in feature_cols:
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR = Q3 - Q1
    lower, upper = Q1 - 1.5*IQR, Q3 + 1.5*IQR
    out_lower = (df[col] < lower).sum()
    out_upper = (df[col] > upper).sum()
    print(f"  {col}: Q1={Q1}, Q3={Q3}, IQR={IQR}, 범위=[{lower:.1f}, {upper:.1f}], 하한이상치={out_lower}, 상한이상치={out_upper}, 총={out_lower+out_upper}")

# --- 5-3: 이상치 시각화 ---
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(16, 5))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    Q1 = df[col].quantile(0.25)
    Q3 = df[col].quantile(0.75)
    IQR_val = Q3 - Q1
    colors = ['red' if (x < Q1 - 1.5*IQR_val or x > Q3 + 1.5*IQR_val) else 'steelblue'
              for x in df[col]]
    axes[i].scatter(range(len(df)), df[col], c=colors, alpha=0.5, s=10)
    axes[i].axhline(Q1 - 1.5*IQR_val, color='red', linestyle='--', alpha=0.5)
    axes[i].axhline(Q3 + 1.5*IQR_val, color='red', linestyle='--', alpha=0.5)
    axes[i].set_title(f'{col} Outliers', fontsize=11)
plt.tight_layout()
plt.savefig("output/figures/06_outliers.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/06_outliers.png")

# ==============================================================
# 단계 6: 상관관계 및 관계 분석
# ==============================================================
print("\n" + "=" * 60)
print("[단계 6] 상관관계 분석")
print("=" * 60)

numeric_df = df[feature_cols + [class_col]].select_dtypes(include=[np.number])

# --- 6-1: Pearson ---
print("\n--- 6-1: Pearson 상관계수 ---")
pearson_corr = numeric_df.corr(method='pearson')
print(pearson_corr)
plt.figure(figsize=(10, 8))
sns.heatmap(pearson_corr, annot=True, fmt='.3f', cmap='RdBu_r', center=0,
            square=True, linewidths=0.5, vmin=-1, vmax=1)
plt.title('Pearson Correlation Heatmap', fontsize=15)
plt.tight_layout()
plt.savefig("output/figures/07_pearson_heatmap.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/07_pearson_heatmap.png")

# --- 6-2: Spearman ---
print("\n--- 6-2: Spearman 상관계수 ---")
spearman_corr = numeric_df.corr(method='spearman')
print(spearman_corr)
plt.figure(figsize=(10, 8))
sns.heatmap(spearman_corr, annot=True, fmt='.3f', cmap='RdBu_r', center=0,
            square=True, linewidths=0.5, vmin=-1, vmax=1)
plt.title('Spearman Correlation Heatmap', fontsize=15)
plt.tight_layout()
plt.savefig("output/figures/08_spearman_heatmap.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/08_spearman_heatmap.png")

# --- 6-2b: Pearson vs Spearman 해석 ---
print("\n--- 6-2b: Pearson vs Spearman 해석 ---")
print("  [참고] 정규성 검정 결과 대부분 feature가 비정규 분포를 따르므로,")
print("  Spearman 상관계수가 더 신뢰할 수 있는 순위 기반 상관관계를 제공합니다.")
for i_idx in range(len(feature_cols)):
    for j_idx in range(i_idx+1, len(feature_cols)):
        c1, c2 = feature_cols[i_idx], feature_cols[j_idx]
        p_r = pearson_corr.loc[c1, c2]
        s_r = spearman_corr.loc[c1, c2]
        diff = abs(p_r - s_r)
        if diff > 0.1:
            print(f"  {c1}-{c2}: Pearson={p_r:.3f}, Spearman={s_r:.3f} (차이={diff:.3f}) -> 비선형 관계 가능성")

# --- 6-3: 공분산 ---
print("\n--- 6-3: 공분산 행렬 ---")
cov_matrix = numeric_df.cov()
print(cov_matrix)
plt.figure(figsize=(10, 8))
sns.heatmap(cov_matrix, annot=True, fmt='.1f', cmap='YlOrRd', square=True, linewidths=0.5)
plt.title('Covariance Matrix', fontsize=15)
plt.tight_layout()
plt.savefig("output/figures/09_covariance_heatmap.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/09_covariance_heatmap.png")

# --- 6-4: Pair Plot ---
print("\n--- 6-4: 산점도 행렬 (Pair Plot) ---")
pair_df = df[feature_cols + [class_col]].copy()
pair_df[class_col] = pair_df[class_col].astype(int).astype(str)
g = sns.pairplot(pair_df, hue=class_col, diag_kind='kde',
                 palette=CLASS_COLORS, plot_kws={'alpha': 0.5, 's': 20})
g.fig.suptitle('Pairwise Scatter Matrix (by Class)', y=1.02, fontsize=15)
plt.tight_layout()
plt.savefig("output/figures/10_pairplot.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/10_pairplot.png")

# --- 6-4b: 주요 변수 간 산점도 + 추세선 ---
print("\n--- 6-4b: 산점도 + 추세선 ---")
# 주요 변수 쌍 선택 (상관계수가 높은 쌍 + 타겟과 관련 높은 쌍)
scatter_pairs = []
for i_idx in range(len(feature_cols)):
    for j_idx in range(i_idx+1, len(feature_cols)):
        scatter_pairs.append((feature_cols[i_idx], feature_cols[j_idx]))

n_pairs = len(scatter_pairs)
n_scatter_cols = min(3, n_pairs)
n_scatter_rows = (n_pairs + n_scatter_cols - 1) // n_scatter_cols
fig, axes = plt.subplots(n_scatter_rows, n_scatter_cols, figsize=(6 * n_scatter_cols, 5 * n_scatter_rows))
axes = np.array(axes).flatten() if n_pairs > 1 else [axes]
for idx, (x_col, y_col) in enumerate(scatter_pairs):
    ax = axes[idx]
    for cls in class_unique:
        mask = df[class_col] == cls
        ax.scatter(df.loc[mask, x_col], df.loc[mask, y_col], alpha=0.4, s=15,
                   color=CLASS_COLORS.get(str(int(cls)), '#999'),
                   label=CLASS_LABELS.get(str(int(cls)), str(cls)))
    # 추세선 (전체 데이터) + R-squared 적합도 표시
    z = np.polyfit(df[x_col], df[y_col], 1)
    p = np.poly1d(z)
    y_pred = p(df[x_col])
    ss_res = np.sum((df[y_col] - y_pred) ** 2)
    ss_tot = np.sum((df[y_col] - df[y_col].mean()) ** 2)
    r_squared = 1 - (ss_res / ss_tot) if ss_tot != 0 else 0
    x_line = np.linspace(df[x_col].min(), df[x_col].max(), 100)
    # R² < 0.1 이면 회색 점선으로 약한 관계임을 시각적 전달
    line_color = 'k' if r_squared >= 0.1 else 'gray'
    line_alpha = 0.7 if r_squared >= 0.1 else 0.3
    ax.plot(x_line, p(x_line), '--', color=line_color, linewidth=2, alpha=line_alpha,
            label=f'Trend: R\u00b2={r_squared:.3f}')
    ax.set_xlabel(x_col); ax.set_ylabel(y_col)
    ax.set_title(f'{x_col} vs {y_col}', fontsize=11)
    ax.legend(fontsize=8)
# 빈 subplot 숨기기
for idx in range(n_pairs, n_scatter_rows * n_scatter_cols):
    axes[idx].set_visible(False)
plt.suptitle('Scatter Plots with Trend Lines', fontsize=14, y=1.01)
plt.tight_layout()
plt.savefig("output/figures/10c_scatter_trendline.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/10c_scatter_trendline.png")

# --- 6-4c: 다중공선성 확인 ---
print("\n--- 6-4c: 다중공선성 확인 ---")
high_corr_pairs = []
for i_idx in range(len(feature_cols)):
    for j_idx in range(i_idx+1, len(feature_cols)):
        r = pearson_corr.loc[feature_cols[i_idx], feature_cols[j_idx]]
        if abs(r) > 0.7:
            high_corr_pairs.append((feature_cols[i_idx], feature_cols[j_idx], r))
            print(f"  다중공선성 의심: {feature_cols[i_idx]} - {feature_cols[j_idx]} (r={r:.3f})")
if not high_corr_pairs:
    print("  |r| > 0.7 인 변수 쌍 없음 -> 다중공선성 위험 낮음")
else:
    print(f"  결론: {len(high_corr_pairs)}개 변수 쌍에서 다중공선성 의심 -> 모델링 시 변수 제거 고려")

# --- 6-4d: VIF (Variance Inflation Factor) ---
print("\n--- 6-4d: VIF (다중공선성 정밀 확인) ---")
try:
    from statsmodels.stats.outliers_influence import variance_inflation_factor
    if len(feature_cols) >= 2:
        vif_data = df[feature_cols].dropna()
        vif_df = pd.DataFrame({
            '변수': feature_cols,
            'VIF': [variance_inflation_factor(vif_data.values, i) for i in range(len(feature_cols))]
        })
        print(vif_df.to_string(index=False))
        high_vif = vif_df[vif_df['VIF'] > 10]
        if len(high_vif) > 0:
            print(f"  경고: VIF > 10인 변수 {high_vif['변수'].tolist()} -> 다중공선성 심각")
        elif len(vif_df[vif_df['VIF'] > 5]) > 0:
            print(f"  주의: VIF > 5인 변수 존재 -> 다중공선성 가능성")
        else:
            print("  VIF 모두 10 이하 -> 다중공선성 문제 없음")
except ImportError:
    print("  statsmodels 미설치 -> VIF 계산 건너뜀 (pip install statsmodels)")

# --- 6-5: 독립 바이올린 플롯 ---
print("\n--- 6-5: 독립 바이올린 플롯 ---")
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(16, 5))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    sns.violinplot(y=df[col], ax=axes[i], color='lightblue', inner='quartile')
    axes[i].set_title(f'{col} Distribution (Violin)', fontsize=12)
plt.tight_layout()
plt.savefig("output/figures/10b_violin_overall.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/10b_violin_overall.png")

# --- 6-6: Class별 Box + Violin ---
print("\n--- 6-6: Class별 Box + Violin ---")
fig, axes = plt.subplots(2, max(len(feature_cols), 2), figsize=(20, 10))
if len(feature_cols) == 1:
    axes = axes.reshape(-1, 1)
for i, col in enumerate(feature_cols):
    sns.boxplot(x=class_col, y=col, data=df, ax=axes[0][i], palette=CLASS_COLORS)
    axes[0][i].set_title(f'{col} by Class (Box)', fontsize=11)
    sns.violinplot(x=class_col, y=col, data=df, ax=axes[1][i], palette=CLASS_COLORS)
    axes[1][i].set_title(f'{col} by Class (Violin)', fontsize=11)
plt.tight_layout()
plt.savefig("output/figures/11_class_comparison.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/11_class_comparison.png")

# --- 6-7: 세그먼트별 비교 (비즈니스 관점) ---
print("\n--- 6-7: 세그먼트별 비교 (비즈니스 관점) ---")

# Class별 기본 비교
segment_stats = df.groupby(class_col)[feature_cols].agg(['mean', 'median', 'std', 'min', 'max'])
print(segment_stats)
segment_stats.to_csv("output/tables/segment_comparison.csv", encoding="utf-8-sig")
print("  -> output/tables/segment_comparison.csv")

# VIP(고빈도) vs 일반 헌혈자 세그먼트 분석
freq_median = df[freq_col].median()  # freq_col: 동적으로 탐지된 Frequency 컬럼
df['Segment'] = df[freq_col].apply(lambda x: 'VIP' if x > freq_median else 'General')
print(f"\n  세그먼트 기준: {freq_col} > {freq_median} -> VIP")
print(f"  VIP: {(df['Segment']=='VIP').sum()}명, General: {(df['Segment']=='General').sum()}명")

seg_compare = df.groupby('Segment')[feature_cols].agg(['mean', 'std'])
print(f"\n  VIP vs General 비교:")
print(seg_compare)

# VIP vs 일반 시각화
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(18, 5))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    sns.boxplot(x='Segment', y=col, data=df, ax=axes[i], palette={'VIP': '#FF5722', 'General': '#2196F3'},
                order=['General', 'VIP'])
    axes[i].set_title(f'{col} by Segment', fontsize=11)
plt.suptitle('VIP vs General Donor Comparison', fontsize=14, y=1.02)
plt.tight_layout()
plt.savefig("output/figures/11b_segment_comparison.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/11b_segment_comparison.png")

# Segment 열 정리 (이후 분석에 영향 주지 않도록)
df.drop('Segment', axis=1, inplace=True)

# ==============================================================
# 단계 7: 스케일링
# ==============================================================
print("\n" + "=" * 60)
print("[단계 7] 스케일링")
print("=" * 60)

# 주의: 아래 스케일링은 EDA 목적으로 전체 데이터에 fit_transform함.
# 실제 ML 파이프라인에서는 train/test 분리 후 train에만 fit하여 data leakage를 방지해야 함.

# --- 7-1: 표준화 ---
scaler_std = StandardScaler()
df_standardized = pd.DataFrame(
    scaler_std.fit_transform(df[feature_cols]),
    columns=[f'{c}_std' for c in feature_cols]
)
print("\n--- 7-1: 표준화 ---")
print(df_standardized.describe())

# --- 7-2: 정규화 ---
scaler_mm = MinMaxScaler()
df_normalized = pd.DataFrame(
    scaler_mm.fit_transform(df[feature_cols]),
    columns=[f'{c}_norm' for c in feature_cols]
)
print("\n--- 7-2: 정규화 ---")
print(df_normalized.describe())

# --- 7-3: 비교 시각화 ---
fig, axes = plt.subplots(3, max(len(feature_cols), 2), figsize=(20, 12))
if len(feature_cols) == 1:
    axes = axes.reshape(-1, 1)
for i, col in enumerate(feature_cols):
    axes[0][i].hist(df[col], bins=30, alpha=0.7, color='steelblue')
    axes[0][i].set_title(f'{col}\n(Original)', fontsize=10)
    axes[1][i].hist(df_standardized[f'{col}_std'], bins=30, alpha=0.7, color='coral')
    axes[1][i].set_title(f'{col}\n(Standardized)', fontsize=10)
    axes[2][i].hist(df_normalized[f'{col}_norm'], bins=30, alpha=0.7, color='mediumpurple')
    axes[2][i].set_title(f'{col}\n(Normalized)', fontsize=10)
plt.suptitle('Scaling Comparison', fontsize=14, y=1.01)
plt.tight_layout()
plt.savefig("output/figures/12_scaling_comparison.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/12_scaling_comparison.png")

# ==============================================================
# 단계 8: 추가 시각화
# ==============================================================
print("\n" + "=" * 60)
print("[단계 8] 추가 시각화")
print("=" * 60)

# --- 8-1: CDF 선 그래프 ---
fig, axes = plt.subplots(1, max(len(feature_cols), 2), figsize=(18, 4))
axes = [axes] if len(feature_cols) == 1 else (axes if hasattr(axes, '__len__') else [axes])
for i, col in enumerate(feature_cols):
    sorted_data = np.sort(df[col].dropna())
    cdf = np.arange(1, len(sorted_data)+1) / len(sorted_data)
    axes[i].plot(sorted_data, cdf, color='steelblue', linewidth=2)
    axes[i].set_title(f'{col} CDF (Line Plot)', fontsize=11)
    axes[i].set_xlabel(col); axes[i].set_ylabel('Cumulative Probability')
plt.tight_layout()
plt.savefig("output/figures/13_cdf_lineplot.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/13_cdf_lineplot.png")

# --- 8-2: Class별 프로파일 플롯 ---
class_means = df.groupby(class_col)[feature_cols].mean()
fig, ax = plt.subplots(figsize=(10, 6))
for cls in class_means.index:
    vals = class_means.loc[cls]
    # 전체 데이터 범위 기준 정규화 (클래스 간 절대값 비교 가능)
    global_min = df[feature_cols].min()
    global_max = df[feature_cols].max()
    normalized = (vals - global_min) / (global_max - global_min + 1e-10)
    ax.plot(feature_cols, normalized, marker='o', linewidth=2,
            label=CLASS_LABELS.get(str(int(cls)), str(cls)),
            color=CLASS_COLORS.get(str(int(cls)), None))
ax.set_title('Class-wise Variable Mean Profile (Normalized)', fontsize=14)
ax.set_ylabel('Normalized Mean')
ax.legend(); ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig("output/figures/14_class_profile_lineplot.png", dpi=150, bbox_inches='tight')
plt.close()
print("  -> output/figures/14_class_profile_lineplot.png")

# ==============================================================
# 완료
# ==============================================================
print("\n" + "=" * 60)
print("EDA 분석 완료!")
fig_count = len([f for f in os.listdir('output/figures') if f.endswith('.png')])
tbl_count = len([f for f in os.listdir('output/tables') if f.endswith('.csv')])
print(f"시각화 파일: {fig_count}개 생성")
print(f"통계표 파일: {tbl_count}개 생성")
print("=" * 60)
