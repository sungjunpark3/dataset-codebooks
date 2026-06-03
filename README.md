# Data Codebook Viewer

연구용 공공·행정 데이터셋의 변수 설명과 코드값을 웹에서 쉽게 확인할 수 있도록 정리한 HTML 코드북 저장소입니다.

## 현재 포함된 문서

- `index.html`: 표본코호트 맞춤형 제공 컬럼 코드북
- 원본 엑셀: `assets/맞춤형 자료 제공 컬럼 레이아웃_2026_v1.xlsx`

## 사용 방법

브라우저에서 `index.html` 파일을 열면 됩니다.

- 테이블 구분별 필터
- 변수명별 필터
- 전체 변수 보기
- 라이트/다크 모드
- 모바일/PC 반응형 화면

## 파일 구조

```text
.
├── assets/
│   └── 맞춤형 자료 제공 컬럼 레이아웃_2026_v1.xlsx
├── tools/
│   └── build_nsc_custom_columns.R
├── index.html
└── README.md
```

## 업데이트

원본 엑셀 파일이 바뀌면 `tools/build_nsc_custom_columns.R`를 실행해 HTML을 다시 생성합니다.

```bash
Rscript tools/build_nsc_custom_columns.R
```
