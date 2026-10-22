# PDF → PPT 변환기

PDF를 PowerPoint(.pptx)로 변환하는 웹사이트입니다. 변환은 모두 브라우저 안에서 이루어지므로 파일이 서버로 업로드되지 않고, 별도 백엔드 없이 정적 호스팅(GitHub Pages 등)만으로 운영할 수 있습니다.

## 변환 방식

| 방식 | 결과 | 용도 |
| --- | --- | --- |
| **편집 가능** (기본) | 텍스트를 지운 페이지 이미지를 배경으로 깔고, 그 위에 원래 위치·크기·색상·굵기·기울임을 맞춘 텍스트 상자를 올림 | 문구 수정이 필요한 경우 |
| **원본 그대로** | 각 페이지를 고해상도 이미지 한 장으로 넣음 | 모양이 100% 같아야 하는 경우 |

편집 가능 모드의 동작 원리:

1. pdf.js로 페이지를 한 번 렌더링하면서 캔버스에 그려지는 글자 하나하나의 위치와 색을 기록합니다.
2. `getTextContent()`로 얻은 정확한 유니코드 텍스트와 1의 글자 위치를 맞춰, 줄 단위 텍스트 상자와 색·글꼴별 구간(run)으로 묶습니다.
3. 텍스트 상자로 옮긴 글자만 빼고 페이지를 다시 렌더링해 배경 이미지를 만듭니다. 도형·표 테두리·사진·그라데이션은 배경에 그대로 남습니다.
4. 대체 폰트를 써도 줄 폭이 원본과 같도록 자간을 보정하고, PDF의 웹 링크는 슬라이드 하이퍼링크로 옮깁니다.

## 한계

- 도형·표·차트는 배경 이미지에 포함되며, PowerPoint 도형으로 편집할 수는 없습니다.
- 회전된 텍스트, 세로쓰기, Type3 글꼴, 윤곽선만 있는 텍스트는 편집 가능한 텍스트로 바꾸지 않고 배경 이미지에 그대로 남깁니다.
- 스캔한 PDF(텍스트가 이미지인 경우)는 OCR을 하지 않으므로 이미지로만 변환됩니다.
- PDF에 내장된 글꼴은 PPTX에 넣을 수 없습니다. 같은 이름의 글꼴이 PowerPoint를 여는 PC에 없으면 비슷한 기본 글꼴(맑은 고딕, Arial, Times New Roman, Courier New)로 바뀝니다.
- 암호가 걸린 PDF는 지원하지 않습니다.

## 개발

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # dist/ 에 정적 파일 생성
npm run preview  # 빌드 결과 확인
```

`npm run dev`/`build` 전에 pdf.js의 CMap(한글 등 CJK 글꼴용), 표준 글꼴, WASM 디코더가 `public/pdfjs/`로 자동 복사됩니다.

## 배포

`main` 브랜치에 push하면 `.github/workflows/deploy.yml`이 GitHub Pages에 배포합니다. 저장소 **Settings → Pages → Source**를 **GitHub Actions**로 설정해야 합니다.

## 구조

- `src/converter.js` — PDF 렌더링, 텍스트 추출·병합, PPTX 생성
- `src/fonts.js` — PDF 글꼴 이름을 PowerPoint 글꼴로 매핑, 설치 여부 확인, 텍스트 폭 측정
- `src/main.js` — 업로드·옵션·진행률 UI
