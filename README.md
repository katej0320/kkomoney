# kkomoney
# 💗 kkomoney

귀엽고 사랑스러운 개인 가계부 웹앱

## 기능
- 수입/지출 기록 및 카테고리 관리
- 엑셀(.xlsx, .csv) 파일 가져오기 (컬럼 매핑 + 중복 감지)
- 월별/카테고리별 예산 관리
- 저축 목표 추적
- 대출 관리 (원리금균등 / 원금균등 상환 계산)
- 월별 상환 스케줄 테이블
- 분석 차트 (월별, 카테고리, 연간)
- Supabase 연동 (다기기 동기화) 또는 로컬 모드

---

## 설치 없이 시작하는 방법

### 1단계 — GitHub 저장소에 파일 올리기

1. `github.com` 에서 저장소 열기
2. URL을 `github.dev` 로 변경 (또는 `.` 키 누르기)
3. `index.html` 파일 생성 후 코드 붙여넣기
4. `supabase_schema.sql` 파일도 동일하게 추가
5. **Commit & Push**

### 2단계 — GitHub Pages 배포

1. 저장소 → Settings → Pages
2. Source: **main 브랜치** 선택 → Save
3. 몇 분 후 `https://[유저명].github.io/[저장소명]` 에서 접속 가능

---

## Supabase 연결하기 (다기기 동기화)

### 1. Supabase 프로젝트 생성
1. [supabase.com](https://supabase.com) 가입 (무료, 신용카드 불필요)
2. New Project 생성
3. 프로젝트 로딩 완료 후 → **SQL Editor** 탭
4. `supabase_schema.sql` 내용 전체 복사 → 붙여넣기 → **Run**

### 2. API 키 확인
1. 좌측 메뉴 → **Settings** → **API**
2. **Project URL** 복사
3. **anon / public key** 복사

### 3. 앱에서 연결
1. 앱 첫 화면에서 URL과 Key 입력
2. **연결하기** 클릭
3. 이메일/비밀번호로 회원가입 후 시작!

---

## 로컬 모드 (Supabase 없이 사용)

연결 화면에서 **"로컬 모드로 시작"** 클릭  
→ 브라우저에 저장 (같은 브라우저에서만 데이터 유지)

---

## 엑셀 가져오기 사용법

1. 홈 또는 내역 탭 하단의 **엑셀 가져오기** 클릭
2. 은행에서 내려받은 `.xlsx` 또는 `.csv` 파일 선택
3. 컬럼 매핑 (자동 추천됨, 필요 시 수정)
4. 미리보기 확인 후 **가져오기** 클릭
5. 중복 내역은 자동으로 건너뜀

### 지원 은행 형식
- 국민은행, 신한은행, 카카오뱅크 등 대부분의 은행 엑셀/CSV 내보내기 지원
- 날짜, 금액, 내용 열만 있으면 OK

---

## 기술 스택
- HTML/CSS/JS (프레임워크 없음, 파일 1개)
- Supabase (PostgreSQL 기반 무료 DB)
- Chart.js (차트)
- SheetJS/XLSX (엑셀 파싱)
- GitHub Pages (무료 호스팅)

**모두 무료, 설치 없음** 🎉