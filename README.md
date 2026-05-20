# Math Washing

Gemini AI 기반 수학 문제 자동화 도구

---

## 다른 PC에서 설치하는 방법

### 1. 필수 프로그램 설치
- [Python 3.13+](https://www.python.org/downloads/) — 설치 시 **"Add Python to PATH"** 반드시 체크
- [Google Cloud SDK](https://cloud.google.com/sdk/docs/install)

### 2. 코드 다운로드
```powershell
git clone https://github.com/lkh0117-arch/math-washing.git
cd math-washing
```

### 3. 라이브러리 한번에 설치
```powershell
pip install -r requirements.txt
```

### 4. Google Cloud 인증
```powershell
gcloud auth application-default login
```
브라우저에서 lkh0117@gmail.com 계정으로 로그인

### 5. 실행
```powershell
python math_washing.py
```

---

## GCP 설정 정보
- **프로젝트 ID**: project-202ee3d7-34d9-426a-be5
- **리전**: us-central1
- **모델**: gemini-2.5-pro / gemini-2.5-flash
