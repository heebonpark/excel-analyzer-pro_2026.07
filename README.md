# Excel 분석기 PRO v5.1
## 맥북(Mac) 설치 및 실행 안내

### 폴더 구조
```
자동프로그램/
├── Excel_Analyzer_PRO.py   ← 메인 프로그램
├── run_mac.sh              ← 실행 스크립트 (더블클릭 or 터미널)
├── build_mac.sh            ← .app 빌드 스크립트
└── README.md               ← 이 파일
```

---

### 1단계: Python 설치 확인

터미널(Terminal) 앱을 열고:
```bash
python3 --version
```
→ `Python 3.x.x` 가 나오면 OK
→ 없으면 https://www.python.org/downloads/ 에서 설치

---

### 2단계: 실행

**방법 A — 터미널에서 실행 (권장)**
```bash
# 폴더로 이동
cd ~/Downloads/자동프로그램

# 실행 권한 부여 (최초 1회만)
chmod +x run_mac.sh

# 실행
./run_mac.sh
```

**방법 B — Python 직접 실행**
```bash
cd ~/Downloads/자동프로그램
pip3 install pandas openpyxl xlrd matplotlib xlwings
python3 Excel_Analyzer_PRO.py
```

---

### 자주 묻는 문제

**Q: "tkinter가 없다"고 나올 때**
```bash
brew install python-tk
```
Homebrew가 없으면: https://brew.sh

**Q: "보안 정책으로 열 수 없습니다" 팝업**
→ 시스템 설정 → 개인 정보 보호 및 보안 → "확인 없이 열기" 클릭

**Q: 패키지 설치 오류**
```bash
pip3 install pandas openpyxl xlrd matplotlib xlwings --break-system-packages
```

**Q: 열린 엑셀 불러오기가 안 될 때**
→ Mac은 xlwings만 지원 (win32com 불필요)
→ Excel for Mac이 설치되어 있어야 함

---

### .app 파일로 만들기 (선택사항)
```bash
chmod +x build_mac.sh
./build_mac.sh
# → dist/Excel분석기PRO 생성됨
```
