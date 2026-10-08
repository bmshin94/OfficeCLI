# OfficeCLI 전수조사 분석 및 활용 가이드 (한국어)

> 이 문서는 OfficeCLI 저장소를 전수조사하여 **"무엇인지 / 언제 쓰는지 / 어떤 도움이 되는지 /
> 어떻게 설치·활용하는지 / 어떻게 수익화할 수 있는지"** 를 정리한 분석 문서입니다.
>
> - 작성일: 2026-10-08
> - 분석 대상 커밋: `5385dda` (버전 1.0.151)
> - 분석 범위: 전체 1,209개 파일 / C# 소스 360개 파일 약 27.9만 줄

---

## 📌 관련 링크 (GitHub 주소 포함)

| 구분 | 주소 |
|---|---|
| **이 저장소 (포크)** | https://github.com/bmshin94/OfficeCLI |
| **원본 저장소 (업스트림)** | https://github.com/iOfficeAI/OfficeCLI |
| 공식 웹사이트 | https://officecli.ai |
| 릴리스 (바이너리 다운로드) | https://github.com/iOfficeAI/OfficeCLI/releases |
| 위키 (명령어 상세 문서) | https://github.com/iOfficeAI/OfficeCLI/wiki |
| 이슈 트래커 | https://github.com/iOfficeAI/OfficeCLI/issues |
| GUI 데스크톱 앱 (AionUi) | https://github.com/iOfficeAI/AionUi |
| 커뮤니티 (Discord) | https://discord.gg/2QAwJn7Egx |
| npm 패키지 (CLI) | `@officecli/officecli` |
| npm 패키지 (Node SDK) | `@officecli/sdk` |
| PyPI 패키지 (Python SDK) | `officecli-sdk` |
| 스킬 파일 직접 받기 | https://officecli.ai/SKILL.md |
| 설치 스크립트 | https://d.officecli.ai/install.sh / https://d.officecli.ai/install.ps1 |

---

## 1. 한 줄 요약

> **AI가 Word·Excel·PowerPoint 파일을 직접 만들고 읽고 고칠 수 있게 해주는 명령줄(CLI) 도구.**

- Microsoft Office **설치 불필요**
- 화면 없는 서버(헤드리스) / Docker / CI 에서도 동작
- 단일 자체 완결형 바이너리 (.NET 런타임 내장, 의존성 0)
- 라이선스: **Apache 2.0** (상업적 이용·수정·재배포·SaaS화 모두 허용, 저작권 고지 유지 의무)
- 만든 사람: goworm / OfficeCLI (© 2026)
- 기술 스택: C# / .NET 10, `DocumentFormat.OpenXml` 3.4.1, `System.CommandLine`

---

## 2. 폴더 전수조사 결과

| 폴더 / 파일 | 규모 | 내용 |
|---|---|---|
| `src/officecli/` | .cs 360개 / 27.9만 줄 | 본체 소스 |
| ├ `Core/` | 158개 | 수식엔진(`Formula/`), 피벗(`PivotTableHelper.*` 7개), 차트(`Chart/`), 머메이드 다이어그램(`Diagram/`), 렌더링(`Rendering/`), 라이브 미리보기(`Watch/`), 템플릿 병합(`TemplateMerger.cs`), 단위변환(`EmuConverter.cs`/`Units.cs`), 보안(`SsrfGuard.cs`), 원자적 저장(`AtomicPackageWriter.cs`), 자동설치(`SkillInstaller.cs`/`Installer.cs`), 플러그인(`Plugins/`) |
| ├ `Handlers/Pptx/` | 95개 | 도형·표·차트·3D(.glb)·애니메이션·모프전환·OLE·하이퍼링크·HTML/SVG 미리보기. `EffectTemplates/`에 애니메이션 효과 XML 수십 개 |
| ├ `Handlers/Word/`, `Handlers/Excel/` | — | Word / Excel 처리 로직 |
| ├ `McpServer.cs` | 1개 | **MCP 서버** (JSON-RPC 2.0 over stdio) |
| ├ `CommandBuilder.*.cs` | 9개 | CLI 명령 정의 (Add / Save / Check / GetQuery / View / Refresh …) |
| └ `ResidentClient.cs` | 1개 | **상주 모드** (문서를 메모리에 유지, 명명된 파이프 IPC) |
| `skills/` | 11개 스킬 | AI 에이전트용 지침서 (아래 §7) |
| `schemas/help/` | 수백 개 JSON | 내장 계층형 도움말 스키마 (요소별 속성 정의) |
| `examples/` | **381개** | 예제마다 `.md`(설명)+`.sh`(쉘)+`.py`(파이썬)+결과파일 4종 세트 |
| `sdk/node/` | 6개 | Node.js SDK (`@officecli/sdk`, `index.d.ts` 타입 포함) |
| `sdk/python/` | 5개 | Python SDK (`officecli-sdk`, 표준 라이브러리만) |
| `npm/` | 5개 | npm 배포 래퍼 (설치 시 플랫폼별 바이너리 자동 다운로드) |
| `plugins/plugin-protocol.md` | 1개 | **플러그인 프로토콜 v1** — `.doc`/`.rtf`/`.odt`/**`.hwpx`/`.hwp`**/`.pdf`/`.epub` 확장 규격 |
| `assets/` | 30+ | README 데모 GIF 14개 + `showcase/` 실제 생성 결과물(.docx/.xlsx/.pptx + PNG) |
| `.github/workflows/` | 6개 | 8플랫폼 빌드, npm/PyPI/SDK 배포, SDK 스모크, **스킬-CLI 정합성 검사** |
| `SKILL.md` (루트) | 26KB | AI에게 주는 메인 사용설명서 |
| `README.md` / `_ko` / `_ja` / `_zh` | 4개 | 영/한/일/중 README (한국어 37KB) |
| `build.sh` / `install.sh` / `install.ps1` / `dev-install.sh` | 4개 | 빌드·설치 스크립트 |
| `LICENSE` / `NOTICE` / `THIRD-PARTY-NOTICES.txt` | 3개 | Apache 2.0 + 저작자 표시 |

---

## 3. 핵심 차별점 4가지

### ① 렌더링 엔진 내장 = AI에게 "눈"을 준다 ⭐
`.docx`/`.xlsx`/`.pptx`를 **HTML / PNG 로 직접 그려냅니다.** Office도 브라우저도 필수가 아닙니다.

```bash
officecli view deck.pptx html -o /tmp/deck.html   # 독립 HTML (에셋 인라인)
officecli view deck.pptx screenshot -o out.png    # 페이지별 PNG (멀티모달 AI용)
officecli watch deck.pptx                          # localhost:26315 자동 새로고침
```

AI가 **"제목이 넘쳤나? 도형이 겹쳤나?"를 눈으로 확인하고 스스로 고칩니다**
(= 렌더링 → 보기 → 수정 루프를 닫음). 차트 추세선·오차막대·워터폴·캔들스틱·스파크라인,
수식(OMML → LaTeX → KaTeX), Three.js 기반 3D `.glb` 모델, 모프 전환, 슬라이드 줌까지 렌더링합니다.
Excel `watch`는 **인라인 셀 편집 + 차트 드래그 재배치**도 지원합니다.

> python-docx / openpyxl 같은 라이브러리와 결정적으로 갈리는 지점입니다.

### ② Excel 수식 엔진 내장 (350+ 함수)
`=SUM(A1:A2)`를 써넣으면 **즉시 계산해서 값까지 파일에 기록**합니다. Excel로 열어 재계산할 필요가 없습니다.
동적 배열(`FILTER`/`SORT`/`UNIQUE`/`SEQUENCE`/`LET`/`LAMBDA`, `_xlfn.` 자동 접두),
`VLOOKUP`/`XLOOKUP`/`INDEX`/`MATCH`, 재무·채권, 통계 분포·검정·회귀, 날짜·텍스트 함수.
**네이티브 OOXML 피벗테이블**을 단일 명령으로 생성합니다.

```bash
officecli add sales.xlsx '/Sheet1' --type pivottable \
  --prop source='Data!A1:E10000' --prop rows='Region,Category' \
  --prop cols=Quarter --prop values='Revenue:sum,Units:avg' \
  --prop showDataAs=percentOfTotal
```

### ③ `merge` — 한 번 설계, N번 찍어내기
템플릿의 `{{key}}` 자리에 JSON 데이터를 꽂습니다. 단락·표셀·도형·머리글/바닥글·차트제목 전부에서 작동.

```bash
officecli merge invoice-template.docx out-001.docx --data '{"client":"Acme","total":"$5,200"}'
officecli merge q4-template.pptx q4-acme.pptx --data data.json
```

> **AI는 레이아웃을 1회만 설계(비싼 작업), 이후 N건은 코드가 채움(공짜·결정론적·토큰 0).**
> AI가 매 보고서를 처음부터 재생성해 N개의 불일치 레이아웃을 만드는 실패 모드를 방지합니다.

### ④ `dump` ↔ `batch` — 기존 문서에서 학습
```bash
officecli dump existing.docx -o blueprint.json          # 전체 문서
officecli dump existing.docx /body/tbl[1] -o table.json # 임의 서브트리
officecli dump existing.xlsx /Sheet1 -o sheet.json      # 단일 워크시트
officecli batch new.docx --input blueprint.json         # 그대로 재생
```
> "회사에서 쓰던 양식 그대로 100개 만들어줘"가 가능해집니다.
> AI가 raw OOXML XML이 아닌 **구조화된 사양**을 읽기 때문에 사람이 만든 샘플에서 학습할 수 있습니다.

---

## 4. 3계층 아키텍처

```
L1 (읽기)    view  → text / annotated / outline / stats / issues / html / svg / screenshot
L2 (DOM)     get, query, set, add, remove, move, swap        ← 평소엔 여기만 사용
L3 (원시XML) raw, raw-set, add-part, validate                 ← 막혔을 때의 탈출구
```

AI는 L1부터 쓰고, 안 되면 L2, 그래도 안 되면 L3로 내려갑니다. → **토큰(비용) 최소화 구조.**
L3가 있기 때문에 "이건 못 한다"가 거의 없습니다.

### 주소(경로) 문법 — XPath가 아닌 자체 문법, 1-based
```
/slide[1]              1번 슬라이드
/slide[1]/shape[2]     1번 슬라이드의 2번 도형
/body/p[5]             본문 5번째 단락
/body/p[1]/r[1]        1번째 단락의 1번째 run
/Sheet1                Sheet1 워크시트
$Sheet1:A1             Sheet1의 A1 셀
```
→ AI가 OOXML 네임스페이스(`w:`, `a:`, `p:`)를 전혀 몰라도 문서를 탐색할 수 있습니다.

### 단위·색 입력이 관대함
| 유형 | 지원 형식 |
|---|---|
| 치수 | `2cm` `1in` `72pt` `96px` `914400`(EMU) |
| 색 | `#FF0000` `FF0000` `red` `rgb(255,0,0)` `accent1`(테마색) |
| 글꼴 크기 | `14` `14pt` `10.5pt` |
| 간격 | `12pt` `0.5cm` `1.5x` `150%` |

### 상주 모드 / 배치 모드
```bash
# 상주 모드 — 모든 명령이 첫 접근 시 자동 시작(60초 유휴), 명시적 open은 12분 유휴
officecli open report.docx
officecli set report.docx /body/p[1]/r[1] --prop bold=true
officecli close report.docx

# 배치 모드 — 한 번의 open/save 사이클에서 원자적 다중 실행 (기본 첫 오류 중단, --force로 계속)
echo '[{"command":"set","path":"/slide[1]/shape[1]","props":{"text":"Hello"}},
       {"command":"set","path":"/slide[1]/shape[2]","props":{"fill":"FF0000"}}]' \
  | officecli batch deck.pptx --json
```

---

## 5. 언제 쓰는가 / 어떤 도움이 되는가

### 활용 상황
| 상황 | 활용 |
|---|---|
| AI에게 PPT를 맡길 때 | "Q4 실적 보고서 10장" → AI가 만들고 렌더링해 보고 스스로 수정 |
| 보고서 대량 생산 | DB/API → 템플릿 `merge` → 거래처 200곳 맞춤 제안서 |
| CI/CD 문서 파이프라인 | 데이터 변경 시 문서 자동 재생성 |
| 문서 품질 검사 | `view issues`, `validate`로 서식 깨짐·스타일 불일치 자동 검출 |
| 문서에서 데이터 추출 | 수백 개 파일 → 구조화된 JSON |
| Docker/서버 자동화 | Office 설치 불가 환경에서 헤드리스 문서 생성 |

### 개인에게 돌아오는 이득
1. 보고서 작성 시간이 **시간 → 분 단위**로 단축 (서식 다듬기 노동 소멸)
2. Claude Code / Cursor 등에 붙여 **말로 시킬 수 있음** (`officecli install` 한 번)
3. **전문 템플릿이 기본 제공** — 투자 피치덱 / 재무모델(3단표·DCF·LBO) / 학술논문(APA·IEEE) /
   KPI 대시보드 / 작성가능 Word 양식. 업계 표준 규칙이 코드화되어 있음
   (예: 재무모델 스킬은 "CFO 4색 규칙 — 입력=파랑 / 수식=검정 / 시트간참조=초록 / 가정=노랑배경"을 강제)
4. **PPT 디자인 52종 스타일 라이브러리** (`skills/morph-ppt/reference/styles/`) —
   `dark--investor-pitch`, `bw--swiss-bauhaus`, `light--glassmorphism-vc` 등.
   각 스타일마다 `style.md`(색·폰트·레이아웃 규칙) + 실제 `.pptx` 샘플
5. **예제 381개**가 그대로 학습자료 (쉘·파이썬 양쪽 제공)
6. 무료 · 오픈소스(Apache 2.0) — 회사 업무, 상업 제품에도 사용 가능
7. 의존성 지옥 없음 (단일 실행파일)
8. **한글(.hwp) 지원 가능성** — 플러그인 프로토콜에 `.hwpx`/`.hwp`가 명시적 타깃으로 기재

---

## 6. 설치 및 사용법

### 설치 — 5가지 방법
```bash
# ① 원라인 (가장 쉬움)
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash   # macOS/Linux
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex          # Windows

# ② 패키지 매니저
brew install officecli                 # macOS / Linux
scoop install officecli                # Windows
npm install -g @officecli/officecli    # 모든 플랫폼

# ③ 수동 다운로드 → officecli install  (또는 그냥 officecli 실행도 설치 트리거)

# ④ AI 에이전트에게 한 줄 던지기 (제작자 추천)
curl -fsSL https://officecli.ai/SKILL.md

# ⑤ 소스 빌드 (.NET 10 SDK 필요)
./build.sh        # 현재 플랫폼
./build.sh all    # 8개 플랫폼 전부
```

`install.sh` 동작: 미러(`d.officecli.ai`) 우선 → 실패 시 GitHub Releases 폴백 →
**SHA256 체크섬 검증** → PATH 등록(쉘 rc 파일 추가).

지원 바이너리: `mac-arm64`, `mac-x64`, `linux-x64`, `linux-arm64`,
`linux-alpine-x64`(musl), `linux-alpine-arm64`(musl), `win-x64.exe`, `win-arm64.exe`

확인: `officecli --version`

### `officecli install`의 자동 감지
`Core/SkillInstaller.cs`가 설치된 AI 도구를 자동 감지해 스킬을 주입합니다 —
Claude Code(`~/.claude/skills/`), Cursor, Windsurf, GitHub Copilot, Codex CLI(`.agents/skills`) 등.

```bash
officecli install          # 전체
officecli install claude   # Claude Code만
```

### 전체 명령어
| 명령 | 설명 |
|---|---|
| `create` | 빈 .docx/.xlsx/.pptx 생성 (확장자로 판단) |
| `view` | 보기 — `outline`/`text`/`annotated`/`stats`/`issues`/`html`/`svg`/`screenshot` |
| `get` | 요소 가져오기 (`--depth N`, `--json`) |
| `query` | CSS 유사 쿼리 (`[attr=value]`, `:contains()`, `:has()`) |
| `set` | 속성 수정 |
| `add` | 요소 추가 (`--from <path>`로 복제) |
| `remove` | 삭제 |
| `move` | 이동 (`--to` / `--index` / `--after` / `--before`) |
| `swap` | 두 요소 맞교환 |
| `validate` | OpenXML 스키마 검증 |
| `batch` | 다중 작업 원자적 실행 (stdin / `--input` / `--commands`) |
| `merge` | 템플릿 `{{key}}` ← JSON 데이터 |
| `dump` | 문서/서브트리 → 재생 가능 batch JSON |
| `watch` | 브라우저 라이브 미리보기 (자동 새로고침) |
| `mcp` | MCP 서버 시작 / 에디터 등록 |
| `raw` / `raw-set` / `add-part` | 원시 XML 접근 (XPath) |
| `open` / `close` | 상주 모드 시작 / 종료 |
| `install` | 바이너리 + 스킬 + MCP 설치 (`all`, `claude`, `cursor` …) |
| `config` | 설정 (`~/.officecli/config.json`) |
| `help <format> <cmd>` | 내장 계층형 도움말 |

### 30초 체험
```bash
officecli create deck.pptx
officecli watch deck.pptx                      # 브라우저 열림 (localhost:26315)
# 다른 터미널에서 ↓ 실행하면 브라우저가 즉시 갱신
officecli add deck.pptx / --type slide --prop title="Hello, World!"
```

### 실전 전체 흐름
```bash
officecli create report.pptx
officecli add report.pptx / --type slide --prop title="Q4 Results" --prop background=1A1A2E
officecli add report.pptx '/slide[1]' --type shape \
  --prop text="Revenue: \$4.2M" --prop x=2cm --prop y=5cm \
  --prop font=Arial --prop size=28 --prop color=FFFFFF
officecli view report.pptx outline
officecli view report.pptx screenshot -o check.png     # 눈으로 확인
officecli view report.pptx issues --json               # 문제 자동 검출
officecli validate report.pptx
officecli set report.pptx '/slide[1]/shape[1]' --prop font=Arial
```

### 모를 때는 `help` (추측 금지)
```bash
officecli help                         # 전체 개요
officecli help pptx                    # pptx 요소 전부
officecli help pptx set shape          # shape에 set 가능한 속성만
officecli help docx paragraph --json   # 기계 판독용 스키마
```
별칭: `word`→`docx`, `excel`→`xlsx`, `ppt`/`powerpoint`→`pptx`

---

## 7. 플러그인? 스킬? MCP? → **4가지 얼굴을 가진 하나의 바이너리**

| 질문 | 답 |
|---|---|
| 플러그인인가? | ❌ 아니다 — 오히려 **플러그인을 받는 호스트** |
| 스킬인가? | ✅ **그렇다** — 11개 스킬 제공, 주력 방식 |
| MCP인가? | ✅ **그렇다** — MCP 서버 내장, 한 줄 등록 |
| 근본은? | **CLI 실행파일** (+ Node / Python SDK) |

### ① CLI 실행파일 (뿌리)
C# / .NET 10 단일 네이티브 바이너리. 나머지는 모두 이걸 감싼 겉옷.

### ② 스킬 (Skill) — 주력 ⭐
`skills/` 의 11개 `SKILL.md` + 루트 `SKILL.md`(26KB).
스킬은 **프로그램이 아니라 "AI용 사용설명서 + 작업 규칙" 마크다운**입니다.
Claude Code Agent Skills 규격(`name` + `description` 프론트매터)을 따릅니다.
CI의 `skill-parity.yml`이 **스킬 문서와 실제 CLI 기능의 정합성을 자동 검사**합니다.

```bash
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md
```

**11개 스킬 목록**

| 스킬 | 용도 |
|---|---|
| `officecli` | 메인 지침서 (전체 명령·전략) |
| `officecli-docx` | Word 범용 (보고서/편지/메모/제안서) |
| `officecli-xlsx` | Excel 범용 (시트/수식/차트/피벗) |
| `officecli-pptx` | PPT 범용 (12컬럼 그리드, 제목≥36pt / 본문≥18pt "시각 하한선" 강제) |
| `officecli-pitch-deck` | 투자 피치덱 전용 (시드~시리즈C, 10개 핵심 슬라이드 레시피, VC 체크) |
| `officecli-academic-paper` | 학술논문 (APA/Chicago/IEEE/MLA, 수식번호, SEQ·PAGEREF 교차참조, 참고문헌) |
| `officecli-financial-model` | 재무모델 (3단표/DCF/LBO/민감도·시나리오/IRR, CFO 4색 규칙) |
| `officecli-data-dashboard` | Excel KPI 대시보드 (카드+차트+스파크라인+조건부서식) |
| `officecli-word-form` | 작성가능 Word 양식 (콘텐츠컨트롤 SDT + 체크박스 + MERGEFIELD + 문서보호) |
| `morph-ppt` | 키노트식 모프 전환 애니메이션 PPT (+ 52종 스타일 라이브러리) |
| `morph-ppt-3d` | 3D 모델(.glb) + 시네마틱 카메라 PPT |

**계층(scene layer) 구조**
```
officecli-pptx (기본: 시각 하한선, 12컬럼 그리드, 4개 표준 팔레트, 연결선 규칙, Delivery Gate 1~5a)
   ├─ officecli-pitch-deck   (+ 라운드 진단, 10개 슬라이드 레시피, VC 체크, Gate 6 "fresh-eyes")
   ├─ morph-ppt              (+ 슬라이드간 모프, 고스트 규칙, 52종 스타일)
   └─ morph-ppt-3d           (+ GLB 3D, 카메라 무브)
```
규칙: **한 산출물에 스킬 하나만 로드, 절대 쌓지 말 것.**
각 스킬에 **역방향 라우팅**도 명시 (예: 피치덱 스킬 → "세일즈덱/보드리뷰면 `pptx`로 보내라").
스킬 내부에 **Delivery Gate(납품 전 품질 게이트)** 가 정의되어 있어 일반 AI 생성물과 품질이 갈립니다.

### ③ MCP 서버 (내장)
`src/officecli/McpServer.cs` — JSON-RPC 2.0 over stdio.
```bash
officecli mcp claude     # Claude Code
officecli mcp cursor     # Cursor
officecli mcp vscode     # VS Code / Copilot
officecli mcp lmstudio   # LM Studio
officecli mcp list       # 등록 상태 확인
```
도구 파라미터가 `command` 문자열 **딱 하나**이고 CLI에 그대로 전달하는 단순 설계.
**셸 접근 권한 없이** 문서 작업이 가능해 보안상 유리합니다.

### ④ 플러그인 (OfficeCLI가 *호스트*)
`plugins/plugin-protocol.md` (v1 최종 초안). OfficeCLI**에** 포맷 확장을 꽂는 규격입니다.
- 종류: `dump-reader`(외부 포맷 → 네이티브 변환) 등 3종
- 타깃: `.doc`, `.rtf`, `.odt`, **`.hwpx`, `.hwp`(한글)**, `.pdf`, `.epub`
- 설계 동기: 본체 바이너리 비대화 방지 + **Apache 라이선스 본체와 독점 구현의 분리**

### ⑤ 보너스: SDK 2종
```bash
npm install @officecli/sdk   # Node.js (TypeScript 타입 index.d.ts 포함)
pip install officecli-sdk    # Python (표준 라이브러리만, 의존성 0)
```
상주 파이프로 명령 전달 → 프로세스 매번 생성 안 함(빠름).
⚠️ **CLI 바이너리는 별도 설치되어 PATH에 있어야 합니다** (pip이 바이너리를 설치해주지 않음).

---

## 8. API 토큰이 필요한가? → **아니요, 전혀 불필요** ✅

코드베이스 전체를 `api[_ -]key` / `apikey` / `bearer` / `openai` / `anthropic` / `token` /
`license key` 로 전수 검색한 결과 **단 1건**이며, 그것도 `SkillInstaller.cs:28`의
`"openai-codex"` — **Codex CLI 감지용 폴더명 문자열**로 API와 무관합니다.

### 이유
OfficeCLI는 **AI가 아니라 AI가 쓰는 도구**입니다. 자체적으로 LLM을 호출하지 않습니다.

```
[AI (Claude / GPT) — 여기에 API 토큰 필요]
          ↓ 명령 생성
[OfficeCLI — 토큰 불필요, 100% 로컬]
          ↓
[.pptx / .xlsx / .docx 파일]
```

→ OfficeCLI 자체는 **완전 무료, 무제한, 로그인·계정·사용량 제한 없음.**

### 네트워크 사용 지점 (전부 선택적)
| 파일 | 용도 | 필수 |
|---|---|---|
| `Core/UpdateChecker.cs` | 자동 업데이트 확인 (24시간 디바운스) | ❌ 끌 수 있음 |
| `Core/ImageSource.cs`, `Core/FileSource.cs` | URL로 이미지/파일 삽입 시에만 | ❌ |
| `Handlers/Pptx/PowerPointHandler.Background.cs` | URL 배경 이미지 | ❌ |
| `Core/Diagram/MermaidImageRenderer.cs` | 머메이드 다이어그램 PNG 렌더 | ❌ |
| `Core/SsrfGuard.cs` | **SSRF 공격 방어 가드** (보안 장치) | — |

```bash
officecli config autoUpdate false        # 자동 업데이트 영구 중지
OFFICECLI_SKIP_UPDATE=1 officecli ...    # 1회만 건너뛰기
```

### 보안·프라이버시
- 문서가 외부로 나가지 않음 → **완전 오프라인 / 에어갭 환경에서 동작**
- `SsrfGuard.cs` — SSRF(내부망 요청 위조) 방어 구현
- `AtomicPackageWriter.cs` — **원자적 저장**, 저장 중단에도 파일 손상 없음
- 커밋 로그: *"a refused Add/Set/Remove/RawSet leaves the file byte-identical"*
  → **거부된 작업은 파일을 1바이트도 변경하지 않음**
- 사내 기밀 문서 처리에도 적합한 설계

---

## 9. AI 에이전트 구축에 도움이 되는가? → **핵심 부품급** ⭐⭐⭐⭐⭐

### 에이전트 친화 설계 (단순 "사용 가능"이 아니라 "에이전트를 위해 설계됨")
| 특징 | 에이전트 관점 이점 |
|---|---|
| 모든 명령 `--json` + 일관 스키마 | 정규식 파싱 / stdout 스크래핑 불필요 |
| 경로 기반 주소 | XML 네임스페이스 지식 0으로 탐색 |
| L1→L2→L3 점진적 복잡도 | **토큰 비용 최소화** |
| 구조화된 에러 + suggestion | `not_found` / `invalid_value` / `unsupported_property` + 유효범위 → **사람 개입 없는 자가 수정** |
| 내장 렌더링 | 자기 출력을 **보고** 레이아웃 수정 (CI / Docker / 헤드리스 전부) |
| 내장 계층형 help | "추측-실패-재시도 루프" 제거 |
| 수식·피벗 자동 평가 | Office 왕복 없이 계산값·집계 즉시 읽기 |
| `merge` | 레이아웃 1회 설계 + N회 채우기 → 토큰 소모 없음 |
| `dump` | 사람이 만든 샘플에서 학습 (raw XML 아닌 구조화 사양) |
| 자동 설치 | 에이전트 도구 감지 + 자가 구성 |

### 토큰 경제 (실무 핵심)
```
❌ 나쁜 패턴: 보고서 100개를 AI가 100번 생성 → 토큰 100배 + 레이아웃 100종류(불일치)
✅ 좋은 패턴: AI가 템플릿 1번 설계 → merge로 100번 채움 → 토큰 1배 + 레이아웃 100% 동일
```
README/스킬 문서가 이 **실패 모드를 명시적으로 경고**하고 있습니다.

### 에이전트 아키텍처 제안
```
A. 보고서 생성 에이전트
   DB/API → 집계 → AI 인사이트 작성 → OfficeCLI 생성 → view screenshot 자가검수 → 발송

B. 문서 분석 에이전트
   문서 100개 → dump/get --json → AI 분석 → 요약 보고서를 다시 OfficeCLI로 생성

C. 문서 QA 에이전트
   제출 문서 → view issues --json + validate → 문제 리포트 → 자동 수정

D. 멀티 에이전트 분업
   [기획] 구조 설계 → [작성] 집필 → [디자인] OfficeCLI 서식 → [검수] screenshot 보고 반려/승인
```

### 통합 방법 3가지
```python
# ① 서브프로세스 (가장 쉬움)
import subprocess, json
r = subprocess.run(['officecli','get','deck.pptx','/slide[1]','--json'],
                   capture_output=True, text=True)
data = json.loads(r.stdout)
```
```bash
# ② SDK (빠름 — 상주 파이프)
pip install officecli-sdk      /      npm install @officecli/sdk
```
```bash
# ③ MCP (셸 권한 없이, 가장 안전)
officecli mcp claude
```

### 한계 (정직하게)
- Office/LibreOffice 대체 **변환기가 아님** (`.pptx`→`.pdf` 완벽 변환은 범위 밖, PDF는 플러그인 타깃)
- 렌더링은 고충실도지만 **Office 100% 픽셀 동일은 아님** (스킬 문서에 renderer quirks 명시)
- **`.hwp`(한글) 미지원** — 플러그인 규격만 존재
- 매우 복잡한 레이아웃은 L3(raw XML)로 내려가야 할 수 있음

---

## 10. React / PHP 로 만들 수 있는가? → **"만든다"가 아니라 "감싼다"** ✅

### 재작성은 비현실적
27.9만 줄 + OOXML 스펙(수천 페이지) + 350+ Excel 함수 엔진 + 렌더링 엔진.
JS/PHP에 `DocumentFormat.OpenXml` 급 성숙 라이브러리가 없습니다. 1인 기준 수년치 작업.

### ⚠️ React 단독 불가 — 백엔드 필수
브라우저는 보안상 로컬 실행파일을 실행할 수 없습니다.
```
❌ React(브라우저) ──X──> officecli 바이너리
✅ React(브라우저) ──HTTP──> 백엔드(Node/PHP) ──exec──> officecli 바이너리
```

### 권장 아키텍처
```
┌──────────────────────────────────────────┐
│ React 프론트엔드                           │
│ - 폼/위자드 (클릭으로 문서 설계)             │
│ - view html 결과를 iframe에 실시간 미리보기   │
│ - 완성 파일 다운로드                        │
└──────────────┬───────────────────────────┘
               │ REST / WebSocket
┌──────────────▼───────────────────────────┐
│ 백엔드 (Node.js 또는 PHP)                  │
│ - officecli 실행 / SDK 호출                │
│ - 작업 큐 (Redis / BullMQ / Laravel Queue) │
│ - 파일 저장 (S3) · 인증 · 결제 · 요금제      │
└──────────────┬───────────────────────────┘
               │ exec / 상주 파이프
┌──────────────▼───────────────────────────┐
│ officecli 바이너리 (Docker 컨테이너)         │
└──────────────────────────────────────────┘
```

### Node.js (가장 추천 — 공식 SDK 존재)
```javascript
const officecli = require('@officecli/sdk');          // 상주 파이프 (빠름)

// 또는 가장 단순하게
const { execFile } = require('child_process');
execFile('officecli', ['add', file, '/', '--type','slide',
                       '--prop', `title=${title}`], cb);
```
`sdk/node/index.d.ts` 타입 제공, `sdk/node/demo.js` / `smoke.js` 가 즉시 활용 가능한 예제.

### PHP (가능, 공식 SDK 없음)
```php
<?php
// ① 단순 실행 — 인자는 반드시 escapeshellarg!
exec('officecli view ' . escapeshellarg($file) . ' outline --json 2>&1', $out);
$data = json_decode(implode("\n", $out), true);

// ② 배치 (권장 — 한 번의 열기/저장)
$batch = json_encode([
  ['command'=>'set','path'=>'/slide[1]/shape[1]','props'=>['text'=>$title]],
  ['command'=>'set','path'=>'/slide[1]/shape[2]','props'=>['fill'=>'FF0000']],
]);
$p = proc_open('officecli batch ' . escapeshellarg($file) . ' --json',
               [['pipe','r'],['pipe','w'],['pipe','w']], $pipes);
fwrite($pipes[0], $batch); fclose($pipes[0]);
$result = json_decode(stream_get_contents($pipes[1]), true);
?>
```
**PHP 주의사항**
- 공유호스팅은 `exec`/`proc_open`이 막혀 있을 수 있음 → **VPS / Docker 필요**
- **반드시 `escapeshellarg()`** — 사용자 입력 직접 삽입은 셸 인젝션
- `max_execution_time` 초과 주의 → **작업 큐로 비동기 처리**
- Laravel이면 `Symfony\Process` + Queue(Job) 조합

### 난이도 / 기간 (1인 기준)
| 만들 것 | 난이도 | 기간 |
|---|---|---|
| 문서 생성 웹폼 (템플릿 선택 → 입력 → 다운로드) | ⭐⭐ | 1~2주 |
| `merge` 기반 대량 생성 (CSV 업로드 → N개 파일) | ⭐⭐ | 1~2주 |
| 실시간 미리보기 에디터 (`view html` → iframe) | ⭐⭐⭐ | 3~4주 |
| AI 챗 + 문서생성 SaaS (LLM API 결합) | ⭐⭐⭐⭐ | 2~3개월 |
| OfficeCLI 자체를 JS/PHP로 재작성 | ⭐⭐⭐⭐⭐ | **권하지 않음** |

### 실전 팁
- **`merge`를 최대한 활용** — 백엔드는 `{{key}}`만 채우면 됨 (가장 안전·빠름·결정론적)
- **미리보기는 `view html -o`** → 독립 HTML(에셋 인라인) → iframe 직접 삽입
- **동시 접속** — 사용자별 작업 디렉터리 격리 + 큐
  (상주 모드가 파일락 충돌을 자동 회피하지만 경로는 분리해야 안전)
- **Docker 패키징** — Alpine(musl) 바이너리가 있어 이미지가 가벼움

---

## 11. 유튜브 강의 영상 제작 가능성 → **소재가 넘침** ⭐⭐⭐⭐

### 유리한 이유
1. 결과가 즉각 시각적 → 썸네일·훅이 저절로 나옴
2. **`watch` 라이브 미리보기가 영상에 최적** (명령 치면 브라우저가 실시간 변화)
3. 한국어 README 37KB 존재 → 번역 작업 거의 불필요
4. 예제 381개 + 스타일 52종이 통째로 대본 재료
5. 한국어 커뮤니티 콘텐츠가 희박한 **블루오션**
6. AI 자동화 + 사무직 업무 자동화 = 시청자 층이 두꺼운 주제

### 시리즈 커리큘럼 (12편)
| # | 제목 | 길이 | 핵심 |
|---|---|---|---|
| 1 | AI가 PPT를 만들어준다고? OfficeCLI 충격 데모 | 8분 | 훅 · before/after |
| 2 | 5분 설치 완전정복 (Win/Mac/Linux) | 7분 | 원라인 설치 + 자동감지 |
| 3 | 첫 PPT 만들기 — `create`/`add`/`view` | 12분 | 기본 3종 |
| 4 | **`watch` 라이브 미리보기** 실시간 편집 | 10분 | ⭐ 임팩트 최고 |
| 5 | Excel 완전정복 — 수식 자동계산 + 피벗 한 줄 | 15분 | 350+ 함수 |
| 6 | Word 보고서 자동생성 — 목차·머리글·스타일 | 15분 | TOC 자동 |
| 7 | **Claude Code 연결** — 스킬 & MCP | 12분 | ⭐ AI 연동 |
| 8 | `merge`로 거래처 100곳 제안서 한 번에 | 12분 | 실무 임팩트 최고 |
| 9 | `dump`로 기존 양식 복제 — 회사 템플릿 재활용 | 12분 | 실무 수요 큼 |
| 10 | 투자 피치덱 자동생성 (pitch-deck 스킬) | 15분 | 스타트업 타깃 |
| 11 | 재무모델 (3단표/DCF) 자동생성 | 18분 | 고부가 시청자 |
| 12 | React+Node로 문서생성 SaaS 만들기 | 25분 | 개발자 · 수익화 연결 |

### 타깃별 분기
- 일반 직장인 → AionUi(GUI) + 말로 시키기 중심, 터미널 최소화
- 개발자 → CLI / SDK / MCP / 배치 / CI 중심
- 스타트업·기획자 → 피치덱 · 재무모델 · 대시보드 스킬 중심

### 제작 팁
- `watch` 화면을 전체화면 녹화 + 터미널 작게 오버레이 → 변화가 극적으로 보임
- before/after 분할화면: "손으로 30분" vs "명령 한 줄 3초"
- `assets/*.gif` 에 **공식 데모 GIF 14개** 존재 (Apache 2.0, 출처 표기하면 사용 가능)
- `assets/showcase/` 에 완성 결과물(.docx/.xlsx/.pptx + PNG) 존재 → 결과 보여주기 용이

### ⚠️ 법적 체크
- Apache 2.0 → **영상 제작·수익창출 자유** ✅
- **저작자 표시는 라이선스 4조 의무.** 영상 설명란에 반드시:
  ```
  OfficeCLI (Apache License 2.0)
  © 2026 OfficeCLI — Created and maintained by goworm
  원본: https://github.com/iOfficeAI/OfficeCLI
  공식: https://officecli.ai
  ```
- "내가 만들었다" ❌ → "소개합니다 / 활용법" ⭕
- 코드 수정 배포 시 `NOTICE` 파일 유지 의무
- 로고·브랜드 사용은 별도 확인 (Apache 2.0은 **상표권을 부여하지 않음** — 5조)

---

## 12. 수익화 아이디어 (상세)

### 법적 토대
| 가능 | 항목 |
|---|---|
| ✅ | 상업적 사용 / 수정 / 재배포 |
| ✅ | **클로즈드 소스 제품에 포함** (GPL과 달리 소스 공개 의무 없음) |
| ✅ | SaaS 서비스화 |
| ✅ | 특허 라이선스 부여 (3조) |
| ⚠️ | **저작권 / `NOTICE` 고지 유지 의무** (4조), 수정 시 "변경했음" 명시 |
| ❌ | 상표(로고·이름) 사용권 **불포함** (5조) |
| ❌ | 보증 없음 (면책) |

→ **"OfficeCLI를 엔진으로 쓰는 유료 제품"은 완전히 합법.**
제품 내 Apache 2.0 고지 페이지를 두고, 제품명에 "OfficeCLI"를 브랜드처럼 쓰는 건 피하세요
("powered by OfficeCLI" 수준은 무난).

### 핵심 통찰 — 돈이 나오는 자리
```
OfficeCLI의 능력은 압도적 ────┐
                            ├──> 이 간극이 전부 사업 기회
일반인은 터미널을 못 쓴다 ────┘
```
4개의 간극: **UI 간극** / **지식 간극**(좋은 문서는 도구가 아니라 디자인·내용 역량) /
**운영 간극**(서버·큐·저장·동시성) / **현지화 간극**(한국 양식, `.hwp`, 정부지원사업)

---

### 🥇 티어 1 — 가장 현실적 (3개월 내 수익 가능)

#### 1. 한국형 문서 템플릿 SaaS ⭐ 최고 추천
- **무엇**: 한국 기업 양식 템플릿 + 웹 입력폼 → 클릭으로 완성 문서 다운로드
- **왜 되나**: 한국 기업 문서는 양식이 극도로 정형화 (기안서·품의서·지출결의서·사업계획서·
  정부지원사업 신청서(창업패키지·R&D)·IR덱·주간보고). `merge` 템플릿화 시 **생성 비용 ≈ 0**
- **구현**: 양식 50종 템플릿화 → React 웹폼 → 백엔드 `officecli merge` → 파일 반환
- **가격**: 무료 3건/월 → Pro ₩9,900/월 무제한 → 팀 ₩49,000/월
- **투입** 1~2개월 · **난이도** ⭐⭐ · **예상** 유료 500명 × ₩9,900 = **월 495만원**
- **moat**: 개발이 아니라 **템플릿 콘텐츠의 질**

#### 2. AI 보고서 생성 SaaS
- **무엇**: "2024 4분기 매출 보고서" + CSV 업로드 → 완성 PPT/Excel
- **흐름**: 요청+데이터 → LLM으로 구조·인사이트 → OfficeCLI 생성 → `view screenshot` 자가검수 → 다운로드
- **차별점**: 경쟁(Gamma·Tome·Beautiful.ai)은 웹 전용 또는 PPT 내보내기 품질 저하.
  OfficeCLI는 **네이티브 파일 + 살아있는 Excel 수식** → "다운로드해서 바로 수정 가능"
- **가격**: 크레딧 ₩500/문서 또는 ₩29,000/월 50건
- **투입** 2~3개월 · **난이도** ⭐⭐⭐⭐ · **예상** 1,000명 × ₩29,000 = **월 2,900만원**
- **리스크**: LLM API 비용 → **`merge` 패턴 필수**

#### 3. 교육 콘텐츠 (유튜브 + 온라인 강의)
- 투입 자본 ≈ 0, 리스크 ≈ 0, 다른 아이디어의 마케팅 채널
| 경로 | 예상 |
|---|---|
| 유튜브 애드센스 (구독 1만 / 월 10만뷰) | 월 50~150만원 |
| 온라인 강의 ₩99,000 × 100명 | 990만원 (반복 판매) |
| 기업 출강 1일 ₩100~300만원 × 월 1~3회 | 월 100~900만원 |
| 전자책 / 노션 템플릿 | 월 50~200만원 |
| 멤버십 ₩9,900 × 300명 | 월 297만원 |
- **난이도** ⭐⭐ · **예상** 월 100~1,000만원
- **추천 순서**: 유튜브(신뢰) → 강의 판매 → 기업 출강(단가 최고) → SaaS 확장

#### 4. 프리랜서 / 자동화 구축 대행
- 수요 예: 일일보고 PPT 자동생성 ₩500~1,500만 / 거래처 200곳 맞춤 제안서 ₩300~800만 /
  월말 재무보고 자동화 ₩500~1,000만 / 문서 품질검사 봇 ₩300~600만
- **왜 유리**: OfficeCLI가 90%를 이미 해줌 → 조립만 하면 됨 → **시간 대비 단가 압도적**
- **투입** 학습 2~4주 · **난이도** ⭐⭐⭐ · **예상** 월 300~3,000만원
- **팁**: 첫 1~2건 저가 수주 → 포트폴리오 + **재사용 코드 자산** 확보 → 이후 실질 마진 급등

---

### 🥈 티어 2 — 중기 (6~12개월)

#### 5. 한글(.hwp/.hwpx) 플러그인 ⭐ 전략적 가치 최고
- `plugins/plugin-protocol.md`에 `.hwpx`/`.hwp`가 **명시적 타깃**, 아직 아무도 미구현
- 한국 공공기관·학교·정부는 거의 전부 한글 사용 → OfficeCLI가 못 건드리는 시장
- `.hwpx`는 **OOXML 유사 XML+ZIP** 구조로 접근 가능 (`.hwp`는 바이너리 복합문서로 훨씬 어려움 —
  `Core/CompoundFile.cs` 존재가 힌트)
- **플러그인은 Apache 2.0 본체 밖 → 독자(유료) 라이선스 판매 가능** (프로토콜 문서의 설계 동기)
- **수익**: 기업 라이선스 ₩300~1,000만원/년 또는 SaaS 종량제
- **투입** 6개월~1년 · **난이도** ⭐⭐⭐⭐⭐ · **예상** 공공 SI 연계 시 **연 1억+**
- **현실적 1단계**: `.hwpx` → `.docx` **단방향 변환만** 먼저 출시

#### 6. 프리미엄 템플릿 / 스킬 마켓플레이스
- 산업별 PPT 스타일팩(의료·법률·금융·교육·제조) ₩49,000/팩
- 기업 CI 적용 커스텀 스킬 ₩200~500만원 (B2B)
- 업종별 재무모델(SaaS·제조·유통·부동산) ₩99,000
- "한국형 IR덱" 스킬 (국내 VC 취향) ₩149,000
- **투입** 2~4개월 · **난이도** ⭐⭐⭐ · **예상** 월 200~1,000만원 · 한계비용 0

#### 7. 데스크톱 앱 (GUI 래퍼)
- ⚠️ 원본 팀이 이미 [AionUi](https://github.com/iOfficeAI/AionUi) 보유 → **정면 경쟁 불리**
- 차별화: **한국어 완전 현지화 + 한국 양식 내장** / 직군 특화(회계사·학원장·공무원) /
  **오프라인 전용**(보안 민감 기업·공공 — 문서 외부 유출 없음이 강력한 셀링포인트)
- **가격** ₩99,000 영구 또는 ₩9,900/월 · **투입** 3~6개월 · **난이도** ⭐⭐⭐⭐ · **예상** 월 300~2,000만원

#### 8. API 서비스 (개발자용 B2B)
```
POST /v1/documents
{ "template": "invoice", "format": "docx", "data": {...} }
→ { "url": "https://.../out.docx" }
```
- **타깃**: ERP·CRM·HR SaaS 업체 (송장·계약서·급여명세서 자동생성)
- **가격** ₩50/건 또는 월 10,000건 ₩290,000
- **투입** 2~3개월 · **난이도** ⭐⭐⭐ · **예상** 고객 20곳 × 월 50만원 = **월 1,000만원**
- **장점**: B2B → 이탈률 낮고 매출 안정
- **경쟁**: DocRaptor / Documentero / CraftMyPDF (해외) → **한국어 폰트·양식 대응이 틈새**

---

### 🥉 티어 3 — 장기 / 대형

#### 9. 수직 통합(버티컬) SaaS — 가장 방어력 높음
| 버티컬 | 제품 | 가격 |
|---|---|---|
| 회계·세무 | 결산보고서·재무제표 자동생성 | ₩99,000/월 |
| 학원·교육 | 성적표·학습보고서·학부모 리포트 대량생성 | ₩49,000/월 |
| 부동산 | 물건 소개서·시장분석 리포트 | ₩79,000/월 |
| 병원 | 진료보고서·건강검진 결과지 | ₩149,000/월 |
| 변호사·법무 | 계약서·의견서 (조항 라이브러리) | ₩199,000/월 |
| 정부지원사업 | 사업계획서 작성 지원 (창업패키지·R&D) | ₩59,000/월 |
- **예상** 단일 버티컬 **연 1~10억** · **난이도** ⭐⭐⭐⭐⭐ (도메인 전문성 필수 → 파트너 확보가 빠름)

#### 10. 기업 내부 자동화 플랫폼 (엔터프라이즈)
온프레미스 + 사내 템플릿 관리 + SSO + 감사로그.
**"문서가 외부로 나가지 않음"** 이 핵심 셀링포인트 (금융·공공·의료).
라이선스 **연 3,000만~2억** · 난이도 ⭐⭐⭐⭐⭐

#### 11. 기존 SaaS에 기능으로 끼워넣기
운영 중인 서비스에 **"보고서 PPT 내보내기"를 프리미엄 기능으로** 추가.
가장 적은 투입으로 ARPU 상승. 난이도 ⭐⭐

---

### 📊 종합 비교
| # | 아이디어 | 투입 | 난이도 | 월 예상 | 시작 추천 |
|---|---|---|---|---|---|
| 3 | **교육 콘텐츠** | 0원 | ⭐⭐ | 100~1,000만 | 🥇 즉시 |
| 4 | **자동화 대행** | 0원 | ⭐⭐⭐ | 300~3,000만 | 🥇 즉시 |
| 1 | **한국형 템플릿 SaaS** | 1~2개월 | ⭐⭐ | 495만~ | 🥈 1개월 후 |
| 8 | API 서비스 | 2~3개월 | ⭐⭐⭐ | 1,000만 | 🥈 |
| 6 | 템플릿 마켓 | 2~4개월 | ⭐⭐⭐ | 200~1,000만 | 🥈 |
| 2 | AI 보고서 SaaS | 2~3개월 | ⭐⭐⭐⭐ | 2,900만 | 🥉 |
| 7 | 데스크톱 앱 | 3~6개월 | ⭐⭐⭐⭐ | 300~2,000만 | 🥉 |
| 5 | **한글 플러그인** | 6~12개월 | ⭐⭐⭐⭐⭐ | 연 1억+ | 🥉 전략적 |
| 9 | 버티컬 SaaS | 6~12개월 | ⭐⭐⭐⭐⭐ | 연 1~10억 | 장기 |
| 10 | 엔터프라이즈 | 12개월+ | ⭐⭐⭐⭐⭐ | 연 3천만~2억 | 장기 |

> ⚠️ 금액은 시장 상식에 기반한 **추정치**이며 보장치가 아닙니다. 실행력·마케팅·타이밍에 따라 크게 달라집니다.

### 🎯 추천 실행 로드맵
```
[1개월]   유튜브 5편 + 자동화 대행 1건 수주
          → 리스크 0, 현금흐름 + 신뢰 + 시장 피드백 동시 확보
[2~3개월] 한국형 템플릿 SaaS MVP (양식 10종)
          → 유튜브 시청자가 그대로 첫 고객
[4~6개월] 온라인 강의 출시 + 템플릿 50종 확장 + API 서비스 베타
          → 매출 다각화
[7~12개월] 버티컬 하나 집중 OR .hwpx 플러그인 착수
          → 방어 가능한 moat 구축
```

### 핵심 원칙 4가지
1. **교육 + 대행으로 시작** — 선투자 0, 즉시 현금, 시장 학습
2. **템플릿 / 스킬이 진짜 moat** — OfficeCLI는 누구나 쓸 수 있음.
   차별점은 **당신이 쌓은 템플릿과 도메인 지식**
3. **`merge` 패턴으로 LLM 비용 통제** — 매번 AI 생성하면 마진 소멸. 설계 1회 + 채우기 N회
4. **한국 현지화가 가장 큰 빈자리** — 한국 양식, 한국어 폰트·조판, `.hwp`, 정부지원사업.
   글로벌 플레이어가 영원히 안 채울 자리

### 리스크 (정직하게)
- 원본 팀이 직접 SaaS/GUI 출시 가능 (AionUi 존재) → **현지화·버티컬로 회피**
- Microsoft Copilot이 이 영역 흡수 가능 → **Copilot이 못 하는 것**(`.hwp`, 대량생성, 온프레미스,
  커스텀 양식)에 집중
- 오픈소스이므로 경쟁자도 동일 도구 사용 가능 → **콘텐츠 / 도메인 / 고객관계**로 쌓기

---

## 13. 빠른 참조 (치트시트)

```bash
# 설치 & 확인
curl -fsSL https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.sh | bash
officecli --version
officecli install                       # 바이너리 + 스킬 + MCP 자동 설치

# 생성 → 편집 → 확인
officecli create deck.pptx
officecli add deck.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli view deck.pptx outline
officecli view deck.pptx screenshot -o check.png
officecli view deck.pptx issues --json
officecli validate deck.pptx

# 라이브 미리보기
officecli watch deck.pptx               # http://localhost:26315

# 템플릿 대량 생성
officecli merge template.docx out-001.docx --data '{"client":"Acme"}'

# 기존 문서 복제 학습
officecli dump existing.pptx -o blueprint.json
officecli batch new.pptx --input blueprint.json

# 상주 모드 (빠른 연속 편집)
officecli open report.docx && officecli set report.docx ... && officecli close report.docx

# AI 연동
officecli mcp claude                    # MCP 등록
curl -fsSL https://officecli.ai/SKILL.md -o ~/.claude/skills/officecli.md

# 모를 때
officecli help pptx set shape
officecli help docx paragraph --json

# 업데이트 끄기
officecli config autoUpdate false
```

---

## 14. 라이선스 고지

```
OfficeCLI
Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
Created and maintained by goworm.

Licensed under the Apache License, Version 2.0.
http://www.apache.org/licenses/LICENSE-2.0

원본 저장소: https://github.com/iOfficeAI/OfficeCLI
이 저장소(포크): https://github.com/bmshin94/OfficeCLI
```

> Apache License 2.0 4조에 따라 재배포 시 저작권 고지와 `NOTICE` 파일을 유지해야 합니다.
> 5조에 따라 상표(로고·이름) 사용권은 라이선스에 포함되지 않습니다.
