# Dooray Wiki Export Tool

Dooray Wiki 페이지를 로컬 파일로 백업하는 스크립트입니다.

프로젝트를 선택한 뒤 위키 페이지 본문, 메타데이터, 첨부 파일, 인라인 이미지를 내려받아 로컬에 저장합니다.

## Requirements

- Python 3
- `requests`
- `selenium`
- Chrome
- 원격 디버깅 모드로 실행된 Chrome 브라우저

## config.json

`config.json`은 예제 템플릿입니다. 실제 실행 전 플레이스홀더를 자신의 환경 값으로 바꿔야 합니다.

```json
{
    "token": "{YOUR_DOORAY_API_TOKEN}",
    "base_url": "https://api.dooray.com/wiki/v1",
    "domain": "https://{YOUR_SUBDOMAIN}.dooray.com",
    "page_limit": -1
}
```

- `token`: Dooray API 토큰
- `base_url`: Dooray Wiki API 기본 주소
- `domain`: Dooray 서비스 도메인
- `page_limit`: 백업할 페이지 수 제한 (`-1`이면 전체)

`project_id`, `wiki_id`는 실행 중 프로젝트를 선택한 뒤 코드에서 채워집니다.

## Run

1. `config.json`의 플레이스홀더를 실제 값으로 바꿉니다.
2. Chrome을 원격 디버깅 모드로 실행합니다.
   - 예: `run_chrome_debug.bat`
3. 의존성을 설치합니다.

```bash
pip install requests selenium
```

4. 스크립트를 실행합니다.

```bash
python dooray_wiki_backup.py
```

5. 백업할 프로젝트를 선택합니다.

결과물은 `backups/` 아래에 생성되고, 로그는 `logs/` 아래에 저장됩니다.

## Notes

- 이 도구는 실행 중인 Chrome에 `127.0.0.1:9222`로 연결합니다.
- 다운로드 파일은 기본적으로 Windows `Downloads` 폴더를 잠깐 사용합니다.
- 현재 요청은 SSL 검증을 끄고 수행합니다.
- 실행 과정에서 `backups/`, `downloads/`, `logs/` 폴더가 생성됩니다.
