import requests
import os
from datetime import datetime
import json
from collections import defaultdict
import logging
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import sys
import time
import shutil
import re
from selenium import webdriver

SCRIPT_DIR = ""

if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 로그 디렉토리 생성
LOGS_DIR = os.path.join(SCRIPT_DIR, 'logs')
os.makedirs(LOGS_DIR, exist_ok=True)

# 로그 파일 경로 설정
log_file = os.path.join(LOGS_DIR, f"wiki_backup_{datetime.now().strftime('%Y%m%d')}.log")

# Logging 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(log_file, encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

@dataclass
class WikiConfig:
    """Wiki API 설정 정보"""
    token: str
    base_url: str
    domain: str
    page_limit: int
    project_id: str = ""
    wiki_id: str = ""

class DoorayAPIClient:
    """Dooray API 호출 클라이언트"""

    def __init__(self, config: WikiConfig):
        self.config = config
        self.headers = {"Authorization": f"dooray-api {config.token}"}

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        """API 요청 공통 처리"""
        if method.upper() == 'POST':
            raise ValueError("POST API는 사용할 수 없습니다.")
        
        # SSL 검증 비활성화 옵션 추가
        kwargs['verify'] = False
        
        # SSL 경고 메시지 무시
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        response = requests.request(method, url, **kwargs)
        response.raise_for_status()
        return response

    def get_projects(self, params: Optional[Dict] = None) -> List[Dict]:
        """프로젝트 목록 조회 (페이징 처리)"""
        url = f"{self.config.base_url.replace('/wiki/v1', '/project/v1')}/projects"
        default_params = {
            #"member": "me",
            "state": "active",
            "size": 20
        }
        if params:
            default_params.update(params)
        
        all_projects = []
        page = 0
        
        while True:
            default_params["page"] = page
            response = self._request('GET', url, headers=self.headers, params=default_params)
            result = response.json()
            
            if not result.get("result"):
                break
                
            all_projects.extend(result["result"])
            total_count = result.get("totalCount", 0)
            
            # 모든 프로젝트를 가져왔는지 확인
            if len(all_projects) >= total_count:
                break
                
            page += 1
        
        return all_projects

    def get_pages(self, parent_page_id: Optional[str] = None) -> Dict:
        """위키 페이지 목록 조회"""
        url = f"{self.config.base_url}/wikis/{self.config.wiki_id}/pages"
        params = {"parentPageId": parent_page_id} if parent_page_id else {}
        
        response = self._request('GET', url, headers=self.headers, params=params)
        return response.json()

    def get_page_content(self, page_id: str) -> Dict:
        """특정 페이지 내용 조회"""
        url = f"{self.config.base_url}/wikis/{self.config.wiki_id}/pages/{page_id}"
        response = self._request('GET', url, headers=self.headers)
        return response.json()
   
class SeleniumDownloader:
    """Selenium을 사용한 파일 다운로더"""

    def __init__(self, config: WikiConfig):
        self.config = config
        self.download_dir = os.path.join(SCRIPT_DIR, 'downloads')
        os.makedirs(self.download_dir, exist_ok=True)
        
        # Chrome 옵션 설정
        self.chrome_options = webdriver.ChromeOptions()
        # 기존 실행 중인 Chrome에 연결
        self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        
        self.driver = None

    def start(self):
        """기존 브라우저에 연결"""
        if self.driver is None:
            try:
                self.driver = webdriver.Chrome(options=self.chrome_options)
                logger.info("기존 Chrome 브라우저에 연결 성공")
            except Exception as e:
                logger.error(f"Chrome 연결 실패. Chrome이 디버그 모드로 실행되어 있는지 확인하세요. 오류: {str(e)}")
                raise

    def download_file(self, file_id: str, file_name: str, is_inline: bool = False) -> str:
        """파일 다운로드"""
        if self.driver is None:
            self.start()
        
        try:
            # 현재 다운로드 폴더의 파일 목록 (다운로드 전)
            downloads_path = os.path.expanduser("~\\Downloads")
            before_files = set(os.listdir(downloads_path))
            
            # 파일 다운로드 URL 생성
            if is_inline:
                file_url = f"{self.config.domain}/wikis/{self.config.wiki_id}/files/{file_id}?disposition=attachment"
            else:
                file_url = f"{self.config.domain}/page-files/{file_id}?disposition=attachment"
                
            logger.info(f"파일 다운로드 페이지 접근: {file_url}")
            self.driver.get(file_url)
            
            # 페이지 로드 대기
            # time.sleep(1)

            try:
                # 다운로드 완료 대기 (최대 10초)
                for _ in range(10):
                    current_files = set(os.listdir(downloads_path))
                    new_files = current_files - before_files
                    
                    if new_files:
                        new_file = list(new_files)[0]
                        # .crdownload 파일은 건너뛰기
                        if new_file.endswith('.crdownload') or new_file.endswith('.tmp'):
                            time.sleep(1)
                            continue
                        downloaded_file = os.path.join(downloads_path, new_file)
                        target_path = os.path.join(self.download_dir, file_name)
                        
                        if os.path.exists(target_path):
                            os.remove(target_path)
                        
                        shutil.move(downloaded_file, target_path)
                        logger.info(f"파일 다운로드 성공: {file_name}")
                        return target_path
                    
                    time.sleep(1)
                
                logger.error(f"파일 다운로드 시간 초과: {file_name}")
                return None
                
            except Exception as e:
                logger.error(f"다운로드 처리 중 실패: {str(e)}")
                return None
                
        except Exception as e:
            logger.error(f"파일 다운로드 실패: {str(e)}")
            return None

    def close(self):
        """브라우저 연결 종료 (브라우저는 닫지 않음)"""
        if self.driver:
            self.driver = None

class PageCounter:
    """페이지 카운터"""

    def __init__(self, limit: int):
        self.count = 0
        self.limit = limit
        self.level_counters = defaultdict(int)

    def increment(self) -> bool:
        """전체 카운트 증가"""
        self.count += 1
        return self.limit == -1 or self.count <= self.limit

    def get_next_number(self, parent_id: str) -> int:
        """각 레벨별 번호 생성"""
        self.level_counters[parent_id] += 1
        return self.level_counters[parent_id]

class WikiBackupManager:
    """위키 백업 관리자"""

    def __init__(self, config: WikiConfig, project_code: str):
        self.config = config
        self.project_code = project_code  # 프로젝트 코드 추가
        self.api_client = DoorayAPIClient(config)
        self.page_counter = PageCounter(config.page_limit)
        self.backup_dir = self._create_backup_dir()
        self.downloader = None

    def _create_backup_dir(self) -> str:
        """백업 디렉토리 생성"""
        # backups 디렉토리 생성
        backups_dir = os.path.join(SCRIPT_DIR, 'backups')
        os.makedirs(backups_dir, exist_ok=True)
        
        # timestamp_projectcode 형식의 디렉토리 생성
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_dir = os.path.join(backups_dir, f"{timestamp}")
        os.makedirs(backup_dir, exist_ok=True)
        
        return backup_dir

    def _create_numbered_dir(self, base_path: str, subject: str, number: int) -> str:
        """번호가 포함된 디렉토리 생성"""
        safe_subject = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in subject)
        dir_name = f"{number:02d} {safe_subject}"
        full_path = os.path.join(base_path, dir_name)
        os.makedirs(full_path, exist_ok=True)
        return full_path
   
    def _process_attachments(self, files: List[Dict], page_path: str) -> List[str]:
        """첨부 파일 처리"""
        if not files:
            return []
        
        logger.info(f"첨부 파일 정보: {json.dumps(files, ensure_ascii=False, indent=2)}")
        
        # attachments 디렉토리 생성
        attachments_dir = os.path.join(page_path, "attachments")
        os.makedirs(attachments_dir, exist_ok=True)
        
        # Selenium 다운로더 초기화 (필요한 경우)
        if self.downloader is None:
            self.downloader = SeleniumDownloader(self.config)
        
        attachment_links = []
        for file in files:
            try:
                original_name = file['name']
                file_id = file['id']
                
                # 파일명과 확장자 분리
                base_name, ext = os.path.splitext(original_name)
                # 공백을 언더바로 치환하고 특수문자 처리
                base_name = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in base_name)
                # 타임스탬프가 포함된 파일명 생성
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
                unique_filename = f"{base_name}_{timestamp}{ext}"
                
                target_path = os.path.join(attachments_dir, unique_filename)
                
                # 파일 다운로드
                logger.info(f"파일 다운로드 시도: {original_name}")
                downloaded_file = self.downloader.download_file(file_id, unique_filename)
                
                if downloaded_file and os.path.exists(downloaded_file):
                    # 다운로드된 파일을 attachments 디렉토리로 이동
                    shutil.move(downloaded_file, target_path)
                    logger.info(f"파일 가져오기 성공: {original_name} -> {unique_filename}")
                    attachment_links.append(f"- [{original_name}](attachments/{unique_filename}) (크기: {file['size']} bytes)\n")
                else:
                    logger.error(f"파일 다운로드 실패: {original_name}")
                    attachment_links.append(f"- {original_name} (다운로드 실패) (크기: {file['size']} bytes)\n")
                    
            except Exception as e:
                logger.error(f"파일 처리 중 오류 발생: {str(e)}")
                attachment_links.append(f"- {file['name']} (처리 오류) (크기: {file['size']} bytes)\n")
        
        return attachment_links

    def _process_inline_images(self, content: str, page_path: str) -> str:
        """인라인 이미지를 처리하고 다운로드합니다."""
        pattern = r'!\[(.*?)\]\(/wikis/(\d+)/files/(\d+)\)'
        
        # images 디렉토리 생성
        images_dir = os.path.join(page_path, "images")
        os.makedirs(images_dir, exist_ok=True)
        
        def replace_image(match) -> str:
            """찾은 이미지를 다운로드하고 경로를 수정합니다."""
            alt_text = match.group(1)
            wiki_id = match.group(2)
            file_id = match.group(3)
            
            # 원본 파일명 사용
            original_filename = alt_text.strip()
            if not original_filename:
                original_filename = f"image_{file_id}.png"
                
            # 파일명과 확장자 분리
            base_name, ext = os.path.splitext(original_filename)
            # 공백을 언더바로 치환하고 특수문자 처리
            base_name = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in base_name)
            # 타임스탬프가 포함된 파일명 생성
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            unique_filename = f"{base_name}_{timestamp}{ext}"
            
            try:
                # Selenium 다운로더 초기화 (필요한 경우)
                if self.downloader is None:
                    self.downloader = SeleniumDownloader(self.config)
                
                target_path = os.path.join(images_dir, unique_filename)
                logger.info(f"인라인 이미지 다운로드 시도: {original_filename} -> {unique_filename} (Wiki ID: {wiki_id}, File ID: {file_id})")
                
                downloaded_file = self.downloader.download_file(file_id, unique_filename, is_inline=True)

                if downloaded_file and os.path.exists(downloaded_file):
                    if os.path.exists(target_path):
                        os.remove(target_path)
                    shutil.move(downloaded_file, target_path)
                    logger.info(f"인라인 이미지 다운로드 성공: {original_filename} -> {unique_filename}")
                    # 마크다운에서 사용할 상대 경로로 변환
                    return f"![{alt_text}](images/{unique_filename})"
                else:
                    logger.error(f"인라인 이미지 다운로드 실패: {original_filename}")
                    return f"![{alt_text}] (다운로드 실패 - Wiki ID: {wiki_id}, File ID: {file_id})"
                    
            except Exception as e:
                logger.error(f"인라인 이미지 처리 중 오류 발생: {str(e)}")
                return f"![{alt_text}] (처리 오류 - Wiki ID: {wiki_id}, File ID: {file_id})"
        
        # 모든 인라인 이미지 처리
        processed_content = re.sub(pattern, replace_image, content)
        return processed_content

    def _save_page(self, page_data: Dict, content_data: Dict, page_path: str):
        """페이지 정보 저장"""
        # result에서 body 가져오기
        body = content_data.get("result", {}).get("body", {})
        content = body.get("content", "")
        mime_type = body.get("mimeType", "")
        files = content_data.get("result", {}).get("files", [])
        
        # 페이지 제목 추가
        title = page_data.get("subject", "")
        content = f"# {title}\n\n{content}"
        
        # 인라인 이미지 처리
        content = self._process_inline_images(content, page_path)
        
        # 첨부 파일 링크 추가
        attachment_links = self._process_attachments(files, page_path)
        if attachment_links:
            content += "\n\n## 첨부 파일\n" + "".join(attachment_links)
        
        # 메타데이터 저장
        metadata = {
            "id": page_data["id"],
            "subject": title,
            "mimeType": mime_type,
            "createdAt": content_data["result"].get("createdAt", ""),
            "parentPageId": page_data.get("parentPageId"),
            "attachments": files
        }
        
        with open(os.path.join(page_path, "metadata.json"), "w", encoding='utf-8') as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
        
        # 내용 저장
        with open(os.path.join(page_path, "content.md"), "w", encoding='utf-8') as f:
            f.write(content)

    def backup_recursive(self, parent_page_id: Optional[str], base_path: str):
        """재귀적으로 위키 페이지 백업"""
        try:
            if self.page_counter.count >= self.page_counter.limit and self.page_counter.limit != -1:
                return

            pages = self.api_client.get_pages(parent_page_id)
            
            for page in pages["result"]:
                if not self.page_counter.increment():
                    return

                current_number = self.page_counter.get_next_number(parent_page_id or 'root')
                logger.info(f"Backing up page {self.page_counter.count}: "
                            f"{current_number:02d} {page['subject']}")
                
                # 페이지 정보 저장
                page_path = self._create_numbered_dir(base_path, page["subject"], current_number)
                page_content = self.api_client.get_page_content(page["id"])
                self._save_page(page, page_content, page_path)
                
                # 하위 페이지 처리
                self.backup_recursive(page["id"], page_path)
                
        except Exception as e:
            logger.error(f"백업 중 오류 발생: {str(e)}")
            raise
       
    def backup(self):
        """백업 실행"""
        try:
            logger.info("Wiki backup started")
            
            # 최상위 페이지 조회
            root_pages = self.api_client.get_pages()
            if not root_pages.get("result"):
                raise ValueError("최상위 페이지를 찾을 수 없습니다.")
            
            root_page = root_pages["result"][0]
            root_content = self.api_client.get_page_content(root_page["id"])
            
            # 최상위 페이지 저장 (project code 사용)
            logger.info(f"Backing up root page: {self.project_code}")
            root_path = self._create_numbered_dir(self.backup_dir, self.project_code, 1)
            
            # 최상위 페이지 정보도 동일한 방식으로 저장
            self.page_counter.increment()
            self.page_counter.get_next_number('root')
            self._save_page(root_page, root_content, root_path)
            
            # 하위 페이지 처리
            self.backup_recursive(root_page["id"], root_path)
            
            logger.info(f"Backup completed in directory: {self.backup_dir}")
            logger.info(f"Total pages backed up: {self.page_counter.count}")
            
        except Exception as e:
            logger.error(f"백업 실패: {str(e)}")
            raise

def select_projects() -> List[Dict]:
    """여러 프로젝트 선택"""
    try:
        config = load_config()
        api_client = DoorayAPIClient(config)
        
        # 프로젝트 목록 조회 (전체)
        print("프로젝트 목록을 가져오는 중...")
        projects = api_client.get_projects()
        
        if not projects:
            print("사용 가능한 프로젝트가 없습니다.")
            return []
        
        print("\n사용 가능한 프로젝트 목록:")
        print("-" * 50)
        for idx, project in enumerate(projects, 1):
            print(f"{idx}. {project['code']}")
        print("-" * 50)
        print(f"총 {len(projects)}개의 프로젝트가 있습니다.")
        print("\n옵션:")
        print("- 쉼표로 구분된 번호 입력: 해당 프로젝트들 선택")
        print("- a 또는 all 입력: 모든 프로젝트 선택")
        print("- 0 입력: 종료")
        
        selected_projects = []
        while True:
            try:
                choice = input("\n백업할 프로젝트를 선택하세요: ").strip().lower()
                if choice == "0":
                    break
                
                # 모든 프로젝트 선택
                if choice in ('a', 'all'):
                    selected_projects = [p for p in projects if "wiki" in p and p["wiki"].get("id")]
                    if not selected_projects:
                        print("위키가 있는 프로젝트가 없습니다.")
                        continue
                    print(f"\n위키가 있는 모든 프로젝트가 선택되었습니다 ({len(selected_projects)}개):")
                    for p in selected_projects:
                        print(f"- {p['code']}")
                else:
                    # 개별 프로젝트 선택
                    for num in choice.split(","):
                        try:
                            idx = int(num.strip())
                            if 1 <= idx <= len(projects):
                                project = projects[idx - 1]
                                if "wiki" not in project or not project["wiki"].get("id"):
                                    print(f"프로젝트 {project['code']}에 위키가 없습니다.")
                                    continue
                                if project not in selected_projects:
                                    selected_projects.append(project)
                                    print(f"프로젝트 {project['code']} 추가됨")
                            else:
                                print(f"잘못된 번호입니다: {idx}")
                        except ValueError:
                            print(f"잘못된 입력값: {num}")
                
                if selected_projects:
                    confirm = input("\n선택된 프로젝트들을 백업하시겠습니까? (y/n): ")
                    if confirm.lower() == 'y':
                        break
                    selected_projects = []  # 다시 선택
                    
            except Exception as e:
                print(f"입력 처리 중 오류 발생: {str(e)}")
        
        return selected_projects
            
    except Exception as e:
        logger.error(f"프로젝트 선택 중 오류 발생: {str(e)}")
        return []

def load_config() -> WikiConfig:
    """설정 파일 로드"""
    try:
        config_path = os.path.join(SCRIPT_DIR, 'config.json')
        if not os.path.exists(config_path):
            logger.error(f"config.json 파일이 없습니다. 다음 경로에 생성해주세요: {config_path}")
            logger.error("필요한 설정: token, base_url, domain, page_limit")
            raise FileNotFoundError(f"config.json not found at {config_path}")
            
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        return WikiConfig(**config_data)
    except FileNotFoundError:
        logger.error(f"config.json 파일을 찾을 수 없습니다. 경로: {config_path}")
        raise
    except json.JSONDecodeError:
        logger.error(f"config.json 파일 형식이 잘못되었습니다. 경로: {config_path}")
        raise
    except Exception as e:
        logger.error(f"설정 로드 중 오류 발생: {str(e)}")
        raise

def main():
    try:
        # 여러 프로젝트 선택 (또는 전체 선택)
        selected_projects = select_projects()
        if not selected_projects:
            print("프로젝트가 선택되지 않았습니다.")
            return
        
        print(f"\n총 {len(selected_projects)}개의 프로젝트를 백업합니다.")
        
        # 설정 로드
        config = load_config()
        
        # 각 프로젝트 백업
        for project in selected_projects:
            print(f"\n프로젝트 {project['code']} 백업 시작...")
            
            # 프로젝트의 wiki id로 config 업데이트
            config.project_id = project['id']
            config.wiki_id = project['wiki']['id']
            
            # 백업 실행
            backup_manager = WikiBackupManager(config, project['code'])
            backup_manager.backup()
            
            print(f"프로젝트 {project['code']} 백업 완료")
        
        print("\n모든 프로젝트 백업이 완료되었습니다.")
        
    except Exception as e:
        logger.error(f"프로그램 실행 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    main()