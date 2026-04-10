"""
소프트웨어 검사결과 자동화 스크립트
====================================
1. 구글 시트에서 폼 응답(제출내역)과 제출 대상자 목록 조회
2. 미제출자 검출 → Jandi 웹훅 독촉 메시지 발송
3. 구글 드라이브 API로 .ipt 파일을 다운로드하여 제출자별 ZIP 압축
4. ZIP 파일을 GitHub Releases로 업로드 (로그인 없이 다운로드 가능)

변경사항 (GitHub Actions / Public 저장소 대응):
- 모든 민감정보(SHEET_URL, JANDI_WEBHOOK_URL 등) → 환경변수로 분리
- ZIP 업로드: Google Drive API → GitHub Releases (로그인 없이 다운로드 가능)
- OUTPUT_DIR → /tmp 임시 경로 사용
- 미제출자 있을 때 sys.exit(2) 반환 → Actions 재시도 트리거용

파일명 패턴: {장비코드}_{이름}_{YYYYMMDD}.ipt
  예) 개21-06_손만식_20260303.ipt → 장비코드=개21-06, 이름=손만식, 날짜=20260303
"""

import os
import re
import time
import zipfile
import argparse
import logging
import requests
import gspread
import sys
import tempfile
from datetime import datetime, timedelta
import shutil

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

# ==========================================
# ⚙️ 설정 (Configuration) — 모두 환경변수로 관리
# ==========================================
# 1. 시트 정보 — GitHub Secret: SHEET_URL
CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credentials.json')
SHEET_URL     = os.environ.get('SHEET_URL', '')
TAB_RESPONSES = '설문지 응답 시트1'
TAB_MEMBERS   = '제출 대상자'

# 2. Google Drive 소스 폴더 ID — GitHub Secret: GDRIVE_SOURCE_FOLDER_ID
#    Drive URL에서 확인: https://drive.google.com/drive/folders/{FOLDER_ID}
GDRIVE_SOURCE_FOLDER_ID = os.environ.get('GDRIVE_SOURCE_FOLDER_ID', '')

# 3. 잔디 웹훅 — GitHub Secret: JANDI_WEBHOOK_URL
JANDI_WEBHOOK_URL = os.environ.get('JANDI_WEBHOOK_URL', '')

# 4. 구글 폼 URL — GitHub Secret: FORM_URL (독촉 메시지 링크용)
FORM_URL = os.environ.get('FORM_URL', '')

# 5. GitHub Releases 업로드용 — Actions에서 자동 제공되므로 별도 Secret 불필요
#    GITHUB_TOKEN: secrets.GITHUB_TOKEN으로 yml에서 주입
#    GITHUB_REPOSITORY: "owner/repo" 형식 (Actions 환경에서 자동 제공)
GITHUB_TOKEN      = os.environ.get('GITHUB_TOKEN', '')
GITHUB_REPOSITORY = os.environ.get('GITHUB_REPOSITORY', '')

# 6. 임시 출력 경로 (GitHub Actions: /tmp, 로컬: 시스템 tempdir)
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), 'sw_inspector_output')

# 7. .ipt 파일명 파싱 패턴
#    날짜(8자리) 이후 접미사는 개수 제한 없이 허용
#    예) 공26-13_손만식_20260425 - test - Mansik Sohn.ipt
IPT_FILENAME_PATTERN = re.compile(
    r'^(?P<equip_code>[가-힣a-zA-Z0-9]+(?:-\d+)?\d*)_(?P<n>[가-힣]+)_(?P<date>\d{8})(?:\s*-\s*.+)?\.ipt$'
)
# 날짜(8자리) 이후의 모든 접미사 제거 → 원본 파일명으로 복원
GDRIVE_SUFFIX_PATTERN = re.compile(r'(_\d{8}).*$')

# ==========================================
# 로깅 설정
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)


# ==========================================
# 🔑 Google API 클라이언트 공통 인증
# ==========================================
def get_google_credentials():
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/drive',
    ]
    return Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)


def get_drive_service():
    creds = get_google_credentials()
    return build('drive', 'v3', credentials=creds)


# ==========================================
# 🚀 1. GSpread 클라이언트 로그인
# ==========================================
def get_gsheet_client():
    logger.info("▶ 구글 시트에 연결 중...")
    if not SHEET_URL:
        logger.error("❌ SHEET_URL 환경변수가 설정되지 않았습니다.")
        return None
    try:
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
        doc = gc.open_by_url(SHEET_URL)
        logger.info("✅ 구글 시트 연결 성공")
        return doc
    except Exception as e:
        logger.error(f"❌ 구글 시트 접근 실패. credentials.json 파일과 시트 공유 권한을 확인하세요.\n({e})")
        return None


# ==========================================
# 🚀 2. 제출 대상자 목록 가져오기
# ==========================================
def get_target_members(doc):
    member_sheet = doc.worksheet(TAB_MEMBERS)
    members = member_sheet.get_all_records()

    target_members = []
    for member in members:
        name = str(member.get('이름', '')).strip()
        eligible_raw = str(member.get('대상여부', '')).strip().upper()
        is_eligible = eligible_raw in ('O', 'TRUE', '예', 'Y', '1', 'YES')
        if name and is_eligible:
            target_members.append(name)

    logger.info(f"📋 제출 대상자: {len(target_members)}명 → {target_members}")
    return target_members


# ==========================================
# 🚀 3. 이번 달 폼 응답 건수 확인
# ==========================================
def get_form_submission_count(doc, target_yyyymm):
    resp_sheet = doc.worksheet(TAB_RESPONSES)
    responses = resp_sheet.get_all_records()

    count = 0
    for row in responses:
        timestamp_str = str(row.get('타임스탬프', '')).strip()
        if not timestamp_str:
            continue
        try:
            digits_only = re.sub(r'[^0-9.]', '', timestamp_str)
            parts = digits_only.split('.')
            if len(parts) >= 2:
                row_yyyymm = f"{parts[0]}{int(parts[1]):02d}"
                if target_yyyymm and row_yyyymm != target_yyyymm:
                    continue
        except Exception:
            pass
        file_url = str(row.get('결과 파일 제출', '')).strip()
        if file_url:
            count += 1
    return count


# ==========================================
# 🚀 4. Google Drive에서 .ipt 파일 목록 조회
# ==========================================
def list_drive_ipt_files(folder_id):
    service = get_drive_service()
    query = f"'{folder_id}' in parents and name contains '.ipt' and trashed = false"
    results = service.files().list(
        q=query,
        fields='files(id, name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        pageSize=200
    ).execute()
    files = results.get('files', [])
    logger.info(f"📂 Drive .ipt 파일: {len(files)}개")
    return files


def download_drive_file(service, file_id, dest_path):
    request = service.files().get_media(
        fileId=file_id,
        supportsAllDrives=True
    )
    with open(dest_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def get_or_create_archive_folder(service, parent_folder_id):
    """소스 폴더 하위에 Archive 폴더를 찾거나 없으면 생성합니다."""
    try:
        query = (
            f"'{parent_folder_id}' in parents"
            " and name = 'Archive'"
            " and mimeType = 'application/vnd.google-apps.folder'"
            " and trashed = false"
        )
        results = service.files().list(
            q=query,
            fields='files(id, name)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        folders = results.get('files', [])

        if folders:
            logger.info(f"📂 Archive 폴더 확인: {folders[0]['id']}")
            return folders[0]['id']

        # 없으면 생성
        folder_metadata = {
            'name': 'Archive',
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = service.files().create(
            body=folder_metadata,
            fields='id',
            supportsAllDrives=True
        ).execute()
        logger.info(f"📂 Archive 폴더 생성 완료: {folder['id']}")
        return folder['id']

    except Exception as e:
        logger.warning(f"⚠️ Archive 폴더 생성/조회 실패: {e}")
        return None


# ==========================================
# 🚀 5. 미제출자 체크
# ==========================================
def check_unsubmitted(doc, target_yyyymm, folder_id=None):
    folder_id = folder_id or GDRIVE_SOURCE_FOLDER_ID
    logger.info(f"\n▶ {target_yyyymm} 미제출자 확인 중...")
    logger.info(f"📂 Drive 소스 폴더 ID: {folder_id}")

    target_members = get_target_members(doc)
    form_count = get_form_submission_count(doc, target_yyyymm)
    logger.info(f"📩 시트 응답 건수 ({target_yyyymm}): {form_count}건")

    drive_files = list_drive_ipt_files(folder_id)
    submitted_names = set()
    parsed_count = 0

    for f in drive_files:
        parsed = parse_ipt_filename(f['name'])
        if parsed:
            submitted_names.add(parsed['n'])
            parsed_count += 1

    logger.info(f"📂 Drive .ipt 파일: {len(drive_files)}개 (파싱 성공: {parsed_count}개)")
    logger.info(f"✅ 제출 확인: {len(submitted_names)}명 — {sorted(list(submitted_names))}")

    unsubmitted = [name for name in target_members if name not in submitted_names]
    logger.info(f"❌ 미제출자: {len(unsubmitted)}명 — {unsubmitted}")
    return unsubmitted


# ==========================================
# 🚀 6. Jandi 웹훅 — 미제출 독촉
# ==========================================
def send_jandi_reminder(unsubmitted, target_yyyymm):
    if not JANDI_WEBHOOK_URL:
        logger.warning("⚠️ JANDI_WEBHOOK_URL이 설정되지 않았습니다.")
        return False
    if not unsubmitted:
        logger.info("✅ 미제출자가 없어 Jandi 메시지를 발송하지 않습니다.")
        return True

    today_str = datetime.now().strftime("%m/%d")
    names_text = ', '.join([f"@{name}" for name in unsubmitted]) + " 님"

    payload = {
        "body": "소프트웨어 검사 현황 업데이트",
        "connectColor": "#FF6600",
        "connectInfo": [
            {
                "title": f"🔍 소프트웨어 검사 현황 업데이트 ({today_str})",
                "description": (
                    f"현재까지 리스트에 확인되지 않은 분이 총 {len(unsubmitted)}분 계십니다.\n"
                    "바쁘시겠지만 잠시만 시간을 내어 마무리를 부탁드려요!\n\n"
                    f"미제출자: {names_text}\n"
                    "(맥 사용자는 제외했습니다. 최신 검사기가 동작하는지 확인해주세요)\n\n"
                    f"📍 [여기서 바로 제출하기]({FORM_URL})\n"
                    "💡 이미 제출 하셨다면 답글 남겨주세요"
                )
            }
        ]
    }

    try:
        headers = {
            'Accept': 'application/vnd.tosslab.jandi-v2+json',
            'Content-Type': 'application/json'
        }
        resp = requests.post(JANDI_WEBHOOK_URL, json=payload, headers=headers, timeout=10)
        resp.raise_for_status()
        logger.info(f"✅ Jandi 미제출 독촉 메시지 발송 완료 (대상: {len(unsubmitted)}명)")
        return True
    except Exception as e:
        logger.error(f"❌ Jandi 메시지 발송 실패: {e}")
        return False


# ==========================================
# 🚀 7. Jandi 웹훅 — 업로드 완료 알림
# ==========================================
def send_jandi_upload_link(link_url, target_yyyymm, original_filename=None, excluded_files=None):
    if not JANDI_WEBHOOK_URL:
        return False

    display_month = f"{target_yyyymm[:4]}년 {int(target_yyyymm[4:])}월"
    rename_note = f"\n📝 다운로드 후 파일명을 `{original_filename}`으로 변경 후 제출해주세요." if original_filename else ""

    # 제외된 파일이 있을 경우 경고 섹션 추가
    excluded_note = ""
    if excluded_files:
        file_list = "\n".join([f"  - {f['name']}" for f in excluded_files])
        excluded_note = (
            f"\n\n⚠️ **아래 파일은 처리에서 제외됐으니 확인이 필요합니다** ({len(excluded_files)}개)\n"
            f"{file_list}\n"
            "(비대상자 파일, 파싱 불가 파일, 또는 중복 파일일 수 있습니다.)"
        )

    payload = {
        "body": "소프트웨어 검사 마무리",
        "connectColor": "#00C300" if not excluded_files else "#FFA500",
        "connectInfo": [
            {
                "title": f"🎉 {display_month} 소프트웨어 검사 취합 완료",
                "description": (
                    "이번 달 점검도 무사히 끝났습니다. 모두 고생 많으셨습니다!\n"
                    "제출해주신 모든 검사 결과 압축본이 준비되었습니다.\n\n"
                    f"📦 [ZIP 파일 다운로드]({link_url})\n"
                    f"{rename_note}\n\n"
                    "📌 **최종 제출 방법 (참조용)**\n"
                    "압축 파일을 다운로드한 후, [소프트웨어 검사제출 게시판](https://gw.mailplug.com/board/24445)의 안내에 따라 게시글에 답글로 제출합니다."
                    f"{excluded_note}"
                )
            }
        ]
    }

    try:
        headers = {
            'Accept': 'application/vnd.tosslab.jandi-v2+json',
            'Content-Type': 'application/json'
        }
        resp = requests.post(JANDI_WEBHOOK_URL, json=payload, headers=headers, timeout=10)
        resp.raise_for_status()
        logger.info("✅ Jandi 완료 알림 및 링크 발송 완료")
        return True
    except Exception as e:
        logger.error(f"❌ Jandi 파일 제출완료 메시지 발송 실패: {e}")
        return False


# ==========================================
# 🚀 8. .ipt 파일명 파싱
# ==========================================
def parse_ipt_filename(filename):
    match = IPT_FILENAME_PATTERN.match(filename)
    if match:
        return match.groupdict()
    return None


# ==========================================
# 🚀 9. 파일명 정리 (구글 드라이브 접미사 제거)
# ==========================================
def clean_filename(basename):
    """날짜(8자리) 이후의 모든 접미사를 제거합니다.
    
    예) 공26-13_손만식_20260425 - test - Mansik Sohn.ipt
      → 공26-13_손만식_20260425.ipt
    """
    name_part, ext = os.path.splitext(basename)
    cleaned = GDRIVE_SUFFIX_PATTERN.sub(r'\1', name_part)
    return cleaned + ext


# ==========================================
# 🚀 10. Drive에서 .ipt 파일 다운로드 후 ZIP 압축
# ==========================================
def create_submission_zip(doc, target_yyyymm, folder_id=None,
                          team_name='AI에듀테크', submitter='손만식',
                          dry_run=False):
    """Drive .ipt 파일을 필터링하여 ZIP으로 압축합니다.
    
    - 대상자 목록에 없는 파일은 제외 (경고)
    - 동일인 파일이 여러 개면 가장 최신 날짜 파일만 포함 (나머지 경고)
    - 처리에서 제외된 파일은 Archive로 이동하지 않음
    - 반환값: (zip_path, excluded_files) 튜플
    """
    folder_id = folder_id or GDRIVE_SOURCE_FOLDER_ID
    logger.info(f"\n▶ 파일 {'미리보기(dry-run)' if dry_run else '압축'} 시작...")
    logger.info(f"📂 Drive 소스 폴더 ID: {folder_id}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    drive_files = list_drive_ipt_files(folder_id)
    if not drive_files:
        logger.warning("⚠️ .ipt 파일이 없습니다. Drive 폴더 ID를 확인하세요.")
        return None, []

    # ── 파일 분류 ──────────────────────────────────────────
    target_members = get_target_members(doc)
    target_set = set(target_members)

    unparseable   = []   # 파일명 파싱 실패
    non_target    = []   # 대상자 목록에 없는 사람
    by_name       = {}   # {이름: [(날짜, 파일), ...]}

    for f in drive_files:
        parsed = parse_ipt_filename(f['name'])
        if not parsed:
            logger.warning(f"⚠️ 파일명 파싱 실패 → 제외: {f['name']}")
            unparseable.append(f)
            continue
        name = parsed['n']
        date = parsed['date']
        if name not in target_set:
            logger.warning(f"⚠️ 대상자 아님 → 제외: {f['name']} (이름: {name})")
            non_target.append(f)
            continue
        by_name.setdefault(name, []).append((date, f))

    # 중복 처리: 최신 날짜 1개만 유지
    valid_files       = []
    skipped_duplicate = []
    for name, entries in by_name.items():
        entries.sort(key=lambda x: x[0], reverse=True)  # 날짜 내림차순
        valid_files.append(entries[0][1])                # 최신 파일 선택
        for _, f in entries[1:]:
            logger.warning(f"⚠️ 중복 파일 (최신 아님) → 제외: {f['name']} (이름: {name})")
            skipped_duplicate.append(f)

    excluded_files = unparseable + non_target + skipped_duplicate

    # ── 요약 로그 ──────────────────────────────────────────
    target_zip_name = f"{target_yyyymm}_{team_name}_{submitter}.zip"
    target_zip_path = os.path.join(OUTPUT_DIR, target_zip_name)

    logger.info(f"\n{'📋 미리보기' if dry_run else '📦 압축 대상'}: {target_zip_name}")
    logger.info(f"  ✅ 포함: {len(valid_files)}개 / 제외: {len(excluded_files)}개")
    for f in valid_files:
        cleaned = clean_filename(f['name'])
        label = f" → {cleaned}  (접미사 제거)" if f['name'] != cleaned else ""
        logger.info(f"    ↳ {f['name']}{label}")
    if excluded_files:
        logger.info("  ⚠️ 제외 파일:")
        for f in excluded_files:
            logger.info(f"    ✖ {f['name']}")

    if dry_run:
        logger.info(f"\n🔍 미리보기 완료: {len(valid_files)}개 파일 → {target_zip_name}")
        return len(valid_files), excluded_files

    if not valid_files:
        logger.warning("⚠️ 압축할 유효한 파일이 없습니다.")
        return None, excluded_files

    service = get_drive_service()
    tmp_dir = tempfile.mkdtemp(prefix='sw_inspector_dl_')

    try:
        with zipfile.ZipFile(target_zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for f in valid_files:
                tmp_path = os.path.join(tmp_dir, f['name'])
                logger.info(f"  ⬇️  다운로드: {f['name']}")
                download_drive_file(service, f['id'], tmp_path)
                cleaned = clean_filename(f['name'])
                zipf.write(tmp_path, cleaned)

        logger.info(f"\n✅ 압축 완료: {target_zip_name} ({len(valid_files)}개 파일)")
        logger.info(f"📁 결과물: {target_zip_path}")

        # ── Archive 이동 (처리된 파일만) ────────────────────
        archive_folder_id = get_or_create_archive_folder(service, folder_id)
        if archive_folder_id:
            archived = 0
            for f in valid_files:   # excluded_files는 이동하지 않음
                try:
                    service.files().update(
                        fileId=f['id'],
                        addParents=archive_folder_id,
                        removeParents=folder_id,
                        supportsAllDrives=True,
                        fields='id'
                    ).execute()
                    archived += 1
                except Exception as e:
                    logger.warning(f"⚠️ Archive 이동 실패 ({f['name']}): {e}")
            logger.info(f"📂 {archived}개 파일을 Archive 폴더로 이동했습니다.")
            if excluded_files:
                logger.info(f"📂 제외된 {len(excluded_files)}개 파일은 소스 폴더에 유지합니다.")

        return target_zip_path, excluded_files

    except Exception as e:
        logger.error(f"❌ 압축 실패: {e}")
        return None, excluded_files
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ==========================================
# 🚀 11. GitHub Releases 업로드
#   - Public 저장소에서 로그인 없이 누구나 다운로드 가능
#   - GITHUB_TOKEN, GITHUB_REPOSITORY는 Actions에서 자동 제공
# ==========================================
def upload_to_github_release(file_path, target_yyyymm):
    """ZIP 파일을 GitHub Release로 업로드하고 다운로드 URL을 반환합니다."""
    if not GITHUB_TOKEN or not GITHUB_REPOSITORY:
        logger.error("❌ GITHUB_TOKEN 또는 GITHUB_REPOSITORY 환경변수가 없습니다.")
        return None

    filename     = os.path.basename(file_path)
    # 업로드용 영문 파일명 (GitHub API 한글 오류 우회)
    upload_filename = f"{target_yyyymm}_SW-Inspection.zip"
    
    tag          = f"release-{target_yyyymm}"
    release_name = f"{target_yyyymm[:4]}년 {int(target_yyyymm[4:])}월 소프트웨어 검사 결과"
    api_base     = f"https://api.github.com/repos/{GITHUB_REPOSITORY}"
    headers      = {
        'Authorization': f'Bearer {GITHUB_TOKEN}',
        'Accept': 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
    }

    logger.info(f"\n▶ GitHub Release 업로드 시작: {filename} (업로드명: {upload_filename})")

    # 1. 동일 태그 Release가 있으면 삭제 후 재생성 (매월 덮어쓰기)
    existing = requests.get(f"{api_base}/releases/tags/{tag}", headers=headers)
    if existing.status_code == 200:
        release_id = existing.json()['id']
        requests.delete(f"{api_base}/releases/{release_id}", headers=headers)
        requests.delete(f"{api_base}/git/refs/tags/{tag}", headers=headers)
        time.sleep(2)  # 기존 Release 삭제 후 태그 동기화 대기
        logger.info(f"  🗑️  기존 Release({tag}) 삭제 완료")

    # 2. 새 Release 생성
    release_resp = requests.post(
        f"{api_base}/releases",
        headers=headers,
        json={
            'tag_name':    tag,
            'name':        release_name,
            'body':        f"{target_yyyymm[:4]}년 {int(target_yyyymm[4:])}월 소프트웨어 검사 결과 파일입니다.",
            'draft':       False,
            'prerelease':  False,
            'make_latest': 'true'   # GitHub API 요구사항: bool이 아닌 문자열 전달
        }
    )
    if release_resp.status_code != 201:
        logger.error(f"❌ Release 생성 실패: {release_resp.text}")
        return None

    upload_url = release_resp.json()['upload_url'].replace('{?name,label}', '')

    # 3. ZIP 파일 업로드
    with open(file_path, 'rb') as f:
        upload_resp = requests.post(
            f"{upload_url}?name={upload_filename}",
            headers={**headers, 'Content-Type': 'application/zip'},
            data=f
        )

    if upload_resp.status_code != 201:
        logger.error(f"❌ 파일 업로드 실패: {upload_resp.text}")
        return None

    download_url = upload_resp.json().get('browser_download_url')
    logger.info(f"✅ GitHub Release 업로드 완료!")
    logger.info(f"🔗 다운로드 링크: {download_url}")
    logger.info(f"📝 다운로드 후 파일명을 '{filename}'으로 변경해주세요.")
    return download_url


# ==========================================
# 📅 12. 월 전환 로직
# ==========================================
def resolve_target_month(explicit_month):
    if explicit_month:
        return explicit_month
    now = datetime.now()
    if now.day <= 5:
        prev = now.replace(day=1) - timedelta(days=1)
        return prev.strftime("%Y%m")
    return now.strftime("%Y%m")


# ==========================================
# 🏃 메인 실행
# ==========================================
def run(args):
    target_yyyymm = resolve_target_month(args.month)
    logger.info(f"═══════════════════════════════════════")
    logger.info(f"  소프트웨어 검사결과 자동화 ({target_yyyymm})")
    logger.info(f"═══════════════════════════════════════")

    doc = get_gsheet_client()
    if not doc:
        sys.exit(1)

    folder_id = args.source or GDRIVE_SOURCE_FOLDER_ID

    if args.command == 'check':
        unsubmitted = check_unsubmitted(doc, target_yyyymm, folder_id=folder_id)
        if args.notify and unsubmitted:
            send_jandi_reminder(unsubmitted, target_yyyymm)

    elif args.command == 'zip':
        create_submission_zip(
            doc, target_yyyymm,
            folder_id=folder_id,
            team_name=args.team,
            submitter=args.submitter,
            dry_run=args.dry_run,
        )

    elif args.command == 'upload':
        zip_path, excluded_files = create_submission_zip(
            doc, target_yyyymm,
            folder_id=folder_id,
            team_name=args.team,
            submitter=args.submitter,
            dry_run=args.dry_run,
        )
        if zip_path and not args.dry_run:
            original_filename = os.path.basename(zip_path)
            link = upload_to_github_release(zip_path, target_yyyymm)
            if link and args.notify:
                send_jandi_upload_link(link, target_yyyymm, original_filename,
                                       excluded_files=excluded_files or None)

    elif args.command == 'all':
        # ─────────────────────────────────────────────
        # 미제출자 있으면 exit(2) → GitHub Actions 재시도
        # ─────────────────────────────────────────────
        unsubmitted = check_unsubmitted(doc, target_yyyymm, folder_id=folder_id)

        if unsubmitted:
            logger.info("⚠️ 미제출자가 존재합니다. ZIP 생성을 건너뜁니다.")
            if args.notify:
                send_jandi_reminder(unsubmitted, target_yyyymm)
            logger.info("🔁 GitHub Actions 재시도를 위해 exit code 2 반환")
            sys.exit(2)

        logger.info("✅ 모든 대상자 제출 완료. 압축 및 업로드를 진행합니다.")
        zip_path, excluded_files = create_submission_zip(
            doc, target_yyyymm,
            folder_id=folder_id,
            team_name=args.team,
            submitter=args.submitter,
            dry_run=args.dry_run,
        )
        if zip_path and not args.dry_run:
            original_filename = os.path.basename(zip_path)
            link = upload_to_github_release(zip_path, target_yyyymm)
            if link and args.notify:
                send_jandi_upload_link(link, target_yyyymm, original_filename,
                                       excluded_files=excluded_files or None)

    else:
        logger.info("사용법: python sw_inspector_auto.py {check|zip|upload|all} [옵션]")
        logger.info("  --help 으로 전체 옵션을 확인하세요.")


def main():
    parser = argparse.ArgumentParser(
        description='소프트웨어 검사결과 자동화 스크립트',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python sw_inspector_auto.py check                  # 미제출자 확인
  python sw_inspector_auto.py check --notify         # 미제출자 확인 + Jandi 독촉
  python sw_inspector_auto.py zip --dry-run           # ZIP 압축 미리보기
  python sw_inspector_auto.py zip                     # ZIP 압축 실행
  python sw_inspector_auto.py upload                  # ZIP 및 GitHub Release 업로드
  python sw_inspector_auto.py all --month 202603      # 전체 실행 (특정 월)
        """
    )

    parser.add_argument(
        'command',
        choices=['check', 'zip', 'upload', 'all'],
        help='실행할 기능: check(미제출자 확인), zip(파일 압축), upload(업로드), all(전체 실행)'
    )
    parser.add_argument('--month', '-m', type=str, default=None,
                        help='대상 연월 (YYYYMM 형식, 예: 202603). 미지정시 현재 월')
    parser.add_argument('--source', '-s', type=str, default=None,
                        help='소스 Drive 폴더 ID (기본값은 환경변수 GDRIVE_SOURCE_FOLDER_ID)')
    parser.add_argument('--team', '-t', type=str, default='AI에듀테크',
                        help='팀 이름 (ZIP 파일명에 사용, 기본: AI에듀테크)')
    parser.add_argument('--submitter', type=str, default='손만식',
                        help='제출 담당자 이름 (ZIP 파일명에 사용, 기본: 손만식)')
    parser.add_argument('--notify', '-n', action='store_true',
                        help='미제출자 발견 시 Jandi 웹훅으로 독촉 메시지 발송')
    parser.add_argument('--dry-run', '-d', action='store_true',
                        help='실제 파일 조작 없이 결과만 미리보기')

    args = parser.parse_args()
    run(args)


if __name__ == '__main__':
    main()
