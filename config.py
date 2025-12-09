import os
import json
import logging
from dataclasses import dataclass, field
from typing import Dict, Optional
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

@dataclass
class Config:
    client_secret_path: str
    token_path: str
    folder_id: str  # 하위 호환용 (기본 폴더 ID)
    folders: Dict[str, str] = field(default_factory=dict)  # 별칭 → 폴더ID 매핑
    default_folder: str = ""  # 기본 폴더 별칭

    @classmethod
    def from_env(cls) -> 'Config':
        load_dotenv()

        client_secret_path = os.getenv('MCPGD_CLIENT_SECRET_PATH')
        if not client_secret_path:
            raise ValueError("MCPGD_CLIENT_SECRET_PATH environment variable is required")

        token_path = os.getenv('MCPGD_TOKEN_PATH')
        if not token_path:
            home_dir = os.path.expanduser('~')
            token_path = os.path.join(home_dir, '.mcp_google_spreadsheet.json')

        # 새로운 다중 폴더 설정 파싱
        folders_str = os.getenv('MCPGD_FOLDERS', '')
        folders = {}
        if folders_str:
            for item in folders_str.split(','):
                item = item.strip()
                if ':' in item:
                    alias, fid = item.split(':', 1)
                    folders[alias.strip()] = fid.strip()

        default_folder = os.getenv('MCPGD_DEFAULT_FOLDER', '')

        # 하위 호환: 기존 MCPGD_FOLDER_ID 지원
        folder_id = os.getenv('MCPGD_FOLDER_ID', '')
        if not folders and folder_id:
            # 기존 설정만 있는 경우 → 자동 변환
            folders = {'default': folder_id}
            default_folder = 'default'
        elif folders and not folder_id:
            # 새 설정만 있는 경우 → 기본 폴더 ID 설정
            if default_folder and default_folder in folders:
                folder_id = folders[default_folder]
            elif folders:
                # default_folder 미설정 시 첫 번째 폴더 사용
                first_alias = list(folders.keys())[0]
                folder_id = folders[first_alias]
                if not default_folder:
                    default_folder = first_alias

        if not folder_id and not folders:
            raise ValueError("MCPGD_FOLDERS or MCPGD_FOLDER_ID environment variable is required")

        logger.info(json.dumps({
            "event": "config_loaded",
            "config": {
                "client_secret_path": client_secret_path,
                "token_path": token_path,
                "folder_id": folder_id,
                "folders": folders,
                "default_folder": default_folder,
                "client_secret_exists": os.path.exists(client_secret_path)
            }
        }))

        return cls(
            client_secret_path=client_secret_path,
            token_path=token_path,
            folder_id=folder_id,
            folders=folders,
            default_folder=default_folder
        )

    def get_folder_id(self, folder: Optional[str] = None) -> str:
        """별칭으로 폴더 ID 조회. 미지정 시 기본 폴더 반환."""
        if not folder:
            return self.folder_id
        if folder in self.folders:
            return self.folders[folder]
        # 별칭이 아닌 폴더 ID 직접 입력된 경우
        return folder

    def get_folder_name(self, folder_id: str) -> str:
        """폴더 ID로 별칭 조회."""
        for alias, fid in self.folders.items():
            if fid == folder_id:
                return alias
        return "" 