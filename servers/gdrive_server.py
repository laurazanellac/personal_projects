import io
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
HERE = Path(__file__).parent
mcp = MCPServer("gdrive")

def drive():
    token = HERE / "token.json"
    creds = Credentials.from_authorized_user_file(token, SCOPES) if token.exists() else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:  # first run: opens a browser to log in
            creds = InstalledAppFlow.from_client_secrets_file(
                HERE / "credentials.json", SCOPES).run_local_server(port=0)
        token.write_text(creds.to_json())
    return build("drive", "v3", credentials=creds)

@mcp.tool()
def search_files(query: str, max_results: int = 10) -> list[dict]:
    """Search Google Drive by file name. Returns id, name, type and last-modified date."""
    q = f"name contains '{query}' and trashed = false"
    res = drive().files().list(q=q, pageSize=max_results,
                               fields="files(id,name,mimeType,modifiedTime)").execute()
    return res.get("files", [])

@mcp.tool()
def read_file(file_id: str) -> str:
    """Read a Drive file's text content. Google Docs are exported as plain text."""
    svc = drive()
    meta = svc.files().get(fileId=file_id, fields="mimeType").execute()
    if meta["mimeType"] == "application/vnd.google-apps.document":
        return svc.files().export(fileId=file_id, mimeType="text/plain").execute().decode()
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue().decode(errors="replace")

if __name__ == "__main__":
    mcp.run()  # stdio transport by default
