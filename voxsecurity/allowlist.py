import requests
from urllib.parse import urlparse

ALLOWED_DOMAINS = {
    "api.elevenlabs.io",
    "api.elevenlabs.io.edgekey.net",  # CDN alias
    "update.voxsmith.app"
}


class DomainNotAllowed(Exception):
    pass

def make_voxsmith_session():
    s = requests.Session()
    old_request = s.request

    def checked_request(method, url, *a, **kw):
        host = urlparse(url).hostname
        if host not in ALLOWED_DOMAINS:
            raise DomainNotAllowed(f"Blocked domain: {host}")
        return old_request(method, url, *a, **kw)

    s.request = checked_request
    return s
