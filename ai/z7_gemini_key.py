# Backward-compatible wrapper — redirects to z7_api_key (OpenRouter).
# Kept so that any legacy imports still resolve correctly.

from z7_api_key import (  # noqa: F401
    get_api_key,
    delete_api_key,
    read_stored_api_key,
    write_api_key,
)