"""IMAP connector: fetch emails directly from an IMAP server and index them."""
import email
import hashlib
import imaplib
import logging
import re
import threading
from email.utils import parseaddr, parsedate_to_datetime
from typing import Optional

from modules import indexer, nlp_engine, semantic_search
from modules.eml_parser import (
    _strip_html,
    _decode_header_str as _decode_str,
    _parse_address_list,
    _decode_payload,
)

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# OAuth2 constants (Microsoft / Outlook)
# ---------------------------------------------------------------------------

OUTLOOK_IMAP_SCOPE = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/consumers"
OUTLOOK_IMAP_HOST   = "imap-mail.outlook.com"


def _xoauth2_string(username: str, access_token: str) -> bytes:
    """Build the XOAUTH2 SASL string required by Microsoft IMAP.
    Returns raw bytes — imaplib.authenticate() handles base64 encoding itself."""
    auth_str = f"user={username}\x01auth=Bearer {access_token}\x01\x01"
    return auth_str.encode("ascii")


def _parse_message(msg: email.message.Message, uid: str, mailbox: str, server: str) -> dict:
    """Convert an imaplib Message into the same dict shape as eml_parser.parse_eml()."""
    subject = _decode_str(msg.get("Subject", ""))

    from_raw = _decode_str(msg.get("From", ""))
    sender_name, sender_email = parseaddr(from_raw)
    sender_email = sender_email.lower()
    sender_name = sender_name or (sender_email.split("@")[0] if sender_email else "")

    recipients = _parse_address_list(msg.get("To", ""))
    cc = _parse_address_list(msg.get("CC", ""))

    date_str = msg.get("Date", "")
    try:
        date = parsedate_to_datetime(date_str).isoformat() if date_str else None
    except Exception:
        date = None

    message_id = msg.get("Message-ID", "").strip("<>").strip()
    in_reply_to = msg.get("In-Reply-To", "").strip("<>").strip()

    refs_raw = msg.get("References", "").strip()
    refs = [r.strip("<>").strip() for r in refs_raw.split() if r.strip()]
    thread_id = refs[0] if refs else (in_reply_to or message_id)

    body_text = ""
    body_html = ""
    attachments: list[str] = []

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition", ""))

            if "attachment" in disposition:
                fname = part.get_filename()
                if fname:
                    attachments.append(_decode_str(fname))
            elif content_type == "text/plain" and not body_text:
                body_text = _decode_payload(part)
            elif content_type == "text/html" and not body_html:
                body_html = _decode_payload(part)
    else:
        content_type = msg.get_content_type()
        body_text = _decode_payload(msg)
        if content_type == "text/html":
            body_html = body_text
            body_text = ""

    if not body_text and body_html:
        body_text = _strip_html(body_html)

    body_text = re.sub(r"\s+", " ", body_text).strip()

    # Stable ID: hash of server + mailbox + uid (or message_id if present)
    id_source = message_id if message_id else f"{server}/{mailbox}/{uid}"
    file_id = hashlib.md5(id_source.encode()).hexdigest()

    # Virtual path so the rest of the app knows where this came from
    file_path = f"imap://{server}/{mailbox}/{uid}"

    return {
        "id": file_id,
        "file_path": file_path,
        "message_id": message_id,
        "subject": subject,
        "sender_name": sender_name,
        "sender_email": sender_email,
        "recipients": recipients,
        "cc": cc,
        "date": date,
        "body_text": body_text,
        "has_attachments": len(attachments) > 0,
        "attachment_names": attachments,
        "thread_id": thread_id,
        "in_reply_to": in_reply_to,
    }


# ---------------------------------------------------------------------------
# IMAPConnector
# ---------------------------------------------------------------------------

class IMAPConnector:
    """
    Fetches emails from an IMAP server and runs them through the same indexing
    pipeline used by EmailWatcher (eml_parser → indexer → NER → embedding).

    Usage (one-shot):
        conn = IMAPConnector("imap.example.com", "user@example.com", "secret")
        result = conn.fetch_and_index(mailbox="INBOX")  # handles 15k+ emails

    Usage (background polling):
        conn = IMAPConnector(...)
        conn.start(mailbox="INBOX", interval=300)
        ...
        conn.stop()
    """

    def __init__(
        self,
        host: str,
        username: str,
        password: str = "",
        port: int = 993,
        use_ssl: bool = True,
        # OAuth2 — used instead of password when provided
        access_token: str = "",
        refresh_token: str = "",
        client_id: str = "",
        # Optional callback invoked whenever tokens are auto-refreshed.
        # Signature: callback({"access_token": str, "refresh_token": str})
        token_save_callback=None,
    ):
        self.host = host
        self.username = username
        self.password = password
        self.port = port
        self.use_ssl = use_ssl
        self.access_token = access_token
        self.refresh_token = refresh_token
        self.client_id = client_id
        self.token_save_callback = token_save_callback

        self.status = "idle"
        self.last_indexed = 0
        self.last_deleted = 0
        self.new_tokens: Optional[dict] = None  # set when tokens auto-refreshed
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None

    # ------------------------------------------------------------------
    # Low-level IMAP helpers
    # ------------------------------------------------------------------

    def _do_token_refresh(self) -> bool:
        """
        Use the stored refresh_token to get a new access_token via MSAL.
        Updates self.access_token / self.refresh_token / self.new_tokens in place.
        Returns True on success.
        """
        try:
            import msal
        except ImportError:
            logger.error("msal not installed — cannot refresh OAuth2 token. Run: pip install msal")
            return False
        try:
            app = msal.PublicClientApplication(self.client_id, authority=MICROSOFT_AUTHORITY)
            result = app.acquire_token_by_refresh_token(self.refresh_token, scopes=OUTLOOK_IMAP_SCOPE)
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.refresh_token = result.get("refresh_token", self.refresh_token)
                self.new_tokens = {
                    "access_token": self.access_token,
                    "refresh_token": self.refresh_token,
                }
                logger.info("OAuth2 access token refreshed successfully.")
                if self.token_save_callback:
                    try:
                        self.token_save_callback(self.new_tokens)
                    except Exception as cb_exc:
                        logger.warning("token_save_callback failed: %s", cb_exc)
                return True
            logger.warning("Token refresh error: %s", result.get("error_description", result))
            return False
        except Exception as exc:
            logger.warning("Token refresh failed: %s", exc)
            return False

    def _connect(self) -> imaplib.IMAP4_SSL | imaplib.IMAP4:
        conn = (
            imaplib.IMAP4_SSL(self.host, self.port)
            if self.use_ssl
            else imaplib.IMAP4(self.host, self.port)
        )

        if self.access_token:
            # OAuth2 / XOAUTH2
            auth_bytes = _xoauth2_string(self.username, self.access_token)
            try:
                conn.authenticate("XOAUTH2", lambda x: auth_bytes)
            except imaplib.IMAP4.error as exc:
                # Token likely expired — attempt automatic refresh
                logger.warning("XOAUTH2 failed (%s) — attempting token refresh…", exc)
                if self.refresh_token and self.client_id and self._do_token_refresh():
                    conn = (
                        imaplib.IMAP4_SSL(self.host, self.port)
                        if self.use_ssl
                        else imaplib.IMAP4(self.host, self.port)
                    )
                    conn.authenticate("XOAUTH2", lambda x: _xoauth2_string(self.username, self.access_token))
                else:
                    raise RuntimeError(
                        "OAuth2 token expired and could not be refreshed. "
                        "Please re-authenticate in Settings → IMAP connection."
                    ) from exc
        else:
            # Basic auth (password)
            conn.login(self.username, self.password)

        return conn

    def refresh_access_token(self) -> Optional[dict]:
        """
        Manually refresh the OAuth2 access token.
        Returns {"access_token": ..., "refresh_token": ...} on success, None on failure.
        """
        return self.new_tokens if self._do_token_refresh() else None

    def _sync_deletions(self, server_uid_set: set[str], mailbox: str) -> int:
        """
        Compare server UIDs against the DB and delete any emails that no longer
        exist on the server. Returns the number of emails removed.
        """
        known_paths = indexer.get_indexed_imap_paths(self.host, mailbox)
        server_paths = {f"imap://{self.host}/{mailbox}/{uid}" for uid in server_uid_set}
        deleted_paths = list(known_paths - server_paths)
        if deleted_paths:
            count = indexer.delete_emails_by_paths(deleted_paths)
            logger.info("Deletion sync: removed %d email(s) from %s", count, mailbox)
            return count
        return 0

    def _fetch_uids(self, conn: imaplib.IMAP4_SSL | imaplib.IMAP4, mailbox: str) -> list[str]:
        """
        Return UIDs not yet indexed for *mailbox*.

        On the first run we ask the server for ALL UIDs and cache the highest
        seen UID in the DB (key: imap_max_uid:<host>/<mailbox>). On subsequent
        runs we only ask for UIDs *above* the cached max, which is a single
        cheap IMAP command regardless of mailbox size.

        We still cross-check against indexed paths so that a reset DB correctly
        re-indexes everything even when the cached max UID is set.
        """
        conn.select(mailbox, readonly=True)

        cache_key = f"imap_max_uid:{self.host}/{mailbox}"
        cached_max = indexer.get_meta(cache_key)
        known = indexer.get_indexed_imap_paths(self.host, mailbox)

        if cached_max and known:
            # Fast path: only fetch UIDs newer than the last seen
            _, data = conn.uid("search", None, f"UID {int(cached_max) + 1}:*")
        else:
            # First run or reset DB: fetch everything
            _, data = conn.uid("search", None, "ALL")

        all_uids = data[0].decode().split() if data[0] else []

        if all_uids:
            # Update cached max — UIDs are monotonically increasing
            new_max = max(int(u) for u in all_uids)
            if cached_max is None or new_max > int(cached_max):
                indexer.set_meta(cache_key, str(new_max))

        return [uid for uid in all_uids if f"imap://{self.host}/{mailbox}/{uid}" not in known]

    def _fetch_messages_bulk(
        self,
        conn: imaplib.IMAP4_SSL | imaplib.IMAP4,
        uids: list[str],
    ) -> list[tuple[str, email.message.Message]]:
        """
        Fetch multiple UIDs in a single IMAP round trip.
        Returns a list of (uid, message) pairs for successfully fetched messages.
        """
        if not uids:
            return []
        uid_set = b",".join(u.encode() for u in uids)
        try:
            _, data = conn.uid("fetch", uid_set, "(RFC822)")
        except Exception as exc:
            logger.warning("Bulk fetch failed for %d UIDs: %s", len(uids), exc)
            return []

        # imaplib returns alternating (header_tuple, b')') items
        results = []
        uid_iter = iter(uids)
        for item in data:
            if isinstance(item, tuple) and len(item) == 2 and isinstance(item[1], bytes):
                try:
                    msg = email.message_from_bytes(item[1])
                    uid = next(uid_iter)
                    results.append((uid, msg))
                except Exception:
                    next(uid_iter, None)  # keep iterators in sync
        return results

    # ------------------------------------------------------------------
    # Indexing pipeline (same steps as watcher._scan)
    # ------------------------------------------------------------------

    def _index_parsed_batch(self, parsed_list: list[dict]) -> tuple[int, int]:
        """
        Insert, NER, and embed a batch of parsed emails.
        Batches the embedding step for efficiency.
        Returns (inserted_count, skipped_count).
        """
        to_embed: list[dict] = []

        for parsed in parsed_list:
            inserted = indexer.insert_email(parsed)
            if not inserted:
                continue

            text = f"{parsed.get('subject', '')} {parsed.get('body_text', '')}"
            entities = nlp_engine.extract_entities(text)
            entities += nlp_engine.extract_orgs_from_email_addrs(parsed)
            if entities:
                indexer.insert_entities(parsed["id"], entities)

            to_embed.append(parsed)

        if not to_embed:
            return 0, len(parsed_list)

        # Batch embed — single model pass for all emails in this chunk
        texts = [
            f"{p.get('subject', '')} {p.get('body_text', '')[:400]}"
            for p in to_embed
        ]
        try:
            vecs = semantic_search.embed_batch(texts)
            for parsed, vec in zip(to_embed, vecs):
                indexer.insert_embedding(parsed["id"], vec)
        except Exception as exc:
            # Fall back to one-at-a-time if embed_batch fails
            logger.warning("embed_batch failed, falling back to single embeds: %s", exc)
            for parsed in to_embed:
                try:
                    vec = semantic_search.embed_text(
                        f"{parsed.get('subject', '')} {parsed.get('body_text', '')[:400]}"
                    )
                    indexer.insert_embedding(parsed["id"], vec)
                except Exception as e:
                    logger.warning("embed_text failed for %s: %s", parsed["id"], e)

        return len(to_embed), len(parsed_list) - len(to_embed)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def fetch_and_index(
        self,
        mailbox: str = "INBOX",
        batch_size: int = 20000,
        imap_chunk: int = 100,
        embed_chunk: int = 64,
        sync_deletions: bool = True,
    ) -> dict:
        """
        Synchronous one-shot fetch: connect, find new UIDs, index up to
        *batch_size* messages, optionally sync deletions, disconnect.

        Args:
            mailbox:        IMAP mailbox to fetch from.
            batch_size:     Max new UIDs to process per call (default 20 000).
            imap_chunk:     UIDs fetched per IMAP round trip (default 100).
            embed_chunk:    Emails per embedding model pass (default 64).
            sync_deletions: If True, emails deleted on the server are also
                            removed from the local index each poll.

        Returns:
            {"indexed", "deleted", "skipped", "errors", "total_in_db", "new_tokens"}
        """
        indexer.init_db()
        indexed = skipped = errors = deleted = 0

        try:
            conn = self._connect()
        except Exception as exc:
            self.status = f"connection failed: {exc}"
            logger.error("IMAP connection failed: %s", exc)
            return {
                "indexed": 0, "deleted": 0, "skipped": 0,
                "errors": 1, "total_in_db": indexer.get_email_count(),
                "new_tokens": self.new_tokens,
            }

        try:
            conn.select(mailbox, readonly=True)

            # Always fetch ALL server UIDs — needed for deletion sync and cheap
            # (SEARCH returns only UIDs, not bodies).
            _, data = conn.uid("search", None, "ALL")
            all_server_uids: list[str] = data[0].decode().split() if data[0] else []
            server_uid_set = set(all_server_uids)

            # Update cached max UID
            if all_server_uids:
                cache_key = f"imap_max_uid:{self.host}/{mailbox}"
                new_max = max(int(u) for u in all_server_uids)
                cached_max = indexer.get_meta(cache_key)
                if cached_max is None or new_max > int(cached_max):
                    indexer.set_meta(cache_key, str(new_max))

            # Deletion sync — runs before indexing so counts are accurate
            if sync_deletions:
                self.status = "checking for deletions…"
                deleted = self._sync_deletions(server_uid_set, mailbox)

            # Find new UIDs not yet in the DB
            known_paths = indexer.get_indexed_imap_paths(self.host, mailbox)
            new_uids = [
                uid for uid in all_server_uids
                if f"imap://{self.host}/{mailbox}/{uid}" not in known_paths
            ]
            to_process = new_uids[:batch_size]
            self.status = f"fetching {len(to_process):,} of {len(new_uids):,} new email(s)…"
            logger.info("IMAP: %d new UIDs to index in %s", len(to_process), mailbox)

            # Process in imap_chunk-sized round trips
            for chunk_start in range(0, len(to_process), imap_chunk):
                uid_chunk = to_process[chunk_start : chunk_start + imap_chunk]
                fetched = self._fetch_messages_bulk(conn, uid_chunk)

                parsed_chunk: list[dict] = []
                for uid, msg in fetched:
                    try:
                        parsed_chunk.append(_parse_message(msg, uid, mailbox, self.host))
                    except Exception as exc:
                        logger.warning("Failed to parse UID %s: %s", uid, exc)
                        errors += 1

                for embed_start in range(0, len(parsed_chunk), embed_chunk):
                    sub = parsed_chunk[embed_start : embed_start + embed_chunk]
                    try:
                        ins, skp = self._index_parsed_batch(sub)
                        indexed += ins
                        skipped += skp
                    except Exception as exc:
                        logger.warning("Batch index failed: %s", exc)
                        errors += len(sub)

                self.status = f"indexed {indexed:,} / {len(to_process):,}…"
                logger.info("IMAP progress: %d / %d", chunk_start + len(uid_chunk), len(to_process))

        finally:
            try:
                conn.logout()
            except Exception:
                pass

        self.last_indexed = indexed
        self.last_deleted = deleted
        _del_str = f", {deleted:,} deleted" if deleted else ""
        self.status = f"idle ({indexed:,} indexed{_del_str})"
        return {
            "indexed": indexed,
            "deleted": deleted,
            "skipped": skipped,
            "errors": errors,
            "total_in_db": indexer.get_email_count(),
            "new_tokens": self.new_tokens,
        }

    # ------------------------------------------------------------------
    # Background polling (optional)
    # ------------------------------------------------------------------

    def start(
        self,
        mailbox: str = "INBOX",
        interval: int = 300,
        batch_size: int = 20000,
        sync_deletions: bool = True,
    ) -> None:
        """Start a background daemon thread that polls IMAP every *interval* seconds."""
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._run,
            kwargs={
                "mailbox": mailbox,
                "interval": interval,
                "batch_size": batch_size,
                "sync_deletions": sync_deletions,
            },
            daemon=True,
            name="IMAPConnector",
        )
        self._thread.start()

    def stop(self) -> None:
        """Signal the background thread to stop after the current poll."""
        self._stop_event.set()

    def _run(self, mailbox: str, interval: int, batch_size: int, sync_deletions: bool) -> None:
        self.fetch_and_index(mailbox=mailbox, batch_size=batch_size, sync_deletions=sync_deletions)
        while not self._stop_event.wait(interval):
            self.fetch_and_index(mailbox=mailbox, batch_size=batch_size, sync_deletions=sync_deletions)
