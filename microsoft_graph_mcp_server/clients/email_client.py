"""Email client for Microsoft Graph API."""

import asyncio
import base64
import logging
import os
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Optional, List, Dict, Any
from zoneinfo import ZoneInfo

from .base_client import BaseGraphClient
from ..utils import DateHandler as date_handler
from ..utils.image_utils import compress_base64_image
from ..config import settings

logger = logging.getLogger(__name__)

MAX_EMAIL_SEARCH_LIMIT = 1000


class EmailClient(BaseGraphClient):
    """Client for email-related operations."""

    def __init__(self):
        super().__init__()
        self._folder_cache: Dict[str, str] = {}
        self._well_known_folders = {
            "inbox": "Inbox",
            "sent": "SentItems",
            "sent items": "SentItems",
            "drafts": "Drafts",
            "deleted": "DeletedItems",
            "deleted items": "DeletedItems",
            "archive": "Archive",
            "junk": "JunkEmail",
            "junk email": "JunkEmail",
        }
        self._user_timezone_cache: Optional[str] = None
        self._user_email_cache: Optional[str] = None

    async def get_user_email(self) -> str:
        """Get user's email address with caching."""
        if self._user_email_cache is not None:
            return self._user_email_cache

        try:
            result = await self.get("/me", params={"$select": "mail"})
            user_email = result.get("mail", "")
            if user_email:
                self._user_email_cache = user_email
                return user_email
        except Exception:
            pass

        return ""

    async def get_user_timezone(self) -> str:
        """Get user's timezone identifier. Uses server local timezone with caching."""
        if self._user_timezone_cache is not None:
            return self._user_timezone_cache

        try:
            local_tz = datetime.now().astimezone().tzinfo
            if local_tz:
                tz_str = str(local_tz)
                if tz_str and tz_str != "UTC":
                    self._user_timezone_cache = date_handler.convert_to_iana_timezone(
                        tz_str
                    )
                    return self._user_timezone_cache
        except Exception:
            pass

        self._user_timezone_cache = date_handler.convert_to_iana_timezone(
            settings.user_timezone
        )
        return self._user_timezone_cache

    async def get_messages(
        self, folder: str = "Inbox", top: int = 10, filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Get messages from specified folder."""
        params = {"$top": top}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get(f"/me/mailFolders/{folder}/messages", params=params)
        return result.get("value", [])

    async def list_mail_folders(self) -> List[Dict[str, Any]]:
        """List all mail folders with their paths (including all levels)."""

        async def fetch_all_folders(
            folders_list: List[Dict[str, Any]], parent_path: str = ""
        ) -> List[Dict[str, Any]]:
            """Recursively fetch all folders and their children."""
            all_folders = []

            for folder in folders_list:
                folder_name = folder.get("displayName", "")
                current_path = (
                    f"{parent_path}/{folder_name}" if parent_path else folder_name
                )

                all_folders.append(
                    {
                        "path": current_path,
                        "emailCount": folder.get("totalItemCount", 0),
                    }
                )

                if folder.get("childFolderCount", 0) > 0:
                    child_folders_result = await self.get(
                        f"/me/mailFolders/{folder.get('id')}/childFolders"
                    )
                    child_folders = child_folders_result.get("value", [])
                    if child_folders:
                        all_folders.extend(
                            await fetch_all_folders(child_folders, current_path)
                        )

            return all_folders

        result = await self.get("/me/mailFolders")
        folders = result.get("value", [])

        if not folders:
            return []

        all_folders = await fetch_all_folders(folders)

        existing_paths = {folder["path"] for folder in all_folders}

        try:
            test_folders_result = await self.get(
                "/me/mailFolders",
                params={"$filter": "startswith(displayName, 'TestFolder')"},
            )
            test_folders = test_folders_result.get("value", [])

            for test_folder in test_folders:
                folder_name = test_folder.get("displayName", "")
                if folder_name not in existing_paths:
                    all_folders.append(
                        {
                            "path": folder_name,
                            "emailCount": test_folder.get("totalItemCount", 0),
                        }
                    )
        except Exception:
            pass

        return sorted(all_folders, key=lambda x: x["path"])

    async def _get_folder_id_by_path(
        self, folder_path: str, max_retries: int = 10, retry_delay: float = 2.0
    ) -> str:
        """Get folder ID by path. Supports nested folders like 'Archive/2024'.

        Args:
            folder_path: Path to folder
            max_retries: Maximum number of retries when folder is not found
            retry_delay: Delay between retries in seconds

        Returns:
            Folder ID

        Raises:
            ValueError: If folder is not found after all retries
        """
        if not folder_path:
            folder_path = "Inbox"

        cache_key = folder_path.lower()
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]

        parts = folder_path.split("/")
        current_folder_id = None

        for part in parts:
            part = part.strip()
            if not part:
                continue

            folder = None
            for attempt in range(max_retries):
                if current_folder_id is None:
                    result = await self.get(
                        "/me/mailFolders",
                        params={"$filter": f"displayName eq '{part}'"},
                    )
                    folders = result.get("value", [])
                    folder = next(
                        (
                            f
                            for f in folders
                            if f.get("displayName", "").lower() == part.lower()
                        ),
                        None,
                    )
                else:
                    result = await self.get(
                        f"/me/mailFolders/{current_folder_id}/childFolders",
                        params={"$filter": f"displayName eq '{part}'"},
                    )
                    folders = result.get("value", [])
                    folder = next(
                        (
                            f
                            for f in folders
                            if f.get("displayName", "").lower() == part.lower()
                        ),
                        None,
                    )

                if folder:
                    break

                if attempt < max_retries - 1:
                    await asyncio.sleep(retry_delay)

            if not folder:
                raise ValueError(f"Folder not found: {part}")

            current_folder_id = folder.get("id")

        self._folder_cache[cache_key] = current_folder_id
        return current_folder_id

    async def _get_or_create_templates_folder(self) -> str:
        """Get or create the Templates folder.

        Returns:
            Folder ID of the Templates folder
        """
        try:
            return await self._get_folder_id_by_path("Templates")
        except ValueError:
            await self.create_folder("Templates")
            await asyncio.sleep(2.0)
            return await self._get_folder_id_by_path("Templates")

    async def load_emails_by_folder(
        self,
        folder: str = "Inbox",
        days: Optional[int] = None,
        top: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Load emails from a folder with optional days or top parameter (mutually exclusive)."""
        if days is not None and top is not None:
            raise ValueError(
                "Cannot specify both 'days' and 'top' parameters simultaneously"
            )

        if days is not None and days >= 30:
            raise ValueError("Days parameter must be less than 30")

        if top is not None and top > MAX_EMAIL_SEARCH_LIMIT:
            raise ValueError(
                f"Maximum number of emails per search is {MAX_EMAIL_SEARCH_LIMIT}"
            )

        folder_id = await self._get_folder_id_by_path(folder)
        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        params = {}
        filter_parts = []

        if days is not None:
            now_local = datetime.now(user_tz)
            cutoff_local = now_local - timedelta(days=days)
            cutoff_utc = cutoff_local.astimezone(timezone.utc)
            filter_parts.append(
                f"receivedDateTime ge {cutoff_utc.isoformat().replace('+00:00', 'Z')}"
            )

        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)

        if top is not None:
            params["$top"] = top
        else:
            params["$top"] = 100

        params["$select"] = (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,inferenceClassification"
        )

        result = await self.get(f"/me/mailFolders/{folder_id}/messages", params=params)
        emails = result.get("value", [])

        summaries = [
            self._create_email_summary(email, idx + 1, user_tz)
            for idx, email in enumerate(emails)
        ]

        sorted_summaries = sorted(
            summaries, key=lambda x: x.get("receivedDateTimeOriginal", ""), reverse=True
        )

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(
            sorted_summaries, user_timezone_str
        )
        filter_date_range = (
            date_handler.format_filter_date_range(days, user_timezone_str)
            if days
            else None
        )

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range,
        }

    def _create_email_summary(
        self, email: Dict[str, Any], index: int, user_tz: ZoneInfo
    ) -> Dict[str, Any]:
        """Create a summary of an email with timezone-aware timestamps."""
        from_email = email.get("from", {}).get("emailAddress", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("address", "")

        to_recipients = [
            r.get("emailAddress", {}).get("address", "")
            for r in email.get("toRecipients", [])
        ]
        cc_recipients = [
            r.get("emailAddress", {}).get("address", "")
            for r in email.get("ccRecipients", [])
        ]

        received_datetime = email.get("receivedDateTime", "")
        sent_datetime = email.get("sentDateTime", "")

        received_datetime_original = received_datetime
        received_datetime_display = date_handler.convert_utc_to_timezone(
            received_datetime, user_tz
        )
        sent_datetime_display = date_handler.convert_utc_to_timezone(
            sent_datetime, user_tz
        )

        # Process attachments if available
        attachments = []
        for attachment in email.get("attachments", []):
            if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
                attachments.append({
                    "id": attachment.get("id", ""),
                    "name": attachment.get("name", ""),
                    "size": attachment.get("size", 0),
                    "contentType": attachment.get("contentType", ""),
                    "isInline": attachment.get("isInline", False),
                })

        return {
            "number": index,
            "id": email.get("id", ""),
            "subject": email.get("subject", ""),
            "from": {"name": from_name, "email": from_address},
            "to": to_recipients,
            "cc": cc_recipients,
            "receivedDateTime": received_datetime_display,
            "receivedDateTimeOriginal": received_datetime_original,
            "sentDateTime": sent_datetime_display,
            "isRead": email.get("isRead", False),
            "hasAttachments": email.get("hasAttachments", False),
            "attachments": attachments,
            "importance": email.get("importance", "normal"),
            "inferenceClassification": email.get("inferenceClassification", "focused"),
            "bodyPreview": email.get("bodyPreview", ""),
            "conversationId": email.get("conversationId", ""),
            "parentFolderId": email.get("parentFolderId", ""),
        }

    def _extract_text_from_html(self, html_content: str) -> str:
        """Extract plain text from HTML content with aggressive cleanup (optimized)."""
        import re

        text = html_content

        text = re.sub(
            r"<(head|style|script).*?>.*?</\1>",
            "",
            text,
            flags=re.DOTALL | re.IGNORECASE,
        )
        text = re.sub(
            r'\s*(style|class|id|data-outlook-trace)="[^"]*"',
            "",
            text,
            flags=re.IGNORECASE,
        )
        text = re.sub(r"<(img|hr)[^>]*>", "", text, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        text = re.sub(
            r"&(nbsp|amp|lt|gt|quot|#39);",
            lambda m: {
                "nbsp": " ",
                "amp": "&",
                "lt": "<",
                "gt": ">",
                "quot": '"',
                "#39": "'",
            }[m.group(1)],
            text,
        )
        text = re.sub(r"\s+", " ", text)
        text = re.sub(r"(\n\s*){3,}", "\n\n", text)

        return text.strip()

    async def get_email_count(
        self, folder: str = "Inbox", filter_query: Optional[str] = None
    ) -> int:
        """Get count of messages in folder."""
        params = {}
        if filter_query:
            params["$filter"] = filter_query

        result = await self.get(
            f"/me/mailFolders/{folder}/messages/$count",
            params=params,
            headers={"Accept": "text/plain"},
        )
        return int(result) if result else 0

    async def get_email(
        self,
        email_id: str,
        emailNumber: int = 0,
        return_html: bool = False,
        download_attachments: bool = False,
        download_path: Optional[str] = None,
        attachment_names: Optional[List[str]] = None,
        multimodal_supported: bool = False,
    ) -> Dict[str, Any]:
        """Get full email by ID with optimized field selection.
        
        Args:
            email_id: The email message ID
            emailNumber: Cache number for reference
            return_html: If True, return full HTML body. If False (default), extract plain text from HTML body.
            download_attachments: If True, download attachments to disk
            download_path: Custom path for downloads (default: workspace/attachments)
            attachment_names: Specific attachment names to download (default: all)
            multimodal_supported: If True, include base64 content for image attachments
        
        Returns:
            Dictionary with email content and metadata
        """
        user_timezone_str = await self.get_user_timezone()

        params = {
            "$select": "subject,from,toRecipients,ccRecipients,bccRecipients,body,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,isDraft,internetMessageId,conversationId,parentFolderId,flag",
            "$expand": "attachments($select=id,name,size,contentType,isInline)",
        }

        email = await self.get(f"/me/messages/{email_id}", params=params)

        from_email = email.get("from", {}).get("emailAddress", {})
        to_recipients = [
            r.get("emailAddress", {}) for r in email.get("toRecipients", [])
        ]
        cc_recipients = [
            r.get("emailAddress", {}) for r in email.get("ccRecipients", [])
        ]
        bcc_recipients = [
            r.get("emailAddress", {}) for r in email.get("bccRecipients", [])
        ]

        body_content = email.get("body", {}).get("content", "")
        body_type = email.get("body", {}).get("contentType", "Text")

        if not return_html and body_type.lower() == "html":
            body_content = self._extract_text_from_html(body_content)

        received_datetime = email.get("receivedDateTime", "")
        sent_datetime = email.get("sentDateTime", "")

        attachments = []
        for attachment in email.get("attachments", []):
            if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
                att_info = {
                    "id": attachment.get("id", ""),
                    "name": attachment.get("name", ""),
                    "size": attachment.get("size", 0),
                    "contentType": attachment.get("contentType", ""),
                    "isInline": attachment.get("isInline", False),
                }
                attachments.append(att_info)

        # Download attachments if requested
        downloaded_attachments = []
        if download_attachments and attachments:
            downloaded_attachments = await self._download_attachments(
                email_id=email_id,
                attachments=attachments,
                download_path=download_path,
                attachment_names=attachment_names,
                multimodal_supported=multimodal_supported,
            )
            # Update attachment info with download status
            for att in attachments:
                for dl in downloaded_attachments:
                    if att["id"] == dl["id"]:
                        att["downloaded"] = True
                        att["file_path"] = dl["file_path"]
                        # Add base64 content for image attachments if multimodal is supported
                        if multimodal_supported and dl.get("base64_content"):
                            att["base64_content"] = dl["base64_content"]

        # Handle inline images for multimodal support (independent of download_attachments)
        logger.info(f"[MULTIMODAL DEBUG] multimodal_supported={multimodal_supported}, attachments count={len(attachments)}")
        if multimodal_supported:
            inline_images = [a for a in attachments if a.get("isInline") and a.get("contentType", "").startswith("image/")]
            logger.info(f"[MULTIMODAL DEBUG] Found {len(inline_images)} inline images")
            for inline_image in inline_images:
                logger.info(f"[MULTIMODAL DEBUG] Processing inline image: {inline_image.get('name')}, isInline={inline_image.get('isInline')}, contentType={inline_image.get('contentType')}")
                # Skip if already processed in download_attachments
                if inline_image.get("downloaded"):
                    logger.info(f"[MULTIMODAL DEBUG] Skipping already processed inline image: {inline_image['name']}")
                    continue
                # Fetch the inline image content
                try:
                    attachment_data = await self.get(f"/me/messages/{email_id}/attachments/{inline_image['id']}")
                    content_bytes = attachment_data.get("contentBytes", "")
                    logger.info(f"[MULTIMODAL DEBUG] Fetched inline image content, length={len(content_bytes) if content_bytes else 0}")
                    if content_bytes:
                        # Compress image if needed
                        compressed_content = compress_base64_image(
                            base64_content=content_bytes,
                            max_size_bytes=settings.image_max_size,
                            max_dimension=settings.image_max_dimension,
                            quality=settings.image_quality,
                        )
                        inline_image["base64_content"] = compressed_content
                        inline_image["downloaded"] = True
                        logger.info(f"Got and compressed base64 content for inline image (multimodal): {inline_image['name']}")
                except Exception as e:
                    logger.warning(f"Failed to get inline image {inline_image['name']}: {e}")

        return {
            "content": {
                "subject": email.get("subject", ""),
                "from": {
                    "name": from_email.get("name", ""),
                    "email": from_email.get("address", ""),
                },
                "to": [
                    {"name": r.get("name", ""), "email": r.get("address", "")}
                    for r in to_recipients
                ],
                "cc": [
                    {"name": r.get("name", ""), "email": r.get("address", "")}
                    for r in cc_recipients
                ],
                "bcc": [
                    {"name": r.get("name", ""), "email": r.get("address", "")}
                    for r in bcc_recipients
                ],
                "body": body_content,
                "bodyType": body_type,
                "attachments": attachments,
            },
            "metadata": {
                "number": emailNumber,
                "receivedDateTime": date_handler.convert_utc_to_user_timezone(
                    received_datetime, user_timezone_str
                ),
                "sentDateTime": date_handler.convert_utc_to_user_timezone(
                    sent_datetime, user_timezone_str
                ),
                "isRead": email.get("isRead", False),
                "hasAttachments": email.get("hasAttachments", False),
                "importance": email.get("importance", "normal"),
                "isDraft": email.get("isDraft", False),
                "internetMessageId": email.get("internetMessageId", ""),
                "conversationId": email.get("conversationId", ""),
                "parentFolderId": email.get("parentFolderId", ""),
                "flag": email.get("flag", {}),
            },
        }

    async def _download_attachments(
        self,
        email_id: str,
        attachments: List[Dict[str, Any]],
        download_path: Optional[str] = None,
        attachment_names: Optional[List[str]] = None,
        multimodal_supported: bool = False,
    ) -> List[Dict[str, Any]]:
        """Download attachments to local disk.
        
        Args:
            email_id: The email message ID
            attachments: List of attachment metadata
            download_path: Custom download path (default: workspace/attachments)
            attachment_names: Specific attachment names to download (default: all)
            multimodal_supported: If True, include base64 content for image attachments
        
        Returns:
            List of downloaded attachment info with file paths
        """
        # Determine download directory
        if download_path:
            download_dir = Path(download_path)
        else:
            # Default to workspace/attachments
            download_dir = Path.cwd() / "attachments"
        
        # Create directory if not exists
        download_dir.mkdir(parents=True, exist_ok=True)
        
        downloaded = []
        
        # Filter attachments to download
        to_download = attachments
        if attachment_names:
            to_download = [
                a for a in attachments 
                if a["name"] in attachment_names
            ]
        
        # Skip inline attachments unless multimodal is supported and it's an image
        to_download = [
            a for a in to_download 
            if not a.get("isInline", False) or (multimodal_supported and a.get("contentType", "").startswith("image/"))
        ]
        
        for attachment in to_download:
            attachment_id = attachment.get("id", "")
            attachment_name = attachment.get("name", "")
            is_inline = attachment.get("isInline", False)
            content_type = attachment.get("contentType", "")
            
            if not attachment_id or not attachment_name:
                continue
            
            try:
                # Fetch full attachment with contentBytes
                attachment_data = await self.get(
                    f"/me/messages/{email_id}/attachments/{attachment_id}"
                )
                
                content_bytes = attachment_data.get("contentBytes", "")
                if not content_bytes:
                    logger.warning(f"No content bytes for attachment: {attachment_name}")
                    continue
                
                # Build attachment info
                att_info = {
                    "id": attachment_id,
                    "name": attachment_name,
                    "contentType": content_type,
                    "isInline": is_inline,
                }
                
                # For inline images with multimodal support, only get base64 (no file download)
                if is_inline and multimodal_supported and content_type.startswith("image/"):
                    # Compress image for multimodal support
                    compressed_content = compress_base64_image(
                        base64_content=content_bytes,
                        max_size_bytes=settings.image_max_size,
                        max_dimension=settings.image_max_dimension,
                        quality=settings.image_quality,
                    )
                    att_info["base64_content"] = compressed_content
                    att_info["size"] = len(base64.b64decode(content_bytes))
                    att_info["downloaded"] = True
                    logger.info(f"Got and compressed base64 content for inline image (multimodal): {attachment_name}")
                else:
                    # For non-inline attachments, download to file
                    file_content = base64.b64decode(content_bytes)
                    
                    # Generate unique filename if file already exists
                    file_path = download_dir / attachment_name
                    if file_path.exists():
                        base_name = file_path.stem
                        suffix = file_path.suffix
                        counter = 1
                        while file_path.exists():
                            file_path = download_dir / f"{base_name}_{counter}{suffix}"
                            counter += 1
                    
                    # Write file
                    await asyncio.to_thread(file_path.write_bytes, file_content)
                    
                    att_info["file_path"] = str(file_path)
                    att_info["size"] = len(file_content)
                    att_info["downloaded"] = True
                    
                    # Add base64 content for image attachments if multimodal is supported
                    if multimodal_supported and content_type.startswith("image/"):
                        # Compress image for multimodal support
                        compressed_content = compress_base64_image(
                            base64_content=content_bytes,
                            max_size_bytes=settings.image_max_size,
                            max_dimension=settings.image_max_dimension,
                            quality=settings.image_quality,
                        )
                        att_info["base64_content"] = compressed_content
                        logger.info(f"Added and compressed base64 content for image attachment: {attachment_name}")
                    
                    logger.info(f"Downloaded attachment: {attachment_name} to {file_path}")
                
                downloaded.append(att_info)
                
            except Exception as e:
                logger.error(f"Failed to download attachment {attachment_name}: {e}")
        
        return downloaded

    async def create_template_from_email(
        self,
        email_id: str,
        template_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a template by copying an email to the Templates folder.

        Args:
            email_id: ID of the email to copy
            template_name: Optional name for the template (defaults to email subject)

        Returns:
            Created template information
        """
        templates_folder_id = await self._get_folder_id_by_path("Templates")

        email = await self.get(
            f"/me/messages/{email_id}",
            params={
                "$select": "subject,from,toRecipients,ccRecipients,bccRecipients,body,attachments",
                "$expand": "attachments",
            },
        )

        template_data = {
            "subject": template_name or email.get("subject", ""),
            "body": email.get("body", {}),
            "isDraft": True,
            "toRecipients": email.get("toRecipients", []),
            "ccRecipients": email.get("ccRecipients", []),
            "bccRecipients": email.get("bccRecipients", []),
        }

        result = await self.post(
            f"/me/mailFolders/{templates_folder_id}/messages", json=template_data
        )

        template_id = result.get("id", "")
        attachments = email.get("attachments", [])

        for attachment in attachments:
            if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
                attachment_data = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name"),
                    "contentBytes": attachment.get("contentBytes"),
                    "isInline": attachment.get("isInline", False),
                }
                await self.post(
                    f"/me/messages/{template_id}/attachments", json=attachment_data
                )

        return {
            "id": template_id,
            "subject": template_data["subject"],
            "folder": "Templates",
        }

    async def list_templates(
        self,
        top: int = 50,
    ) -> List[Dict[str, Any]]:
        """List all templates in the Templates folder.

        Args:
            top: Maximum number of templates to return

        Returns:
            List of template summaries
        """
        templates_folder_id = await self._get_folder_id_by_path("Templates")
        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        params = {
            "$top": top,
            "$filter": "isDraft eq true",
            "$select": "id,subject,createdDateTime,lastModifiedDateTime,toRecipients,ccRecipients,hasAttachments,bodyPreview",
            "$orderby": "lastModifiedDateTime desc",
        }

        result = await self.get(
            f"/me/mailFolders/{templates_folder_id}/messages", params=params
        )

        templates = result.get("value", [])

        summaries = []
        for idx, template in enumerate(templates):
            created_datetime = date_handler.convert_utc_to_timezone(
                template.get("createdDateTime", ""), user_tz
            )
            modified_datetime = date_handler.convert_utc_to_timezone(
                template.get("lastModifiedDateTime", ""), user_tz
            )

            summaries.append(
                {
                    "number": idx + 1,
                    "id": template.get("id", ""),
                    "subject": template.get("subject", ""),
                    "createdDateTime": created_datetime,
                    "lastModifiedDateTime": modified_datetime,
                    "toRecipients": [
                        r.get("emailAddress", {}).get("address", "")
                        for r in template.get("toRecipients", [])
                    ],
                    "ccRecipients": [
                        r.get("emailAddress", {}).get("address", "")
                        for r in template.get("ccRecipients", [])
                    ],
                    "hasAttachments": template.get("hasAttachments", False),
                    "bodyPreview": template.get("bodyPreview", ""),
                }
            )

        return summaries

    async def get_template(
        self,
        template_id: str,
        return_html: bool = False,
    ) -> Dict[str, Any]:
        """Get a template by ID.

        Args:
            template_id: ID of the template
            return_html: If True, returns full HTML body. If False (default), returns simple text body.

        Returns:
            Template details
        """
        params = {
            "$select": "subject,toRecipients,ccRecipients,bccRecipients,body,createdDateTime,lastModifiedDateTime,hasAttachments",
            "$expand": "attachments($select=name,size,contentType,isInline)",
        }

        template = await self.get(f"/me/messages/{template_id}", params=params)

        body_content = template.get("body", {}).get("content", "")
        body_type = template.get("body", {}).get("contentType", "Text")

        if not return_html and body_type.lower() == "html":
            body_content = self._extract_text_from_html(body_content)

        to_recipients = [
            {
                "name": r.get("emailAddress", {}).get("name", ""),
                "email": r.get("emailAddress", {}).get("address", ""),
            }
            for r in template.get("toRecipients", [])
        ]
        cc_recipients = [
            {
                "name": r.get("emailAddress", {}).get("name", ""),
                "email": r.get("emailAddress", {}).get("address", ""),
            }
            for r in template.get("ccRecipients", [])
        ]
        bcc_recipients = [
            {
                "name": r.get("emailAddress", {}).get("name", ""),
                "email": r.get("emailAddress", {}).get("address", ""),
            }
            for r in template.get("bccRecipients", [])
        ]

        attachments = []
        for attachment in template.get("attachments", []):
            if attachment.get("@odata.type") == "#microsoft.graph.fileAttachment":
                attachments.append(
                    {
                        "name": attachment.get("name"),
                        "size": attachment.get("size"),
                        "contentType": attachment.get("contentType"),
                        "isInline": attachment.get("isInline", False),
                    }
                )

        user_timezone_str = await self.get_user_timezone()

        return {
            "content": {
                "subject": template.get("subject", ""),
                "to": to_recipients,
                "cc": cc_recipients,
                "bcc": bcc_recipients,
                "body": body_content,
                "bodyType": body_type,
                "attachments": attachments,
            },
            "metadata": {
                "id": template.get("id", ""),
                "createdDateTime": date_handler.convert_utc_to_user_timezone(
                    template.get("createdDateTime", ""), user_timezone_str
                ),
                "lastModifiedDateTime": date_handler.convert_utc_to_user_timezone(
                    template.get("lastModifiedDateTime", ""), user_timezone_str
                ),
                "hasAttachments": template.get("hasAttachments", False),
            },
        }

    async def update_template(
        self,
        template_id: str,
        subject: Optional[str] = None,
        body: Optional[str] = None,
        to_recipients: Optional[List[Dict[str, str]]] = None,
        cc_recipients: Optional[List[Dict[str, str]]] = None,
        bcc_recipients: Optional[List[Dict[str, str]]] = None,
    ) -> Dict[str, Any]:
        """Update a template.

        Args:
            template_id: ID of the template to update
            subject: New subject (optional)
            body: New body content (optional)
            to_recipients: New to recipients (optional)
            cc_recipients: New cc recipients (optional)
            bcc_recipients: New bcc recipients (optional)

        Returns:
            Updated template information
        """
        update_data = {}

        if subject is not None:
            update_data["subject"] = subject

        if body is not None:
            update_data["body"] = {"contentType": "HTML", "content": body}

        if to_recipients is not None:
            update_data["toRecipients"] = [
                {"emailAddress": {"address": r.get("email", r.get("address"))}}
                for r in to_recipients
            ]

        if cc_recipients is not None:
            update_data["ccRecipients"] = [
                {"emailAddress": {"address": r.get("email", r.get("address"))}}
                for r in cc_recipients
            ]

        if bcc_recipients is not None:
            update_data["bccRecipients"] = [
                {"emailAddress": {"address": r.get("email", r.get("address"))}}
                for r in bcc_recipients
            ]

        result = await self.patch(f"/me/messages/{template_id}", json=update_data)

        return {
            "id": template_id,
            "subject": result.get("subject", ""),
            "updated": True,
        }

    async def delete_template(
        self,
        template_id: str,
    ) -> Dict[str, Any]:
        """Delete a template by moving it to Deleted Items (soft delete).

        Args:
            template_id: ID of the template to delete

        Returns:
            Deletion confirmation
        """
        deleted_items_id = await self._get_folder_id_by_path("Deleted Items")
        move_data = {"destinationId": deleted_items_id}
        await self.post(f"/me/messages/{template_id}/move", data=move_data)
        await asyncio.sleep(2.0)

        return {
            "id": template_id,
            "deleted": True,
            "message": "Template moved to Deleted Items",
        }

    async def send_template(
        self,
        template_id: str,
        to: Optional[List[str]] = None,
        cc: Optional[List[str]] = None,
        bcc: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Send a template (creates a copy and sends it, preserving the original template).

        Args:
            template_id: ID of the template to send
            to: Optional list of recipient email addresses (overrides template's to)
            cc: Optional list of cc recipient email addresses (overrides template's cc)
            bcc: Optional list of bcc recipient email addresses (overrides template's bcc)

        Returns:
            Sent email information including the saved copy ID
        """
        template = await self.get(
            f"/me/messages/{template_id}",
            params={
                "$select": "subject,body,toRecipients,ccRecipients,bccRecipients,attachments",
                "$expand": "attachments",
            },
        )

        email_data = {
            "subject": template.get("subject", ""),
            "body": template.get("body", {}),
            "toRecipients": template.get("toRecipients", []),
            "ccRecipients": template.get("ccRecipients", []),
            "bccRecipients": template.get("bccRecipients", []),
        }

        if to:
            email_data["toRecipients"] = [
                {"emailAddress": {"address": email}} for email in to
            ]

        if cc:
            email_data["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc
            ]

        if bcc:
            email_data["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc
            ]

        templates_folder_id = await self._get_or_create_templates_folder()

        copy_data = {
            "destinationId": templates_folder_id,
        }

        copy_result = await self.post(
            f"/me/messages/{template_id}/copy", json=copy_data
        )

        saved_copy_id = copy_result.get("id")

        await self.post(f"/me/messages/{saved_copy_id}/send")

        return {
            "subject": email_data["subject"],
            "to": [
                r.get("emailAddress", {}).get("address", "")
                for r in email_data["toRecipients"]
            ],
            "sent": True,
            "savedCopyId": saved_copy_id,
        }

    async def search_emails_by_sender(
        self,
        sender: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Search emails by sender."""
        if top > MAX_EMAIL_SEARCH_LIMIT:
            raise ValueError(
                f"Maximum number of emails per search is {MAX_EMAIL_SEARCH_LIMIT}"
            )

        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        filter_parts = [f"from/emailAddress/address eq '{sender}'"]

        if days is not None:
            start_date, _ = date_handler.get_filter_date_range(days)
            filter_parts.append(f"receivedDateTime ge {start_date}")

        params = {
            "$filter": " and ".join(filter_parts),
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview",
        }

        endpoint = "/me/messages"
        if folder:
            folder_lower = folder.lower()
            if folder_lower in self._well_known_folders:
                endpoint = (
                    f"/me/mailFolders/{self._well_known_folders[folder_lower]}/messages"
                )
            else:
                folder_id = await self._get_folder_id_by_path(folder)
                endpoint = f"/me/mailFolders/{folder_id}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        filter_parts = [f"toRecipients/any(r: r/emailAddress/address eq '{recipient}')"]

        if days is not None:
            start_date, _ = date_handler.get_filter_date_range(days)
            filter_parts.append(f"receivedDateTime ge {start_date}")

        params = {
            "$filter": " and ".join(filter_parts),
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview",
        }

        endpoint = "/me/messages"
        if folder:
            folder_lower = folder.lower()
            if folder_lower in self._well_known_folders:
                endpoint = (
                    f"/me/mailFolders/{self._well_known_folders[folder_lower]}/messages"
                )
            else:
                folder_id = await self._get_folder_id_by_path(folder)
                endpoint = f"/me/mailFolders/{folder_id}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])
        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        escaped_subject = subject.replace("'", "''")
        filter_parts = [f"contains(subject, '{escaped_subject}')"]

        if days is not None:
            start_date, _ = date_handler.get_filter_date_range(days)
            filter_parts.append(f"receivedDateTime ge {start_date}")

        params = {
            "$filter": " and ".join(filter_parts),
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview",
        }

        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])

        # Client-side filtering for inferenceClassification when $search was used
        if needs_client_classification_filter:
            emails = [
                email for email in emails
                if email.get("inferenceClassification", "focused") == inference_classification
            ]

        summaries = [
            self._create_email_summary(email, idx + 1, user_tz)
            for idx, email in enumerate(emails)
        ]

        sorted_summaries = sorted(
            summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True
        )

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(
            sorted_summaries, user_timezone_str
        )
        filter_date_range = (
            date_handler.format_filter_date_range(days, user_timezone_str)
            if days is not None
            else None
        )

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range,
        }

    async def search_emails_by_body(
        self,
        body: str,
        folder: str = "Inbox",
        top: int = 10,
        days: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Search emails by body content."""
        if top > MAX_EMAIL_SEARCH_LIMIT:
            raise ValueError(
                f"Maximum number of emails per search is {MAX_EMAIL_SEARCH_LIMIT}"
            )

        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        escaped_body = body.replace("'", "''")
        filter_parts = [f"contains(body, '{escaped_body}')"]

        if days is not None:
            start_date, _ = date_handler.get_filter_date_range(days)
            filter_parts.append(f"receivedDateTime ge {start_date}")

        params = {
            "$filter": " and ".join(filter_parts),
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview",
        }

        endpoint = "/me/messages"
        if folder:
            folder_id = await self._get_folder_id_by_path(folder)
            endpoint = f"/me/mailFolders/{folder_id}/messages"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])

        # Client-side filtering for inferenceClassification when $search was used
        if needs_client_classification_filter:
            emails = [
                email for email in emails
                if email.get("inferenceClassification", "focused") == inference_classification
            ]

        summaries = [
            self._create_email_summary(email, idx + 1, user_tz)
            for idx, email in enumerate(emails)
        ]

        sorted_summaries = sorted(
            summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True
        )

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        date_range = date_handler.format_email_date_range(
            sorted_summaries, user_timezone_str
        )
        filter_date_range = (
            date_handler.format_filter_date_range(days, user_timezone_str)
            if days is not None
            else None
        )

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "date_range": date_range,
            "filter_date_range": filter_date_range,
        }

    async def send_message(self, message_data: Dict[str, Any]) -> Dict[str, Any]:
        """Send an email message."""
        return await self.post("/me/sendMail", data={"message": message_data})

    async def search_emails(
        self,
        query: Optional[str] = None,
        search_type: Optional[str] = None,
        start_date: Optional[str] = None,
        end_date: Optional[str] = None,
        folder: str = "Inbox",
        top: int = 20,
        inference_classification: str = "focused",
    ) -> Dict[str, Any]:
        """Search or list emails by keywords, sender, recipient, subject, or body.
        Follows the same pattern as search_events for consistency.

        Args:
            query: Search query (keywords for general search, or specific value for search_type)
            search_type: Type of search (sender, recipient, subject, body). If None, does general search.
            start_date: Start date in UTC ISO format (converted from user local time by handler)
            end_date: End date in UTC ISO format (converted from user local time by handler)
            folder: Folder to search in (default: "Inbox")
            top: Number of results to return
            inference_classification: Filter by Focused Inbox classification ('focused', 'other', or 'all'). Default: 'focused'

        Returns:
            Dictionary with email summaries, count, and timezone
        """
        if top > MAX_EMAIL_SEARCH_LIMIT:
            raise ValueError(
                f"Maximum number of emails per search is {MAX_EMAIL_SEARCH_LIMIT}"
            )

        user_timezone_str = await self.get_user_timezone()
        user_tz = date_handler.get_user_timezone_object(user_timezone_str)

        params = {
            "$top": top,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,hasAttachments,importance,bodyPreview,inferenceClassification",
            "$expand": "attachments($select=id,name,size,contentType,isInline)",
        }

        endpoint = "/me/messages"
        if folder:
            folder_lower = folder.lower()
            if folder_lower in self._well_known_folders:
                endpoint = (
                    f"/me/mailFolders/{self._well_known_folders[folder_lower]}/messages"
                )
            else:
                folder_id = await self._get_folder_id_by_path(folder)
                endpoint = f"/me/mailFolders/{folder_id}/messages"

        filter_parts = []

        # Track if we need client-side filtering for inferenceClassification
        # (needed when using $search since it doesn't support $filter)
        needs_client_classification_filter = False

        # Add inferenceClassification filter for Focused Inbox
        # Will be applied via $filter or client-side filtering depending on search usage
        if inference_classification in ("focused", "other"):
            filter_parts.append(f"inferenceClassification eq '{inference_classification}'")

        # Helper function to convert ISO date to KQL date format (YYYY-MM-DD)
        def iso_to_kql_date(iso_date: str) -> str:
            """Convert ISO date string to KQL date format (YYYY-MM-DD)."""
            if not iso_date:
                return ""
            # Extract date part from ISO format (e.g., "2026-02-26T10:30:00Z" -> "2026-02-26")
            return iso_date[:10] if len(iso_date) >= 10 else iso_date

        # Build KQL date filter for $search queries
        # KQL syntax: The entire search expression must be wrapped in double quotes
        # Date filter format: received:YYYY-MM-DD..YYYY-MM-DD (inside the quotes)
        kql_date_filter = ""
        if start_date or end_date:
            start_kql = iso_to_kql_date(start_date) if start_date else None
            end_kql = iso_to_kql_date(end_date) if end_date else None
            
            if start_kql and end_kql:
                # Date range: received:2026-02-01..2026-02-26
                kql_date_filter = f" received:{start_kql}..{end_kql}"
            elif start_kql:
                # Date from: received:>=2026-02-01
                kql_date_filter = f" received>={start_kql}"
            elif end_kql:
                # Date until: received:<=2026-02-26
                kql_date_filter = f" received<={end_kql}"

        # Use server-side filtering optimized for performance
        # KQL syntax: entire expression wrapped in double quotes
        # Format: "property:value AND property:value"
        if search_type:
            if search_type == "subject" and query:
                # Use $search for subject with KQL date filter (server-side filtering)
                # KQL: "subject:value received:YYYY-MM-DD..YYYY-MM-DD"
                params["$search"] = f'"subject:{query}{kql_date_filter}"'
            elif search_type == "body" and query:
                # Use $search for body with KQL date filter (server-side filtering)
                params["$search"] = f'"{query}{kql_date_filter}"'
            elif search_type == "sender" and query:
                # For sender: use $filter for exact email (supports date filtering)
                # Use $search for fuzzy name match with KQL date filter
                if "@" in query:
                    # Exact email match - use $filter for better performance with date filtering
                    filter_parts.append(f"from/emailAddress/address eq '{query}'")
                else:
                    # Fuzzy name match - use $search with KQL date filter
                    params["$search"] = f'"from:{query}{kql_date_filter}"'
        elif query:
            # Default: search subject with KQL date filter
            params["$search"] = f'"subject:{query}{kql_date_filter}"'

        # Add $filter for date range only when NOT using $search
        # (when using $search, date is already included in KQL query)
        if not params.get("$search"):
            if start_date and end_date:
                filter_parts.append(f"receivedDateTime ge {start_date}")
                filter_parts.append(f"receivedDateTime le {end_date}")
            elif start_date:
                filter_parts.append(f"receivedDateTime ge {start_date}")
            elif end_date:
                filter_parts.append(f"receivedDateTime le {end_date}")

        if filter_parts and not params.get("$search"):
            params["$filter"] = " and ".join(filter_parts)
        elif params.get("$search") and inference_classification in ("focused", "other"):
            # When using $search, we can't use $filter, so mark for client-side filtering
            needs_client_classification_filter = True

        # Add $orderby only when NOT using $search or $filter (Graph API limitations)
        # $search doesn't support $orderby
        # $filter with date filters doesn't support $orderby (InefficientFilter error)
        if not params.get("$search") and not filter_parts:
            params["$orderby"] = "receivedDateTime desc"

        result = await self.get(endpoint, params=params)

        emails = result.get("value", [])

        # Client-side filtering for inferenceClassification when $search was used
        if needs_client_classification_filter:
            emails = [
                email for email in emails
                if email.get("inferenceClassification", "focused") == inference_classification
            ]

        summaries = [
            self._create_email_summary(email, idx + 1, user_tz)
            for idx, email in enumerate(emails)
        ]

        # Note: Client-side date filtering is no longer needed because
        # KQL date filters are now embedded in $search query for server-side filtering.
        # KQL uses date format (YYYY-MM-DD) which filters at day granularity.
        # This provides good performance while ensuring emails are not missed.

        # Sort by receivedDateTime descending
        sorted_summaries = sorted(
            summaries, key=lambda x: x.get("receivedDateTime", ""), reverse=True
        )

        for idx, summary in enumerate(sorted_summaries):
            summary["number"] = idx + 1

        return {
            "metadata": sorted_summaries,
            "count": len(sorted_summaries),
            "timezone": user_timezone_str,
        }

    async def forward_message(
        self,
        message_id: str,
        body: str,
        body_content_type: str = "HTML",
        to_recipients: Optional[List[str]] = None,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        subject: Optional[str] = None,
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Forward an email message using Microsoft Graph API.

        This method appends the original email thread to the user's forward body.
        The LLM must generate HTML directly when calling this tool.
        Inline images from the original email are re-attached to preserve display.

        IMPORTANT HTML FORMATTING:
        - The body parameter MUST be valid HTML content (for body_content_type="HTML")
        - Do NOT convert newlines to <br> - the body is treated as raw HTML
        - Use <p> tags for paragraphs, not <br> between paragraphs
        - The normalize_email_html function will clean up any whitespace issues

        Args:
            message_id: The ID of the message to forward
            body: Forward body content (must be HTML)
            body_content_type: Content type for body (always 'HTML')
            to_recipients: Optional list of recipient email addresses (defaults to original sender)
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            subject: Optional subject for the forward (defaults to "FW: " + original subject)
            importance: Optional importance level ('normal', 'high', 'low')

        Returns:
            Response from the Graph API
        """
        original_email = await self.get_email(message_id, return_html=True)

        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", "")

        forward_subject = subject if subject else f"FW: {original_subject}"

        original_body_content = original_body

        body_match = re.search(
            r"<body[^>]*>(.*?)</body>", original_body, re.DOTALL | re.IGNORECASE
        )
        if body_match:
            original_body_content = body_match.group(1)
        else:
            html_match = re.search(
                r"<html[^>]*>(.*?)</html>", original_body, re.DOTALL | re.IGNORECASE
            )
            if html_match:
                original_body_content = html_match.group(1)

        quoted_forward = f"""
<br><br>
<hr>
<div>
    <b>From:</b> {from_name} &lt;{from_address}&gt;<br>
    <b>Sent:</b> {sent_date}<br>
    <b>Subject:</b> {original_subject}
</div>
<br>
{original_body_content}
"""
        forward_body = body + quoted_forward

        message_data = {
            "subject": forward_subject,
            "body": {"contentType": body_content_type, "content": forward_body},
        }

        if to_recipients:
            message_data["toRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ]

        if cc_recipients:
            message_data["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_recipients
            ]

        if bcc_recipients:
            message_data["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc_recipients
            ]

        params = {"$expand": "attachments($select=id,name,contentType,isInline)"}
        email_with_attachments = await self.get(
            f"/me/messages/{message_id}", params=params
        )
        attachments = email_with_attachments.get("attachments", [])
        inline_attachments = []

        for attachment in attachments:
            if attachment.get("isInline", False):
                attachment_id = attachment.get("id", "")

                try:
                    attachment_with_content = await self.get(
                        f"/me/messages/{message_id}/attachments/{attachment_id}"
                    )
                    content_bytes = attachment_with_content.get("contentBytes", "")
                    content_id = attachment_with_content.get("contentId", "")
                except Exception as e:
                    logger.warning(f"Could not fetch attachment content for {attachment.get('name')}: {e}")
                    content_bytes = ""
                    content_id = ""

                if content_id.startswith("<") and content_id.endswith(">"):
                    content_id = content_id[1:-1]

                inline_attachments.append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": attachment.get("name", ""),
                        "contentType": attachment.get("contentType", ""),
                        "contentBytes": content_bytes,
                        "isInline": True,
                        "id": attachment.get("id", ""),
                        "contentId": content_id,
                    }
                )

        if inline_attachments:
            message_data["attachments"] = inline_attachments

        if importance:
            message_data["importance"] = importance

        return await self.send_message(message_data)

    async def reply_to_message(
        self,
        message_id: str,
        body: str,
        body_content_type: str = "HTML",
        to_recipients: Optional[List[str]] = None,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        subject: Optional[str] = None,
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Reply to an email message using Microsoft Graph API.

        This method appends the original email thread to the user's reply body.
        The LLM must generate HTML directly when calling this tool.
        Inline images from the original email are re-attached to preserve display.

        IMPORTANT HTML FORMATTING:
        - The body parameter MUST be valid HTML content (for body_content_type="HTML")
        - Do NOT convert newlines to <br> - the body is treated as raw HTML
        - Use <p> tags for paragraphs, not <br> between paragraphs
        - The normalize_email_html function will clean up any whitespace issues

        Args:
            message_id: The ID of the message to reply to
            body: Reply body content (must be HTML for body_content_type="HTML")
            body_content_type: Content type for body (always 'HTML')
            to_recipients: Optional list of recipient email addresses (defaults to original sender)
            cc_recipients: Optional list of CC recipient email addresses (defaults to original CC)
            bcc_recipients: Optional list of BCC recipient email addresses
            subject: Optional subject for the reply (defaults to original subject)
            importance: Optional importance level ('normal', 'high', 'low')

        Returns:
            Response from the Graph API
        """
        original_email = await self.get_email(message_id, return_html=True)

        email_content = original_email.get("content", {})
        email_metadata = original_email.get("metadata", {})

        from_email = email_content.get("from", {})
        from_name = from_email.get("name", "")
        from_address = from_email.get("email", "")
        sent_date = email_metadata.get("sentDateTime", "")
        original_subject = email_content.get("subject", "")
        original_body = email_content.get("body", "")

        # Get user's email to filter out from recipients
        user_email = await self.get_user_email()

        # Outlook behavior: If to is None, use Reply All behavior
        if to_recipients is None:
            # Reply All: TO includes sender + original TO recipients
            # Skip adding sender if it's user's own email (e.g., from sent items)
            to_recipients = [from_address] if from_address != user_email else []

            # Add original TO recipients (these should stay in TO, not move to CC)
            original_to = email_content.get("to", [])
            for recipient in original_to:
                recipient_email = recipient.get("email", "")
                # Filter out user's own email
                if recipient_email and recipient_email not in to_recipients and recipient_email != user_email:
                    to_recipients.append(recipient_email)

        # Outlook behavior: If cc is None, keep original CC recipients
        if cc_recipients is None:
            original_cc = email_content.get("cc", [])
            cc_recipients = [recipient.get("email", "") for recipient in original_cc if recipient.get("email") and recipient.get("email") != user_email]

        reply_subject = subject if subject else f"RE: {original_subject}"

        original_body_content = original_body

        body_match = re.search(
            r"<body[^>]*>(.*?)</body>", original_body, re.DOTALL | re.IGNORECASE
        )
        if body_match:
            original_body_content = body_match.group(1)
        else:
            html_match = re.search(
                r"<html[^>]*>(.*?)</html>", original_body, re.DOTALL | re.IGNORECASE
            )
            if html_match:
                original_body_content = html_match.group(1)

        if body_content_type == "Text":
            original_body_content = re.sub(r"<[^>]+>", "", original_body_content)
            original_body_content = re.sub(r"\s+", " ", original_body_content).strip()
            quoted_reply = f"""

-----
From: {from_name} <{from_address}>
Sent: {sent_date}
Subject: {original_subject}

{original_body_content}
"""
            reply_body = body + quoted_reply
        else:
            quoted_reply = f"""
<br><br>
<hr>
<div>
    <b>From:</b> {from_name} &lt;{from_address}&gt;<br>
    <b>Sent:</b> {sent_date}<br>
    <b>Subject:</b> {original_subject}
</div>
<br>
{original_body_content}
"""
            # body is already HTML - do not convert newlines to <br>
            # as it would create excessive spacing when HTML already has proper formatting
            reply_body = body + quoted_reply

        # Use Graph's createReply to obtain a draft carrying In-Reply-To,
        # References, and the original conversationId. PATCH our composed
        # body/recipients onto the draft, then POST /send. This is required
        # for Outlook conversation threading; /me/sendMail with a manually
        # prefixed "RE:" subject does NOT thread (createReply is the only
        # supported way to bind a reply to the original message envelope).
        draft = await self.post(
            f"/me/messages/{message_id}/createReply", data={}
        )
        draft_id = draft.get("id")
        if not draft_id:
            raise RuntimeError(
                f"createReply did not return a draft id: {draft!r}"
            )

        patch_data: Dict[str, Any] = {
            "subject": reply_subject,
            "body": {"contentType": body_content_type, "content": reply_body},
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_recipients
            ],
        }
        if cc_recipients:
            patch_data["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_recipients
            ]
        if bcc_recipients:
            patch_data["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc_recipients
            ]
        if importance:
            patch_data["importance"] = importance

        await self.patch(f"/me/messages/{draft_id}", data=patch_data)
        await self.post(f"/me/messages/{draft_id}/send", data={})
        return {"status": "accepted", "id": draft_id}

    async def send_email(
        self,
        to_recipients: Optional[List[str]] = None,
        subject: str = "",
        body: str = "",
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        reply_to_message_id: Optional[str] = None,
        forward_to_message_id: Optional[str] = None,
        body_content_type: str = "Text",
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Unified backend function to send emails (send_new, reply, or forward).

        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            reply_to_message_id: Optional message ID to reply to
            forward_to_message_id: Optional message ID to forward
            body_content_type: Content type for body ('Text' or 'HTML')
            importance: Optional importance level ('normal', 'high', 'low')

        Returns:
            Response from the Graph API
        """

        if reply_to_message_id:
            logger.debug(f"send_email: Routing to reply_to_message, body first 100 chars: {repr(body[:100])}")
            logger.debug(f"send_email: body has {repr(body.count(chr(10)))} newlines, body_content_type={repr(body_content_type)}")

            return await self.reply_to_message(
                message_id=reply_to_message_id,
                body=body,
                body_content_type=body_content_type,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                subject=subject,
                importance=importance,
            )

        if forward_to_message_id:
            return await self.forward_message(
                message_id=forward_to_message_id,
                body=body,
                body_content_type=body_content_type,
                to_recipients=to_recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                subject=subject,
                importance=importance,
            )

        message_data = {
            "subject": subject,
            "body": {"contentType": body_content_type, "content": body},
            "toRecipients": [
                {"emailAddress": {"address": email}} for email in to_recipients
            ],
        }

        if cc_recipients:
            message_data["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_recipients
            ]

        if bcc_recipients:
            message_data["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_recipients
            ]

        if importance:
            message_data["importance"] = importance

        return await self.post("/me/sendMail", data={"message": message_data})

    async def batch_forward_emails(
        self,
        to_recipients: Optional[List[str]] = None,
        subject: str = "",
        body: str = "",
        email_ids: List[str] = [],
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        body_content_type: str = "Text",
        importance: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Backend function to forward multiple emails to recipients with batch BCC support.

        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            email_ids: List of email IDs to forward (only first email is used)
            cc_recipients: Optional list of CC recipient email addresses
            bcc_recipients: Optional list of BCC recipient email addresses
            body_content_type: Content type for body ('Text' or 'HTML')
            importance: Optional importance level ('normal', 'high', 'low')

        Returns:
            Response from the Graph API
        """
        return await self.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            forward_to_message_id=email_ids[0] if email_ids else None,
            body_content_type=body_content_type,
            importance=importance,
        )

    async def create_folder(
        self, folder_name: str, parent_folder: Optional[str] = None
    ) -> Dict[str, Any]:
        """Create a new mail folder.

        Args:
            folder_name: Name of the folder to create
            parent_folder: Optional parent folder path (e.g., 'Inbox', 'Archive/2024')

        Returns:
            Created folder information
        """
        folder_data = {"displayName": folder_name}

        if parent_folder:
            parent_folder_id = await self._get_folder_id_by_path(parent_folder)
            result = await self.post(
                f"/me/mailFolders/{parent_folder_id}/childFolders", data=folder_data
            )
        else:
            result = await self.post("/me/mailFolders", data=folder_data)

        await asyncio.sleep(2.0)

        return {
            "id": result.get("id"),
            "displayName": result.get("displayName"),
            "parentFolderId": result.get("parentFolderId"),
            "childFolderCount": result.get("childFolderCount", 0),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
        }

    async def delete_folder(self, folder_path: str) -> Dict[str, Any]:
        """Delete a mail folder by moving it to Deleted Items.

        Args:
            folder_path: Path of the folder to delete (e.g., 'Inbox/Projects')

        Returns:
            Success message
        """
        folder_id = await self._get_folder_id_by_path(folder_path)

        folder_name = folder_path.split("/")[-1]

        deleted_items_id = await self._get_folder_id_by_path("Deleted Items")

        try:
            child_folders_result = await self.get(
                f"/me/mailFolders/{deleted_items_id}/childFolders"
            )
            child_folders = child_folders_result.get("value", [])

            existing_folder = next(
                (f for f in child_folders if f.get("displayName") == folder_name), None
            )

            if existing_folder:
                existing_folder_id = existing_folder.get("id")
                await self.delete(f"/me/mailFolders/{existing_folder_id}")
                await asyncio.sleep(1.0)
        except Exception:
            pass

        move_data = {"destinationId": deleted_items_id}
        await self.post(f"/me/mailFolders/{folder_id}/move", data=move_data)
        await asyncio.sleep(2.0)
        return {
            "status": "success",
            "message": f"Folder '{folder_path}' moved to Deleted Items",
        }

    async def rename_folder(self, folder_path: str, new_name: str) -> Dict[str, Any]:
        """Rename a mail folder.

        Args:
            folder_path: Path of the folder to rename (e.g., 'Inbox/OldName')
            new_name: New name for the folder

        Returns:
            Updated folder information
        """
        folder_id = await self._get_folder_id_by_path(folder_path)
        folder_data = {"displayName": new_name}
        result = await self.patch(f"/me/mailFolders/{folder_id}", data=folder_data)

        await asyncio.sleep(2.0)

        return {
            "id": result.get("id"),
            "displayName": result.get("displayName"),
            "parentFolderId": result.get("parentFolderId"),
            "childFolderCount": result.get("childFolderCount", 0),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
        }

    async def get_folder_details(self, folder_path: str) -> Dict[str, Any]:
        """Get detailed information about a specific folder.

        Args:
            folder_path: Path of the folder (e.g., 'Inbox', 'Archive/2024')

        Returns:
            Detailed folder information
        """
        folder_id = await self._get_folder_id_by_path(folder_path)
        result = await self.get(f"/me/mailFolders/{folder_id}")

        return {
            "id": result.get("id"),
            "displayName": result.get("displayName"),
            "parentFolderId": result.get("parentFolderId"),
            "childFolderCount": result.get("childFolderCount", 0),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
            "wellKnownName": result.get("wellKnownName"),
            "sizeInBytes": result.get("sizeInBytes", 0),
        }

    async def move_email_to_folder(
        self, email_id: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Move an email to a different folder.

        Args:
            email_id: ID of the email to move
            destination_folder: Path of the destination folder (e.g., 'Archive/2024')

        Returns:
            Success message
        """
        destination_folder_id = await self._get_folder_id_by_path(destination_folder)
        move_data = {"destinationId": destination_folder_id}
        await self.post(f"/me/messages/{email_id}/move", data=move_data)
        await asyncio.sleep(2.0)
        return {
            "status": "success",
            "message": f"Email moved to '{destination_folder}'",
        }

    async def copy_email_to_folder(
        self, email_id: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Copy an email to a different folder.

        Args:
            email_id: ID of the email to copy
            destination_folder: Path of the destination folder (e.g., 'Archive/2024')

        Returns:
            Copied email information
        """
        destination_folder_id = await self._get_folder_id_by_path(destination_folder)
        copy_data = {"destinationId": destination_folder_id}
        result = await self.post(f"/me/messages/{email_id}/copy", data=copy_data)
        await asyncio.sleep(2.0)
        return {
            "id": result.get("id"),
            "status": "success",
            "message": f"Email copied to '{destination_folder}'",
        }

    async def move_all_emails_from_folder(
        self, source_folder: str, destination_folder: str
    ) -> Dict[str, Any]:
        """Move all emails from one folder to another using batch operations.

        Args:
            source_folder: Path of the source folder (e.g., 'Inbox')
            destination_folder: Path of the destination folder (e.g., 'Archive/2024')

        Returns:
            Summary of moved emails
        """
        source_folder_id = await self._get_folder_id_by_path(
            source_folder, max_retries=2, retry_delay=0.2
        )
        destination_folder_id = await self._get_folder_id_by_path(
            destination_folder, max_retries=2, retry_delay=0.2
        )

        params = {"$select": "id", "$top": MAX_EMAIL_SEARCH_LIMIT}

        result = await self.get(
            f"/me/mailFolders/{source_folder_id}/messages", params=params
        )
        emails = result.get("value", [])

        if not emails:
            return {
                "status": "success",
                "message": f"Moved 0 emails from '{source_folder}' to '{destination_folder}'",
                "moved_count": 0,
                "failed_count": 0,
                "errors": None,
            }

        batch_size = 20
        moved_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_emails: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (moved_count, failed_count, errors)."""
            batch_moved = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email in enumerate(batch_emails):
                email_id = email.get("id")
                if not email_id:
                    batch_failed += 1
                    continue

                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "POST",
                        "url": f"/me/messages/{email_id}/move",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"destinationId": destination_folder_id},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_moved += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_moved, batch_failed, batch_errors)

        batches = [
            emails[i : i + batch_size] for i in range(0, len(emails), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_moved, batch_failed, batch_errors in batch_results:
            moved_count += batch_moved
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Moved {moved_count} emails from '{source_folder}' to '{destination_folder}'",
            "moved_count": moved_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }

    async def delete_email(self, email_id: str) -> Dict[str, Any]:
        """Delete an email by moving it to Deleted Items.

        Args:
            email_id: ID of the email to delete

        Returns:
            Success message
        """
        deleted_items_id = await self._get_folder_id_by_path("Deleted Items")
        move_data = {"destinationId": deleted_items_id}
        await self.post(f"/me/messages/{email_id}/move", data=move_data)
        await asyncio.sleep(2.0)
        return {"status": "success", "message": "Email moved to Deleted Items"}

    async def batch_delete_emails(self, email_ids: List[str]) -> Dict[str, Any]:
        """Delete multiple emails using batch operations.

        Args:
            email_ids: List of email IDs to delete

        Returns:
            Result with deleted count and any errors
        """
        deleted_items_id = await self._get_folder_id_by_path("Deleted Items")

        batch_size = 20
        deleted_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_ids: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (deleted_count, failed_count, errors)."""
            batch_deleted = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email_id in enumerate(batch_ids):
                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "POST",
                        "url": f"/me/messages/{email_id}/move",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"destinationId": deleted_items_id},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_deleted += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_deleted, batch_failed, batch_errors)

        batches = [
            email_ids[i : i + batch_size] for i in range(0, len(email_ids), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_deleted, batch_failed, batch_errors in batch_results:
            deleted_count += batch_deleted
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Deleted {deleted_count} emails",
            "deleted_count": deleted_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }

    async def delete_all_emails_from_folder(self, folder_path: str) -> Dict[str, Any]:
        """Delete all emails from a folder using batch operations.

        Args:
            folder_path: Path of the folder to clear (e.g., 'Inbox', 'Archive/2024')

        Returns:
            Result with deleted count and any errors
        """
        folder_id = await self._get_folder_id_by_path(folder_path)
        deleted_items_id = await self._get_folder_id_by_path("Deleted Items")

        params = {"$select": "id", "$top": MAX_EMAIL_SEARCH_LIMIT}

        result = await self.get(f"/me/mailFolders/{folder_id}/messages", params=params)
        emails = result.get("value", [])

        if not emails:
            return {
                "status": "success",
                "message": f"Deleted 0 emails from '{folder_path}'",
                "deleted_count": 0,
                "failed_count": 0,
                "errors": None,
            }

        batch_size = 20
        deleted_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_emails: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (deleted_count, failed_count, errors)."""
            batch_deleted = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email in enumerate(batch_emails):
                email_id = email.get("id")
                if not email_id:
                    batch_failed += 1
                    continue

                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "POST",
                        "url": f"/me/messages/{email_id}/move",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"destinationId": deleted_items_id},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_deleted += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_deleted, batch_failed, batch_errors)

        batches = [
            emails[i : i + batch_size] for i in range(0, len(emails), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_deleted, batch_failed, batch_errors in batch_results:
            deleted_count += batch_deleted
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Deleted {deleted_count} emails from '{folder_path}'",
            "deleted_count": deleted_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }

    async def move_folder(
        self, folder_path: str, destination_parent: str
    ) -> Dict[str, Any]:
        """Move a folder to a different parent folder.

        Args:
            folder_path: Path of the folder to move (e.g., 'Inbox/Projects')
            destination_parent: Path of the destination parent folder (e.g., 'Archive')

        Returns:
            Moved folder information
        """
        folder_id = await self._get_folder_id_by_path(folder_path)
        destination_parent_id = await self._get_folder_id_by_path(destination_parent)
        move_data = {"destinationId": destination_parent_id}
        result = await self.post(f"/me/mailFolders/{folder_id}/move", data=move_data)

        await asyncio.sleep(2.0)

        return {
            "id": result.get("id"),
            "displayName": result.get("displayName"),
            "parentFolderId": result.get("parentFolderId"),
            "childFolderCount": result.get("childFolderCount", 0),
            "totalItemCount": result.get("totalItemCount", 0),
            "unreadItemCount": result.get("unreadItemCount", 0),
        }

    async def archive_email(self, email_id: str) -> Dict[str, Any]:
        """Archive an email by moving it to the Archive folder.

        Args:
            email_id: ID of the email to archive

        Returns:
            Success message
        """
        archive_folder_id = await self._get_folder_id_by_path("Archive")
        move_data = {"destinationId": archive_folder_id}
        await self.post(f"/me/messages/{email_id}/move", data=move_data)
        await asyncio.sleep(2.0)
        return {"status": "success", "message": "Email archived"}

    async def batch_archive_emails(self, email_ids: List[str]) -> Dict[str, Any]:
        """Archive multiple emails using batch operations.

        Args:
            email_ids: List of email IDs to archive

        Returns:
            Result with archived count and any errors
        """
        archive_folder_id = await self._get_folder_id_by_path("Archive")

        batch_size = 20
        archived_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_ids: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (archived_count, failed_count, errors)."""
            batch_archived = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email_id in enumerate(batch_ids):
                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "POST",
                        "url": f"/me/messages/{email_id}/move",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"destinationId": archive_folder_id},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_archived += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_archived, batch_failed, batch_errors)

        batches = [
            email_ids[i : i + batch_size] for i in range(0, len(email_ids), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_archived, batch_failed, batch_errors in batch_results:
            archived_count += batch_archived
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Archived {archived_count} emails",
            "archived_count": archived_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }

    async def flag_email(self, email_id: str, flag_status: str) -> Dict[str, Any]:
        """Flag or unflag an email.

        Args:
            email_id: ID of the email to flag
            flag_status: Flag status ('flagged' or 'complete')

        Returns:
            Success message
        """
        flag_data = {"flag": {"flagStatus": flag_status}}
        await self.patch(f"/me/messages/{email_id}", data=flag_data)
        await asyncio.sleep(2.0)
        return {"status": "success", "message": f"Email {flag_status}"}

    async def batch_flag_emails(
        self, email_ids: List[str], flag_status: str
    ) -> Dict[str, Any]:
        """Flag multiple emails using batch operations.

        Args:
            email_ids: List of email IDs to flag
            flag_status: Flag status ('flagged' or 'complete')

        Returns:
            Result with flagged count and any errors
        """
        batch_size = 20
        flagged_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_ids: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (flagged_count, failed_count, errors)."""
            batch_flagged = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email_id in enumerate(batch_ids):
                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "PATCH",
                        "url": f"/me/messages/{email_id}",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"flag": {"flagStatus": flag_status}},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_flagged += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_flagged, batch_failed, batch_errors)

        batches = [
            email_ids[i : i + batch_size] for i in range(0, len(email_ids), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_flagged, batch_failed, batch_errors in batch_results:
            flagged_count += batch_flagged
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Flagged {flagged_count} emails as {flag_status}",
            "flagged_count": flagged_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }

    async def categorize_email(
        self, email_id: str, categories: List[str]
    ) -> Dict[str, Any]:
        """Add categories to an email.

        Args:
            email_id: ID of the email to categorize
            categories: List of category names to apply

        Returns:
            Success message
        """
        category_data = {"categories": categories}
        await self.patch(f"/me/messages/{email_id}", data=category_data)
        await asyncio.sleep(2.0)
        return {
            "status": "success",
            "message": f"Email categorized with: {', '.join(categories)}",
        }

    async def batch_categorize_emails(
        self, email_ids: List[str], categories: List[str]
    ) -> Dict[str, Any]:
        """Categorize multiple emails using batch operations.

        Args:
            email_ids: List of email IDs to categorize
            categories: List of category names to apply

        Returns:
            Result with categorized count and any errors
        """
        batch_size = 20
        categorized_count = 0
        failed_count = 0
        errors = []

        async def process_batch(batch_ids: list, batch_idx: int) -> tuple:
            """Process a batch of emails and return (categorized_count, failed_count, errors)."""
            batch_categorized = 0
            batch_failed = 0
            batch_errors = []
            requests = []

            for idx, email_id in enumerate(batch_ids):
                request_id = f"req_{batch_idx * batch_size + idx}"
                requests.append(
                    {
                        "id": request_id,
                        "method": "PATCH",
                        "url": f"/me/messages/{email_id}",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"categories": categories},
                    }
                )

            if not requests:
                return (0, 0, [])

            batch_data = {"requests": requests}

            try:
                batch_result = await self.post("/$batch", data=batch_data)
                responses = batch_result.get("responses", [])

                for response in responses:
                    status = response.get("status", 0)
                    request_id = response.get("id", "")

                    if 200 <= status < 300:
                        batch_categorized += 1
                    else:
                        batch_failed += 1
                        body = response.get("body", {})
                        error_msg = body.get("error", {}).get(
                            "message", f"HTTP {status}"
                        )
                        batch_errors.append(f"Request {request_id}: {error_msg}")
            except Exception as e:
                batch_failed += len(requests)
                batch_errors.append(f"Batch {batch_idx + 1} failed: {str(e)}")

            return (batch_categorized, batch_failed, batch_errors)

        batches = [
            email_ids[i : i + batch_size] for i in range(0, len(email_ids), batch_size)
        ]
        batch_tasks = [process_batch(batch, idx) for idx, batch in enumerate(batches)]
        batch_results = await asyncio.gather(*batch_tasks)

        for batch_categorized, batch_failed, batch_errors in batch_results:
            categorized_count += batch_categorized
            failed_count += batch_failed
            errors.extend(batch_errors)

        return {
            "status": "success",
            "message": f"Categorized {categorized_count} emails with: {', '.join(categories)}",
            "categorized_count": categorized_count,
            "failed_count": failed_count,
            "errors": errors if errors else None,
        }
