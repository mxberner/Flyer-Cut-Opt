#!/usr/bin/env python3
"""
Register a generated flyer stack output bundle with the HTMDEC data portal.

This script:
1. Loads one generated stack output bundle.
2. Creates a child stack IGSN via POST /deposition/{id}/split?suffix=<stack_id>.
3. Uploads the generated layout, inventory, and metadata files into
   Collection AIMD-L/Projects/Flyer-Machining/StackIGSN/{stack_id}/.
4. Annotates the uploaded items with stack/deposition metadata.
5. Writes a local registration receipt next to the generated files.
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urljoin


DEFAULT_API_URL = "https://data.htmdec.org/api/v1"
DEFAULT_TARGET_PATH = ("AIMD-L", "Projects", "Flyer-Machining", "StackIGSN")
DEFAULT_COLLECTION_NAME = "AIMD-L"
DEFAULT_CHUNK_SIZE = 1024 * 1024


class ApiError(RuntimeError):
    """Raised when the HTMDEC API returns an unexpected error."""


@dataclass
class OutputBundle:
    """Generated files that belong to one stack output."""

    directory: Path
    metadata_path: Path
    inventory_path: Path
    layout_path: Path
    metadata: dict[str, Any]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Register one generated flyer stack output.")
    parser.add_argument(
        "path",
        help="Path to a stack metadata JSON file or to the output directory that contains it.",
    )
    parser.add_argument(
        "--api-url",
        default=os.environ.get("HTMDEC_API_URL", DEFAULT_API_URL),
        help=f"HTMDEC API base URL. Default: {DEFAULT_API_URL}",
    )
    parser.add_argument(
        "--api-key",
        default=os.environ.get("HTMDEC_API_KEY"),
        help="HTMDEC API key. Can also be set via HTMDEC_API_KEY.",
    )
    parser.add_argument(
        "--parent-deposition-id",
        default=None,
        help="Override the parent foil deposition id. If omitted, it is inferred from the local IGSN config link.",
    )
    parser.add_argument(
        "--collection-name",
        default=DEFAULT_COLLECTION_NAME,
        help=f"Top-level collection name for uploads. Default: {DEFAULT_COLLECTION_NAME}",
    )
    parser.add_argument(
        "--target-leaf",
        default=None,
        help="Override the leaf folder name under StackIGSN. Default: stack ID.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the actions that would be taken without creating/uploading anything.",
    )
    return parser.parse_args()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def normalize_base_url(base_url: str) -> str:
    return base_url.rstrip("/")


def build_bundle(path_str: str) -> OutputBundle:
    raw_path = Path(path_str).expanduser().resolve()
    if raw_path.is_dir():
        metadata_candidates = sorted(raw_path.glob("stack*-metadata.json"))
        require(metadata_candidates, f"No stack metadata JSON found in directory: {raw_path}")
        require(len(metadata_candidates) == 1, f"Expected exactly one metadata JSON in {raw_path}.")
        metadata_path = metadata_candidates[0]
    else:
        metadata_path = raw_path

    require(metadata_path.exists(), f"Metadata file not found: {metadata_path}")
    metadata = load_json(metadata_path)
    bundle_dir = metadata_path.parent

    prefix = metadata_path.name.removesuffix("-metadata.json")
    inventory_path = bundle_dir / f"{prefix}-inventory.csv"
    layout_candidates = list(bundle_dir.glob(f"{prefix}-layout.*"))
    require(inventory_path.exists(), f"Inventory CSV not found: {inventory_path}")
    require(layout_candidates, f"Layout file not found for prefix: {prefix}")
    require(len(layout_candidates) == 1, f"Expected exactly one layout file for prefix: {prefix}")

    require(str(metadata.get("ID", "")).strip(), "Metadata JSON missing ID.")
    material = metadata.get("material", {}) or {}
    require(str(material.get("igsn", "")).strip(), "Metadata JSON missing material.igsn.")

    return OutputBundle(
        directory=bundle_dir,
        metadata_path=metadata_path,
        inventory_path=inventory_path,
        layout_path=layout_candidates[0],
        metadata=metadata,
    )


def parse_deposition_id_from_link(link: str) -> str | None:
    if "#deposition/" not in link:
        return None
    return link.rsplit("#deposition/", 1)[-1].strip() or None


def find_repo_root(start: Path) -> Path:
    for path in [start, *start.parents]:
        if (path / "inputs" / "igsn").exists():
            return path
    return Path.cwd()


def resolve_parent_deposition_id(bundle: OutputBundle, repo_root: Path, override: str | None) -> str:
    if override:
        return override

    material = bundle.metadata.get("material", {}) or {}
    config_name = str(material.get("config_name", "")).strip()
    require(config_name, "Metadata JSON missing material.config_name; cannot infer parent deposition id.")

    igsn_path = repo_root / "inputs" / "igsn" / config_name
    require(igsn_path.exists(), f"IGSN config not found for metadata material.config_name: {igsn_path}")

    igsn_cfg = load_json(igsn_path)
    link = str((igsn_cfg.get("material", {}) or {}).get("link", "")).strip()
    deposition_id = parse_deposition_id_from_link(link)
    require(
        deposition_id,
        f"Could not parse deposition id from IGSN config material.link: {link or '<empty>'}",
    )
    return deposition_id


class HtmdecClient:
    def __init__(self, base_url: str, api_key: str | None, dry_run: bool = False):
        self.base_url = normalize_base_url(base_url)
        self.api_key = api_key
        self.dry_run = dry_run
        self.token = None if dry_run else self._exchange_api_key()

    def _exchange_api_key(self) -> str:
        require(self.api_key, "An HTMDEC API key is required.")
        response = self.request(
            "POST",
            "/api_key/token",
            query={"key": self.api_key},
            include_token=False,
        )
        token = ((response or {}).get("authToken") or {}).get("token")
        require(token, "API key exchange did not return an auth token.")
        return token

    def request(
        self,
        method: str,
        path: str,
        *,
        query: dict[str, Any] | None = None,
        data: bytes | None = None,
        json_body: dict[str, Any] | None = None,
        headers: dict[str, str] | None = None,
        include_token: bool = True,
        expect_json: bool = True,
    ) -> Any:
        url = urljoin(self.base_url + "/", path.lstrip("/"))
        if query:
            encoded = urlencode({k: v for k, v in query.items() if v is not None}, doseq=True)
            url = f"{url}?{encoded}"

        request_headers = dict(headers or {})
        body = data
        if json_body is not None:
            body = json.dumps(json_body).encode("utf-8")
            request_headers.setdefault("Content-Type", "application/json")
        if include_token and getattr(self, "token", None):
            request_headers["Girder-Token"] = self.token

        command = [
            "curl",
            "-L",
            "--silent",
            "--show-error",
            "-X",
            method.upper(),
            "-w",
            "\n%{http_code}",
        ]
        for key, value in request_headers.items():
            command.extend(["-H", f"{key}: {value}"])
        if body is not None:
            command.extend(["--data-binary", "@-"])
        command.append(url)

        try:
            completed = subprocess.run(
                command,
                input=body,
                capture_output=True,
                check=False,
            )
        except OSError as exc:
            raise ApiError(f"Failed to run curl for {method.upper()} {url}: {exc}") from exc

        if completed.returncode != 0:
            stderr = completed.stderr.decode("utf-8", errors="replace")
            raise ApiError(f"{method.upper()} {url} failed via curl: {stderr}")

        stdout = completed.stdout
        body_bytes, _, status_bytes = stdout.rpartition(b"\n")
        status_text = status_bytes.decode("utf-8", errors="replace").strip()
        if not status_text.isdigit():
            raise ApiError(f"{method.upper()} {url} returned an unreadable HTTP status.")
        status_code = int(status_text)

        if status_code >= 400:
            message = body_bytes.decode("utf-8", errors="replace")
            raise ApiError(f"{method.upper()} {url} failed with HTTP {status_code}: {message}")

        if not expect_json:
            return body_bytes
        if not body_bytes:
            return None
        return json.loads(body_bytes.decode("utf-8"))

    def split_deposition(self, parent_deposition_id: str, suffix: str) -> dict[str, Any]:
        if self.dry_run:
            return {
                "_id": "<dry-run-child-deposition-id>",
                "igsn": f"<dry-run-parent>-{suffix}",
                "parentId": parent_deposition_id,
                "metadata": {"alternateIdentifiers": []},
            }
        response = self.request(
            "POST",
            f"/deposition/{parent_deposition_id}/split",
            query={"suffix": suffix},
        )
        require(isinstance(response, dict), "Split response was not a JSON object.")
        return response

    def find_collection(self, name: str) -> dict[str, Any]:
        if self.dry_run:
            return {"_id": f"<dry-run-collection:{name}>", "name": name}
        collections = self.request("GET", "/collection", query={"text": name})
        require(isinstance(collections, list), "Collection lookup returned unexpected data.")
        for collection in collections:
            if str(collection.get("name", "")).strip() == name:
                return collection
        raise ApiError(f"Collection '{name}' was not found or is not accessible.")

    def create_folder(self, parent_type: str, parent_id: str, name: str) -> dict[str, Any]:
        if self.dry_run:
            return {"_id": f"<dry-run-folder:{parent_type}:{name}>", "name": name}
        response = self.request(
            "POST",
            "/folder",
            query={
                "parentType": parent_type,
                "parentId": parent_id,
                "name": name,
                "reuseExisting": "true",
            },
        )
        require(isinstance(response, dict), "Folder creation returned unexpected data.")
        return response

    def set_metadata(self, resource_type: str, resource_id: str, metadata: dict[str, Any]) -> dict[str, Any] | None:
        if self.dry_run:
            return None
        return self.request(
            "PUT",
            f"/{resource_type}/{resource_id}/metadata",
            json_body=metadata,
        )

    def create_item(self, folder_id: str, name: str) -> dict[str, Any]:
        if self.dry_run:
            return {"_id": f"<dry-run-item:{name}>", "name": name}
        response = self.request(
            "POST",
            "/item",
            query={
                "folderId": folder_id,
                "name": name,
                "reuseExisting": "true",
            },
        )
        require(isinstance(response, dict), "Item creation returned unexpected data.")
        return response

    def upload_file(self, item_id: str, file_path: Path) -> dict[str, Any]:
        mime_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        size = file_path.stat().st_size
        if self.dry_run:
            return {
                "_id": f"<dry-run-upload:{file_path.name}>",
                "itemId": item_id,
                "name": file_path.name,
                "size": size,
                "mimeType": mime_type,
            }

        upload = self.request(
            "POST",
            "/file",
            query={
                "parentType": "item",
                "parentId": item_id,
                "name": file_path.name,
                "size": size,
                "mimeType": mime_type,
            },
        )
        require(isinstance(upload, dict) and upload.get("_id"), "Upload creation did not return an upload id.")

        upload_id = str(upload["_id"])
        offset = 0
        with file_path.open("rb") as handle:
            while True:
                chunk = handle.read(DEFAULT_CHUNK_SIZE)
                if not chunk:
                    break
                self.request(
                    "POST",
                    "/file/chunk",
                    query={"uploadId": upload_id, "offset": offset},
                    data=chunk,
                    headers={"Content-Type": "application/octet-stream"},
                )
                offset += len(chunk)

        file_doc = self.request("GET", f"/item/{item_id}/files")
        require(isinstance(file_doc, list) and file_doc, f"No files were found on uploaded item {item_id}.")
        return file_doc[0]


def build_target_folder_path(stack_id: str, target_leaf: str | None) -> tuple[str, ...]:
    leaf = target_leaf or stack_id
    return (*DEFAULT_TARGET_PATH, leaf)


def _upsert_local_identifier(
    alternate_identifiers: list[dict[str, Any]],
    local_identifier: str,
) -> list[dict[str, Any]]:
    filtered = [
        entry
        for entry in alternate_identifiers
        if not (
            str(entry.get("alternateIdentifierType", "")).strip().lower() == "local"
            and str(entry.get("alternateIdentifier", "")).strip() == local_identifier
        )
    ]
    filtered.append(
        {
            "alternateIdentifier": local_identifier,
            "alternateIdentifierType": "local",
        }
    )
    return filtered


def patch_child_local_identifier(
    client: HtmdecClient,
    child_deposition: dict[str, Any],
    stack_id: str,
) -> None:
    local_identifier = f"stack_{stack_id}"
    child_id = str(child_deposition.get("_id", "")).strip()

    if client.dry_run:
        return

    response = client.request("GET", f"/deposition/{child_id}")
    require(isinstance(response, dict), "Child deposition lookup returned unexpected data.")

    child_metadata = dict((response.get("metadata") or {}))
    child_alternate = list(child_metadata.get("alternateIdentifiers") or [])
    child_metadata["alternateIdentifiers"] = _upsert_local_identifier(
        child_alternate,
        local_identifier,
    )
    client.set_metadata("deposition", child_id, child_metadata)


def ensure_folder_tree(client: HtmdecClient, collection_name: str, path_parts: tuple[str, ...]) -> dict[str, Any]:
    collection = client.find_collection(collection_name)
    current_type = "collection"
    current_id = str(collection["_id"])
    current_folder: dict[str, Any] = collection

    for part in path_parts:
        current_folder = client.create_folder(current_type, current_id, part)
        current_type = "folder"
        current_id = str(current_folder["_id"])

    return current_folder


def artifact_type_for(path: Path) -> str:
    name = path.name.lower()
    if name.endswith("-layout.lbrn2"):
        return "layout"
    if name.endswith("-inventory.csv"):
        return "inventory"
    if name.endswith("-metadata.json"):
        return "metadata"
    return "artifact"


def upload_bundle_files(
    client: HtmdecClient,
    target_folder_id: str,
    bundle: OutputBundle,
    child_deposition: dict[str, Any],
) -> list[dict[str, Any]]:
    files = [bundle.layout_path, bundle.inventory_path, bundle.metadata_path]
    uploaded: list[dict[str, Any]] = []

    child_igsn = str(child_deposition.get("igsn", "")).strip()
    child_id = str(child_deposition.get("_id", "")).strip()
    foil_igsn = str((bundle.metadata.get("material", {}) or {}).get("igsn", "")).strip()
    stack_id = str(bundle.metadata.get("ID", "")).strip()

    for file_path in files:
        item = client.create_item(target_folder_id, file_path.name)
        file_doc = client.upload_file(str(item["_id"]), file_path)
        item_metadata = {
            "igsn": child_igsn,
            "stackIgsn": child_igsn,
            "stackDepositionId": child_id,
            "foilIgsn": foil_igsn,
            "stackId": stack_id,
            "artifactType": artifact_type_for(file_path),
        }
        client.set_metadata("item", str(item["_id"]), item_metadata)
        uploaded.append(
            {
                "path": str(file_path),
                "itemId": str(item["_id"]),
                "fileId": str(file_doc.get("_id", "")),
                "artifactType": item_metadata["artifactType"],
            }
        )

    return uploaded


def build_receipt(
    bundle: OutputBundle,
    parent_deposition_id: str,
    child_deposition: dict[str, Any],
    target_folder: dict[str, Any],
    uploaded_files: list[dict[str, Any]],
    dry_run: bool,
) -> dict[str, Any]:
    return {
        "dry_run": dry_run,
        "stack_id": str(bundle.metadata.get("ID", "")).strip(),
        "stack_local_identifier": f"stack_{str(bundle.metadata.get('ID', '')).strip()}",
        "foil_igsn": str((bundle.metadata.get("material", {}) or {}).get("igsn", "")).strip(),
        "parent_deposition_id": parent_deposition_id,
        "stack_igsn": str(child_deposition.get("igsn", "")).strip(),
        "stack_deposition_id": str(child_deposition.get("_id", "")).strip(),
        "target_folder_id": str(target_folder.get("_id", "")).strip(),
        "target_folder_name": str(target_folder.get("name", "")).strip(),
        "uploaded_files": uploaded_files,
    }


def main() -> int:
    args = parse_args()
    bundle = build_bundle(args.path)
    repo_root = find_repo_root(bundle.directory)
    parent_deposition_id = resolve_parent_deposition_id(bundle, repo_root, args.parent_deposition_id)
    stack_id = str(bundle.metadata.get("ID", "")).strip()

    print(f"Bundle directory: {bundle.directory}")
    print(f"Parent foil IGSN: {(bundle.metadata.get('material', {}) or {}).get('igsn', '')}")
    print(f"Stack ID: {stack_id}")
    print(f"Parent deposition id: {parent_deposition_id}")
    if args.dry_run:
        print("Dry run enabled; no remote changes will be made.")

    client = HtmdecClient(args.api_url, args.api_key, dry_run=args.dry_run)
    child_deposition = client.split_deposition(parent_deposition_id, stack_id)
    patch_child_local_identifier(client, child_deposition, stack_id)

    target_path = build_target_folder_path(stack_id, args.target_leaf)
    target_folder = ensure_folder_tree(client, args.collection_name, target_path)
    folder_metadata = {
        "igsn": str(child_deposition.get("igsn", "")).strip(),
        "stackIgsn": str(child_deposition.get("igsn", "")).strip(),
        "stackDepositionId": str(child_deposition.get("_id", "")).strip(),
        "foilIgsn": str((bundle.metadata.get("material", {}) or {}).get("igsn", "")).strip(),
        "stackId": stack_id,
    }
    client.set_metadata("folder", str(target_folder.get("_id", "")), folder_metadata)

    uploaded_files = upload_bundle_files(client, str(target_folder["_id"]), bundle, child_deposition)

    receipt = build_receipt(
        bundle,
        parent_deposition_id,
        child_deposition,
        target_folder,
        uploaded_files,
        args.dry_run,
    )
    receipt_path = bundle.directory / bundle.metadata_path.name.replace("-metadata.json", "-registration.json")
    write_json(receipt_path, receipt)

    print(f"Stack IGSN: {receipt['stack_igsn']}")
    print(f"Stack deposition id: {receipt['stack_deposition_id']}")
    print(f"Target folder id: {receipt['target_folder_id']}")
    print(f"Registration receipt: {receipt_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ApiError, OSError, ValueError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
