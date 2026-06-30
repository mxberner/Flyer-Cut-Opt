"""
Helpers for validating and generating stack IDs.
"""

from __future__ import annotations

CROCKFORD_ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"
CROCKFORD_WIDTH = 5
AUTO_STACK_ID_TOKENS = {"", "AUTO"}


def normalize_crockford(value: str) -> str:
    """Return a canonical uppercase Crockford string."""
    return str(value or "").strip().upper()


def is_auto_stack_id(value: str) -> bool:
    """Return True when the stack ID should be assigned automatically."""
    return normalize_crockford(value) in AUTO_STACK_ID_TOKENS


def is_legacy_stack_id(value: str) -> bool:
    """Legacy IDs use the form F###."""
    normalized = normalize_crockford(value)
    return len(normalized) == 4 and normalized.startswith("F") and normalized[1:].isdigit()


def is_crockford_id(value: str, width: int = CROCKFORD_WIDTH) -> bool:
    """Return True when the value is a canonical fixed-width Crockford ID."""
    normalized = normalize_crockford(value)
    return (
        len(normalized) == width
        and all(char in CROCKFORD_ALPHABET for char in normalized)
    )


def is_valid_stack_id(value: str) -> bool:
    """Support legacy F### IDs and new fixed-width Crockford IDs."""
    return is_legacy_stack_id(value) or is_crockford_id(value)


def int_to_crockford(value: int, width: int = CROCKFORD_WIDTH) -> str:
    """Encode a non-negative integer as a fixed-width Crockford ID."""
    if value < 0:
        raise ValueError("Stack ID integer must be non-negative.")

    if value == 0:
        encoded = CROCKFORD_ALPHABET[0]
    else:
        encoded_chars: list[str] = []
        base = len(CROCKFORD_ALPHABET)
        current = value
        while current > 0:
            current, remainder = divmod(current, base)
            encoded_chars.append(CROCKFORD_ALPHABET[remainder])
        encoded = "".join(reversed(encoded_chars))

    if len(encoded) > width:
        raise ValueError(f"Encoded value '{encoded}' exceeds width {width}.")

    return encoded.rjust(width, CROCKFORD_ALPHABET[0])


def crockford_to_int(value: str) -> int:
    """Decode a canonical Crockford ID into an integer."""
    normalized = normalize_crockford(value)
    if not normalized:
        raise ValueError("Stack ID was empty.")
    if not all(char in CROCKFORD_ALPHABET for char in normalized):
        raise ValueError(f"Stack ID '{value}' contains unsupported characters.")

    total = 0
    base = len(CROCKFORD_ALPHABET)
    for char in normalized:
        total = (total * base) + CROCKFORD_ALPHABET.index(char)
    return total
