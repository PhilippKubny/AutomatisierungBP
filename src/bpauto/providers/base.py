"""Provider base interfaces and shared types."""

from typing import Protocol, TypedDict


class CompanyRecord(TypedDict, total=False):
    legal_name: str
    register_type: str
    register_no: str
    street: str
    zip: str
    city: str
    country: str
    pdf_path: str | None
    source: str
    notes: str


class Provider(Protocol):
    """Protocol defining a data provider."""

    def fetch(
        self,
        name: str,
        zip_code: str | None = None,
        country: str | None = None,
    ) -> CompanyRecord:
        """Retrieve a company record for the given identifiers."""

        ...
