"""Provider base interfaces and shared types."""

from typing import Optional, Protocol, TypedDict


class CompanyRecord(TypedDict, total=False):
    legal_name: str
    register_type: str
    register_no: str
    street: str
    zip: str
    city: str
    country: str
    pdf_path: Optional[str]
    source: str
    notes: str


class Provider(Protocol):
    """Protocol defining a data provider."""

    def fetch(
        self,
        name: str,
        zip_code: Optional[str] = None,
        country: Optional[str] = None,
    ) -> CompanyRecord:
        """Retrieve a company record for the given identifiers."""

        ...
