"""Document property operations for PowerPoint COM automation.

Handles getting and setting built-in document properties such as
title, author, subject, keywords, comments, category, and company.
"""

from typing import Optional

from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json


READABLE_PROPERTIES = [
    "Title", "Author", "Subject", "Keywords",
    "Comments", "Category", "Company",
    "Last Author", "Creation Date", "Last Save Time",
]


# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class SetPropertiesInput(BaseModel):
    """Input for setting document properties."""
    model_config = ConfigDict(str_strip_whitespace=True)

    title: Optional[str] = Field(default=None, description="Document title")
    author: Optional[str] = Field(default=None, description="Author name")
    subject: Optional[str] = Field(default=None, description="Document subject")
    keywords: Optional[str] = Field(default=None, description="Keywords (comma-separated)")
    comments: Optional[str] = Field(default=None, description="Document comments")
    category: Optional[str] = Field(default=None, description="Document category")
    company: Optional[str] = Field(default=None, description="Company name")


# ---------------------------------------------------------------------------
# COM implementation functions (run on COM thread via ppt.execute)
# ---------------------------------------------------------------------------
def _set_properties_impl(title, author, subject, keywords, comments, category, company):
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    props = pres.BuiltInDocumentProperties

    # Map field names to COM property names
    field_map = {
        "Title": title,
        "Author": author,
        "Subject": subject,
        "Keywords": keywords,
        "Comments": comments,
        "Category": category,
        "Company": company,
    }

    properties_set = 0
    set_names = []
    for prop_name, value in field_map.items():
        if value is not None:
            props(prop_name).Value = value
            properties_set += 1
            set_names.append(prop_name)

    return {
        "success": True,
        "properties_set": properties_set,
        "set_names": set_names,
    }


def _get_properties_impl():
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    props = pres.BuiltInDocumentProperties

    result = {}
    for prop_name in READABLE_PROPERTIES:
        try:
            value = props(prop_name).Value
            # Convert COM date objects to string
            if hasattr(value, "strftime"):
                value = value.strftime("%Y-%m-%d %H:%M:%S")
            elif not isinstance(value, (str, int, float, bool)):
                value = str(value)
            result[prop_name] = value
        except Exception:
            result[prop_name] = None

    return {
        "success": True,
        "properties": result,
    }


# ---------------------------------------------------------------------------
# MCP tool functions (async wrappers that delegate to COM thread)
# ---------------------------------------------------------------------------
def set_properties(params: SetPropertiesInput) -> str:
    """Set built-in document properties.

    Args:
        params: Properties to set. Only provided (non-None) values are updated.

    Returns:
        JSON with count of properties set and their names.
    """
    return execute_json('Failed to set properties: ',
            _set_properties_impl,
            params.title, params.author, params.subject,
            params.keywords, params.comments, params.category,
            params.company,
        )


def get_properties() -> str:
    """Get built-in document properties.

    Returns:
        JSON with all readable document properties.
    """
    return execute_json('Failed to get properties: ', _get_properties_impl)
