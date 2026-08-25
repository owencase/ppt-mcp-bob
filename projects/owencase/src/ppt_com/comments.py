"""Slide comment tools for PowerPoint COM automation.

Handles adding, listing, and deleting comments on slides.
"""


from pydantic import BaseModel, Field, ConfigDict

from utils.com_wrapper import ppt
from utils.tool_result import execute_json
from utils.com_objects import get_slide



# ---------------------------------------------------------------------------
# Pydantic input models
# ---------------------------------------------------------------------------
class AddCommentInput(BaseModel):
    """Input for adding a comment to a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    text: str = Field(..., description="Comment text")
    author: str = Field(
        default="AI Agent", description="Comment author name"
    )
    author_initials: str = Field(
        default="AI", description="Comment author initials"
    )
    left: float = Field(
        default=0, description="Horizontal position in points"
    )
    top: float = Field(
        default=0, description="Vertical position in points"
    )


class ListCommentsInput(BaseModel):
    """Input for listing comments on a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")


class DeleteCommentInput(BaseModel):
    """Input for deleting a comment from a slide."""
    model_config = ConfigDict(str_strip_whitespace=True)

    slide_index: int = Field(..., ge=1, description="1-based slide index")
    comment_index: int = Field(
        ..., ge=1, description="1-based comment index"
    )


# ---------------------------------------------------------------------------
# COM implementation functions
# ---------------------------------------------------------------------------
def _add_comment_impl(slide_index, text, author, author_initials,
                       left, top) -> dict:
    app, pres, slide = get_slide(slide_index)

    # Try Add2 first (newer PowerPoint versions), fall back to Add
    try:
        slide.Comments.Add2(left, top, author, author_initials, text, "AD", "")
    except Exception:
        try:
            slide.Comments.Add(left, top, author, author_initials, text)
        except Exception as e:
            raise RuntimeError(f"Failed to add comment: {e}")

    return {
        "success": True,
        "text": text,
        "author": author,
    }


def _list_comments_impl(slide_index) -> dict:
    app = ppt._get_app_impl()
    pres = ppt._get_pres_impl()
    slide = pres.Slides(slide_index)

    comments_col = slide.Comments
    count = comments_col.Count
    comments = []

    for i in range(1, count + 1):
        comment = comments_col(i)
        try:
            dt_str = str(comment.DateTime)
        except Exception:
            dt_str = ""
        comments.append({
            "index": i,
            "author": comment.Author,
            "author_initials": comment.AuthorInitials,
            "text": comment.Text,
            "datetime": dt_str,
            "left": round(comment.Left, 2),
            "top": round(comment.Top, 2),
        })

    return {
        "slide_index": slide_index,
        "comments_count": count,
        "comments": comments,
    }


def _delete_comment_impl(slide_index, comment_index) -> dict:
    app, pres, slide = get_slide(slide_index)

    comments_col = slide.Comments
    if comment_index < 1 or comment_index > comments_col.Count:
        raise ValueError(
            f"Comment index {comment_index} out of range "
            f"(1-{comments_col.Count})"
        )

    comments_col(comment_index).Delete()

    return {
        "success": True,
    }


# ---------------------------------------------------------------------------
# MCP tool functions
# ---------------------------------------------------------------------------
def add_comment(params: AddCommentInput) -> str:
    """Add a comment to a slide."""
    return execute_json(None,
            _add_comment_impl,
            params.slide_index, params.text, params.author,
            params.author_initials, params.left, params.top,
        )


def list_comments(params: ListCommentsInput) -> str:
    """List all comments on a slide."""
    return execute_json(None,
            _list_comments_impl,
            params.slide_index,
        )


def delete_comment(params: DeleteCommentInput) -> str:
    """Delete a comment from a slide."""
    return execute_json(None,
            _delete_comment_impl,
            params.slide_index, params.comment_index,
        )
