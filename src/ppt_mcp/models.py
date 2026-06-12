"""Content models (DESIGN.md §4.3): the agent describes *what a slide says*;
the writer decides *where it goes* by mapping onto layout placeholders."""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field, model_validator


class Paragraph(BaseModel):
    text: str
    level: Annotated[int, Field(ge=0, le=8, description="Bullet indent level, 0 = top")] = 0


class TableSpec(BaseModel):
    headers: list[str] | None = None
    rows: list[list[str]]

    @model_validator(mode="after")
    def _consistent_columns(self) -> "TableSpec":
        width = len(self.headers) if self.headers else (len(self.rows[0]) if self.rows else 0)
        if width == 0:
            raise ValueError("table needs headers or at least one row")
        for i, row in enumerate(self.rows):
            if len(row) != width:
                raise ValueError(f"row {i} has {len(row)} cells, expected {width}")
        return self


class ChartSeries(BaseModel):
    name: str
    values: list[float]


class ChartSpec(BaseModel):
    chart_type: Literal["column", "bar", "line", "pie", "doughnut", "area"] = "column"
    categories: list[str]
    series: list[ChartSeries]

    @model_validator(mode="after")
    def _series_match_categories(self) -> "ChartSpec":
        for s in self.series:
            if len(s.values) != len(self.categories):
                raise ValueError(
                    f"series '{s.name}' has {len(s.values)} values for "
                    f"{len(self.categories)} categories"
                )
        return self


class ImageRef(BaseModel):
    path: Annotated[str, Field(description="Absolute path to a local image file")]
    alt_text: str | None = None


class ContentSpec(BaseModel):
    """Semantic slide content. Sections that cannot be mapped to a placeholder
    of the chosen layout are reported as unplaced_content — never silently
    dropped, never free-floated onto the slide."""

    title: str | None = None
    subtitle: str | None = None
    body: list[Paragraph] | None = None
    table: TableSpec | None = None
    image: ImageRef | None = None
    chart: ChartSpec | None = None
    notes: str | None = None


class ShapeContent(BaseModel):
    """Content for a single placeholder. Exactly one field must be set."""

    text: str | None = None
    paragraphs: list[Paragraph] | None = None
    table: TableSpec | None = None
    image: ImageRef | None = None
    chart: ChartSpec | None = None

    @model_validator(mode="after")
    def _exactly_one(self) -> "ShapeContent":
        set_fields = [k for k, v in self.__dict__.items() if v is not None]
        if len(set_fields) != 1:
            raise ValueError(
                f"exactly one of text/paragraphs/table/image/chart must be set, got: "
                f"{set_fields or 'none'}"
            )
        return self


# -- text edit operations -------------------------------------------------------


class ReplaceTextOp(BaseModel):
    op: Literal["replace_text"]
    find: str
    replace: str
    match_case: bool = False


class SetParagraphsOp(BaseModel):
    op: Literal["set_paragraphs"]
    paragraphs: list[Paragraph]


class AppendParagraphOp(BaseModel):
    op: Literal["append_paragraph"]
    text: str
    level: Annotated[int, Field(ge=0, le=8)] = 0


EditOp = Annotated[
    Union[ReplaceTextOp, SetParagraphsOp, AppendParagraphOp],
    Field(discriminator="op"),
]


class Position(BaseModel):
    """Explicit geometry in inches — only for gated freeform placement."""

    left: float
    top: float
    width: float
    height: float


class ImagePlaceholderSpec(BaseModel):
    """A governed image slot: intent + prompt (composed from a style profile
    or supplied directly), targeting a picture placeholder or a labeled
    freeform box (DESIGN.md §4.6)."""

    image_intent: Annotated[str, Field(description="What the image should show, e.g. 'vision scene'")]
    profile_id: Annotated[
        str | None, Field(description="Style profile to compose the prompt from")
    ] = None
    prompt: Annotated[
        str | None, Field(description="Explicit prompt (validated against the profile if one is set)")
    ] = None
    negative_prompt: str | None = None
    aspect_ratio: str | None = None
    alt_text: str | None = None
    shape_ref: Annotated[
        str | None, Field(description="Target picture placeholder (preferred)")
    ] = None
    position: Position | None = None
    allow_freeform: bool = False
