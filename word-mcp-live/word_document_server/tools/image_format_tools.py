"""Image formatting tools — wrapping, cropping, brightness/contrast, and position."""

import json
import sys

_PTS_PER_CM = 28.35

# wdWrapType constants
WRAP_TYPE = {
    "inline":      7,   # wdWrapInline (convert to inline shape)
    "square":      0,   # wdWrapSquare
    "tight":       1,   # wdWrapTight
    "through":     2,   # wdWrapThrough
    "top_bottom":  3,   # wdWrapTopBottom
    "behind":      4,   # wdWrapBehind
    "in_front":    5,   # wdWrapFront
}

# wdWrapSideType constants
WRAP_SIDE = {
    "both":      3,   # wdWrapBoth
    "left":      1,   # wdWrapLeft
    "right":     2,   # wdWrapRight
    "largest":   4,   # wdWrapLargest
}


async def word_live_set_image_wrap(
    filename: str = None,
    shape_name: str = None,
    shape_index: int = None,
    wrap_type: str = "square",
    wrap_side: str = "both",
) -> str:
    """Set text wrapping style for a floating image/shape in an open Word document.

    Note: This only works on floating shapes (doc.Shapes), not inline shapes.
    Use word_live_convert_to_floating first if the image is inline.

    Args:
        filename: Document name or path (None = active document).
        shape_name: Shape name (from word_live_list_shapes).
        shape_index: 1-based shape index (alternative to shape_name).
        wrap_type: Wrapping style:
            "inline"     → embedded in text flow (use word_live_convert_to_inline instead)
            "square"     → text wraps around bounding box (default)
            "tight"      → text wraps tightly around image outline
            "through"    → text flows through transparent areas
            "top_bottom" → text above and below only
            "behind"     → image behind text (no wrap)
            "in_front"   → image in front of text (no wrap)
        wrap_side: Which side text wraps on — "both" (default), "left", "right", "largest".

    Returns:
        JSON confirming the new wrap setting.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        if wrap_type not in WRAP_TYPE:
            return json.dumps({
                "error": f"Invalid wrap_type '{wrap_type}'. Use: {list(WRAP_TYPE.keys())}"
            })

        app = get_word_app()
        doc = find_document(app, filename)

        if shape_name:
            shape = doc.Shapes(shape_name)
        elif shape_index is not None:
            shape = doc.Shapes(shape_index)
        else:
            return json.dumps({"error": "Provide shape_name or shape_index"})

        with undo_record(app, "MCP: Set Image Wrap"):
            wf = shape.WrapFormat
            wf.Type = WRAP_TYPE[wrap_type]
            if wrap_type not in ("inline", "behind", "in_front", "top_bottom"):
                wf.Side = WRAP_SIDE.get(wrap_side, 3)

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "shape": shape.Name,
            "wrap_type": wrap_type,
            "wrap_side": wrap_side,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_image_position(
    filename: str = None,
    shape_name: str = None,
    shape_index: int = None,
    horizontal_position_cm: float = None,
    vertical_position_cm: float = None,
    horizontal_relative_to: str = "page",
    vertical_relative_to: str = "page",
) -> str:
    """Set absolute position of a floating shape/image in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        shape_name / shape_index: Target shape identifier.
        horizontal_position_cm: Distance from the left edge of the reference in cm.
        vertical_position_cm: Distance from the top edge of the reference in cm.
        horizontal_relative_to: Reference for horizontal position:
            "page" (default), "margin", "column", "character".
        vertical_relative_to: Reference for vertical position:
            "page" (default), "margin", "paragraph", "line".

    Returns:
        JSON with new position.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        # wdRelativeHorizontalPosition
        H_REL = {"margin": 0, "page": 1, "column": 2, "character": 3}
        # wdRelativeVerticalPosition
        V_REL = {"margin": 0, "page": 1, "paragraph": 2, "line": 3}

        app = get_word_app()
        doc = find_document(app, filename)

        if shape_name:
            shape = doc.Shapes(shape_name)
        elif shape_index is not None:
            shape = doc.Shapes(shape_index)
        else:
            return json.dumps({"error": "Provide shape_name or shape_index"})

        with undo_record(app, "MCP: Set Image Position"):
            if horizontal_position_cm is not None:
                shape.RelativeHorizontalPosition = H_REL.get(horizontal_relative_to, 1)
                shape.Left = horizontal_position_cm * _PTS_PER_CM
            if vertical_position_cm is not None:
                shape.RelativeVerticalPosition = V_REL.get(vertical_relative_to, 1)
                shape.Top = vertical_position_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "shape": shape.Name,
            "left_cm": round(shape.Left / _PTS_PER_CM, 2),
            "top_cm": round(shape.Top / _PTS_PER_CM, 2),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_crop_image(
    filename: str = None,
    image_index: int = 1,
    top_cm: float = 0.0,
    bottom_cm: float = 0.0,
    left_cm: float = 0.0,
    right_cm: float = 0.0,
) -> str:
    """Crop an inline image in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        image_index: 1-based inline image index (use word_live_list_images).
        top_cm: Crop from top in centimetres (positive = crop in).
        bottom_cm: Crop from bottom in centimetres.
        left_cm: Crop from left in centimetres.
        right_cm: Crop from right in centimetres.

    Returns:
        JSON with new visible dimensions after cropping.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.InlineShapes.Count
        if image_index < 1 or image_index > total:
            return json.dumps({"error": f"image_index {image_index} out of range (1-{total})"})

        shape = doc.InlineShapes(image_index)

        with undo_record(app, "MCP: Crop Image"):
            # PictureFormat.CropTop etc. are in points
            pf = shape.PictureFormat
            pf.CropTop = top_cm * _PTS_PER_CM
            pf.CropBottom = bottom_cm * _PTS_PER_CM
            pf.CropLeft = left_cm * _PTS_PER_CM
            pf.CropRight = right_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image_index": image_index,
            "width_cm": round(shape.Width / _PTS_PER_CM, 2),
            "height_cm": round(shape.Height / _PTS_PER_CM, 2),
            "crop": {
                "top_cm": top_cm, "bottom_cm": bottom_cm,
                "left_cm": left_cm, "right_cm": right_cm,
            },
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_set_image_brightness_contrast(
    filename: str = None,
    image_index: int = 1,
    brightness: float = 0.0,
    contrast: float = 0.0,
) -> str:
    """Adjust brightness and contrast of an inline image in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        image_index: 1-based inline image index.
        brightness: Brightness adjustment from -1.0 (darkest) to 1.0 (brightest). 0 = original.
        contrast: Contrast adjustment from -1.0 (least) to 1.0 (most). 0 = original.

    Returns:
        JSON confirming adjustments.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.InlineShapes.Count
        if image_index < 1 or image_index > total:
            return json.dumps({"error": f"image_index {image_index} out of range (1-{total})"})

        brightness = max(-1.0, min(1.0, brightness))
        contrast = max(-1.0, min(1.0, contrast))

        shape = doc.InlineShapes(image_index)
        with undo_record(app, "MCP: Set Image Brightness/Contrast"):
            pf = shape.PictureFormat
            pf.Brightness = brightness
            pf.Contrast = contrast

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image_index": image_index,
            "brightness": brightness,
            "contrast": contrast,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_convert_to_floating(
    filename: str = None,
    image_index: int = 1,
    wrap_type: str = "square",
    width_cm: float = None,
    height_cm: float = None,
) -> str:
    """Convert an inline image to a floating shape in an open Word document.

    After conversion, use word_live_set_image_wrap and word_live_set_image_position
    to control exact positioning.

    Args:
        filename: Document name or path (None = active document).
        image_index: 1-based inline image index (use word_live_list_images).
        wrap_type: Initial wrapping style ("square", "tight", "behind", "in_front", etc.).
        width_cm / height_cm: Optional size override (default = keep current size).

    Returns:
        JSON with new shape name and index.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image format tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        total = doc.InlineShapes.Count
        if image_index < 1 or image_index > total:
            return json.dumps({"error": f"image_index {image_index} out of range (1-{total})"})

        inline = doc.InlineShapes(image_index)
        orig_w = inline.Width
        orig_h = inline.Height

        with undo_record(app, "MCP: Convert Image to Floating"):
            shape = inline.ConvertToShape()
            shape.WrapFormat.Type = WRAP_TYPE.get(wrap_type, 0)
            if width_cm:
                shape.Width = width_cm * _PTS_PER_CM
            if height_cm:
                shape.Height = height_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "shape_name": shape.Name,
            "wrap_type": wrap_type,
            "width_cm": round(shape.Width / _PTS_PER_CM, 2),
            "height_cm": round(shape.Height / _PTS_PER_CM, 2),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
