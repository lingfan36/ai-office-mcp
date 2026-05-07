"""Image and shape management tools for open Word documents via COM."""

import json
import sys

_PTS_PER_CM = 28.35


async def word_live_list_images(filename: str = None) -> str:
    """List all inline images in an open Word document.

    Returns:
        JSON array with index, size, alt text, and paragraph position for each image.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        images = []
        for i in range(1, doc.InlineShapes.Count + 1):
            s = doc.InlineShapes(i)
            para_idx = None
            try:
                r = s.Range
                for p in range(1, doc.Paragraphs.Count + 1):
                    pr = doc.Paragraphs(p).Range
                    if pr.Start <= r.Start <= pr.End:
                        para_idx = p
                        break
            except Exception:
                pass
            images.append({
                "index": i,
                "width_pt": s.Width,
                "height_pt": s.Height,
                "width_cm": round(s.Width / _PTS_PER_CM, 2),
                "height_cm": round(s.Height / _PTS_PER_CM, 2),
                "alt_text": s.AlternativeText or "",
                "paragraph_index": para_idx,
                "type": s.Type,
            })
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "count": len(images),
            "images": images,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_image(
    filename: str = None,
    image_index: int = 1,
) -> str:
    """Delete an inline image from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        image_index: 1-based index of the image (use word_live_list_images to find it).

    Returns:
        JSON confirming deletion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        total = doc.InlineShapes.Count

        if image_index < 1 or image_index > total:
            return json.dumps({
                "error": f"image_index {image_index} out of range (1-{total})"
            })

        with undo_record(app, "MCP: Delete Image"):
            doc.InlineShapes(image_index).Delete()

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "deleted_index": image_index,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_resize_image(
    filename: str = None,
    image_index: int = 1,
    width_cm: float = None,
    height_cm: float = None,
    lock_aspect_ratio: bool = True,
) -> str:
    """Resize an inline image in an open Word document.

    Provide width_cm, height_cm, or both. When lock_aspect_ratio=True and only
    one dimension is given, the other scales proportionally.

    Args:
        filename: Document name or path (None = active document).
        image_index: 1-based image index (use word_live_list_images).
        width_cm: New width in centimetres.
        height_cm: New height in centimetres.
        lock_aspect_ratio: Maintain aspect ratio when only one dimension is set (default True).

    Returns:
        JSON with new dimensions.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Image tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        total = doc.InlineShapes.Count

        if image_index < 1 or image_index > total:
            return json.dumps({"error": f"image_index {image_index} out of range (1-{total})"})

        shape = doc.InlineShapes(image_index)

        with undo_record(app, "MCP: Resize Image"):
            orig_w = shape.Width
            orig_h = shape.Height

            if width_cm and height_cm:
                shape.Width = width_cm * _PTS_PER_CM
                shape.Height = height_cm * _PTS_PER_CM
            elif width_cm:
                new_w = width_cm * _PTS_PER_CM
                if lock_aspect_ratio and orig_w > 0:
                    shape.Height = orig_h * (new_w / orig_w)
                shape.Width = new_w
            elif height_cm:
                new_h = height_cm * _PTS_PER_CM
                if lock_aspect_ratio and orig_h > 0:
                    shape.Width = orig_w * (new_h / orig_h)
                shape.Height = new_h
            else:
                return json.dumps({"error": "Provide width_cm or height_cm"})

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "image_index": image_index,
            "width_cm": round(shape.Width / _PTS_PER_CM, 2),
            "height_cm": round(shape.Height / _PTS_PER_CM, 2),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_list_shapes(filename: str = None) -> str:
    """List all floating shapes (text boxes, drawings, etc.) in an open Word document.

    Returns:
        JSON array with index, name, type, size, position, and text content for each shape.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Shape tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document

        app = get_word_app()
        doc = find_document(app, filename)

        shapes = []
        for i in range(1, doc.Shapes.Count + 1):
            s = doc.Shapes(i)
            text_content = ""
            try:
                text_content = s.TextFrame.TextRange.Text[:200]
            except Exception:
                pass
            shapes.append({
                "index": i,
                "name": s.Name,
                "type": s.Type,
                "left_pt": s.Left,
                "top_pt": s.Top,
                "width_pt": s.Width,
                "height_pt": s.Height,
                "width_cm": round(s.Width / _PTS_PER_CM, 2),
                "height_cm": round(s.Height / _PTS_PER_CM, 2),
                "text_preview": text_content[:100] if text_content else "",
                "visible": s.Visible,
            })
        return json.dumps({
            "success": True,
            "document": doc.Name,
            "count": len(shapes),
            "shapes": shapes,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_delete_shape(
    filename: str = None,
    shape_index: int = None,
    shape_name: str = None,
) -> str:
    """Delete a floating shape from an open Word document.

    Provide shape_index or shape_name (use word_live_list_shapes to find them).

    Args:
        filename: Document name or path (None = active document).
        shape_index: 1-based shape index.
        shape_name: Exact shape name (alternative to index).

    Returns:
        JSON confirming deletion.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Shape tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if shape_name:
            shape = doc.Shapes(shape_name)
        elif shape_index is not None:
            total = doc.Shapes.Count
            if shape_index < 1 or shape_index > total:
                return json.dumps({"error": f"shape_index {shape_index} out of range (1-{total})"})
            shape = doc.Shapes(shape_index)
        else:
            return json.dumps({"error": "Provide shape_index or shape_name"})

        name = shape.Name
        with undo_record(app, "MCP: Delete Shape"):
            shape.Delete()

        return json.dumps({"success": True, "document": doc.Name, "deleted": name})
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_add_text_box(
    filename: str = None,
    text: str = "",
    left_cm: float = 2.0,
    top_cm: float = 2.0,
    width_cm: float = 8.0,
    height_cm: float = 3.0,
    font_name: str = None,
    font_size: float = None,
    bold: bool = None,
    italic: bool = None,
    font_color: str = None,
    border: bool = True,
    fill_color: str = None,
) -> str:
    """Add a floating text box to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Text content of the text box.
        left_cm: Horizontal position from left margin in centimetres.
        top_cm: Vertical position from top margin in centimetres.
        width_cm: Text box width in centimetres.
        height_cm: Text box height in centimetres.
        font_name: Font family for the text.
        font_size: Font size in points.
        bold / italic: Bold/italic formatting.
        font_color: Text color as "#RRGGBB".
        border: Show text box border (default True). False = no border.
        fill_color: Background fill color as "#RRGGBB" (None = transparent).

    Returns:
        JSON with shape index and name.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Shape tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        def hex_to_bgr(h):
            h = h.lstrip("#")
            return int(h[0:2], 16) + (int(h[2:4], 16) << 8) + (int(h[4:6], 16) << 16)

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Add Text Box"):
            shape = doc.Shapes.AddTextbox(
                Orientation=1,   # msoTextOrientationHorizontal
                Left=left_cm * _PTS_PER_CM,
                Top=top_cm * _PTS_PER_CM,
                Width=width_cm * _PTS_PER_CM,
                Height=height_cm * _PTS_PER_CM,
            )

            # Text and formatting
            tf = shape.TextFrame
            tf.TextRange.Text = text

            if font_name or font_size or bold is not None or italic is not None or font_color:
                font = tf.TextRange.Font
                if font_name:
                    font.Name = font_name
                if font_size:
                    font.Size = font_size
                if bold is not None:
                    font.Bold = bold
                if italic is not None:
                    font.Italic = italic
                if font_color:
                    font.Color = hex_to_bgr(font_color)

            # Border
            if not border:
                shape.Line.Visible = False  # msoFalse=0

            # Fill
            if fill_color:
                shape.Fill.ForeColor.RGB = hex_to_bgr(fill_color)
                shape.Fill.Visible = True
            else:
                try:
                    shape.Fill.Visible = False
                except Exception:
                    pass

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "shape_name": shape.Name,
            "shape_index": doc.Shapes.Count,
            "width_cm": width_cm,
            "height_cm": height_cm,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


async def word_live_move_shape(
    filename: str = None,
    shape_index: int = None,
    shape_name: str = None,
    left_cm: float = None,
    top_cm: float = None,
    width_cm: float = None,
    height_cm: float = None,
) -> str:
    """Move or resize a floating shape in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        shape_index / shape_name: Target shape identifier.
        left_cm: New horizontal position from left margin in centimetres.
        top_cm: New vertical position from top margin in centimetres.
        width_cm: New width in centimetres.
        height_cm: New height in centimetres.

    Returns:
        JSON with updated position and size.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Shape tools are Windows only"})
    try:
        from word_document_server.core.word_com import get_word_app, find_document, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if shape_name:
            shape = doc.Shapes(shape_name)
        elif shape_index is not None:
            shape = doc.Shapes(shape_index)
        else:
            return json.dumps({"error": "Provide shape_index or shape_name"})

        with undo_record(app, "MCP: Move Shape"):
            if left_cm is not None:
                shape.Left = left_cm * _PTS_PER_CM
            if top_cm is not None:
                shape.Top = top_cm * _PTS_PER_CM
            if width_cm is not None:
                shape.Width = width_cm * _PTS_PER_CM
            if height_cm is not None:
                shape.Height = height_cm * _PTS_PER_CM

        return json.dumps({
            "success": True,
            "document": doc.Name,
            "shape": shape.Name,
            "left_cm": round(shape.Left / _PTS_PER_CM, 2),
            "top_cm": round(shape.Top / _PTS_PER_CM, 2),
            "width_cm": round(shape.Width / _PTS_PER_CM, 2),
            "height_cm": round(shape.Height / _PTS_PER_CM, 2),
        })
    except Exception as e:
        return json.dumps({"error": str(e)})
