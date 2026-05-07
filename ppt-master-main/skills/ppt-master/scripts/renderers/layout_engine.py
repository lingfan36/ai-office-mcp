"""
Dynamic vertical layout cursor — eliminates hardcoded Y-coordinate overlaps.

Usage
-----
    from layout_engine import LayoutFlow

    flow = LayoutFlow(start_y=100, end_y=680, gap=12)

    y, h = flow.next(measured_height)          # reserve, advance cursor
    y, h = flow.next(measured_height, gap=20)  # custom gap after this element
    flow.skip(8)                               # advance without allocating
    remaining = flow.remaining                 # pixels left

    # Divide remaining space evenly
    slots = flow.allocate_remaining(3)         # [(y0,h0), (y1,h1), (y2,h2)]
"""
from __future__ import annotations


class LayoutFlow:
    """
    Tracks the Y insertion cursor within a vertical region.

    Attributes
    ----------
    y         : int  — current cursor (next available pixel row)
    end_y     : int  — lower bound of the region
    remaining : int  — pixels between cursor and end_y
    """

    def __init__(self, start_y: int, end_y: int, gap: int = 12):
        self.y     = start_y
        self.end_y = end_y
        self._gap  = gap

    # ── Core ──────────────────────────────────────────────────────────────────

    def next(self, height: int, gap=None, min_h: int = 0):
        """
        Reserve *height* pixels, advance cursor, return ``(y, height)``.

        Parameters
        ----------
        height : measured / desired element height
        gap    : gap **after** this element (defaults to self._gap)
        min_h  : floor for height; useful for elements that may be empty

        If both height and min_h are 0, the cursor is unchanged (no-op).
        """
        h = max(height, min_h)
        if h == 0:
            return self.y, 0
        y = self.y
        g = gap if gap is not None else self._gap
        self.y += h + g
        return y, h

    def skip(self, pixels: int):
        """Advance cursor by *pixels* without allocating a named element."""
        self.y += pixels

    # ── Derived ───────────────────────────────────────────────────────────────

    @property
    def remaining(self) -> int:
        """Pixels between cursor and end_y (never negative)."""
        return max(0, self.end_y - self.y)

    def allocate_remaining(self, n: int, gap=None):
        """
        Divide remaining space into *n* equal vertical slots.

        Returns list of ``(y, h)`` tuples.
        """
        if n <= 0:
            return []
        g     = gap if gap is not None else self._gap
        total = max(0, self.remaining - g * max(n - 1, 0))
        h     = max(1, total // n)
        return [self.next(h, gap=g) for _ in range(n)]

    # ── Columns ───────────────────────────────────────────────────────────────

    def columns(self, x: int, total_w: int, n: int, col_gap: int = 20):
        """
        Create *n* independent horizontal LayoutFlow objects all starting at
        the current cursor, covering ``(x, total_w)`` with *col_gap* between.

        Returns list of ``(col_x, col_w, LayoutFlow)`` tuples.
        """
        col_w = (total_w - col_gap * (n - 1)) // n
        return [
            (x + i * (col_w + col_gap), col_w,
             LayoutFlow(self.y, self.end_y, self._gap))
            for i in range(n)
        ]
