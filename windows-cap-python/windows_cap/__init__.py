"""Fastest Windows Screen Capture Library For Python 🔥."""

from typing import List, Optional

import cv2
import numpy

from .windows_capture import (
    NativeDxgiDuplication,
    NativeDxgiDuplicationFrame,
)


class DxgiDuplicationFrame:
    """Represents a CPU-readable DXGI desktop duplication frame."""

    __slots__ = ("_native", "_numpy_cache")

    def __init__(self, native_frame: NativeDxgiDuplicationFrame) -> None:
        self._native = native_frame
        self._numpy_cache: Optional[numpy.ndarray] = None

    @property
    def width(self) -> int:
        return int(self._native.width)

    @property
    def height(self) -> int:
        return int(self._native.height)

    @property
    def color_format(self) -> str:
        return str(self._native.color_format)

    @property
    def bytes_per_pixel(self) -> int:
        return int(self._native.bytes_per_pixel)

    @property
    def bytes_per_row(self) -> int:
        return int(self._native.bytes_per_row)

    def _raw_buffer(self) -> numpy.ndarray:
        memory_bytes = self._native.to_bytes()
        raw = numpy.frombuffer(memory_bytes, dtype=numpy.uint8)
        return raw.reshape(self.height, self.bytes_per_row)

    def to_numpy(self, *, copy: bool = False) -> numpy.ndarray:
        """Returns the frame as a ``numpy.ndarray`` with shape ``(height, width, 4)``.

        The channel order matches the underlying capture format (BGRA or RGBA).
        For ``rgba16f`` frames the returned dtype is ``numpy.float16``; otherwise
        ``numpy.uint8`` is used.
        """

        if self._numpy_cache is not None and not copy:
            return self._numpy_cache

        raw = self._raw_buffer()[:, : self.width * self.bytes_per_pixel]

        if self.color_format == "rgba16f":
            frame = raw.view(numpy.float16).reshape((self.height, self.width, 4))
        else:
            frame = raw.reshape((self.height, self.width, 4))

        if copy:
            return frame.copy()

        self._numpy_cache = frame
        return frame

    def to_bgr(self, *, copy: bool = True) -> numpy.ndarray:
        """Returns the frame converted to BGR ``numpy.uint8`` format."""

        image = self.to_numpy(copy=copy)

        if self.color_format == "bgra8":
            return image[..., :3].copy() if copy else cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)

        if self.color_format == "rgba8":
            return cv2.cvtColor(image, cv2.COLOR_RGBA2BGR) if not copy else image[..., [2, 1, 0]].copy()

        # rgba16f -> convert to 0..255 range before casting
        normalized = numpy.clip(image.astype(numpy.float32), 0.0, 1.0)
        return (normalized[..., [2, 1, 0]] * 255.0).astype(numpy.uint8)

    def to_rgb(self, *, copy: bool = True) -> numpy.ndarray:
        """Returns the frame converted to BGR ``numpy.uint8`` format."""

        image = self.to_numpy(copy=copy)

        if self.color_format == "rgba8":
            return image[..., :3].copy() if copy else cv2.cvtColor(image, cv2.COLOR_BGRA2RGB)

        if self.color_format == "bgra8":
            return cv2.cvtColor(image, cv2.COLOR_RGBA2RGB) if not copy else image[..., [2, 1, 0]].copy()

        # rgba16f -> convert to 0..255 range before casting
        normalized = numpy.clip(image.astype(numpy.float32), 0.0, 1.0)
        return (normalized[..., [2, 1, 0]] * 255.0).astype(numpy.uint8)

    def save_as_image(self, path: str) -> None:
        """Saves the frame to disk using OpenCV."""

        if self.color_format == "rgba16f":
            bgr = self.to_bgr(copy=True)
        else:
            bgr = self.to_bgr(copy=False)

        cv2.imwrite(path, bgr)

    def to_bytes(self) -> bytes:
        """Returns a contiguous copy of the frame bytes."""

        return bytes(self._raw_buffer())


class DxgiDuplicationSession:
    """High-level helper for DXGI desktop duplication captures."""

    __slots__ = ("_native", "_monitor_index")

    def __init__(self, monitor_index: Optional[int] = None) -> None:
        self._native = NativeDxgiDuplication(monitor_index)
        self._monitor_index = monitor_index

    @property
    def monitor_index(self) -> Optional[int]:
        return self._monitor_index

    def acquire_frame(self, timeout_ms: int = 16, area: Optional[List[int]] = None) -> Optional[DxgiDuplicationFrame]:
        native_frame = self._native.acquire_next_frame(timeout_ms, area)
        if native_frame is None:
            return None

        return DxgiDuplicationFrame(native_frame)

    def recreate(self) -> None:
        self._native.recreate()

    def switch_monitor(self, monitor_index: int) -> None:
        self._native.switch_monitor(monitor_index)
        self._monitor_index = monitor_index
