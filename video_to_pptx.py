"""Create a PowerPoint presentation from sampled video frames.

The video is processed only on this computer. Example:
    python video_to_pptx.py "mein-video.mp4" --fps 2

At 2 fps, the script creates one slide every 0.5 seconds.
At 10 fps, it creates one slide every 0.1 seconds.
"""

from __future__ import annotations

import argparse
import io
import math
import os
import re
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import cv2
from pptx.dml.color import RGBColor
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches


MAX_SLIDES = 1000000000
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def positive_fps(value: str) -> float:
    """Validate the frames-per-second argument."""
    try:
        fps = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("fps must be a number") from exc
    if not 0.1 <= fps <= 60:
        raise argparse.ArgumentTypeError("fps must be between 0.1 and 60")
    return fps


def positive_seconds(value: str) -> float:
    try:
        seconds = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("seconds must be a number") from exc
    if not 0.05 <= seconds <= 3600:
        raise argparse.ArgumentTypeError("seconds must be between 0.05 and 3600")
    return seconds


def format_duration(seconds: float) -> str:
    """Return a compact human-readable duration."""
    seconds = max(0.0, float(seconds))
    total_seconds = int(seconds)
    hours, remainder = divmod(total_seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    if minutes:
        return f"{minutes:02d}:{secs:02d}"
    return f"{secs}s"


def set_automatic_advance(pptx_path: Path, milliseconds: int) -> None:
    """Set a PowerPoint auto-advance transition on every slide.

    python-pptx does not expose transitions, so this safely updates the PPTX
    package XML after python-pptx has created the presentation.
    """
    transition = f'<p:transition advClick="1" advTm="{milliseconds}"><p:fade/></p:transition>'
    with zipfile.ZipFile(pptx_path, "r") as source:
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False, dir=pptx_path.parent) as temp_file:
            temporary_path = Path(temp_file.name)
        try:
            with zipfile.ZipFile(temporary_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    content = source.read(item.filename)
                    if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename):
                        xml = content.decode("utf-8")
                        if "</p:cSld>" not in xml:
                            raise RuntimeError(f"Cannot add transition to {item.filename}.")
                        content = xml.replace("</p:cSld>", f"</p:cSld>{transition}", 1).encode("utf-8")
                    target.writestr(item, content)
            os.replace(temporary_path, pptx_path)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise


def add_contained_image(slide, image_source, image_width: int, image_height: int) -> None:
    """Add an image without stretching or cropping it, centered on black."""
    # A black rectangle keeps letterboxing intentional for portrait or square video.
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(0, 0, 0)
    background.line.fill.background()

    image_ratio = image_width / image_height
    slide_ratio = SLIDE_WIDTH / SLIDE_HEIGHT
    if image_ratio > slide_ratio:
        width = SLIDE_WIDTH
        height = int(width / image_ratio)
        left = 0
        top = int((SLIDE_HEIGHT - height) / 2)
    else:
        height = SLIDE_HEIGHT
        width = int(height * image_ratio)
        left = int((SLIDE_WIDTH - width) / 2)
        top = 0
    slide.shapes.add_picture(image_source, left, top, width=width, height=height)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create one PowerPoint slide per sampled video frame. Processing stays local."
    )
    parser.add_argument("video", type=Path, help="Path to the input video, e.g. video.mp4")
    parser.add_argument("--fps", type=positive_fps, default=2.0, help="Sample rate from 0.1 to 60 (default: 2)")
    parser.add_argument("--output", type=Path, help="Output .pptx path (default: beside the video)")
    parser.add_argument("--max-slides", type=int, default=MAX_SLIDES, help=f"Safety limit (default: {MAX_SLIDES})")
    parser.add_argument("--auto-advance", action="store_true", help="Advance slides automatically during a PowerPoint slide show")
    parser.add_argument("--advance-seconds", type=positive_seconds, help="Seconds per slide; defaults to 1/fps and enables --auto-advance")
    return parser.parse_args()


def main() -> int:
    args = parse_arguments()
    video_path: Path = args.video.expanduser().resolve()
    if not video_path.is_file():
        print(f"Error: video not found: {video_path}", file=sys.stderr)
        return 2
    if args.max_slides < 1:
        print("Error: --max-slides must be at least 1.", file=sys.stderr)
        return 2
    if args.advance_seconds is not None:
        args.auto_advance = True

    capture = cv2.VideoCapture(str(video_path))
    if not capture.isOpened():
        print("Error: OpenCV cannot open this video. Try an MP4 with H.264 video.", file=sys.stderr)
        return 2

    source_fps = capture.get(cv2.CAP_PROP_FPS)
    total_frames = int(capture.get(cv2.CAP_PROP_FRAME_COUNT))
    width = int(capture.get(cv2.CAP_PROP_FRAME_WIDTH))
    height = int(capture.get(cv2.CAP_PROP_FRAME_HEIGHT))
    if source_fps <= 0 or total_frames <= 0 or width <= 0 or height <= 0:
        print("Error: the video contains no readable frames or duration metadata.", file=sys.stderr)
        capture.release()
        return 2

    duration = total_frames / source_fps
    # Frame count is more reliable than the container duration at the end of a
    # video. Some codecs report a duration that is a fraction longer than the
    # final decodable frame.
    slide_count = math.floor((total_frames - 1) * args.fps / source_fps) + 1
    print(f"Video duration: {duration:.2f} seconds")
    print(f"Sampling: {args.fps:g} fps, one slide every {1 / args.fps:.3f} seconds")
    print(f"Estimated slides: {slide_count}")
    if args.auto_advance:
        advance_seconds = args.advance_seconds if args.advance_seconds is not None else 1 / args.fps
        print(f"Auto-advance: every {advance_seconds:.3f} seconds")
    if args.max_slides is not None and slide_count > args.max_slides:
        print(
            f"Error: {slide_count} slides exceed the safety limit of {args.max_slides}. "
            "Lower --fps, use a shorter video, or explicitly raise --max-slides.",
            file=sys.stderr,
        )
        capture.release()
        return 2

    output_path = args.output or video_path.with_name(f"{video_path.stem}-folien.pptx")
    output_path = output_path.expanduser().resolve()
    if output_path.suffix.lower() != ".pptx":
        output_path = output_path.with_suffix(".pptx")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT
    presentation.core_properties.title = video_path.stem
    presentation.core_properties.subject = "Slides created from video frames"
    presentation.core_properties.author = "video_to_pptx.py"
    blank_layout = presentation.slide_layouts[6]

    try:
        frame_index = 0
        next_slide_index = 0
        start_time = time.perf_counter()
        while True:
            success, frame = capture.read()
            if not success:
                break

            while next_slide_index < slide_count:
                target_frame = min(round(next_slide_index * source_fps / args.fps), total_frames - 1)
                if target_frame != frame_index:
                    break

                success, encoded = cv2.imencode(".jpg", frame, [cv2.IMWRITE_JPEG_QUALITY, 90])
                if not success:
                    raise RuntimeError(f"Could not encode frame {next_slide_index + 1}.")

                slide = presentation.slides.add_slide(blank_layout)
                add_contained_image(slide, io.BytesIO(encoded.tobytes()), width, height)

                elapsed = time.perf_counter() - start_time
                done = next_slide_index + 1
                remaining = slide_count - done
                eta = (elapsed / done) * remaining if done else 0.0
                print(
                    f"\rPreparing slide {done}/{slide_count} ({done / slide_count:.0%}) | "
                    f"elapsed {format_duration(elapsed)} | ETA {format_duration(eta)}",
                    end="",
                    flush=True,
                )
                next_slide_index += 1

            frame_index += 1

        if next_slide_index < slide_count:
            print(
                f"\nWarning: video ended early after {next_slide_index} slides; expected {slide_count} based on metadata.",
                file=sys.stderr,
            )

        print("\nWriting PowerPoint file …", flush=True)
        presentation.save(str(output_path))
        if args.auto_advance:
            set_automatic_advance(output_path, round(advance_seconds * 1000))
    except (RuntimeError, OSError) as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1
    finally:
        capture.release()

    print(f"Done: {output_path}")
    if args.auto_advance:
        print("Open the file in PowerPoint and start Slide Show mode to use the automatic advance.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
