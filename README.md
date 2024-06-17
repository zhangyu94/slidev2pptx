<a href="https://pypi.org/project/slidev2pptx/">
    <img alt="Newest PyPI version" src="https://img.shields.io/pypi/v/slidev2pptx.svg">
</a>
<a href="https://github.com/psf/black">
    <img alt="Code style: black" src="https://img.shields.io/badge/code%20style-black-000000.svg">
</a>
<a href="http://commitizen.github.io/cz-cli/">
    <img alt="Commitizen friendly" src="https://img.shields.io/badge/commitizen-friendly-brightgreen.svg">
</a>

# slidev2pptx

A Python package for exporting [Slidev](https://github.com/slidevjs/slidev) slides to PPTX.

> [!NOTE]
> This project is archived as Slidev has now officially supported exporting to PPTX ([reference](https://github.com/slidevjs/slidev/pull/1603)).

## Installation

```sh
pip install slidev2pptx
```

## Usage Example

Using `slidev2pptx` as a command line tool ([details](#command-line-tool)):

```sh
slidev2pptx -i ./slidev/ -o ./output.pptx
```

Using `slidev2pptx` as a Python package ([details](#python-api)):

```python
from slidev2pptx import slidev2pptx
slidev2pptx(slidev_path="./slidev/", output_path="./output.pptx")
```

> [!IMPORTANT]
> Make sure you have set up your Slidev repository to support exporting to PNG (see [Slidev document](https://sli.dev/guide/exporting)) before using `slidev2pptx`.

## Command Line Tool

The `slidev2pptx` command line tool exports Slidev slides to PPTX with the following options:

| Option           | Description                           | Default                |
| ---------------- | ------------------------------------- | ---------------------- |
| `-i`, `--input`  | Path to the Slidev project directory. | `./`                   |
| `-o`, `--output` | Path to the output PPTX file.         | `./slides-export.pptx` |
| `-s`, `--scale`  | Scale of Slidev image export.         | `2`                    |

## Python API

### `slidev2pptx.slidev2pptx(slidev_path: str, output_path: str, scale: int = 2) -> None`

Convert a Slidev slide deck stored at `slidev_path` to a PPTX slide deck and store it at `output_path`.
The `scale` argument specifies the scale for the Slidev image export.

### `slidev2pptx.build_pptx(notes_path: str, image_directory: str, output_path: str) -> None`

Build a PPTX slide deck and store it at `output_path`, using the Slidev image export stored at `image_directory` and the presenter notes stored at `notes_path`.
`slidev2pptx.build_pptx` is a lower-level function used by `slidev2pptx.slidev2pptx`.

## Caveats

- **`slidev2pptx` only works for Slidev version >= 0.47.4** as it utilizes `--scale` argument for image export ([reference](https://github.com/slidevjs/slidev/releases/tag/v0.47.4)).
- **Animation and interactivity are not preserved** in the exported PPTX file.

## Motivation

Slidev does not readily support exporting slides to PPTX (as of version 0.48.9), but I find it necessary on several occasions:

- The presentation is required to be conducted on a coordinator's machine.
- A PPTX file of the slides is requested for archival purposes.
- Someone unfamiliar with Slidev needs to present my slides.

## Thanks

This package is inspired by [pdf2pptx](https://github.com/kevinmcguinness/pdf2pptx).
