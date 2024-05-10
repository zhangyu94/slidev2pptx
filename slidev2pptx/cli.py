#!/usr/bin/env python

import click

from . import slidev2pptx


arg = click.argument
opt = click.option


@click.command()
@opt(
    "-i",
    "--input",
    "slidev_path",
    default="./",
    help="path to the Slidev repository (default: ./)",
)
@opt(
    "-o",
    "--output",
    "output_path",
    default="./slides-export.pptx",
    help="path to save the pptx (default: ./slides-export.pptx)",
)
@opt(
    "-s",
    "--scale",
    "scale",
    default=2,
    type=int,
    help="scale of Slidev image export (default: 2)",
)
def cli(slidev_path: str, output_path: str, scale: int):
    """Convert a Slidev slide deck to a PPTX slide deck."""

    slidev2pptx(slidev_path, output_path, scale)


if __name__ == "__main__":
    cli()
