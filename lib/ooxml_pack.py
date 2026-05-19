"""Thin wrapper around an OOXML file as a ZIP: enumerate parts, filter by prefix, read content."""
from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Iterator


class OOXMLPackage:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._zip: zipfile.ZipFile | None = None

    def __enter__(self) -> "OOXMLPackage":
        self._zip = zipfile.ZipFile(self.path, "r")
        return self

    def __exit__(self, *exc):
        if self._zip is not None:
            self._zip.close()
            self._zip = None

    @property
    def zip(self) -> zipfile.ZipFile:
        if self._zip is None:
            raise RuntimeError("OOXMLPackage must be used inside a with block")
        return self._zip

    def names(self) -> list[str]:
        return self.zip.namelist()

    def iter_prefix(self, prefix: str) -> Iterator[str]:
        for name in self.names():
            if name.startswith(prefix):
                yield name

    def read(self, name: str) -> bytes:
        return self.zip.read(name)

    def has(self, name: str) -> bool:
        try:
            self.zip.getinfo(name)
            return True
        except KeyError:
            return False
