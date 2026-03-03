#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Backward-compatible entrypoint.

Core implementation has been consolidated in:
  - pipeline/core.py
"""

from __future__ import annotations

from pipeline.core import main
from pipeline.core import *  # noqa: F401,F403


if __name__ == "__main__":
    main()

