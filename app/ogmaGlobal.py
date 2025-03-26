#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os

APP_VERSION:str = "2.0.0"

# Locked access to word
LOCK_FILE_PATH: str = os.path.join(os.path.abspath("."), os.path.join("tmp", "ogma_lock.lock"))

VERBOSE_LEVEL:int = 0