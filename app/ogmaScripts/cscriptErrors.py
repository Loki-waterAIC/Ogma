#!/usr/bin/env python3.11.11
# -*- coding: utf-8 -*-
'''
 # @ Author: Aaron Shackelford
 # @ Create Time: 2025-03-25 11:39:16
 # @ Modified by: Aaron Shackelford
 # @ Modified time: 2025-03-26 13:10:08
 # @ Description: custom errors for CScript calls
 '''

class cscriptError(Exception):
    """
    cscriptError <br>
    Generic Exception for when a cscript failed

    Args:
        Exception (Exception): parent object
    """
