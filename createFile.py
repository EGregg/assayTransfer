#!/usr/bin/env python3


def something():
    try:
        open(fn,'r')
    except:
        open(fn,'w')

something()
