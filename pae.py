#!/usr/bin/env python
"""Entry point for PAE Automatizacion CLI"""
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from pae_automatizador import main

if __name__ == '__main__':
    main()