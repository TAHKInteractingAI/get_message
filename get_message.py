"""
Entry point for Microsoft Teams Message Collector to Google Sheets.
All core business logic and services are modularized under `src/`.
"""
import gc
from src import run_collector

gc.disable()

if __name__ == "__main__":
    run_collector()
