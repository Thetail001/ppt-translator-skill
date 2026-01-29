"""Entry point for the PPT translator CLI."""
from dotenv import load_dotenv

from ppt_translator.cli import main

load_dotenv()

if __name__ == "__main__":
    main()
