import os
import sys
import argparse
import logging
from docx_translator.config import Config
from docx_translator.pipeline import LocalizationPipeline

def setup_logging(verbose=False):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )

def main():
    parser = argparse.ArgumentParser(
        description="AI-Powered DOCX Translation Tool"
    )
    
    # Core Positional/Required arguments
    parser.add_argument(
        "input",
        help="Path to the input Word Document (.docx)"
    )
    parser.add_argument(
        "target",
        help="Target language code (e.g., 'th', 'es', 'fr', 'ru', 'zh', 'ja')"
    )
    
    # Optional flags
    parser.add_argument(
        "--output", "-o",
        help="Path to save the translated document. Default: <input_name>_<target>.docx"
    )
    parser.add_argument(
        "--source", "-s",
        default="en",
        help="Source language of the document (default: 'en')"
    )
    parser.add_argument(
        "--config", "-c",
        help="Path to custom JSON configuration file"
    )
    parser.add_argument(
        "--backend", "-b",
        choices=["google_translator", "gemini"],
        help="Override the translation backend (google_translator or gemini)"
    )
    parser.add_argument(
        "--tm",
        help="Path to Translation Memory JSON file (default: translation_memory.json)"
    )
    parser.add_argument(
        "--glossary",
        help="Path to Glossary JSON file (default: glossary.json)"
    )
    parser.add_argument(
        "--column-header",
        default="Translation",
        help="The header name of the table column containing text to translate (default: 'Translation')"
    )
    parser.add_argument(
        "--analyze-only",
        action="store_true",
        help="Analyze the document structures, print translation statistics, and exit"
    )
    parser.add_argument(
        "--test-mode",
        action="store_true",
        help="Run translation in test mode (limits execution to 5 tables / 50 cells)"
    )
    parser.add_argument(
        "--review-report",
        help="Path to save the post-translation CSV review report"
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable detailed debug logs"
    )

    args = parser.parse_args()
    setup_logging(args.verbose)
    logger = logging.getLogger("docx_translator.cli")

    if not os.path.exists(args.input):
        logger.error(f"Input file not found: {args.input}")
        sys.exit(1)

    # Initialize configuration
    config = Config(args.config)
    
    # Override translation backend if specified
    if args.backend:
        config.settings["translation_backend"] = args.backend

    # Default output path calculation
    if not args.output and not args.analyze_only:
        base, ext = os.path.splitext(args.input)
        args.output = f"{base}_{args.target}{ext}"

    # Setup pipeline
    pipeline = LocalizationPipeline(
        config=config,
        target_lang=args.target,
        source_lang=args.source,
        tm_path=args.tm,
        glossary_path=args.glossary
    )

    if args.analyze_only:
        # Run report
        report_path = f"analysis_report_{args.target}.json"
        report = pipeline.analyze_document(args.input, args.column_header, report_path)
        print("\n--- Document Analysis Report ---")
        for k, v in report.items():
            print(f"{k.replace('_', ' ').title()}: {v}")
        print("--------------------------------\n")
    else:
        logger.info(f"Translating docx from {args.source} to {args.target}...")
        limit_tables = 5 if args.test_mode else None
        limit_cells = 50 if args.test_mode else None
        
        success = pipeline.translate_document(
            input_path=args.input,
            output_path=args.output,
            column_header=args.column_header,
            limit_tables=limit_tables,
            limit_cells=limit_cells,
            review_report_path=args.review_report
        )
        
        if success:
            logger.info("Translation completed successfully.")
        else:
            logger.error("Translation pipeline failed.")
            sys.exit(1)

if __name__ == "__main__":
    main()
