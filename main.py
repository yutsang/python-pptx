import argparse
from common import assistant
import sys


def main():
    parser = argparse.ArgumentParser(description='Real Estate Due Diligence Report Generator')
    parser.add_argument('-i', '--input', required=True, help='Path to the Excel input file.')
    parser.add_argument('-e', '--entity', required=True, help='Entity name (e.g., Haining, Nanjing, Ningbo).')
    parser.add_argument('--helpers', nargs='+', default=[], help='List of entity helpers.')
    parser.add_argument('--ai', action='store_true', help='Use AI for text generation (default: local/test mode).')
    parser.add_argument('--output', default='output.pptx', help='Output PPTX file path.')
    parser.add_argument('--config', default='utils/config.json', help='Path to config file.')
    parser.add_argument('--mapping', default='utils/mapping.json', help='Path to mapping file.')
    parser.add_argument('--pattern', default='utils/pattern.json', help='Path to pattern file.')
    args = parser.parse_args()

    # Define keys (can be made dynamic if needed)
    keys = [
        "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
        "AP", "Taxes payable", "OP", "Capital", "Reserve"
    ]
    if args.entity in ['Ningbo', 'Nanjing']:
        keys = [key for key in keys if key != "Reserve"]

    # Process keys and generate text
    results = assistant.process_keys(
        keys=keys,
        entity_name=args.entity,
        entity_helpers=args.helpers,
        input_file=args.input,
        mapping_file=args.mapping,
        pattern_file=args.pattern,
        config_file=args.config,
        use_ai=args.ai
    )

    # QA and correction
    qa_agent = assistant.QualityAssuranceAgent()
    for key in results:
        qa_result = qa_agent.validate_content(results[key])
        if qa_result['score'] < 90:
            results[key] = qa_agent.auto_correct(results[key])

    # Print results and save to PPTX (placeholder, to be implemented)
    print("\nGenerated Report Sections:")
    for key, text in results.items():
        print(f"\n[{key}]\n{text}")

    # TODO: Call PPTX export function here, e.g.:
    # assistant.export_to_pptx(results, args.output)
    print(f"\nReport generation complete. (PPTX export to be implemented: {args.output})")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        print("Usage: python main.py -i <excel_file> -e <entity> [--helpers ...] [--ai] [--output <pptx>] [--config <config>] [--mapping <mapping>] [--pattern <pattern>]")
        sys.exit(1)
    main() 