"""Extract failed files from migration.log into categorized lists.

Files that were later successfully processed are excluded from the failure lists.
Re-run this after processing to update the lists.
"""
import re
from collections import defaultdict


# ============================================================================
# CONSTANTS
# ============================================================================

# Pattern to extract file path from ERROR/FALLBACK line
# e.g., "FALLBACK MedicalFiles/path/to/file.ext : error message"
ERROR_PATTERN = re.compile(r'(?:ERROR|FALLBACK)\s+(MedicalFiles/[^:]+?)\s*:')

# Pattern to extract file path from OK line
# e.g., "OK pdf 1 MedicalFiles/path/to/file.ext -> output"
OK_PATTERN = re.compile(r' OK \w+ \d+ (MedicalFiles/.+?) ->')

# Categories and their matching patterns
CATEGORIES = {
    'network_timeout': [
        'Failed to resolve',
        'Read timed out',
        'getaddrinfo failed',
    ],
    'auth_expired': [
        'az login',
    ],
    'msg_com_error': [
        'Call was rejected by callee',
        'Server execution failed',
        'OpenSharedItem',
        'GetNamespace',
    ],
    'password_protected': [
        'Password protected file',
    ],
    'corrupt_image': [
        'cannot identify image file',
        'image file is truncated',
    ],
    'corrupt_office': [
        'Office has detected a problem',
    ],
    'unsupported_format': [
        'Unsupported file extension',
    ],
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def extract_success_path(line):
    """Extract file path from an OK (success) log line."""
    match = OK_PATTERN.search(line)
    if match:
        return match.group(1).strip()
    return None


def extract_error_path(line):
    """Extract file path from an ERROR/FALLBACK log line."""
    match = ERROR_PATTERN.search(line)
    if match:
        return match.group(1).strip()
    return None


def categorize_error(line):
    """Return category for a log line, or None if not categorizable."""
    if 'BlobAlreadyExists' in line:
        return None  # Not a real error
    
    for category, patterns in CATEGORIES.items():
        for pattern in patterns:
            if pattern in line:
                return category
    return None


def parse_log(log_path='migration.log'):
    """Parse migration log and return (failures_by_category, successes).
    
    Returns:
        failures: dict of category -> set of file paths
        successes: set of file paths that completed successfully
    """
    failures = defaultdict(set)
    successes = set()
    
    with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            # Track successful conversions
            if ' OK ' in line:
                file_path = extract_success_path(line)
                if file_path:
                    successes.add(file_path)
                continue
            
            # Track failures
            if 'ERROR' not in line and 'FALLBACK' not in line:
                continue
            
            category = categorize_error(line)
            if category is None:
                continue
            
            file_path = extract_error_path(line)
            if file_path:
                failures[category].add(file_path)
    
    return failures, successes


def write_failure_lists(failures):
    """Write each failure category to a separate file."""
    for category, files in sorted(failures.items()):
        filename = f'failed_{category}.txt'
        with open(filename, 'w', encoding='utf-8') as f:
            for file_path in sorted(files):
                f.write(file_path + '\n')
        print(f'{filename}: {len(files)} files')


# ============================================================================
# MAIN
# ============================================================================

def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Extract and manage failure lists from migration.log"
    )
    parser.add_argument(
        "--extract",
        action="store_true",
        help="Extract failures from log into categorized files (failed_*.txt)"
    )
    parser.add_argument(
        "--update",
        type=str,
        metavar="FILE",
        help="Update a failure list by removing files that were later processed successfully"
    )
    args = parser.parse_args()
    
    if not args.extract and not args.update:
        parser.print_help()
        return
    
    if args.extract:
        # Parse log and write all failures
        failures, successes = parse_log()
        write_failure_lists(failures)
        total = sum(len(files) for files in failures.values())
        print(f'\nTotal: {total} failed files across {len(failures)} categories')
    
    if args.update:
        # Parse log for successes, then update the specified failure file
        _, successes = parse_log()
        
        filename = args.update
        with open(filename, 'r', encoding='utf-8') as f:
            current_failures = {line.strip() for line in f if line.strip()}
        
        remaining = current_failures - successes
        removed = len(current_failures) - len(remaining)
        
        with open(filename, 'w', encoding='utf-8') as f:
            for path in sorted(remaining):
                f.write(path + '\n')
        
        print(f'{filename}: {len(remaining)} remaining ({removed} removed)')


if __name__ == '__main__':
    main()
