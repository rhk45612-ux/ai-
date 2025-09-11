import argparse
import difflib

RED = '\033[91m'
GREEN = '\033[92m'
RESET = '\033[0m'


def color_diff(original: str, modified: str) -> str:
    """Return a word-level diff of two strings with ANSI color codes.

    Removed words are shown in red, added words in green.
    """
    orig_words = original.split()
    mod_words = modified.split()
    diff = difflib.ndiff(orig_words, mod_words)
    colored_words = []
    for token in diff:
        code, word = token[0], token[2:]
        if code == '-':
            colored_words.append(f"{RED}{word}{RESET}")
        elif code == '+':
            colored_words.append(f"{GREEN}{word}{RESET}")
        elif code == ' ':  # unchanged
            colored_words.append(word)
    return ' '.join(colored_words)


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare two text files and highlight word differences.")
    parser.add_argument('original', help='Path to the original text file')
    parser.add_argument('modified', help='Path to the modified text file')
    args = parser.parse_args()

    with open(args.original, encoding='utf-8') as f:
        original = f.read()
    with open(args.modified, encoding='utf-8') as f:
        modified = f.read()

    print(color_diff(original, modified))


if __name__ == '__main__':
    main()
