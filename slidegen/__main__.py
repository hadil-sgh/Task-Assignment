"""Command line: python -m slidegen "rough notes..." [options], or notes piped on stdin.

Exit codes: 0 deck written and checked, 1 deck written but the check found problems,
2 empty notes or bad arguments, 3 API key missing or rejected.
"""

import argparse
import sys
from pathlib import Path

from slidegen.llm import (
    DEFAULT_MODEL,
    MissingKeyError,
    MockClient,
    OpenRouterClient,
    RejectedKeyError,
    load_api_key,
)
from slidegen.pipeline import EmptyNotesError, Result, fallback_deck, generate, to_json
from slidegen.template import TEAMS, load_catalog, resolve_template

GARBAGE = "Sorry, here are some thoughts about your slides instead of JSON."


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    notes = args.notes if args.notes is not None else read_stdin()
    try:
        if not notes.strip():
            raise EmptyNotesError("The notes are empty. Pass them as an argument or on stdin.")
        client = make_client(args, notes)
        result = generate(notes, client, args.out, args.team, args.template)
    except (EmptyNotesError, ValueError, FileNotFoundError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    except (MissingKeyError, RejectedKeyError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 3
    report(result, args.verbose)
    return 1 if result.problems else 0


def parse_args(argv: list[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="python -m slidegen", description="Turn rough notes into a short PowerPoint deck."
    )
    parser.add_argument("notes", nargs="?", help="the notes (default: read from stdin)")
    parser.add_argument("--out", type=Path, default=Path("deck.pptx"), help="output .pptx")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="OpenRouter model id")
    parser.add_argument("--team", choices=[*TEAMS, "auto"],
                        help="team template; 'auto' picks one from the notes")  # fmt: skip
    parser.add_argument("--template", type=Path, help="use this .pptx template instead")
    parser.add_argument("--verbose", action="store_true", help="show calls, tokens and cost")
    parser.add_argument(
        "--mock", nargs="?", const="ok", choices=["ok", "repair", "fallback"],
        help="no key, no network: a valid answer (ok), a bad first answer (repair), "
        "or only bad answers (fallback)",
    )  # fmt: skip
    return parser.parse_args(argv)


def read_stdin() -> str:
    return "" if sys.stdin.isatty() else sys.stdin.read()


def make_client(args: argparse.Namespace, notes: str):
    if not args.mock:
        return OpenRouterClient(load_api_key(), args.model)
    catalog = load_catalog(resolve_template(notes, args.team, args.template))
    good = to_json(fallback_deck(notes, catalog))  # a valid answer for any template
    answers = {"ok": [good], "repair": [GARBAGE, good], "fallback": [GARBAGE]}
    return MockClient(answers[args.mock])


def report(result: Result, verbose: bool) -> None:
    print(f"Deck written to {result.path} ({result.outcome}, template {result.template.name})")
    if verbose:
        for number, usage in enumerate(result.calls, start=1):
            tokens = f"{usage.prompt_tokens} + {usage.completion_tokens} tokens"
            cost = f"${usage.cost_usd:.4f}, {usage.latency_s} s"
            print(f"  call {number}: {usage.model}, {tokens}, {cost}")
        print(f"  total: {result.tokens} tokens, ${result.cost_usd:.4f}, {result.latency_s:.2f} s")
        for warning in result.warnings:
            print(f"  warning: {warning}")
    for problem in result.problems:
        print(f"problem: {problem}", file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
